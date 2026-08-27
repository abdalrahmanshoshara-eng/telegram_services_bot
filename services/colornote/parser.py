from __future__ import annotations

import hashlib
import json
from dataclasses import dataclass
from datetime import datetime, timezone
from typing import Any

from cryptography.hazmat.primitives.ciphers import Cipher, algorithms, modes
from cryptography.hazmat.primitives.padding import PKCS7

COLORNOTE_MAGIC = b"\x00N\x00O\x00T\x00E"
FIXED_SALT = b"ColorNote Fixed Salt"
DEFAULT_PASSWORD = "0000"


class ColorNoteError(Exception):
    code = "COLORNOTE_ERROR"


class UnsupportedBackupError(ColorNoteError):
    code = "UNSUPPORTED_BACKUP"


class DecryptionError(ColorNoteError):
    code = "DECRYPTION_FAILED"


@dataclass(frozen=True)
class ParseResult:
    notes: list[dict[str, Any]]
    format_offset: int
    password_was_default: bool


def _openssl_bytes_to_key(password: str) -> tuple[bytes, bytes]:
    """BouncyCastle OpenSSL-compatible MD5 PBE used by ColorNote backups."""
    password_bytes = password.encode("utf-8")
    derived = bytearray()
    previous = b""
    while len(derived) < 32:
        previous = hashlib.md5(previous + password_bytes + FIXED_SALT).digest()
        derived.extend(previous)
    return bytes(derived[:16]), bytes(derived[16:32])


def _decrypt_payload(data: bytes, password: str, offset: int) -> bytes:
    if offset < 0 or offset >= len(data):
        raise DecryptionError("Invalid backup offset")
    encrypted = data[offset:]
    if not encrypted or len(encrypted) % 16:
        raise DecryptionError("Encrypted payload is not aligned to AES block size")

    key, iv = _openssl_bytes_to_key(password)
    decryptor = Cipher(algorithms.AES(key), modes.CBC(iv)).decryptor()
    padded = decryptor.update(encrypted) + decryptor.finalize()

    try:
        unpadder = PKCS7(128).unpadder()
        return unpadder.update(padded) + unpadder.finalize()
    except ValueError as error:
        raise DecryptionError(
            "Wrong password or unsupported backup version"
        ) from error


def _parse_records_at(plain: bytes, start: int) -> tuple[list[dict[str, Any]], int]:
    records: list[dict[str, Any]] = []
    pos = start
    while pos + 4 <= len(plain):
        length = int.from_bytes(plain[pos : pos + 4], "big")
        if length <= 2 or length > len(plain) - pos - 4:
            break
        pos += 4
        raw = plain[pos : pos + length]
        if not raw.startswith(b"{"):
            break
        try:
            record = json.loads(raw.decode("utf-8", errors="replace"))
        except (UnicodeError, json.JSONDecodeError):
            break
        if not isinstance(record, dict) or "_id" not in record:
            break
        records.append(record)
        pos += length
    return records, pos


def _extract_json_records(plain: bytes) -> list[dict[str, Any]]:
    first_json = plain.find(b'{"_id"')
    candidates = {0, 4, 8, 12, 16, 20}
    if first_json >= 4:
        candidates.add(first_json - 4)

    best: tuple[list[dict[str, Any]], int] = ([], 0)
    for start in sorted(c for c in candidates if 0 <= c < len(plain)):
        records, consumed = _parse_records_at(plain, start)
        if len(records) > len(best[0]) or (
            len(records) == len(best[0]) and consumed > best[1]
        ):
            best = (records, consumed)

    if not best[0]:
        raise DecryptionError(
            "Decryption succeeded but no ColorNote records were found"
        )
    return best[0]


def _legacy_utf16_candidate(value: str) -> str | None:
    if "\x00" not in value:
        return None
    stripped = value.lstrip("\ufffd")
    unicode_tail = "".join(ch for ch in stripped if ord(ch) > 0xFF)
    raw = bytes(ord(ch) for ch in stripped if ord(ch) <= 0xFF)
    if len(raw) < 2:
        return None
    if len(raw) % 2:
        raw = raw[:-1]
    try:
        decoded = raw.decode("utf-16le")
    except UnicodeError:
        return None
    if not decoded:
        return None
    printable = sum(ch.isprintable() or ch in "\r\n\t" for ch in decoded) / len(
        decoded
    )
    return (decoded + unicode_tail) if printable >= 0.85 else None


def _normalize_text(value: Any) -> tuple[str, str, str]:
    text = "" if value is None else str(value)
    utf16 = _legacy_utf16_candidate(text)
    if utf16 is not None:
        signature = utf16[:12].upper()
        if signature.startswith("RIFF") or "CDDA" in signature:
            return (
                "[بيانات ثنائية من نوع RIFF/CDDA محفوظة داخل حقل الملاحظة وليست نصًا عاديًا]",
                "binary",
                text.encode("unicode_escape", errors="backslashreplace").decode(
                    "ascii", errors="ignore"
                ),
            )
        return (
            utf16.replace("\x00", ""),
            "utf16le_repaired",
            text.encode("unicode_escape", errors="backslashreplace").decode(
                "ascii", errors="ignore"
            ),
        )

    cleaned = text.replace("\x00", "")
    upper = cleaned.lstrip("\ufffd").upper()
    if upper.startswith("RIFF") or "CDDA" in upper[:24]:
        return (
            "[بيانات ثنائية من نوع RIFF/CDDA محفوظة داخل حقل الملاحظة وليست نصًا عاديًا]",
            "binary",
            text.encode("unicode_escape", errors="backslashreplace").decode(
                "ascii", errors="ignore"
            ),
        )
    return cleaned, "original", ""


def _epoch_ms_to_iso(value: Any) -> str:
    try:
        milliseconds = int(value or 0)
    except (TypeError, ValueError):
        return ""
    if milliseconds <= 0:
        return ""
    return (
        datetime.fromtimestamp(milliseconds / 1000, tz=timezone.utc)
        .isoformat()
        .replace("+00:00", "Z")
    )


def _normalize_record(record: dict[str, Any]) -> dict[str, Any]:
    note_text, text_status, _ = _normalize_text(record.get("note", ""))
    title, title_status, _ = _normalize_text(record.get("title", ""))
    if not title.strip():
        title = next(
            (line.strip() for line in note_text.splitlines() if line.strip()),
            "ملاحظة بدون عنوان",
        )[:120]

    status = text_status if text_status != "original" else title_status
    return {
        "id": record.get("_id"),
        "uuid": str(record.get("uuid") or record.get("_id") or ""),
        "title": title,
        "text": note_text,
        "created": _epoch_ms_to_iso(record.get("created_date")),
        "modified": _epoch_ms_to_iso(record.get("modified_date")),
        "active_state": record.get("active_state", 0),
        "color_index": record.get("color_index", 0),
        "tags": str(record.get("tags") or ""),
        "text_status": status,
        "encrypted": bool(record.get("encrypted", 0)),
    }


def parse_colornote_backup(
    data: bytes, password: str = DEFAULT_PASSWORD
) -> ParseResult:
    if not data or len(data) < 32:
        raise UnsupportedBackupError(
            "The uploaded file is too small to be a ColorNote backup"
        )

    has_magic = data.startswith(COLORNOTE_MAGIC)
    offsets = [28, 0] if has_magic else [0, 28]
    last_error: Exception | None = None

    for offset in offsets:
        try:
            plain = _decrypt_payload(data, password, offset)
            records = _extract_json_records(plain)
            notes = [_normalize_record(record) for record in records]
            notes.sort(
                key=lambda note: note.get("modified") or note.get("created") or "",
                reverse=True,
            )
            return ParseResult(
                notes=notes,
                format_offset=offset,
                password_was_default=password == DEFAULT_PASSWORD,
            )
        except ColorNoteError as error:
            last_error = error

    if has_magic:
        raise DecryptionError(
            "تعذر فك النسخة الاحتياطية. تحقق من كلمة المرور أو إصدار الملف."
        ) from last_error
    raise UnsupportedBackupError(
        "الملف لا يبدو كنسخة ColorNote مدعومة."
    ) from last_error
