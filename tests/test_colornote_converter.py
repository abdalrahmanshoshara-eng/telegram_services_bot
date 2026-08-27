import hashlib
import json

import pytest
from cryptography.hazmat.primitives.ciphers import Cipher, algorithms, modes
from cryptography.hazmat.primitives.padding import PKCS7

from services.colornote import ColorNoteConversionError, convert_colornote_backups
from services.colornote.merger import merge_backups
from services.colornote.parser import FIXED_SALT, parse_colornote_backup


def _derive(password="0000"):
    output = b""
    previous = b""
    while len(output) < 32:
        previous = hashlib.md5(
            previous + password.encode("utf-8") + FIXED_SALT
        ).digest()
        output += previous
    return output[:16], output[16:32]


def _build_backup(records, password="0000"):
    plain = b"0123456789ABCDEF"
    for record in records:
        raw = json.dumps(
            record,
            ensure_ascii=False,
            separators=(",", ":"),
        ).encode("utf-8")
        plain += len(raw).to_bytes(4, "big") + raw
    padder = PKCS7(128).padder()
    padded = padder.update(plain) + padder.finalize()
    key, iv = _derive(password)
    encryptor = Cipher(algorithms.AES(key), modes.CBC(iv)).encryptor()
    encrypted = encryptor.update(padded) + encryptor.finalize()
    header = b"\x00N\x00O\x00T\x00E" + b"\x00" * (28 - 8)
    return header + encrypted


def _record(note_id, uuid, title, text, modified):
    return {
        "_id": note_id,
        "uuid": uuid,
        "title": title,
        "note": text,
        "created_date": 1_700_000_000_000,
        "modified_date": modified,
        "active_state": 0,
        "color_index": 3,
        "tags": "",
    }


def test_supported_backup_is_decrypted_and_parsed():
    data = _build_backup(
        [_record(1, "u1", "عنوان", "نص تجريبي", 1_700_000_010_000)]
    )

    result = parse_colornote_backup(data)

    assert result.format_offset == 28
    assert result.notes[0]["uuid"] == "u1"
    assert result.notes[0]["text"] == "نص تجريبي"


def test_backups_are_merged_and_exported_to_self_contained_html(tmp_path):
    old_backup = tmp_path / "old.backup"
    new_backup = tmp_path / "new.backup"
    output = tmp_path / "ColorNote_notes.html"
    old_backup.write_bytes(
        _build_backup(
            [_record(1, "same", "عنوان", "النص القديم", 1_700_000_010_000)]
        )
    )
    new_backup.write_bytes(
        _build_backup(
            [_record(1, "same", "عنوان", "النص الجديد", 1_800_000_010_000)]
        )
    )

    summary = convert_colornote_backups(
        [(old_backup.name, old_backup), (new_backup.name, new_backup)],
        output,
        "0000",
    )

    html = output.read_text(encoding="utf-8")
    assert summary == {
        "files": 2,
        "total_input": 2,
        "total": 1,
        "duplicates_removed": 1,
        "changed_notes": 1,
        "repaired": 0,
        "binary": 0,
        "per_file_repaired": [0, 0],
        "per_file_binary": [0, 0],
    }
    assert "النص الجديد" in html
    assert "النص القديم" not in html
    assert 'name="viewport"' in html
    assert "100dvh" in html
    assert 'id="filterToggle"' in html
    assert "<script src=" not in html
    assert "<link rel=" not in html


def test_merge_tie_keeps_first_uploaded_copy():
    timestamp = "2026-01-02T00:00:00Z"
    first = {
        "id": 1,
        "uuid": "same",
        "title": "Note",
        "text": "first",
        "created": timestamp,
        "modified": timestamp,
        "active_state": 0,
        "color_index": 0,
        "tags": "",
        "text_status": "original",
        "encrypted": False,
    }
    second = {**first, "text": "second"}

    result = merge_backups([("first", [first]), ("second", [second])])

    assert result.notes[0]["text"] == "first"
    assert result.notes[0]["changed_across_backups"] is True


def test_wrong_password_returns_safe_error(tmp_path):
    source = tmp_path / "notes.backup"
    source.write_bytes(
        _build_backup(
            [_record(1, "u1", "Title", "Text", 1_700_000_010_000)],
            password="1234",
        )
    )

    with pytest.raises(ColorNoteConversionError, match="تحقق من كلمة المرور"):
        convert_colornote_backups(
            [(source.name, source)],
            tmp_path / "output.html",
            "0000",
        )


def test_script_closer_is_escaped_inside_export(tmp_path):
    source = tmp_path / "notes.backup"
    source.write_bytes(
        _build_backup(
            [
                _record(
                    1,
                    "u1",
                    "Title",
                    "</script><script>alert(1)</script>",
                    1_700_000_010_000,
                )
            ]
        )
    )
    output = tmp_path / "output.html"

    convert_colornote_backups([(source.name, source)], output, "0000")

    html = output.read_text(encoding="utf-8")
    assert "</script><script>alert(1)</script>" not in html
    assert "<\\/script><script>alert(1)<\\/script>" in html
