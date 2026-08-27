"""Stateless ColorNote backup-to-HTML conversion service."""

import os
from collections.abc import Sequence
from pathlib import Path

from .exporter import generate_html
from .merger import merge_backups
from .parser import (
    DEFAULT_PASSWORD,
    ColorNoteError,
    DecryptionError,
    UnsupportedBackupError,
    parse_colornote_backup,
)

MAX_FILES = max(1, int(os.environ.get("MAX_COLORNOTE_FILES", "20")))
MAX_FILE_SIZE = max(1, int(os.environ.get("MAX_UPLOAD_MB", "20"))) * 1024 * 1024
MAX_TOTAL_SIZE = (
    max(1, int(os.environ.get("MAX_TOTAL_UPLOAD_MB", "100"))) * 1024 * 1024
)


class ColorNoteConversionError(ValueError):
    """Expected error that is safe to display to a Telegram user."""


def _display_names(sources: Sequence[tuple[str, Path]]) -> list[str]:
    seen: dict[str, int] = {}
    labels: list[str] = []
    for filename, _ in sources:
        base = filename or "ColorNote"
        seen[base] = seen.get(base, 0) + 1
        labels.append(base if seen[base] == 1 else f"{base} ({seen[base]})")
    return labels


def convert_colornote_backups(
    sources: Sequence[tuple[str, Path]],
    destination: Path,
    password: str = DEFAULT_PASSWORD,
) -> dict:
    if not sources:
        raise ColorNoteConversionError("أرسل ملف ColorNote واحدًا على الأقل.")
    if len(sources) > MAX_FILES:
        raise ColorNoteConversionError(
            f"الحد الأقصى هو {MAX_FILES} ملفات في العملية الواحدة."
        )

    labels = _display_names(sources)
    parsed_backups = []
    total_bytes = 0
    repaired_per_file = []
    binary_per_file = []

    for (filename, path), source_name in zip(sources, labels, strict=True):
        size = path.stat().st_size
        total_bytes += size
        if size == 0:
            raise ColorNoteConversionError(f"الملف «{source_name}» فارغ.")
        if size > MAX_FILE_SIZE:
            raise ColorNoteConversionError(
                f"الملف «{source_name}» يتجاوز الحد الأقصى "
                f"{MAX_FILE_SIZE // (1024 * 1024)} ميغابايت."
            )
        if total_bytes > MAX_TOTAL_SIZE:
            raise ColorNoteConversionError(
                "إجمالي الملفات يتجاوز الحد الأقصى "
                f"{MAX_TOTAL_SIZE // (1024 * 1024)} ميغابايت."
            )

        data = path.read_bytes()
        try:
            result = parse_colornote_backup(data, password=password or DEFAULT_PASSWORD)
        except DecryptionError as error:
            raise ColorNoteConversionError(
                f"تعذر فك «{source_name}». تحقق من كلمة المرور أو إصدار الملف."
            ) from error
        except UnsupportedBackupError as error:
            raise ColorNoteConversionError(
                f"الملف «{source_name}» ليس نسخة ColorNote مدعومة."
            ) from error
        except ColorNoteError as error:
            raise ColorNoteConversionError(
                f"تعذر معالجة «{source_name}»."
            ) from error

        parsed_backups.append((source_name, result.notes))
        repaired_per_file.append(
            sum(
                1
                for note in result.notes
                if note.get("text_status") == "utf16le_repaired"
            )
        )
        binary_per_file.append(
            sum(
                1
                for note in result.notes
                if note.get("text_status") == "binary"
            )
        )

    merged = merge_backups(parsed_backups)
    source_name = (
        labels[0]
        if len(labels) == 1
        else f"ColorNote_Merged_{len(labels)}_Backups"
    )
    destination.write_bytes(generate_html(merged.notes, source_name))
    return {
        "files": len(labels),
        "total_input": merged.total_input,
        "total": merged.unique_notes,
        "duplicates_removed": merged.duplicates_removed,
        "changed_notes": merged.changed_notes,
        "repaired": sum(
            1
            for note in merged.notes
            if note.get("text_status") == "utf16le_repaired"
        ),
        "binary": sum(
            1 for note in merged.notes if note.get("text_status") == "binary"
        ),
        "per_file_repaired": repaired_per_file,
        "per_file_binary": binary_per_file,
    }
