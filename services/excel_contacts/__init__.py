"""Contacts-to-VCF conversion service adapted from the Professional Reports platform.

Accepts one file or several, in any mix of Excel, CSV, TXT, VCF, DOCX and PDF.
Several files are merged into one clean set of contacts.
"""

import os
from collections.abc import Sequence
from pathlib import Path, PurePath

from .domain import normalize_country_code
from .processor import WorkbookValidationError, build_result, read_source

MAX_FILE_SIZE_MB = max(1, int(os.environ.get("MAX_EXCEL_CONTACTS_FILE_MB", "50")))
MAX_FILE_SIZE = MAX_FILE_SIZE_MB * 1024 * 1024
MAX_FILES = max(1, int(os.environ.get("MAX_MULTI_FILES", "10")))
EXCEL_EXTENSIONS = {".xlsx", ".xls"}
TEXT_EXTENSIONS = {".csv", ".txt", ".vcf"}
ALLOWED_EXTENSIONS = EXCEL_EXTENSIONS | TEXT_EXTENSIONS | {".docx", ".pdf"}
ZIP_SIGNATURES = (b"PK\x03\x04", b"PK\x05\x06", b"PK\x07\x08")
XLSX_SIGNATURES = ZIP_SIGNATURES
XLS_SIGNATURE = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"
_KIND = {".xlsx": "Excel", ".xls": "Excel", ".docx": "Word", ".pdf": "PDF"}


class ExcelContactsError(ValueError):
    """Expected validation error that is safe to show to a Telegram user."""


def _read_checked(source: Path, filename: str, name_errors: bool) -> tuple[bytes, str]:
    """The file's bytes and extension, after type, size and content checks.

    With several files, `name_errors` puts the file name in front of each
    message so the user knows which file to fix.
    """
    extension = PurePath(filename).suffix.lower()
    prefix = f"{filename}: " if name_errors else ""
    if extension not in ALLOWED_EXTENSIONS:
        raise ExcelContactsError(
            f"{prefix}الصيغ المدعومة هي XLSX وXLS وCSV وTXT وVCF وDOCX وPDF."
        )
    size = source.stat().st_size
    if size == 0:
        raise ExcelContactsError(f"{prefix}الملف المرفوع فارغ.")
    if size > MAX_FILE_SIZE:
        raise ExcelContactsError(
            f"{prefix}حجم الملف أكبر من الحد المسموح وهو {MAX_FILE_SIZE_MB} ميغابايت."
        )

    data = source.read_bytes()
    if extension in {".xlsx", ".docx"}:
        valid_signature = data.startswith(ZIP_SIGNATURES)
    elif extension == ".xls":
        valid_signature = data.startswith(XLS_SIGNATURE)
    elif extension == ".pdf":
        valid_signature = b"%PDF-" in data[:1024]
    else:
        # Plain-text formats: reject binaries renamed to .csv/.txt/.vcf.
        valid_signature = b"\x00" not in data[:4096] or data.startswith(
            (b"\xff\xfe", b"\xfe\xff")
        )
    if not valid_signature:
        raise ExcelContactsError(
            f"{prefix}محتوى الملف لا يطابق صيغة "
            f"{_KIND.get(extension, 'ملف نصي')} المحددة."
        )
    return data, extension


def convert_contacts_files(
    sources: Sequence[tuple[str, Path]],
    destination: Path,
    country_code: str,
    vcf_destination: Path | None = None,
) -> dict:
    """Merge (display name, path) files into one VCF and one result bundle."""
    if not sources:
        raise ExcelContactsError("أرسل ملفًا واحدًا على الأقل.")
    if len(sources) > MAX_FILES:
        raise ExcelContactsError(f"الحد الأقصى {MAX_FILES} ملفات في العملية الواحدة.")
    try:
        normalized_country_code = normalize_country_code(country_code)
    except ValueError as error:
        raise ExcelContactsError(str(error)) from error

    several = len(sources) > 1
    parsed = []
    for filename, path in sources:
        filename = filename or path.name
        data, extension = _read_checked(path, filename, several)
        try:
            rows, label = read_source(data, extension)
        except (ValueError, WorkbookValidationError) as error:
            message = f"{filename}: {error}" if several else str(error)
            raise ExcelContactsError(message) from error
        parsed.append((filename, rows, label))

    try:
        result = build_result(parsed, normalized_country_code)
    except (ValueError, WorkbookValidationError) as error:
        raise ExcelContactsError(str(error)) from error

    destination.write_bytes(result["zip_buffer"])
    if vcf_destination is not None:
        vcf_destination.write_bytes(result["vcard_buffer"])
    return {**result["summary"], "files": len(sources)}


def convert_excel_contacts(
    source: Path,
    destination: Path,
    country_code: str,
    source_filename: str | None = None,
    vcf_destination: Path | None = None,
) -> dict:
    """One file; kept for callers that send a single file."""
    summary = convert_contacts_files(
        [(source_filename or source.name, source)],
        destination,
        country_code,
        vcf_destination,
    )
    summary.pop("files")
    return summary
