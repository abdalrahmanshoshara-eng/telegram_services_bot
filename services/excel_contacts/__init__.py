"""Excel-to-VCF conversion service adapted from the Professional Reports platform."""

from pathlib import Path, PurePath

from .domain import normalize_country_code
from .processor import WorkbookValidationError, process_workbook

MAX_FILE_SIZE = 10 * 1024 * 1024
ALLOWED_EXTENSIONS = {".xlsx", ".xls"}
XLSX_SIGNATURES = (b"PK\x03\x04", b"PK\x05\x06", b"PK\x07\x08")
XLS_SIGNATURE = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"


class ExcelContactsError(ValueError):
    """Expected validation error that is safe to show to a Telegram user."""


def convert_excel_contacts(
    source: Path,
    destination: Path,
    country_code: str,
    source_filename: str | None = None,
    vcf_destination: Path | None = None,
) -> dict:
    filename = source_filename or source.name
    extension = PurePath(filename).suffix.lower()
    if extension not in ALLOWED_EXTENSIONS:
        raise ExcelContactsError("الصيغ المدعومة هي XLSX وXLS فقط.")

    size = source.stat().st_size
    if size == 0:
        raise ExcelContactsError("ملف Excel المرفوع فارغ.")
    if size > MAX_FILE_SIZE:
        raise ExcelContactsError("حجم ملف Excel أكبر من الحد المسموح وهو 10 ميغابايت.")

    data = source.read_bytes()
    valid_signature = (
        data.startswith(XLSX_SIGNATURES)
        if extension == ".xlsx"
        else data.startswith(XLS_SIGNATURE)
    )
    if not valid_signature:
        raise ExcelContactsError("محتوى الملف لا يطابق صيغة Excel المحددة.")

    try:
        normalized_country_code = normalize_country_code(country_code)
        result = process_workbook(data, extension, filename, normalized_country_code)
    except (ValueError, WorkbookValidationError) as error:
        raise ExcelContactsError(str(error)) from error

    destination.write_bytes(result["zip_buffer"])
    if vcf_destination is not None:
        vcf_destination.write_bytes(result["vcard_buffer"])
    return result["summary"]
