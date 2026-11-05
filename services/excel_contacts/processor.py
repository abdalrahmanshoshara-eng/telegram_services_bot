"""Bounded in-memory Excel parsing and result-archive generation."""

import csv
import io
import os
import re
import zipfile
from datetime import UTC, datetime

import xlrd
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

from .domain import (
    DUPLICATE_COLUMNS,
    EMAIL,
    EXPECTED_COLUMNS,
    FULL_NAME,
    INVALID_COLUMNS,
    REQUIRED_COLUMNS,
    VALID_COLUMNS,
    create_vcard,
    is_empty,
    normalize_header,
    process_contact_rows,
    safe_spreadsheet_value,
)
from .readers import ContactsReadError, read_contacts

MAX_WORKSHEETS = max(1, int(os.environ.get("MAX_EXCEL_CONTACTS_WORKSHEETS", "20")))
MAX_ROWS = max(1, int(os.environ.get("MAX_EXCEL_CONTACTS_ROWS", "10000")))
# Real-world worksheets -- especially files exported from Google Sheets/Forms --
# very often report a huge `max_column`/`ncols` because formatting (borders,
# fill, column width, ...) was applied to a wide tail of otherwise empty
# columns. That declared dimension is not a trustworthy measure of how many
# columns actually hold data, so this limit is enforced against genuinely
# populated columns instead of the raw reported width (see
# `_bounded_xlsx_rows` / `_bounded_xls_rows` below), which lets it stay
# generous without reopening the resource-exhaustion risk it exists for.
MAX_COLUMNS = max(100, int(os.environ.get("MAX_EXCEL_CONTACTS_COLUMNS", "5000")))
MAX_XLSX_ARCHIVE_ENTRIES = 2_000
MAX_XLSX_UNCOMPRESSED_SIZE = 100 * 1024 * 1024
MAX_OUTPUT_SIZE = 25 * 1024 * 1024


class WorkbookValidationError(Exception):
    """Expected, safe-to-report workbook rejection."""


def _validate_xlsx_container(data: bytes) -> None:
    try:
        with zipfile.ZipFile(io.BytesIO(data)) as archive:
            entries = archive.infolist()
            if len(entries) > MAX_XLSX_ARCHIVE_ENTRIES:
                raise WorkbookValidationError(
                    "يحتوي ملف Excel على عدد مفرط من الأجزاء."
                )
            if sum(item.file_size for item in entries) > MAX_XLSX_UNCOMPRESSED_SIZE:
                raise WorkbookValidationError(
                    "حجم محتوى ملف Excel بعد فك الضغط يتجاوز الحد المسموح."
                )
            if any(item.flag_bits & 0x1 for item in entries):
                raise WorkbookValidationError("ملفات Excel المشفرة غير مدعومة.")
    except zipfile.BadZipFile as error:
        raise WorkbookValidationError("ملف Excel تالف أو غير صالح.") from error


def _non_empty_matrix(rows) -> list[list]:
    return [list(row) for row in rows if any(not is_empty(value) for value in row)]


def _first_non_empty_row(rows) -> list:
    for row in rows:
        cells = list(row)
        if any(not is_empty(value) for value in cells):
            return cells
    return []


def _header_match_rank(row: list) -> int:
    """How well a header row matches the contact schema.

    2 -- every expected column is present, 1 -- only the required ones
    (name + phone) are, 0 -- this is not a contacts sheet. A sheet that carries
    the optional columns too is preferred when a workbook holds several
    candidate sheets.
    """
    if not row:
        return 0
    headers = {normalize_header(value) for value in row}
    if not all(column in headers for column in REQUIRED_COLUMNS):
        return 0
    return 2 if all(column in headers for column in EXPECTED_COLUMNS) else 1


def _too_many_columns_error() -> WorkbookValidationError:
    return WorkbookValidationError(
        f"الحد الأعلى المدعوم هو {MAX_COLUMNS} عمود لكل ورقة."
    )


def _bounded_xlsx_rows(sheet, **kwargs):
    """Yield each row's values, measuring width from real data rather than
    the worksheet's declared (and frequently inflated) dimension.

    At most MAX_COLUMNS + 1 columns are ever requested from openpyxl per
    row, regardless of what `sheet.max_column` reports, so a worksheet whose
    reported extent is inflated by stray formatting is neither wrongly
    rejected nor expensive to read. If the extra probe column is ever
    genuinely populated, the sheet really does exceed the configured limit.
    """
    scan_width = MAX_COLUMNS + 1
    for row in sheet.iter_rows(
        min_col=1, max_col=scan_width, values_only=True, **kwargs
    ):
        if len(row) > MAX_COLUMNS and not is_empty(row[MAX_COLUMNS]):
            raise _too_many_columns_error()
        yield row[:MAX_COLUMNS]


def _bounded_xls_rows(sheet):
    """Same real-data width measurement as `_bounded_xlsx_rows`, for legacy .xls."""
    scan_width = MAX_COLUMNS + 1
    for row_index in range(sheet.nrows):
        row = sheet.row_values(row_index, 0, min(sheet.ncols, scan_width))
        if len(row) > MAX_COLUMNS and not is_empty(row[MAX_COLUMNS]):
            raise _too_many_columns_error()
        yield row[:MAX_COLUMNS]


def _read_xlsx(data: bytes) -> tuple[list[list], str]:
    _validate_xlsx_container(data)
    try:
        workbook = load_workbook(io.BytesIO(data), read_only=True, data_only=True)
        if len(workbook.sheetnames) > MAX_WORKSHEETS:
            raise WorkbookValidationError(
                f"الحد الأعلى هو {MAX_WORKSHEETS} ورقة عمل."
            )
        if not workbook.sheetnames:
            raise WorkbookValidationError("ملف Excel لا يحتوي على أوراق عمل.")
        fallback = None
        partial_match = None
        for sheet_name in workbook.sheetnames:
            sheet = workbook[sheet_name]
            if sheet.max_row > MAX_ROWS + 1:
                raise WorkbookValidationError(
                    f"الحد الأعلى المدعوم هو {MAX_ROWS:,} صف لكل ملف."
                )
            header = _first_non_empty_row(_bounded_xlsx_rows(sheet))
            if header and fallback is None:
                fallback = ([header], sheet.title)
            rank = _header_match_rank(header)
            if rank == 2:
                matrix = _non_empty_matrix(_bounded_xlsx_rows(sheet))
                return matrix, sheet.title
            if rank == 1 and partial_match is None:
                partial_match = sheet.title
        if partial_match is not None:
            sheet = workbook[partial_match]
            return _non_empty_matrix(_bounded_xlsx_rows(sheet)), sheet.title
        return fallback or ([], workbook.sheetnames[0])
    except WorkbookValidationError:
        raise
    except Exception as error:
        raise WorkbookValidationError("ملف Excel تالف أو غير صالح.") from error
    finally:
        if "workbook" in locals():
            workbook.close()


def _read_xls(data: bytes) -> tuple[list[list], str]:
    try:
        workbook = xlrd.open_workbook(file_contents=data, on_demand=True)
        if workbook.nsheets > MAX_WORKSHEETS:
            raise WorkbookValidationError(
                f"الحد الأعلى هو {MAX_WORKSHEETS} ورقة عمل."
            )
        if workbook.nsheets == 0:
            raise WorkbookValidationError("ملف Excel لا يحتوي على أوراق عمل.")
        fallback = None
        partial_match = None
        for sheet_index in range(workbook.nsheets):
            sheet = workbook.sheet_by_index(sheet_index)
            if sheet.nrows > MAX_ROWS + 1:
                raise WorkbookValidationError(
                    f"الحد الأعلى المدعوم هو {MAX_ROWS:,} صف لكل ملف."
                )
            header = _first_non_empty_row(_bounded_xls_rows(sheet))
            if header and fallback is None:
                fallback = ([header], sheet.name)
            rank = _header_match_rank(header)
            if rank == 2:
                matrix = _non_empty_matrix(_bounded_xls_rows(sheet))
                return matrix, sheet.name
            if rank == 1 and partial_match is None:
                partial_match = sheet_index
        if partial_match is not None:
            sheet = workbook.sheet_by_index(partial_match)
            return _non_empty_matrix(_bounded_xls_rows(sheet)), sheet.name
        return fallback or ([], workbook.sheet_names()[0])
    except WorkbookValidationError:
        raise
    except Exception as error:
        raise WorkbookValidationError("ملف Excel تالف أو غير صالح.") from error
    finally:
        if "workbook" in locals():
            workbook.release_resources()


def _extract_rows(matrix: list[list]) -> list[dict]:
    if not matrix:
        raise WorkbookValidationError("ورقة العمل الأولى فارغة.")
    headers = [normalize_header(value) for value in matrix[0]]
    missing = [column for column in REQUIRED_COLUMNS if column not in headers]
    if missing:
        raise WorkbookValidationError(
            f"الأعمدة المطلوبة غير موجودة: {'، '.join(missing)}."
        )
    index_by_header = {header: index for index, header in enumerate(headers)}
    # Optional columns the workbook does not carry simply read back as empty,
    # so a file with just a name and a phone number converts normally.
    present = [column for column in EXPECTED_COLUMNS if column in index_by_header]
    return [
        {
            column: (
                cells[index_by_header[column]]
                if index_by_header[column] < len(cells)
                else ""
            )
            for column in present
        }
        for cells in matrix[1:]
    ]


def _create_excel_report(
    sheet_name: str,
    columns: list[str],
    rows: list[dict],
    header_color: str,
) -> bytes:
    workbook = Workbook()
    workbook.creator = "Telegram Services Bot"
    sheet = workbook.active
    sheet.title = sheet_name
    sheet.sheet_view.rightToLeft = True
    sheet.freeze_panes = "A2"
    sheet.append(columns)
    for row in rows:
        sheet.append(
            [safe_spreadsheet_value(row.get(column, "")) for column in columns]
        )
    sheet.auto_filter.ref = (
        f"A1:{get_column_letter(len(columns))}{max(1, len(rows) + 1)}"
    )

    fill = PatternFill(fill_type="solid", fgColor=header_color)
    for cell in sheet[1]:
        cell.font = Font(bold=True, color="1F2937")
        cell.fill = fill
        cell.alignment = Alignment(horizontal="center", vertical="center")
    for column_index, column in enumerate(columns, start=1):
        longest = max(
            [len(str(row.get(column, ""))) for row in rows] + [len(column)]
        )
        sheet.column_dimensions[get_column_letter(column_index)].width = min(
            max(longest + 3, 14), 48
        )

    output = io.BytesIO()
    workbook.save(output)
    workbook.close()
    return output.getvalue()


def _safe_source_filename(filename: str) -> str:
    basename = str(filename).replace("\\", "/").split("/")[-1]
    return re.sub(r"[\x00-\x1f\x7f]", "", basename)[:200] or "contacts.xlsx"


def _summary_text(
    sources: list[tuple[str, str]],
    country_code: str,
    summary: dict,
) -> str:
    created_at = datetime.now(UTC).isoformat(timespec="seconds")
    lines = [
        "تقرير تحويل جهات الاتصال",
        f"تاريخ الإنشاء (UTC): {created_at}",
    ]
    if len(sources) == 1:
        filename, label = sources[0]
        lines += [
            f"الملف المصدر: {_safe_source_filename(filename)}",
            f"ورقة العمل المستخدمة: {label}",
        ]
    else:
        lines.append(f"الملفات المصدر ({len(sources)}):")
        lines += [
            f"  - {_safe_source_filename(filename)}: {label}"
            for filename, label in sources
        ]
    lines += [
        f"إجمالي السجلات: {summary['totalRows']}",
        f"جهات الاتصال الفريدة المقبولة: {summary['validCount']}",
        f"السجلات المكررة التي تم دمجها: {summary['duplicateCount']}",
        f"الصفوف التي تحتاج مراجعة: {summary['invalidCount']}",
        "طريقة إزالة التكرار: السجلات التي تشترك برقم بعد توحيد صيغته جهة واحدة، "
        "يُحتفظ بأول ظهور وتُضاف إليه الأرقام الجديدة",
        f"رمز الدولة الافتراضي للأرقام المحلية: +{country_code}",
        "",
    ]
    return "\n".join(lines)


def _google_csv(contacts: list[dict]) -> bytes:
    """A CSV in Google Contacts' own import format.

    The whole name goes in "First Name": splitting an Arabic full name into
    given and family names cannot be done reliably. No BOM: Google would read
    it as part of the first header.
    """
    phone_columns = max((len(c["phones"]) for c in contacts), default=1)
    header = ["First Name", "Last Name", "E-mail 1 - Label", "E-mail 1 - Value"]
    for number in range(1, phone_columns + 1):
        header += [f"Phone {number} - Label", f"Phone {number} - Value"]
    output = io.StringIO()
    writer = csv.writer(output, lineterminator="\r\n")
    writer.writerow(header)
    for contact in contacts:
        email = contact.get(EMAIL, "")
        row = [contact[FULL_NAME], "", "* Other" if email else "", email]
        for index in range(phone_columns):
            phone = contact["phones"][index] if index < len(contact["phones"]) else ""
            row += ["Mobile" if phone else "", phone]
        writer.writerow(row)
    return output.getvalue().encode("utf-8")


def read_source(data: bytes, extension: str) -> tuple[list[dict], str]:
    """Rows and a source label for one file of any supported format."""
    if extension in {".xlsx", ".xls"}:
        matrix, sheet_name = (
            _read_xlsx(data) if extension == ".xlsx" else _read_xls(data)
        )
        rows = _extract_rows(matrix)
        for index, row in enumerate(rows):
            row["_row"] = index + 2  # the worksheet row, after the header
        return rows, sheet_name
    try:
        return read_contacts(data, extension)
    except ContactsReadError as error:
        raise WorkbookValidationError(str(error)) from error


def build_result(
    sources: list[tuple[str, list[dict], str]],
    country_code: str,
) -> dict:
    """Merge the rows of one or more files into the result bundle.

    `sources` holds (file name, rows, label) for each file, in upload order.
    """
    rows: list[dict] = []
    several = len(sources) > 1
    for index, (filename, file_rows, _) in enumerate(sources):
        name = _safe_source_filename(filename)
        for row in file_rows:
            row["_source"] = index
            if several:
                row["_row"] = f"{name} - {row['_row']}"
            rows.append(row)
    if len({(r["_source"], str(r["_row"])) for r in rows}) > MAX_ROWS:
        raise WorkbookValidationError(
            f"الحد الأعلى المدعوم هو {MAX_ROWS:,} جهة لكل عملية."
        )

    result = process_contact_rows(rows, country_code)
    contacts = result["valid_rows"]
    vcard_buffer = create_vcard(contacts).encode("utf-8")

    archive_buffer = io.BytesIO()
    with zipfile.ZipFile(
        archive_buffer,
        "w",
        compression=zipfile.ZIP_DEFLATED,
        compresslevel=6,
    ) as archive:
        archive.writestr("contacts.vcf", vcard_buffer)
        archive.writestr("google_contacts.csv", _google_csv(contacts))
        archive.writestr(
            "clean_contacts.xlsx",
            _create_excel_report(
                "جهات الاتصال الفريدة",
                VALID_COLUMNS,
                contacts,
                "D9EAD3",
            ),
        )
        archive.writestr(
            "merged_duplicates.xlsx",
            _create_excel_report(
                "مكررات تم دمجها",
                DUPLICATE_COLUMNS,
                result["duplicate_rows"],
                "FCE5CD",
            ),
        )
        archive.writestr(
            "invalid_rows.xlsx",
            _create_excel_report(
                "صفوف تحتاج مراجعة",
                INVALID_COLUMNS,
                result["invalid_rows"],
                "F4CCCC",
            ),
        )
        archive.writestr(
            "summary.txt",
            _summary_text(
                [(filename, label) for filename, _, label in sources],
                country_code,
                result["summary"],
            ),
        )

    zip_buffer = archive_buffer.getvalue()
    if len(zip_buffer) > MAX_OUTPUT_SIZE:
        raise WorkbookValidationError("حجم ملف النتائج يتجاوز الحد المسموح.")
    return {
        **result,
        "zip_buffer": zip_buffer,
        "vcard_buffer": vcard_buffer,
        "source_sheet_name": " | ".join(label for _, _, label in sources),
    }


def process_workbook(
    data: bytes,
    extension: str,
    source_filename: str,
    country_code: str,
) -> dict:
    """One file to the result bundle."""
    rows, label = read_source(data, extension)
    return build_result([(source_filename, rows, label)], country_code)
