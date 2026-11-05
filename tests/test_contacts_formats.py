"""Contacts to VCF from CSV, TXT, VCF, DOCX and PDF files."""

import quopri
import zipfile
from pathlib import Path

import pytest
from docx import Document

from services.excel_contacts import ExcelContactsError, convert_excel_contacts
from services.excel_contacts.processor import process_workbook
from services.excel_contacts.readers import _logical, read_contacts

FIXTURES = Path(__file__).parent / "fixtures" / "contacts"

NAME, PHONE, EMAIL = "الاسم الكامل", "رقم التواصل", "البريد الالكتروني"


def read(data: bytes, extension: str) -> dict:
    return process_workbook(data, extension, f"contacts{extension}", "963")


def valid(result: dict) -> list[tuple[str, str, str]]:
    """(name, all numbers joined by commas, email) for each contact."""
    return [(r[NAME], ",".join(r["phones"]), r[EMAIL]) for r in result["valid_rows"]]


# --- CSV -------------------------------------------------------------------

def test_google_contacts_csv_combines_names_and_splits_numbers():
    data = (
        "First Name,Middle Name,Last Name,E-mail 1 - Label,E-mail 1 - Value,"
        "Phone 1 - Label,Phone 1 - Value,Phone 2 - Label,Phone 2 - Value\n"
        "أحمد,,العلي,* Home,ahmad@example.com,Mobile,+963 933 123 456 ::: 0944 000 111,,\n"
        "Lina,,Haddad,,,Mobile,0955-222-333,Work,011 222 3344\n"
        ",,,,,Mobile,0966111222,,\n"
    ).encode("utf-8-sig")

    result = read(data, ".csv")

    assert valid(result) == [
        ("أحمد العلي", "+963933123456,+963944000111", "ahmad@example.com"),
        ("Lina Haddad", "+963955222333,+963112223344", ""),
    ]
    assert [r["الموقع"] for r in result["invalid_rows"]] == [4]
    assert result["source_sheet_name"] == "جدول CSV"


def test_arabic_windows_csv_with_semicolons_is_read():
    data = "الاسم;الكنية;الموبايل\nسامر;يوسف;0933111222\nهبة;سعيد;944222333\n"

    result = read(data.encode("cp1256"), ".csv")

    assert valid(result) == [
        ("سامر يوسف", "+963933111222", ""),
        ("هبة سعيد", "+963944222333", ""),
    ]


def test_family_name_already_in_full_name_is_not_repeated():
    data = "الاسم الكامل,الكنية,الهاتف\nسامر يوسف,يوسف,0933111222\n"

    assert valid(read(data.encode(), ".csv")) == [
        ("سامر يوسف", "+963933111222", "")
    ]


def test_first_and_family_name_columns_are_joined():
    data = "الاسم الأول,اسم الأب,الكنية,رقم الهاتف\nسامر,محمد,يوسف,0933111222\n"

    assert valid(read(data.encode(), ".csv")) == [
        ("سامر محمد يوسف", "+963933111222", "")
    ]


def test_csv_without_header_is_read_line_by_line():
    data = "رامي,0933555666\nنور,0944555666\n".encode()

    result = read(data, ".csv")

    assert valid(result) == [
        ("رامي", "+963933555666", ""),
        ("نور", "+963944555666", ""),
    ]
    assert result["source_sheet_name"] == "نص CSV"


# --- TXT -------------------------------------------------------------------

def test_free_text_lines_numbering_labels_and_split_lines():
    data = (
        "جهات الاتصال\n\n"
        "1- خالد عمر 0933 777 888\n"
        "فاطمة: +963-944-777-888\n"
        "الاسم: يزن\n"
        "الهاتف: 0955777888\n"
        "البريد: yazan@example.com\n"
        "٠٩٦٦٧٧٧٨٨٨ - علا\n"
    ).encode()

    result = read(data, ".txt")

    assert valid(result) == [
        ("خالد عمر", "+963933777888", ""),
        ("فاطمة", "+963944777888", ""),
        ("يزن", "+963955777888", "yazan@example.com"),
        ("علا", "+963966777888", ""),
    ]
    # Row numbers point at the line in the file; a split contact at its name.
    rows, _ = read_contacts(data, ".txt")
    assert [r["_row"] for r in rows] == [3, 4, 5, 8]


def test_number_without_name_is_reported_not_dropped():
    result = read("0933111222\n".encode(), ".txt")

    assert result["summary"]["invalidCount"] == 1
    assert "الاسم الكامل فارغ" in result["invalid_rows"][0]["سبب الاستبعاد"]


def test_text_without_any_number_is_rejected(tmp_path):
    source = tmp_path / "names.txt"
    source.write_text("أحمد\nليلى\n", encoding="utf-8")

    with pytest.raises(ExcelContactsError, match="لم أجد أي جهة اتصال"):
        convert_excel_contacts(source, tmp_path / "out.zip", "963")


# --- VCF -------------------------------------------------------------------

def qp(text: str) -> str:
    return quopri.encodestring(text.encode("utf-8")).decode().replace("\n", "")


def test_vcf_quoted_printable_multiple_numbers_and_duplicates():
    data = (
        "BEGIN:VCARD\r\nVERSION:2.1\r\n"
        f"N;CHARSET=UTF-8;ENCODING=QUOTED-PRINTABLE:{qp('الحسن')};{qp('محمد')};;;\r\n"
        f"FN;CHARSET=UTF-8;ENCODING=QUOTED-PRINTABLE:{qp('محمد الحسن')}\r\n"
        "TEL;CELL:0933-000-111\r\nEND:VCARD\r\n"
        "BEGIN:VCARD\r\nVERSION:3.0\r\nN:Ali;Sara;;;\r\n"
        "TEL;TYPE=CELL:+963944000222\r\nitem1.TEL:0955000333\r\n"
        "EMAIL;TYPE=INTERNET:sara@example.com\r\nEND:VCARD\r\n"
        "BEGIN:VCARD\r\nVERSION:3.0\r\nFN:نسخة مكررة\r\nTEL:+963933000111\r\nEND:VCARD\r\n"
        "BEGIN:VCARD\r\nVERSION:3.0\r\nFN:بلا رقم\r\nEND:VCARD\r\n"
    ).encode()

    result = read(data, ".vcf")

    assert valid(result) == [
        ("محمد الحسن", "+963933000111", ""),
        ("Sara Ali", "+963944000222,+963955000333", "sara@example.com"),
    ]
    assert result["summary"]["duplicateCount"] == 1
    assert result["invalid_rows"][0][NAME] == "بلا رقم"


def test_vcf_folded_lines_are_joined():
    data = (
        "BEGIN:VCARD\r\nVERSION:3.0\r\nFN:Very Long\r\n  Name\r\n"
        "TEL:0933\r\n 123456\r\nEND:VCARD\r\n"
    ).encode()

    assert valid(read(data, ".vcf")) == [("Very Long Name", "+963933123456", "")]


# --- DOCX ------------------------------------------------------------------

def make_docx_table(path: Path) -> None:
    document = Document()
    document.add_paragraph("قائمة موظفي المديرية")
    table = document.add_table(rows=1, cols=4)
    for cell, text in zip(table.rows[0].cells, ["م", "الاسم", "رقم الهاتف", "البريد الإلكتروني"]):
        cell.text = text
    for row in [
        ("1", "أحمد محمد العلي", "0933 123 456", "ahmad@example.com"),
        ("2", "ليلى حسن", "+963 944 123 456", ""),
        ("3", "أحمد مكرر", "0933123456", ""),
    ]:
        for cell, text in zip(table.add_row().cells, row):
            cell.text = text
    document.save(path)


def test_docx_table_is_read_by_its_header(tmp_path):
    source = tmp_path / "staff.docx"
    make_docx_table(source)

    result = read(source.read_bytes(), ".docx")

    assert valid(result) == [
        ("أحمد محمد العلي", "+963933123456", "ahmad@example.com"),
        ("ليلى حسن", "+963944123456", ""),
    ]
    assert result["duplicate_rows"][0]["موقع المكرر"] == "جدول 1 - صف 4"
    assert result["source_sheet_name"] == "جداول Word"


def test_docx_paragraphs_are_read_as_free_text(tmp_path):
    source = tmp_path / "list.docx"
    document = Document()
    for line in ["قائمة الأرقام", "1. سامر يوسف - 0933 111 222", "Omar Saleh, +963966444555"]:
        document.add_paragraph(line)
    document.save(source)

    assert valid(read(source.read_bytes(), ".docx")) == [
        ("سامر يوسف", "+963933111222", ""),
        ("Omar Saleh", "+963966444555", ""),
    ]


# --- PDF -------------------------------------------------------------------
# Both fixtures were printed to PDF by a browser from right-to-left Arabic
# pages, so their text comes out of the PDF in visual order.

def test_rtl_pdf_table_matches_the_same_table_in_word():
    result = read((FIXTURES / "rtl_table.pdf").read_bytes(), ".pdf")

    assert valid(result) == [
        ("أحمد محمد العلي", "+963933123456", "ahmad@example.com"),
        ("ليلى حسن", "+963944123456", ""),
        ("محمد خير", "+963955123456", "mk@example.com"),
    ]
    assert result["summary"]["duplicateCount"] == 1
    assert result["invalid_rows"][0][NAME] == "بدون رقم"
    assert result["source_sheet_name"] == "جداول PDF"


def test_rtl_pdf_free_text_is_put_back_in_reading_order():
    result = read((FIXTURES / "rtl_text.pdf").read_bytes(), ".pdf")

    assert valid(result) == [
        ("سامر يوسف", "+963933111222", ""),
        ("هبة سعيد", "+963944222333", "heba@example.com"),
        ("ريم خالد", "+963955333444", "reem@example.com"),
        ("Omar Saleh", "+963966444555", ""),
        ("رقم بدون اسم", "+963999888777", ""),
    ]


def test_visual_order_line_is_reversed_without_breaking_numbers():
    # "ﻲﻠﻋ" is "علي" laid out left to right; "ﻻ" must expand after reversal.
    assert _logical("456 123 0933 ﻲﻠﻋ 1") == "1 علي 0933 123 456"
    assert _logical("963+ ﻢﺳﻻا") == "الاسم +963"
    assert _logical("Omar Saleh ﻲﻠﻋ") == "علي Omar Saleh"
    assert _logical("0933 123 456") == "0933 123 456"


def test_pdf_without_text_is_rejected(tmp_path):
    from reportlab.pdfgen import canvas

    source = tmp_path / "scan.pdf"
    page = canvas.Canvas(str(source))
    page.rect(50, 50, 200, 200, fill=1)
    page.save()

    with pytest.raises(ExcelContactsError, match="ممسوح ضوئيًا"):
        convert_excel_contacts(source, tmp_path / "out.zip", "963")


# --- File checks and the full bundle ---------------------------------------

@pytest.mark.parametrize(
    ("name", "content", "message"),
    [
        ("contacts.pdf", b"not a pdf", "صيغة PDF"),
        ("contacts.docx", b"not a docx", "صيغة Word"),
        ("contacts.csv", b"\x00\x01\x02binary", "ملف نصي"),
        ("contacts.doc", b"old word", "الصيغ المدعومة"),
    ],
)
def test_wrong_content_or_type_is_rejected(tmp_path, name, content, message):
    source = tmp_path / name
    source.write_bytes(content)

    with pytest.raises(ExcelContactsError, match=message):
        convert_excel_contacts(source, tmp_path / "out.zip", "963")


def test_csv_produces_the_same_result_bundle_as_excel(tmp_path):
    source = tmp_path / "contacts.csv"
    source.write_text("Name,Phone,Email\nأحمد,0933123456,a@example.com\n", encoding="utf-8")
    destination = tmp_path / "out.zip"
    vcf = tmp_path / "contacts.vcf"

    summary = convert_excel_contacts(source, destination, "963", "contacts.csv", vcf)

    assert summary["validCount"] == 1
    with zipfile.ZipFile(destination) as archive:
        assert set(archive.namelist()) == {
            "contacts.vcf",
            "google_contacts.csv",
            "clean_contacts.xlsx",
            "merged_duplicates.xlsx",
            "invalid_rows.xlsx",
            "summary.txt",
        }
        assert "ورقة العمل المستخدمة: جدول CSV" in archive.read("summary.txt").decode()
    assert "TEL;TYPE=CELL:+963933123456" in vcf.read_text(encoding="utf-8")
    assert "EMAIL;TYPE=INTERNET:a@example.com" in vcf.read_text(encoding="utf-8")
