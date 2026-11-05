"""Several contact files merged into one set; one contact keeps all numbers."""

import csv
import io
import zipfile

import pytest
from openpyxl import Workbook, load_workbook

from services.excel_contacts import ExcelContactsError, convert_contacts_files
from services.excel_contacts.domain import process_contact_rows

NAME, PHONE, EMAIL = "الاسم الكامل", "رقم التواصل", "البريد الالكتروني"


def write_xlsx(path, rows):
    workbook = Workbook()
    sheet = workbook.active
    sheet.append([NAME, PHONE, EMAIL])
    for row in rows:
        sheet.append(row)
    workbook.save(path)
    workbook.close()


@pytest.fixture
def three_files(tmp_path):
    vcf = tmp_path / "phone.vcf"
    vcf.write_text(
        "BEGIN:VCARD\r\nVERSION:3.0\r\nFN:أحمد العلي\r\n"
        "TEL:0933123456\r\nTEL:0944123456\r\nEND:VCARD\r\n"
        "BEGIN:VCARD\r\nVERSION:3.0\r\nFN:ليلى\r\nTEL:0955123456\r\nEND:VCARD\r\n",
        encoding="utf-8",
    )
    google = tmp_path / "google.csv"
    google.write_text(
        "First Name,Last Name,E-mail 1 - Value,Phone 1 - Value\n"
        # Shares 0944... with Ahmad from the phone and adds a third number.
        "أحمد,العلي,ahmad@example.com,+963944123456 ::: 0966123456\n"
        "سامر,يوسف,,0977123456\n",
        encoding="utf-8",
    )
    work = tmp_path / "work.xlsx"
    write_xlsx(work, [["ليلى حسن", "00963955123456", ""], ["هبة", "0988123456", ""]])
    return [("phone.vcf", vcf), ("google.csv", google), ("work.xlsx", work)]


def test_files_of_different_formats_become_one_clean_list(tmp_path, three_files):
    destination = tmp_path / "out.zip"
    vcf = tmp_path / "contacts.vcf"

    summary = convert_contacts_files(three_files, destination, "963", vcf)

    assert summary == {
        "totalRows": 6,
        "validCount": 4,
        "duplicateCount": 2,
        "invalidCount": 0,
        "files": 3,
    }
    vcard = vcf.read_text(encoding="utf-8")
    assert vcard.count("BEGIN:VCARD") == 4
    ahmad = vcard.split("BEGIN:VCARD")[1]
    # One contact, three numbers: two from the phone, one added by Google.
    assert ahmad.count("TEL;TYPE=CELL:") == 3
    assert "TEL;TYPE=CELL:+963966123456" in ahmad
    assert "EMAIL;TYPE=INTERNET:ahmad@example.com" in ahmad

    with zipfile.ZipFile(destination) as archive:
        duplicates = load_workbook(io.BytesIO(archive.read("merged_duplicates.xlsx")))
        rows = list(duplicates.active.iter_rows(min_row=2, values_only=True))
        duplicates.close()
        summary_text = archive.read("summary.txt").decode("utf-8")
    locations = {row[0]: row[4] for row in rows}
    assert locations == {"google.csv - 2": "phone.vcf - 1", "work.xlsx - 2": "phone.vcf - 2"}
    assert "الملفات المصدر (3):" in summary_text
    assert "work.xlsx" in summary_text


def test_google_csv_lists_every_number_of_a_contact(tmp_path, three_files):
    destination = tmp_path / "out.zip"
    convert_contacts_files(three_files, destination, "963")

    with zipfile.ZipFile(destination) as archive:
        text = archive.read("google_contacts.csv").decode("utf-8")
    rows = list(csv.reader(io.StringIO(text)))
    assert not text.startswith("﻿")
    assert rows[0][:4] == ["First Name", "Last Name", "E-mail 1 - Label", "E-mail 1 - Value"]
    assert rows[0][-2:] == ["Phone 3 - Label", "Phone 3 - Value"]
    ahmad = rows[1]
    assert ahmad[0] == "أحمد العلي"
    assert ahmad[3] == "ahmad@example.com"
    assert ahmad[5::2] == ["+963933123456", "+963944123456", "+963966123456"]


def test_clean_excel_shows_extra_numbers(tmp_path, three_files):
    destination = tmp_path / "out.zip"
    convert_contacts_files(three_files, destination, "963")

    with zipfile.ZipFile(destination) as archive:
        workbook = load_workbook(io.BytesIO(archive.read("clean_contacts.xlsx")))
    sheet = workbook.active
    assert [cell.value for cell in sheet[1]] == [NAME, PHONE, EMAIL, "أرقام إضافية"]
    # The leading apostrophe stops Excel reading "+963..." as a formula.
    assert sheet["D2"].value == "'+963944123456، +963966123456"
    workbook.close()


def test_a_bad_file_is_named_in_the_error(tmp_path, three_files):
    broken = tmp_path / "broken.pdf"
    broken.write_bytes(b"not a pdf")

    with pytest.raises(ExcelContactsError, match="^broken.pdf: "):
        convert_contacts_files(three_files + [("broken.pdf", broken)], tmp_path / "o.zip", "963")


def test_single_file_errors_are_not_prefixed(tmp_path):
    broken = tmp_path / "broken.pdf"
    broken.write_bytes(b"not a pdf")

    with pytest.raises(ExcelContactsError, match="^محتوى الملف"):
        convert_contacts_files([("broken.pdf", broken)], tmp_path / "o.zip", "963")


def test_one_bad_number_does_not_lose_the_contact():
    rows = [
        {NAME: "أحمد", PHONE: "0933123456", "_row": 3},
        {NAME: "أحمد", PHONE: "12", "_row": 3},
    ]

    result = process_contact_rows(rows, "963")

    assert [r["phones"] for r in result["valid_rows"]] == [["+963933123456"]]
    assert len(result["invalid_rows"]) == 1
    assert "تم حفظ الجهة بأرقامها الأخرى" in result["invalid_rows"][0]["سبب الاستبعاد"]


def test_excel_rows_with_the_same_name_stay_separate():
    # Two people may share a name; only a shared number makes them one contact.
    rows = [
        {NAME: "محمد", PHONE: "0933123456", EMAIL: ""},
        {NAME: "محمد", PHONE: "0944123456", EMAIL: ""},
    ]

    result = process_contact_rows(rows, "963")

    assert result["summary"]["validCount"] == 2
