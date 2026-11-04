import io
import zipfile

import pytest
from docx import Document
from docx.shared import Inches
from lxml import etree
from PIL import Image
from pypdf import PdfReader

from services.office_to_pdf import (
    OfficeToPdfError,
    libreoffice_path,
    office_to_pdf,
    prepare_docx,
    word_to_pdf,
)
from services.office_to_pdf.docx_fixes import W, fix_part

NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ' \
     'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"'


def body(inner: str) -> bytes:
    return f'<w:document {NS}><w:body>{inner}</w:body></w:document>'.encode()


def fixed(inner: str):
    xml, changes = fix_part(body(inner))
    return etree.fromstring(xml), changes


# --- docx_fixes: pure XML ----------------------------------------------------

def test_size_and_bold_reach_arabic_in_unmarked_runs():
    root, _ = fixed('<w:p><w:r><w:rPr><w:b/><w:sz w:val="52"/></w:rPr><w:t>عنوان</w:t></w:r></w:p>')
    rpr = root.find(f".//{W}rPr")
    assert rpr.find(W + "szCs").get(W + "val") == "52"
    assert rpr.find(W + "bCs") is not None


def test_right_to_left_runs_and_existing_values_are_left_alone():
    root, _ = fixed(
        '<w:p><w:r><w:rPr><w:sz w:val="40"/><w:rtl/></w:rPr><w:t>أ</w:t></w:r>'
        '<w:r><w:rPr><w:sz w:val="40"/><w:szCs w:val="28"/></w:rPr><w:t>ب</w:t></w:r></w:p>'
    )
    first, second = root.findall(f".//{W}rPr")
    assert first.find(W + "szCs") is None
    assert second.find(W + "szCs").get(W + "val") == "28"


def test_toc_wrapper_goes_but_entries_and_page_fields_stay():
    root, _ = fixed(
        '<w:sdt><w:sdtPr><w:alias w:val="المحتويات"/></w:sdtPr><w:sdtContent>'
        '<w:p><w:pPr><w:pStyle w:val="TOC1"/></w:pPr>'
        '<w:r><w:fldChar w:fldCharType="begin"/></w:r>'
        '<w:r><w:instrText>TOC \\h \\o "1-2"</w:instrText></w:r>'
        '<w:r><w:rPr/></w:r>'
        '<w:r><w:fldChar w:fldCharType="separate"/></w:r>'
        '<w:hyperlink w:anchor="_Toc1"><w:r><w:rPr><w:rStyle w:val="Hyperlink"/></w:rPr><w:t>مقدمة</w:t></w:r>'
        '<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText> PAGEREF _Toc1 \\h </w:instrText></w:r>'
        '<w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>3</w:t></w:r>'
        '<w:r><w:fldChar w:fldCharType="end"/></w:r></w:hyperlink></w:p>'
        '<w:p><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>'
        '</w:sdtContent></w:sdt>'
    )
    assert root.find(f".//{W}sdt") is None
    codes = [e.text for e in root.iter(W + "instrText")]
    assert codes == [" PAGEREF _Toc1 \\h "]
    assert [t.text for t in root.iter(W + "t")] == ["مقدمة", "3"]
    assert root.find(f".//{W}rStyle") is None


def test_picture_only_paragraph_is_marked_left_to_right():
    root, _ = fixed(
        '<w:p><w:pPr><w:spacing w:after="0"/></w:pPr><w:r><w:drawing><wp:anchor/></w:drawing></w:r></w:p>'
        '<w:p><w:pPr><w:bidi/></w:pPr><w:r><w:drawing><wp:anchor/></w:drawing></w:r><w:r><w:t>نص</w:t></w:r></w:p>'
    )
    first, second = root.findall(f".//{W}p")
    bidi = first.find(f"{W}pPr/{W}bidi")
    assert bidi.get(W + "val") == "0"
    assert first.find(f"{W}pPr")[0] is bidi  # schema order: before spacing
    assert second.find(f"{W}pPr/{W}bidi").get(W + "val") is None


def arabic_docx(path, with_image=False):
    document = Document()
    title = document.add_paragraph()
    run = title.add_run("تقرير المديرية الشهري LO-check-2026")
    run.bold = True
    table = document.add_table(rows=2, cols=2)
    table.style = "Table Grid"
    for cell, text in zip(table.rows[0].cells + table.rows[1].cells, ["العمل", "الحالة", "منصة الكروت", "منجز"]):
        cell.text = text
    if with_image:
        picture = io.BytesIO()
        Image.new("RGB", (200, 120), "#0a5346").save(picture, "PNG")
        picture.seek(0)
        document.add_picture(picture, width=Inches(2))
    document.save(path)


def test_prepared_copy_still_opens_in_word_tools(tmp_path):
    source, prepared = tmp_path / "a.docx", tmp_path / "b.docx"
    arabic_docx(source)

    changes = prepare_docx(source, prepared)

    assert changes >= 1  # the bold title run gets bCs
    assert Document(str(prepared)).tables[0].cell(1, 0).text == "منصة الكروت"


# --- input checks --------------------------------------------------------------

@pytest.mark.parametrize(
    ("content", "message"),
    [
        (b"", "فارغ"),
        # An old binary .doc renamed to .docx.
        (b"\xd0\xcf\x11\xe0 old binary doc", "ليس ملف Word صالحًا"),
    ],
)
def test_bad_inputs_are_rejected(tmp_path, content, message):
    source = tmp_path / "a.docx"
    source.write_bytes(content)
    with pytest.raises(OfficeToPdfError, match=message):
        word_to_pdf(source, tmp_path / "a.pdf")


def test_password_protected_word_is_rejected(tmp_path):
    source = tmp_path / "locked.docx"
    with zipfile.ZipFile(source, "w") as archive:
        archive.writestr("EncryptionInfo", b"x")
        archive.writestr("EncryptedPackage", b"x")
    with pytest.raises(OfficeToPdfError, match="المحمية بكلمة مرور"):
        word_to_pdf(source, tmp_path / "a.pdf")


# --- real conversion (inside the Docker image, where LibreOffice is) ----------

needs_libreoffice = pytest.mark.skipif(
    libreoffice_path() is None, reason="LibreOffice is installed in the Docker image only"
)


@needs_libreoffice
def test_arabic_document_with_table_and_image_converts(tmp_path):
    source, out = tmp_path / "report.docx", tmp_path / "report.pdf"
    arabic_docx(source, with_image=True)

    result = word_to_pdf(source, out)

    assert result["pages"] == 1
    reader = PdfReader(str(out))
    text = reader.pages[0].extract_text()
    assert "LO-check-2026" in text
    assert reader.pages[0].images, "the picture should be in the PDF"
    fonts = {
        str(font.get_object().get("/BaseFont", ""))
        for font in reader.pages[0]["/Resources"]["/Font"].values()
    }
    # python-docx's template asks for Cambria (-> Caladea) and names no Arabic
    # font, so Arabic must come from the fallback rule, not DejaVu Sans.
    assert any("Caladea" in f for f in fonts), fonts
    assert any("NotoSansArabicUI" in f for f in fonts), fonts
    assert not any("DejaVu" in f for f in fonts), fonts


@needs_libreoffice
def test_two_conversions_can_run_at_once(tmp_path):
    from concurrent.futures import ThreadPoolExecutor

    sources = []
    for name in ("a", "b"):
        arabic_docx(tmp_path / f"{name}.docx")
        sources.append(tmp_path / f"{name}.docx")
    with ThreadPoolExecutor(2) as pool:
        results = list(pool.map(lambda s: word_to_pdf(s, s.with_suffix(".pdf")), sources))
    assert [r["pages"] for r in results] == [1, 1]


@pytest.mark.parametrize("suffix", [".doc", ".xls", ".ppt"])
def test_legacy_office_formats_get_a_clear_message(tmp_path, suffix):
    source = tmp_path / f"old{suffix}"
    source.write_bytes(bytes([0xD0, 0xCF, 0x11, 0xE0]) + b" legacy")
    with pytest.raises(OfficeToPdfError, match="القديمة غير مدعومة"):
        office_to_pdf(source, tmp_path / "out.pdf")


def test_workbook_content_must_match_its_extension(tmp_path):
    source = tmp_path / "report.xlsx"
    arabic_docx(source)  # a Word file renamed to .xlsx
    with pytest.raises(OfficeToPdfError, match="لا يطابق صيغة Excel"):
        office_to_pdf(source, tmp_path / "out.pdf")


def wide_arabic_workbook(path):
    from openpyxl import Workbook

    workbook = Workbook()
    for title in ("الموظفون", "المراسلات"):
        sheet = workbook.active if title == "الموظفون" else workbook.create_sheet(title)
        sheet.title = title
        sheet.sheet_view.rightToLeft = True
        sheet.append([f"العمود {n}" for n in range(1, 31)])
        for row in range(1, 80):
            sheet.append([f"قيمة {row}-{n}" for n in range(1, 31)])
    workbook.save(path)


@needs_libreoffice
def test_excel_fit_puts_each_sheet_on_one_page(tmp_path):
    source = tmp_path / "wide.xlsx"
    wide_arabic_workbook(source)

    fit = office_to_pdf(source, tmp_path / "fit.pdf", sheet_layout="fit")
    printed = office_to_pdf(source, tmp_path / "print.pdf", sheet_layout="print")

    assert fit["pages"] == 2
    assert printed["pages"] > 2


@needs_libreoffice
def test_powerpoint_slides_become_pages(tmp_path):
    from pptx import Presentation

    source = tmp_path / "deck.pptx"
    deck = Presentation()
    for n in range(3):
        slide = deck.slides.add_slide(deck.slide_layouts[1])
        slide.shapes.title.text = f"الشريحة {n + 1} SLIDE-{n + 1}"
        slide.placeholders[1].text = "نقطة عربية"
    deck.save(source)

    result = office_to_pdf(source, tmp_path / "deck.pdf")

    assert result["pages"] == 3
    assert "SLIDE-3" in PdfReader(str(tmp_path / "deck.pdf")).pages[2].extract_text()


def test_unknown_sheet_layout_is_rejected(tmp_path):
    source = tmp_path / "wide.xlsx"
    wide_arabic_workbook(source)
    with pytest.raises(OfficeToPdfError, match="تخطيط"):
        office_to_pdf(source, tmp_path / "x.pdf", sheet_layout="zoom")
