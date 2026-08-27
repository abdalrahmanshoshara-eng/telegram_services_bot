from pathlib import Path

import pytest
from PIL import Image
from pypdf import PdfReader
from reportlab.pdfgen import canvas

from services.file_services import (
    ServiceInputError,
    create_qr_code,
    extract_pdf_pages,
    images_to_pdf,
    manage_pdf_pages,
    merge_pdfs,
    optimize_image,
    parse_page_order,
    parse_page_ranges,
    protect_pdf,
    unprotect_pdf,
    watermark_pdf,
)


def make_pdf(path: Path, labels: list[str]) -> None:
    output = canvas.Canvas(str(path))
    for label in labels:
        output.drawString(72, 720, label)
        output.showPage()
    output.save()


def test_merge_and_extract_pdf_pages(tmp_path):
    first = tmp_path / "first.pdf"
    second = tmp_path / "second.pdf"
    merged = tmp_path / "merged.pdf"
    extracted = tmp_path / "extracted.pdf"
    make_pdf(first, ["A", "B"])
    make_pdf(second, ["C"])

    merge_pdfs([first, second], merged)
    assert len(PdfReader(merged).pages) == 3

    count = extract_pdf_pages(merged, extracted, "١-٢،٣")
    assert count == 3
    assert len(PdfReader(extracted).pages) == 3


def test_page_range_validation():
    assert parse_page_ranges("1-3,3,5", 5) == [0, 1, 2, 4]
    with pytest.raises(ServiceInputError):
        parse_page_ranges("5-2", 5)
    with pytest.raises(ServiceInputError):
        parse_page_ranges("1,9", 5)


def test_manage_pdf_pages_rotate_delete_and_reorder(tmp_path):
    source = tmp_path / "source.pdf"
    rotated = tmp_path / "rotated.pdf"
    deleted = tmp_path / "deleted.pdf"
    reordered = tmp_path / "reordered.pdf"
    make_pdf(source, ["A", "B", "C"])

    result = manage_pdf_pages(source, rotated, "rotate", "2", 90)
    assert result == {"pages": 3, "affected": 1}
    rotated_reader = PdfReader(rotated)
    assert rotated_reader.pages[0].rotation == 0
    assert rotated_reader.pages[1].rotation == 90

    result = manage_pdf_pages(source, deleted, "delete", "2")
    assert result == {"pages": 2, "affected": 1}
    assert [page.extract_text().strip() for page in PdfReader(deleted).pages] == ["A", "C"]

    result = manage_pdf_pages(source, reordered, "reorder", "3,1,2")
    assert result == {"pages": 3, "affected": 3}
    assert [page.extract_text().strip() for page in PdfReader(reordered).pages] == [
        "C",
        "A",
        "B",
    ]


def test_page_order_must_contain_every_page_once():
    assert parse_page_order("3,1,2,5-4", 5) == [2, 0, 1, 4, 3]
    with pytest.raises(ServiceInputError, match="جميع صفحات"):
        parse_page_order("1,2", 3)
    with pytest.raises(ServiceInputError, match="مكررة"):
        parse_page_order("1,2,2", 3)


def test_protect_and_unprotect_pdf(tmp_path):
    source = tmp_path / "source.pdf"
    protected = tmp_path / "protected.pdf"
    restored = tmp_path / "restored.pdf"
    make_pdf(source, ["Secret"])

    assert protect_pdf(source, protected, "safe-password") == 1
    encrypted_reader = PdfReader(protected)
    assert encrypted_reader.is_encrypted
    assert encrypted_reader.decrypt("safe-password") != 0
    assert encrypted_reader.pages[0].extract_text().strip() == "Secret"

    with pytest.raises(ServiceInputError, match="غير صحيحة"):
        unprotect_pdf(protected, restored, "wrong-password")
    assert unprotect_pdf(protected, restored, "safe-password") == 1
    restored_reader = PdfReader(restored)
    assert not restored_reader.is_encrypted
    assert restored_reader.pages[0].extract_text().strip() == "Secret"


def test_images_pdf_conversion_qr_and_optimization(tmp_path):
    first = tmp_path / "one.png"
    second = tmp_path / "two.jpg"
    Image.new("RGBA", (120, 80), (255, 0, 0, 128)).save(first)
    Image.new("RGB", (80, 120), "blue").save(second)

    pdf = tmp_path / "images.pdf"
    images_to_pdf([first, second], pdf)
    assert len(PdfReader(pdf).pages) == 2

    optimized = tmp_path / "optimized.webp"
    optimize_image(second, optimized, "webp")
    with Image.open(optimized) as result:
        assert result.format == "WEBP"

    qr = tmp_path / "qr.png"
    create_qr_code("https://example.com", qr)
    with Image.open(qr) as result:
        assert result.width == result.height
        assert result.width > 100


def test_text_and_logo_watermarks(tmp_path):
    source = tmp_path / "source.pdf"
    text_output = tmp_path / "text-watermark.pdf"
    logo_output = tmp_path / "logo-watermark.pdf"
    logo = tmp_path / "logo.png"
    make_pdf(source, ["Confidential document"])
    Image.new("RGBA", (200, 100), (0, 100, 200, 180)).save(logo)

    watermark_pdf(source, text_output, text="نسخة تجريبية")
    watermark_pdf(source, logo_output, logo=logo)

    assert len(PdfReader(text_output).pages) == 1
    assert len(PdfReader(logo_output).pages) == 1
