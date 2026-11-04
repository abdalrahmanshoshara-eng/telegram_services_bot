import numpy as np
import pytest
from PIL import Image
from pypdf import PdfReader
from reportlab.lib.utils import ImageReader
from reportlab.pdfgen import canvas

from services.file_services import ServiceInputError, compress_pdf


def photo(width=2400, height=1800, alpha=False):
    y, x = np.mgrid[0:height, 0:width]
    pixels = np.stack([x / width * 255, y / height * 255, (x + y) % 256], -1)
    noise = np.random.default_rng(1).normal(0, 8, pixels.shape)
    image = Image.fromarray((pixels + noise).clip(0, 255).astype("uint8"), "RGB")
    if alpha:
        image.putalpha(200)
    return image


def make_pdf(path, image=None, pages=2):
    """Text pages, with an image drawn by reportlab (supports transparency).

    reportlab stores images ASCII85-encoded, which pypdf decodes slowly, so
    keep its images small; use `make_scan` for large photos.
    """
    page = canvas.Canvas(str(path))
    for number in range(pages):
        if image is not None:
            page.drawImage(ImageReader(image), 30, 200, 540, 400, mask="auto")
        for line in range(30):
            page.drawString(30, 800 - line * 15, f"page {number + 1} line {line} " * 3)
        page.showPage()
    page.save()


def make_scan(path, image, pages=2):
    """One high-quality JPEG per page, like a phone or scanner app makes."""
    frames = [image.rotate(180 * n) for n in range(pages)]
    frames[0].save(path, "PDF", save_all=True, append_images=frames[1:], quality=95)


def test_medium_shrinks_photos_and_keeps_pages(tmp_path):
    source, out = tmp_path / "photo.pdf", tmp_path / "out.pdf"
    make_scan(source, photo())

    result = compress_pdf(source, out, "medium")

    assert result["reduced"] and result["images"] == 2
    assert result["after"] < result["before"] / 2
    reader = PdfReader(str(out))
    assert len(reader.pages) == 2
    image = reader.pages[0].images[0].image
    assert max(image.size) == 2000


def test_text_survives_compression(tmp_path):
    source, out = tmp_path / "text.pdf", tmp_path / "out.pdf"
    make_pdf(source, pages=3)

    compress_pdf(source, out, "strong")

    assert "page 3 line 0" in PdfReader(str(out)).pages[2].extract_text()


def test_strong_is_smaller_than_medium(tmp_path):
    source = tmp_path / "photo.pdf"
    make_scan(source, photo())

    medium = compress_pdf(source, tmp_path / "m.pdf", "medium")
    strong = compress_pdf(source, tmp_path / "s.pdf", "strong")

    assert strong["after"] < medium["after"]


def test_lossless_leaves_images_alone(tmp_path):
    source, out = tmp_path / "photo.pdf", tmp_path / "out.pdf"
    make_scan(source, photo(800, 600))

    result = compress_pdf(source, out, "lossless")

    assert result["images"] == 0
    assert PdfReader(str(out)).pages[0].images[0].image.size == (800, 600)


def test_transparent_images_are_not_recompressed(tmp_path):
    source, out = tmp_path / "logo.pdf", tmp_path / "out.pdf"
    make_pdf(source, photo(600, 450, alpha=True), pages=1)

    result = compress_pdf(source, out, "strong")

    assert result["images"] == 0


def test_result_is_never_bigger_than_the_original(tmp_path):
    source, once, twice = tmp_path / "a.pdf", tmp_path / "b.pdf", tmp_path / "c.pdf"
    make_scan(source, photo())
    compress_pdf(source, once, "strong")

    result = compress_pdf(once, twice, "strong")

    assert result["after"] <= result["before"]
    if not result["reduced"]:
        assert twice.read_bytes() == once.read_bytes()


def test_unknown_level_and_encrypted_files_are_rejected(tmp_path):
    from pypdf import PdfWriter

    source = tmp_path / "a.pdf"
    make_pdf(source, pages=1)
    with pytest.raises(ServiceInputError, match="مستوى الضغط"):
        compress_pdf(source, tmp_path / "b.pdf", "extreme")

    writer = PdfWriter(clone_from=str(source))
    writer.encrypt("secret", algorithm="AES-256-R5")
    locked = tmp_path / "locked.pdf"
    with locked.open("wb") as stream:
        writer.write(stream)
    with pytest.raises(ServiceInputError, match="محمي"):
        compress_pdf(locked, tmp_path / "c.pdf", "medium")
