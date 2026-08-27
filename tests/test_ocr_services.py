import sys
from types import SimpleNamespace

from PIL import Image

from services.ocr_services import ocr_image


def test_image_ocr_writes_utf8_text(tmp_path, monkeypatch):
    source = tmp_path / "source.png"
    destination = tmp_path / "text.txt"
    Image.new("RGB", (30, 30), "white").save(source)
    fake = SimpleNamespace(image_to_string=lambda image, lang: "نص عربي English")
    monkeypatch.setitem(sys.modules, "pytesseract", fake)

    ocr_image(source, destination, "ara+eng")

    assert destination.read_text(encoding="utf-8") == "نص عربي English\n"
