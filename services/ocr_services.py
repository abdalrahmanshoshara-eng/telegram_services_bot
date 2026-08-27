import os
from pathlib import Path

from PIL import Image, ImageOps


class OcrError(RuntimeError):
    pass


MAX_IMAGE_PIXELS = int(os.environ.get("MAX_IMAGE_PIXELS", "40000000"))


def ocr_image(source: Path, destination: Path, languages: str) -> None:
    import pytesseract

    with Image.open(source) as image:
        if image.width * image.height > MAX_IMAGE_PIXELS:
            raise OcrError(f"أبعاد الصورة تتجاوز الحد ({MAX_IMAGE_PIXELS:,} بكسل).")
        prepared = ImageOps.exif_transpose(image).convert("RGB")
        text = pytesseract.image_to_string(prepared, lang=languages)
    destination.write_text(text.strip() + "\n", encoding="utf-8")
