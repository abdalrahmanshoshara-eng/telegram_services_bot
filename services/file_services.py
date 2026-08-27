import io
import os
import re
from pathlib import Path

import qrcode
from PIL import Image, ImageOps
from pypdf import PdfReader, PdfWriter
from reportlab.lib.colors import Color
from reportlab.lib.utils import ImageReader
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfgen import canvas


class ServiceInputError(ValueError):
    """Raised when an uploaded file or user option is invalid."""


MAX_IMAGE_PIXELS = int(os.environ.get("MAX_IMAGE_PIXELS", "40000000"))
MAX_PDF_PAGES = int(os.environ.get("MAX_PDF_PAGES", "500"))


def _validate_image(image: Image.Image) -> None:
    if image.width * image.height > MAX_IMAGE_PIXELS:
        raise ServiceInputError(
            f"أبعاد الصورة كبيرة جدًا. الحد الأقصى {MAX_IMAGE_PIXELS:,} بكسل."
        )


def _validate_pdf(reader: PdfReader) -> None:
    if reader.is_encrypted:
        raise ServiceInputError("لا يمكن معالجة ملف PDF محمي بكلمة مرور.")
    if len(reader.pages) > MAX_PDF_PAGES:
        raise ServiceInputError(f"عدد صفحات PDF يتجاوز الحد ({MAX_PDF_PAGES} صفحة).")


def merge_pdfs(sources: list[Path], destination: Path) -> None:
    if len(sources) < 2:
        raise ServiceInputError("يلزم ملفا PDF على الأقل.")
    writer = PdfWriter()
    total_pages = 0
    for source in sources:
        reader = PdfReader(str(source), strict=False)
        _validate_pdf(reader)
        total_pages += len(reader.pages)
        if total_pages > MAX_PDF_PAGES:
            raise ServiceInputError(f"مجموع الصفحات يتجاوز الحد ({MAX_PDF_PAGES} صفحة).")
        writer.append(reader)
    with destination.open("wb") as stream:
        writer.write(stream)


def parse_page_ranges(spec: str, page_count: int) -> list[int]:
    normalized = spec.translate(str.maketrans("٠١٢٣٤٥٦٧٨٩", "0123456789"))
    normalized = normalized.replace("،", ",").replace(" ", "")
    if not normalized or not re.fullmatch(r"\d+(?:-\d+)?(?:,\d+(?:-\d+)?)*", normalized):
        raise ServiceInputError("صيغة الصفحات غير صحيحة. مثال: 1-3,5,8")

    pages: list[int] = []
    for part in normalized.split(","):
        if "-" in part:
            start, end = map(int, part.split("-", 1))
            if start > end:
                raise ServiceInputError("بداية النطاق يجب أن تكون أصغر من نهايته.")
            selected = range(start, end + 1)
        else:
            selected = [int(part)]
        for page_number in selected:
            if not 1 <= page_number <= page_count:
                raise ServiceInputError(f"رقم الصفحة {page_number} خارج الملف ({page_count} صفحة).")
            zero_based = page_number - 1
            if zero_based not in pages:
                pages.append(zero_based)
    return pages


def extract_pdf_pages(source: Path, destination: Path, page_spec: str) -> int:
    reader = PdfReader(str(source), strict=False)
    _validate_pdf(reader)
    pages = parse_page_ranges(page_spec, len(reader.pages))
    writer = PdfWriter()
    for page_index in pages:
        writer.add_page(reader.pages[page_index])
    with destination.open("wb") as stream:
        writer.write(stream)
    return len(pages)


def parse_page_order(spec: str, page_count: int) -> list[int]:
    normalized = spec.translate(str.maketrans("٠١٢٣٤٥٦٧٨٩", "0123456789"))
    normalized = normalized.replace("،", ",").replace(" ", "")
    if not normalized or not re.fullmatch(r"\d+(?:-\d+)?(?:,\d+(?:-\d+)?)*", normalized):
        raise ServiceInputError("صيغة ترتيب الصفحات غير صحيحة. مثال: 3,1,2,5-4")

    pages: list[int] = []
    for part in normalized.split(","):
        if "-" in part:
            start, end = map(int, part.split("-", 1))
            step = 1 if start <= end else -1
            selected = range(start, end + step, step)
        else:
            selected = [int(part)]
        for page_number in selected:
            if not 1 <= page_number <= page_count:
                raise ServiceInputError(
                    f"رقم الصفحة {page_number} خارج الملف ({page_count} صفحة)."
                )
            zero_based = page_number - 1
            if zero_based in pages:
                raise ServiceInputError(f"الصفحة {page_number} مكررة في الترتيب.")
            pages.append(zero_based)

    if len(pages) != page_count or set(pages) != set(range(page_count)):
        raise ServiceInputError(
            "يجب أن يتضمن الترتيب جميع صفحات الملف مرة واحدة دون حذف أو تكرار."
        )
    return pages


def manage_pdf_pages(
    source: Path,
    destination: Path,
    operation: str,
    page_spec: str,
    angle: int = 90,
) -> dict[str, int]:
    reader = PdfReader(str(source), strict=False)
    _validate_pdf(reader)
    page_count = len(reader.pages)
    writer = PdfWriter()

    if operation == "rotate":
        normalized = page_spec.strip().lower()
        if normalized in {"all", "الكل", "كامل", "جميع"}:
            selected = list(range(page_count))
        else:
            selected = parse_page_ranges(page_spec, page_count)
        if angle not in {-90, 90, 180}:
            raise ServiceInputError("زاوية التدوير يجب أن تكون 90 أو -90 أو 180 درجة.")
        for index, page in enumerate(reader.pages):
            if index in selected:
                page.rotate(angle)
            writer.add_page(page)
        affected = len(selected)
    elif operation == "delete":
        selected = set(parse_page_ranges(page_spec, page_count))
        if len(selected) == page_count:
            raise ServiceInputError("لا يمكن حذف جميع صفحات الملف.")
        for index, page in enumerate(reader.pages):
            if index not in selected:
                writer.add_page(page)
        affected = len(selected)
    elif operation == "reorder":
        order = parse_page_order(page_spec, page_count)
        for index in order:
            writer.add_page(reader.pages[index])
        affected = page_count
    else:
        raise ServiceInputError("عملية إدارة صفحات PDF غير معروفة.")

    if reader.metadata:
        writer.add_metadata(reader.metadata)
    with destination.open("wb") as stream:
        writer.write(stream)
    return {"pages": len(writer.pages), "affected": affected}


def protect_pdf(source: Path, destination: Path, password: str) -> int:
    password = password.strip()
    if not 4 <= len(password) <= 128:
        raise ServiceInputError("يجب أن تتكون كلمة المرور من 4 إلى 128 محرفًا.")
    reader = PdfReader(str(source), strict=False)
    if reader.is_encrypted:
        raise ServiceInputError("الملف محمي مسبقًا. استخدم خدمة فك الحماية أولًا.")
    _validate_pdf(reader)
    writer = PdfWriter()
    writer.append(reader)
    writer.encrypt(
        user_password=password,
        owner_password=password,
        algorithm="AES-256-R5",
    )
    with destination.open("wb") as stream:
        writer.write(stream)
    return len(reader.pages)


def unprotect_pdf(source: Path, destination: Path, password: str) -> int:
    if not password:
        raise ServiceInputError("أرسل كلمة مرور الملف.")
    reader = PdfReader(str(source), strict=False)
    if not reader.is_encrypted:
        raise ServiceInputError("الملف غير محمي بكلمة مرور.")
    try:
        result = reader.decrypt(password)
    except Exception as error:
        raise ServiceInputError("تعذر فك الملف. تحقق من كلمة المرور.") from error
    if result == 0:
        raise ServiceInputError("كلمة المرور غير صحيحة.")
    if len(reader.pages) > MAX_PDF_PAGES:
        raise ServiceInputError(f"عدد صفحات PDF يتجاوز الحد ({MAX_PDF_PAGES} صفحة).")

    writer = PdfWriter()
    writer.append(reader)
    with destination.open("wb") as stream:
        writer.write(stream)
    return len(reader.pages)


def images_to_pdf(sources: list[Path], destination: Path) -> None:
    if not sources:
        raise ServiceInputError("أرسل صورة واحدة على الأقل.")
    converted: list[Image.Image] = []
    try:
        for source in sources:
            with Image.open(source) as image:
                _validate_image(image)
                fixed = ImageOps.exif_transpose(image)
                if fixed.mode == "RGBA":
                    background = Image.new("RGB", fixed.size, "white")
                    background.paste(fixed, mask=fixed.getchannel("A"))
                    converted.append(background)
                else:
                    converted.append(fixed.convert("RGB"))
        first, *rest = converted
        first.save(destination, "PDF", save_all=True, append_images=rest, resolution=150)
    finally:
        for image in converted:
            image.close()


def optimize_image(source: Path, destination: Path, operation: str) -> None:
    with Image.open(source) as image:
        _validate_image(image)
        fixed = ImageOps.exif_transpose(image)
        if operation == "half":
            fixed = fixed.resize(
                (max(1, fixed.width // 2), max(1, fixed.height // 2)),
                Image.Resampling.LANCZOS,
            )
            operation = "jpeg"

        if operation == "jpeg":
            if fixed.mode == "RGBA":
                background = Image.new("RGB", fixed.size, "white")
                background.paste(fixed, mask=fixed.getchannel("A"))
                fixed = background
            fixed.convert("RGB").save(destination, "JPEG", quality=82, optimize=True, progressive=True)
        elif operation == "webp":
            fixed.save(destination, "WEBP", quality=82, method=6)
        elif operation == "png":
            fixed.save(destination, "PNG", optimize=True)
        else:
            raise ServiceInputError("خيار تحويل الصورة غير معروف.")


def create_qr_code(data: str, destination: Path) -> None:
    data = data.strip()
    if not data:
        raise ServiceInputError("لا يمكن إنشاء QR من نص فارغ.")
    if len(data) > 2000:
        raise ServiceInputError("النص طويل جدًا. الحد الأقصى 2000 حرف.")
    qr = qrcode.QRCode(version=None, error_correction=qrcode.constants.ERROR_CORRECT_M, box_size=12, border=4)
    qr.add_data(data)
    qr.make(fit=True)
    qr.make_image(fill_color="black", back_color="white").save(destination)


def _arabic_display(text: str) -> str:
    try:
        import arabic_reshaper
        from bidi.algorithm import get_display

        return get_display(arabic_reshaper.reshape(text))
    except Exception:
        return text


def _watermark_page(page, text: str | None = None, logo: Path | None = None):
    width = float(page.mediabox.width)
    height = float(page.mediabox.height)
    packet = io.BytesIO()
    overlay = canvas.Canvas(packet, pagesize=(width, height))
    overlay.saveState()
    if hasattr(overlay, "setFillAlpha"):
        overlay.setFillAlpha(0.14)

    if logo:
        with Image.open(logo) as image:
            _validate_image(image)
            image.thumbnail((width * 0.38, height * 0.38), Image.Resampling.LANCZOS)
            logo_width, logo_height = image.size
            overlay.drawImage(
                ImageReader(image),
                (width - logo_width) / 2,
                (height - logo_height) / 2,
                logo_width,
                logo_height,
                mask="auto",
                preserveAspectRatio=True,
            )
    else:
        font_name = "Helvetica-Bold"
        font_path = Path("/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf")
        if font_path.exists():
            font_name = "DejaVuSans"
            if font_name not in pdfmetrics.getRegisteredFontNames():
                pdfmetrics.registerFont(TTFont(font_name, str(font_path)))
        display_text = _arabic_display(text or "")
        font_size = max(18, min(36, width / max(10, len(display_text)) * 1.35))
        overlay.setFont(font_name, font_size)
        overlay.setFillColor(Color(0.35, 0.35, 0.35))
        overlay.translate(width / 2, height / 2)
        overlay.rotate(35)
        overlay.drawCentredString(0, 0, display_text)
    overlay.restoreState()
    overlay.save()
    packet.seek(0)
    page.merge_page(PdfReader(packet).pages[0])


def watermark_pdf(source: Path, destination: Path, text: str | None = None, logo: Path | None = None) -> None:
    if not text and not logo:
        raise ServiceInputError("يجب إرسال نص أو صورة للعلامة المائية.")
    if text and len(text) > 200:
        raise ServiceInputError("نص العلامة المائية طويل جدًا. الحد الأقصى 200 حرف.")
    reader = PdfReader(str(source), strict=False)
    _validate_pdf(reader)
    writer = PdfWriter()
    for page in reader.pages:
        writer.add_page(page)
        _watermark_page(writer.pages[-1], text=text, logo=logo)
    with destination.open("wb") as stream:
        writer.write(stream)
