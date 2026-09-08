import asyncio
import logging
import sys
import tempfile
from datetime import UTC, datetime
from pathlib import Path

from telegram import InlineKeyboardButton, InlineKeyboardMarkup, Update
from telegram.constants import ChatAction
from telegram.ext import (
    Application,
    CallbackQueryHandler,
    CommandHandler,
    ContextTypes,
    MessageHandler,
    filters,
)

from config import (
    BOT_TOKEN,
    MAX_DECK_SLIDES,
    MAX_CONCURRENT_JOBS,
    MAX_MULTI_FILES,
    MAX_OUTPUT_MB,
    MAX_TOTAL_UPLOAD_MB,
    MAX_UPLOAD_MB,
    OCR_LANG,
    PROCESSING_TIMEOUT_SECONDS,
    SERVICE_CATEGORIES,
    SERVICES,
    WELCOME_MESSAGE,
)
from services.file_services import (
    ServiceInputError,
    create_qr_code,
    extract_pdf_pages,
    images_to_pdf,
    manage_pdf_pages,
    merge_pdfs,
    optimize_image,
    protect_pdf,
    unprotect_pdf,
    watermark_pdf,
)
from services.colornote import (
    MAX_FILES as MAX_COLORNOTE_FILES,
    ColorNoteConversionError,
    convert_colornote_backups,
)
from services.excel_contacts import (
    MAX_FILE_SIZE as EXCEL_CONTACTS_MAX_FILE_SIZE,
    ExcelContactsError,
    convert_excel_contacts,
)
from services.excel_tools import (
    ExcelToolsError,
    clean_excel_workbook,
    merge_excel_workbooks,
    split_excel_workbook,
)
from services.pptx_deck import (
    DEFAULT_SLIDE_COUNT,
    DeckError,
    SLIDE_COUNT_CHOICES,
    api_key_present as gemini_key_present,
    build_deck_outline,
    outline_from_dict,
    outline_preview,
    outline_to_dict,
    render_outline,
)
from services.ocr_services import ocr_image
from services.text_editor import (
    MAX_INPUT_CHARS as TEXT_EDITOR_MAX_CHARS,
    MAX_MESSAGE_CHARS as TEXT_EDITOR_MAX_MESSAGE_CHARS,
    TextEditError,
    polish_text,
)
from services.word_formatter import format_word_tables

logging.basicConfig(
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
    level=logging.INFO,
)
# HTTP client INFO logs contain the Telegram bot token in request URLs.
logging.getLogger("httpx").setLevel(logging.WARNING)
logging.getLogger("httpcore").setLevel(logging.WARNING)
logger = logging.getLogger(__name__)
PROCESSING_SEMAPHORE = asyncio.Semaphore(MAX_CONCURRENT_JOBS)
PROJECT_DIR = Path(__file__).resolve().parent
IMAGE_EXTENSIONS = {".png", ".jpg", ".jpeg", ".webp"}
EXCEL_CONTACTS_TEMPLATE = PROJECT_DIR / "assets" / "contact-template.xlsx"


class ProcessingTimeoutError(Exception):
    pass


def find_category(category_id: str):
    return next(
        (item for item in SERVICE_CATEGORIES if item["id"] == category_id),
        None,
    )


def services_for_category(category_id: str) -> list[dict]:
    return [
        service for service in SERVICES if service.get("category") == category_id
    ]


def build_categories_keyboard() -> InlineKeyboardMarkup:
    return InlineKeyboardMarkup(
        [
            [
                InlineKeyboardButton(
                    f"{category.get('emoji', '')} {category['title']}",
                    callback_data=f"cat:{category['id']}",
                )
            ]
            for category in SERVICE_CATEGORIES
        ]
    )


def categories_text() -> str:
    return "\n".join(
        [WELCOME_MESSAGE, "", "اختر الصنف من الأزرار التالية:"]
    )


def build_services_keyboard(category_id: str) -> InlineKeyboardMarkup:
    services = services_for_category(category_id)
    return InlineKeyboardMarkup(
        [
            [
                InlineKeyboardButton(
                    f"{index}. {service.get('emoji', '')} {service['title']}",
                    callback_data=f"svc:{service['id']}",
                )
            ]
            for index, service in enumerate(services, start=1)
        ]
        + [[InlineKeyboardButton("⬅️ رجوع إلى الأصناف", callback_data="back")]]
    )


def services_text(category_id: str) -> str:
    category = find_category(category_id)
    if category is None:
        return categories_text()
    services = services_for_category(category_id)
    lines = [
        f"{category.get('emoji', '')} {category['title']}",
        "",
        "اضغط على الخدمة أو أرسل رقمها:",
    ]
    lines.extend(
        f"{index}. {service.get('emoji', '')} {service['title']}"
        for index, service in enumerate(services, start=1)
    )
    return "\n".join(lines)


def find_service(service_id: str):
    return next((item for item in SERVICES if item["id"] == service_id), None)


def service_by_category_number(category_id: str, number: int):
    services = services_for_category(category_id)
    return services[number - 1] if 1 <= number <= len(services) else None


def clear_job(context: ContextTypes.DEFAULT_TYPE) -> None:
    context.user_data.pop("job", None)


def job_keyboard(done: bool = False) -> InlineKeyboardMarkup:
    rows = []
    if done:
        rows.append([InlineKeyboardButton("✅ انتهيت - تنفيذ", callback_data="job:done")])
    rows.append([InlineKeyboardButton("❌ إلغاء", callback_data="job:cancel")])
    return InlineKeyboardMarkup(rows)


def attachment_meta(message) -> dict | None:
    if message.document:
        return {
            "kind": "document",
            "file_id": message.document.file_id,
            "name": message.document.file_name or "document",
            "size": message.document.file_size or 0,
            "suffix": Path(message.document.file_name or "").suffix.lower(),
        }
    if message.photo:
        photo = message.photo[-1]
        return {
            "kind": "photo",
            "file_id": photo.file_id,
            "name": "photo.jpg",
            "size": photo.file_size or 0,
            "suffix": ".jpg",
        }
    return None


def validate_attachment(meta: dict, allowed: set[str]) -> None:
    if meta["suffix"] not in allowed:
        expected = "، ".join(sorted(item.upper().lstrip(".") for item in allowed))
        raise ServiceInputError(f"نوع الملف غير مدعوم. الصيغ المقبولة: {expected}.")
    if meta["size"] > MAX_UPLOAD_MB * 1024 * 1024:
        raise ServiceInputError(f"حجم الملف أكبر من الحد المسموح ({MAX_UPLOAD_MB} ميغابايت).")


async def download_meta(context: ContextTypes.DEFAULT_TYPE, meta: dict, target: Path) -> None:
    telegram_file = await context.bot.get_file(meta["file_id"])
    await telegram_file.download_to_drive(custom_path=target)


async def run_blocking(function, *args):
    return await asyncio.wait_for(
        asyncio.to_thread(function, *args), timeout=PROCESSING_TIMEOUT_SECONDS
    )


async def run_background_removal(source: Path, output: Path) -> None:
    process = await asyncio.create_subprocess_exec(
        sys.executable,
        "-m",
        "services.background_remover",
        str(source),
        str(output),
        cwd=PROJECT_DIR,
        stdout=asyncio.subprocess.PIPE,
        stderr=asyncio.subprocess.PIPE,
    )
    try:
        _, stderr = await asyncio.wait_for(
            process.communicate(), timeout=PROCESSING_TIMEOUT_SECONDS
        )
    except TimeoutError as error:
        process.kill()
        await process.wait()
        raise ProcessingTimeoutError from error
    if process.returncode != 0:
        details = stderr.decode("utf-8", errors="replace")[-4000:]
        raise RuntimeError(f"Background removal worker failed: {details}")


async def send_result(message, output: Path, caption: str) -> None:
    if output.stat().st_size > MAX_OUTPUT_MB * 1024 * 1024:
        raise ServiceInputError(
            f"حجم النتيجة تجاوز {MAX_OUTPUT_MB} ميغابايت. جرّب ملفات أصغر."
        )
    with output.open("rb") as stream:
        await message.reply_document(document=stream, filename=output.name, caption=caption)


async def finish_success(message, context: ContextTypes.DEFAULT_TYPE) -> None:
    clear_job(context)
    context.user_data.pop("browse_category", None)
    await message.reply_text(
        "اختر صنفًا لخدمة أخرى:",
        reply_markup=build_categories_keyboard(),
    )


async def show_categories(message) -> None:
    await message.reply_text(
        categories_text(),
        reply_markup=build_categories_keyboard(),
    )


async def start(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    clear_job(context)
    context.user_data.pop("browse_category", None)
    await show_categories(update.effective_message)


async def services_command(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    clear_job(context)
    context.user_data.pop("browse_category", None)
    await show_categories(update.effective_message)


async def cancel(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    clear_job(context)
    context.user_data.pop("browse_category", None)
    await update.effective_message.reply_text(
        "تم الإلغاء. اختر صنف الخدمات:",
        reply_markup=build_categories_keyboard(),
    )


async def present_service(message, service, context: ContextTypes.DEFAULT_TYPE) -> None:
    title = f"{service.get('emoji', '')} {service['title']}".strip()
    if service["kind"] == "link":
        await message.reply_text(
            f"{title}\n\n{service['desc']}",
            reply_markup=InlineKeyboardMarkup(
                [
                    [InlineKeyboardButton("🚀 الدخول إلى المنصة", url=service["url"])],
                    [
                        InlineKeyboardButton(
                            "⬅️ رجوع إلى خدمات الصنف",
                            callback_data=f"cat:{service['category']}",
                        )
                    ],
                ]
            ),
            disable_web_page_preview=True,
        )
        return

    service_id = service["id"]
    if service_id in {
        "merge_pdf",
        "images_to_pdf",
        "colornote_to_html",
        "merge_excel",
    }:
        stage = "collecting"
        if service_id == "colornote_to_html":
            instruction = (
                "أرسل ملف Backup واحدًا أو عدة ملفات واحدًا تلو الآخر، "
                "ثم اضغط «انتهيت»."
            )
        elif service_id == "merge_excel":
            instruction = (
                "أرسل ملفات XLSX واحدًا تلو الآخر، ثم اضغط «انتهيت» "
                "لاختيار طريقة الدمج."
            )
        else:
            instruction = (
                "أرسل الملفات واحدًا تلو الآخر حسب الترتيب المطلوب، "
                "ثم اضغط «انتهيت»."
            )
    elif service_id == "create_qr":
        stage = "waiting_text"
        instruction = "أرسل الرابط أو النص الذي تريد تحويله إلى QR Code."
    elif service_id == "text_editor":
        stage = "waiting_text"
        instruction = (
            "أرسل النص العربي الآن كرسالة واحدة "
            f"(حتى {TEXT_EDITOR_MAX_CHARS} حرف).\n"
            "سأعيده مصححاً ومنسقاً دون تغيير المعنى."
        )
    elif service_id == "excel_to_vcf":
        stage = "waiting_file"
        instruction = (
            "أرسل ملف XLSX أو XLS الآن. يجب أن يحتوي الصف الأول على الأعمدة:\n"
            "الاسم الكامل | رقم التواصل | البريد الالكتروني"
        )
    elif service_id in {"clean_excel", "split_excel"}:
        stage = "waiting_file"
        instruction = "أرسل ملف XLSX الآن."
    elif service_id == "pptx_deck":
        stage = "waiting_file"
        instruction = (
            "أرسل ملف PDF أو DOCX الآن.\n"
            "سأستخرج محتواه ثم أعرض عليك مخطط الشرائح للموافقة قبل بناء العرض."
        )
    else:
        stage = "waiting_file"
        instruction = "أرسل الملف الآن."

    context.user_data["job"] = {"service": service_id, "stage": stage, "files": []}
    reply_markup = job_keyboard(done=stage == "collecting")
    if service_id == "excel_to_vcf":
        reply_markup = InlineKeyboardMarkup(
            [
                [
                    InlineKeyboardButton(
                        "📥 تنزيل قالب Excel",
                        callback_data="excel:template",
                    )
                ],
                [InlineKeyboardButton("❌ إلغاء", callback_data="job:cancel")],
            ]
        )
    await message.reply_text(
        f"{title}\n\n{service['desc']}\n\n{instruction}\n"
        "يمكنك إرسال /cancel للإلغاء.",
        reply_markup=reply_markup,
    )


async def on_service_selected(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    query = update.callback_query
    await query.answer()
    service = find_service(query.data.split(":", 1)[1])
    if service is None:
        await query.edit_message_text("عذرًا، هذه الخدمة لم تعد متاحة.")
        return
    clear_job(context)
    context.user_data["browse_category"] = service["category"]
    await present_service(query.message, service, context)


async def on_category_selected(
    update: Update,
    context: ContextTypes.DEFAULT_TYPE,
) -> None:
    query = update.callback_query
    await query.answer()
    category_id = query.data.split(":", 1)[1]
    if find_category(category_id) is None:
        await query.edit_message_text(
            "عذرًا، هذا الصنف لم يعد متاحًا.",
            reply_markup=build_categories_keyboard(),
        )
        return
    clear_job(context)
    context.user_data["browse_category"] = category_id
    await query.edit_message_text(
        services_text(category_id),
        reply_markup=build_services_keyboard(category_id),
    )


async def on_back(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    query = update.callback_query
    await query.answer()
    clear_job(context)
    context.user_data.pop("browse_category", None)
    await query.edit_message_text(
        categories_text(),
        reply_markup=build_categories_keyboard(),
    )


async def _handle_processing_error(message, status, service_id: str, error: Exception) -> None:
    if isinstance(
        error,
        (
            ServiceInputError,
            ExcelContactsError,
            ExcelToolsError,
            ColorNoteConversionError,
            DeckError,
            TextEditError,
        ),
    ):
        await message.reply_text(str(error))
    elif isinstance(error, (TimeoutError, ProcessingTimeoutError)):
        await message.reply_text("انتهت مهلة المعالجة. جرّب ملفًا أصغر أو أعد المحاولة لاحقًا.")
    else:
        logger.exception("Service failed: %s", service_id, exc_info=error)
        await message.reply_text("تعذر تنفيذ الخدمة. تأكد من سلامة الملف ثم حاول مجددًا.")
    if status:
        try:
            await status.delete()
        except Exception:
            pass


async def process_single_file(message, context, job: dict, meta: dict) -> None:
    service_id = job["service"]
    status = await message.reply_text("⏳ جارٍ تنزيل الملف...")
    try:
        with tempfile.TemporaryDirectory(prefix="telegram_service_") as temp_dir:
            root = Path(temp_dir)
            source = root / f"input{meta['suffix']}"
            await download_meta(context, meta, source)
            async with PROCESSING_SEMAPHORE:
                if service_id == "format_word":
                    output = root / "formatted_document.docx"
                    await status.edit_text("⏳ جارٍ تنسيق ملف Word...")
                    await run_blocking(format_word_tables, source, output)
                    caption = "✅ تم تنسيق جداول Word بنجاح."
                elif service_id == "remove_background":
                    output = root / "no_background.png"
                    await status.edit_text("⏳ جارٍ إزالة الخلفية...")
                    await run_background_removal(source, output)
                    caption = "✅ تمت إزالة الخلفية بصيغة PNG."
                elif service_id == "ocr_image":
                    output = root / "extracted_text.txt"
                    await status.edit_text("⏳ جارٍ استخراج النص من الصورة...")
                    await run_blocking(ocr_image, source, output, OCR_LANG)
                    caption = "✅ تم استخراج النص."
                elif service_id == "clean_excel":
                    output = root / "cleaned_workbook.xlsx"
                    await status.edit_text("⏳ جارٍ تنظيف ملف Excel...")
                    summary = await run_blocking(
                        clean_excel_workbook, source, output
                    )
                    caption = (
                        "✅ اكتمل تنظيف ملف Excel.\n"
                        f"الصفوف الفارغة المحذوفة: {summary['blank_rows']} | "
                        f"المكررة المحذوفة: {summary['duplicate_rows']} | "
                        f"الخلايا المنظفة: {summary['trimmed_cells']}"
                    )
                else:
                    raise ServiceInputError("الخدمة المحددة لا تعالج ملفًا مباشرًا.")
            await send_result(message, output, caption)
        await status.delete()
        await finish_success(message, context)
    except Exception as error:
        await _handle_processing_error(message, status, service_id, error)


async def collect_file(message, context, job: dict, meta: dict) -> None:
    service_id = job["service"]
    if service_id == "colornote_to_html":
        if meta.get("kind") != "document":
            raise ServiceInputError("أرسل نسخة ColorNote كملف، وليس كصورة.")
        if meta["size"] > MAX_UPLOAD_MB * 1024 * 1024:
            raise ServiceInputError(
                f"حجم الملف أكبر من الحد المسموح ({MAX_UPLOAD_MB} ميغابايت)."
            )
        file_limit = MAX_COLORNOTE_FILES
    else:
        if service_id == "merge_pdf":
            allowed = {".pdf"}
        elif service_id == "merge_excel":
            allowed = {".xlsx"}
        else:
            allowed = IMAGE_EXTENSIONS
        validate_attachment(meta, allowed)
        file_limit = MAX_MULTI_FILES
    if len(job["files"]) >= file_limit:
        raise ServiceInputError(f"الحد الأقصى {file_limit} ملفات لكل عملية.")
    total_size = sum(item["size"] for item in job["files"]) + meta["size"]
    if total_size > MAX_TOTAL_UPLOAD_MB * 1024 * 1024:
        raise ServiceInputError(
            f"إجمالي الملفات تجاوز {MAX_TOTAL_UPLOAD_MB} ميغابايت."
        )
    job["files"].append(meta)
    await message.reply_text(
        f"تمت إضافة الملف رقم {len(job['files'])}: {meta['name']}\n"
        "أرسل ملفًا آخر أو اضغط «انتهيت».",
        reply_markup=job_keyboard(done=True),
    )


async def finish_collection(message, context: ContextTypes.DEFAULT_TYPE) -> None:
    job = context.user_data.get("job")
    if not job or job.get("stage") != "collecting":
        await message.reply_text("لا توجد ملفات قيد التجميع.")
        return
    minimum = 2 if job["service"] in {"merge_pdf", "merge_excel"} else 1
    if len(job["files"]) < minimum:
        await message.reply_text(f"أرسل {minimum} ملف/ملفات على الأقل قبل التنفيذ.")
        return

    if job["service"] == "colornote_to_html":
        job["stage"] = "waiting_colornote_password"
        await message.reply_text(
            "أرسل كلمة مرور النسخ الاحتياطية الآن. إذا لم تغيّر كلمة المرور "
            "في ColorNote فاستخدم القيمة الافتراضية 0000.",
            reply_markup=InlineKeyboardMarkup(
                [
                    [
                        InlineKeyboardButton(
                            "🔐 استخدام كلمة المرور 0000",
                            callback_data="colornote:default_password",
                        )
                    ],
                    [InlineKeyboardButton("❌ إلغاء", callback_data="job:cancel")],
                ]
            ),
        )
        return

    if job["service"] == "merge_excel":
        job["stage"] = "waiting_excel_merge_mode"
        await message.reply_text(
            "اختر طريقة الدمج:\n"
            "• دمج الصفوف: يتطلب تطابق عناوين الأعمدة في الأوراق النشطة.\n"
            "• جمع الأوراق: يضع جميع أوراق الملفات داخل مصنف واحد.",
            reply_markup=InlineKeyboardMarkup(
                [
                    [
                        InlineKeyboardButton(
                            "📋 دمج الصفوف في ورقة واحدة",
                            callback_data="excelmerge:rows",
                        )
                    ],
                    [
                        InlineKeyboardButton(
                            "📚 جمع الملفات كأوراق",
                            callback_data="excelmerge:sheets",
                        )
                    ],
                    [InlineKeyboardButton("❌ إلغاء", callback_data="job:cancel")],
                ]
            ),
        )
        return

    status = await message.reply_text("⏳ جارٍ تنزيل الملفات وتجهيزها...")
    try:
        with tempfile.TemporaryDirectory(prefix="telegram_multi_") as temp_dir:
            root = Path(temp_dir)
            sources = []
            for index, meta in enumerate(job["files"], start=1):
                target = root / f"input_{index:02d}{meta['suffix']}"
                await download_meta(context, meta, target)
                sources.append(target)
                await status.edit_text(f"⏳ تم تنزيل {index} من {len(job['files'])}...")
            async with PROCESSING_SEMAPHORE:
                if job["service"] == "merge_pdf":
                    output = root / "merged_documents.pdf"
                    await status.edit_text("⏳ جارٍ دمج ملفات PDF...")
                    await run_blocking(merge_pdfs, sources, output)
                    caption = f"✅ تم دمج {len(sources)} ملفات PDF."
                else:
                    output = root / "images.pdf"
                    await status.edit_text("⏳ جارٍ تحويل الصور إلى PDF...")
                    await run_blocking(images_to_pdf, sources, output)
                    caption = f"✅ تم تحويل {len(sources)} صور إلى PDF."
            await send_result(message, output, caption)
        await status.delete()
        await finish_success(message, context)
    except Exception as error:
        await _handle_processing_error(message, status, job["service"], error)


async def process_split(message, context, job: dict, page_spec: str) -> None:
    status = await message.reply_text("⏳ جارٍ استخراج الصفحات...")
    try:
        with tempfile.TemporaryDirectory(prefix="telegram_split_") as temp_dir:
            root = Path(temp_dir)
            source = root / "input.pdf"
            output = root / "selected_pages.pdf"
            await download_meta(context, job["source"], source)
            async with PROCESSING_SEMAPHORE:
                page_count = await run_blocking(
                    extract_pdf_pages, source, output, page_spec
                )
            await send_result(message, output, f"✅ تم استخراج {page_count} صفحة/صفحات.")
        await status.delete()
        await finish_success(message, context)
    except Exception as error:
        await _handle_processing_error(message, status, job["service"], error)


async def process_pdf_page_management(
    message,
    context,
    operation: str,
    page_spec: str,
    angle: int = 90,
) -> None:
    job = context.user_data.get("job")
    if not job or job.get("service") != "manage_pdf_pages" or "source" not in job:
        await message.reply_text("ابدأ خدمة إدارة صفحات PDF من القائمة أولًا.")
        return
    status = await message.reply_text("⏳ جارٍ معالجة صفحات PDF...")
    try:
        with tempfile.TemporaryDirectory(prefix="telegram_pdf_pages_") as temp_dir:
            root = Path(temp_dir)
            source = root / "input.pdf"
            output = root / "managed_pages.pdf"
            await download_meta(context, job["source"], source)
            async with PROCESSING_SEMAPHORE:
                summary = await run_blocking(
                    manage_pdf_pages,
                    source,
                    output,
                    operation,
                    page_spec,
                    angle,
                )
            operation_text = {
                "rotate": "تدوير الصفحات",
                "delete": "حذف الصفحات",
                "reorder": "إعادة ترتيب الصفحات",
            }[operation]
            await send_result(
                message,
                output,
                f"✅ اكتملت عملية {operation_text}. "
                f"عدد صفحات النتيجة: {summary['pages']}.",
            )
        await status.delete()
        await finish_success(message, context)
    except Exception as error:
        await _handle_processing_error(message, status, "manage_pdf_pages", error)


async def process_pdf_password(message, context, password: str) -> None:
    job = context.user_data.get("job")
    if not job or job.get("service") != "pdf_password" or "source" not in job:
        await message.reply_text("ابدأ خدمة حماية PDF من القائمة أولًا.")
        return
    action = job.get("pdf_password_action")
    if action not in {"protect", "unprotect"}:
        await message.reply_text("اختر حماية الملف أو فك الحماية أولًا.")
        return
    status = await message.reply_text("⏳ جارٍ معالجة حماية ملف PDF...")
    try:
        with tempfile.TemporaryDirectory(prefix="telegram_pdf_password_") as temp_dir:
            root = Path(temp_dir)
            source = root / "input.pdf"
            output = root / (
                "protected_document.pdf"
                if action == "protect"
                else "unprotected_document.pdf"
            )
            await download_meta(context, job["source"], source)
            async with PROCESSING_SEMAPHORE:
                if action == "protect":
                    pages = await run_blocking(
                        protect_pdf, source, output, password
                    )
                    caption = f"✅ تم تشفير ملف PDF وعدد صفحاته {pages}."
                else:
                    pages = await run_blocking(
                        unprotect_pdf, source, output, password
                    )
                    caption = f"✅ تم فك حماية ملف PDF وعدد صفحاته {pages}."
            await send_result(message, output, caption)
        await status.delete()
        await finish_success(message, context)
    except Exception as error:
        await _handle_processing_error(message, status, "pdf_password", error)


async def process_excel_split(message, context, column_selector: str) -> None:
    job = context.user_data.get("job")
    if not job or job.get("service") != "split_excel" or "source" not in job:
        await message.reply_text("ابدأ خدمة تقسيم Excel من القائمة أولًا.")
        return
    status = await message.reply_text("⏳ جارٍ تقسيم ملف Excel...")
    try:
        with tempfile.TemporaryDirectory(prefix="telegram_excel_split_") as temp_dir:
            root = Path(temp_dir)
            source = root / "input.xlsx"
            output = root / "split_workbook.zip"
            await download_meta(context, job["source"], source)
            async with PROCESSING_SEMAPHORE:
                summary = await run_blocking(
                    split_excel_workbook, source, output, column_selector
                )
            caption = (
                "✅ اكتمل تقسيم ملف Excel.\n"
                f"الملفات الناتجة: {summary['groups']} | "
                f"صفوف البيانات: {summary['rows']} | "
                f"صفوف بلا قيمة: {summary['skipped']}"
            )
            await send_result(message, output, caption)
        await status.delete()
        await finish_success(message, context)
    except Exception as error:
        await _handle_processing_error(message, status, "split_excel", error)


async def process_excel_merge(message, context, mode: str) -> None:
    job = context.user_data.get("job")
    if (
        not job
        or job.get("service") != "merge_excel"
        or job.get("stage") != "waiting_excel_merge_mode"
        or len(job.get("files", [])) < 2
    ):
        await message.reply_text("ابدأ خدمة دمج Excel وأرسل ملفين على الأقل.")
        return
    if mode not in {"rows", "sheets"}:
        await message.reply_text("طريقة الدمج غير معروفة.")
        return

    status = await message.reply_text("⏳ جارٍ تنزيل ملفات Excel...")
    try:
        with tempfile.TemporaryDirectory(prefix="telegram_excel_merge_") as temp_dir:
            root = Path(temp_dir)
            sources = []
            for index, meta in enumerate(job["files"], start=1):
                target = root / f"input_{index:02d}.xlsx"
                await download_meta(context, meta, target)
                sources.append((meta["name"], target))
                await status.edit_text(
                    f"⏳ تم تنزيل {index} من {len(job['files'])}..."
                )
            output = root / "merged_workbooks.xlsx"
            async with PROCESSING_SEMAPHORE:
                summary = await run_blocking(
                    merge_excel_workbooks, sources, output, mode
                )
            if mode == "rows":
                details = f"الصفوف المدمجة: {summary['rows']}"
            else:
                details = f"الأوراق المدمجة: {summary['sheets']}"
            await send_result(
                message,
                output,
                f"✅ تم دمج {summary['files']} ملفات Excel. {details}.",
            )
        await status.delete()
        await finish_success(message, context)
    except Exception as error:
        await _handle_processing_error(message, status, "merge_excel", error)


async def process_image_option(message, context, operation: str) -> None:
    job = context.user_data.get("job")
    if not job or job.get("service") != "optimize_image" or "source" not in job:
        await message.reply_text("ابدأ خدمة ضغط الصور من القائمة أولًا.")
        return
    extensions = {"jpeg": ".jpg", "half": ".jpg", "png": ".png", "webp": ".webp"}
    if operation not in extensions:
        await message.reply_text("خيار غير معروف.")
        return
    status = await message.reply_text("⏳ جارٍ معالجة الصورة...")
    try:
        with tempfile.TemporaryDirectory(prefix="telegram_image_") as temp_dir:
            root = Path(temp_dir)
            source = root / f"input{job['source']['suffix']}"
            output = root / f"processed_image{extensions[operation]}"
            await download_meta(context, job["source"], source)
            async with PROCESSING_SEMAPHORE:
                await run_blocking(optimize_image, source, output, operation)
            await send_result(message, output, "✅ تمت معالجة الصورة.")
        await status.delete()
        await finish_success(message, context)
    except Exception as error:
        await _handle_processing_error(message, status, job["service"], error)


async def process_watermark(message, context, text: str | None = None, logo_meta: dict | None = None) -> None:
    job = context.user_data.get("job")
    if not job or job.get("service") != "watermark_pdf" or "source" not in job:
        await message.reply_text("ابدأ خدمة العلامة المائية من القائمة أولًا.")
        return
    status = await message.reply_text("⏳ جارٍ إضافة العلامة المائية...")
    try:
        with tempfile.TemporaryDirectory(prefix="telegram_watermark_") as temp_dir:
            root = Path(temp_dir)
            source = root / "input.pdf"
            output = root / "watermarked_document.pdf"
            await download_meta(context, job["source"], source)
            logo = None
            if logo_meta:
                logo = root / f"logo{logo_meta['suffix']}"
                await download_meta(context, logo_meta, logo)
            async with PROCESSING_SEMAPHORE:
                await run_blocking(watermark_pdf, source, output, text, logo)
            await send_result(message, output, "✅ تمت إضافة العلامة المائية إلى جميع الصفحات.")
        await status.delete()
        await finish_success(message, context)
    except Exception as error:
        await _handle_processing_error(message, status, job["service"], error)


async def process_qr(message, context, data: str) -> None:
    status = await message.reply_text("⏳ جارٍ إنشاء QR Code...")
    try:
        with tempfile.TemporaryDirectory(prefix="telegram_qr_") as temp_dir:
            output = Path(temp_dir) / "qr_code.png"
            await run_blocking(create_qr_code, data, output)
            await send_result(message, output, "✅ تم إنشاء QR Code.")
        await status.delete()
        await finish_success(message, context)
    except Exception as error:
        await _handle_processing_error(message, status, "create_qr", error)


async def process_text_editor(message, context, text: str) -> None:
    status = await message.reply_text("⏳ جارٍ تدقيق النص...")
    try:
        async with PROCESSING_SEMAPHORE:
            result = await run_blocking(polish_text, text)

        footer = ""
        if result.notes:
            footer = "\n\n📝 أهم التصحيحات:\n" + "\n".join(
                f"• {note}" for note in result.notes
            )

        if len(result.text) <= TEXT_EDITOR_MAX_MESSAGE_CHARS:
            await message.reply_text(f"✅ النص بعد التدقيق:\n\n{result.text}{footer}")
        else:
            # Too long for one Telegram message, so it goes out as a file.
            with tempfile.TemporaryDirectory(prefix="telegram_text_editor_") as temp_dir:
                output = Path(temp_dir) / "corrected_text.txt"
                output.write_text(result.text, encoding="utf-8")
                await send_result(
                    message,
                    output,
                    f"✅ تم تدقيق النص.{footer}"[:1024],
                )
        await status.delete()
        await finish_success(message, context)
    except Exception as error:
        await _handle_processing_error(message, status, "text_editor", error)


async def process_excel_contacts(message, context, country_code: str) -> None:
    job = context.user_data.get("job")
    if (
        not job
        or job.get("service") != "excel_to_vcf"
        or job.get("stage") != "waiting_country_code"
        or "source" not in job
    ):
        await message.reply_text("ابدأ خدمة تحويل Excel إلى VCF من القائمة أولًا.")
        return

    status = await message.reply_text("⏳ جارٍ تنزيل ملف Excel...")
    try:
        with tempfile.TemporaryDirectory(prefix="telegram_excel_contacts_") as temp_dir:
            root = Path(temp_dir)
            meta = job["source"]
            source = root / f"input{meta['suffix']}"
            timestamp = datetime.now(UTC).strftime("%Y-%m-%dT%H-%M-%S")
            output = root / f"contacts-output-{timestamp}.zip"
            vcf_output = root / "contacts.vcf"
            await download_meta(context, meta, source)
            async with PROCESSING_SEMAPHORE:
                await status.edit_text("⏳ جارٍ تنظيف جهات الاتصال وإنشاء ملف VCF...")
                summary = await run_blocking(
                    convert_excel_contacts,
                    source,
                    output,
                    country_code,
                    meta["name"],
                    vcf_output,
                )
            caption = (
                "✅ اكتمل تحويل جهات الاتصال.\n"
                f"الإجمالي: {summary['totalRows']} | "
                f"المقبولة: {summary['validCount']} | "
                f"المكررة: {summary['duplicateCount']} | "
                f"تحتاج مراجعة: {summary['invalidCount']}\n"
                "تحتوي الحزمة على contacts.vcf وتقارير Excel والتقرير الملخص."
            )
            await send_result(
                message,
                vcf_output,
                "✅ ملف VCF المباشر جاهز للاستيراد إلى جهات الاتصال.",
            )
            await send_result(message, output, caption)
        await status.delete()
        await finish_success(message, context)
    except Exception as error:
        await _handle_processing_error(message, status, "excel_to_vcf", error)


async def process_colornote(message, context, password: str) -> None:
    job = context.user_data.get("job")
    if (
        not job
        or job.get("service") != "colornote_to_html"
        or job.get("stage") != "waiting_colornote_password"
        or not job.get("files")
    ):
        await message.reply_text("ابدأ خدمة تحويل ColorNote من القائمة أولًا.")
        return

    status = await message.reply_text("⏳ جارٍ تنزيل نسخ ColorNote...")
    try:
        with tempfile.TemporaryDirectory(prefix="telegram_colornote_") as temp_dir:
            root = Path(temp_dir)
            sources = []
            for index, meta in enumerate(job["files"], start=1):
                target = root / f"backup_{index:02d}{meta['suffix']}"
                await download_meta(context, meta, target)
                sources.append((meta["name"], target))
                await status.edit_text(
                    f"⏳ تم تنزيل {index} من {len(job['files'])}..."
                )

            output = root / "ColorNote_notes.html"
            async with PROCESSING_SEMAPHORE:
                await status.edit_text(
                    "⏳ جارٍ فك النسخ ودمج الملاحظات وإنشاء HTML..."
                )
                summary = await run_blocking(
                    convert_colornote_backups,
                    sources,
                    output,
                    password,
                )
            caption = (
                "✅ اكتمل تحويل نسخ ColorNote إلى HTML.\n"
                f"الملفات: {summary['files']} | "
                f"الملاحظات: {summary['total']} | "
                f"المكررات المحذوفة: {summary['duplicates_removed']} | "
                f"الملاحظات المتغيرة: {summary['changed_notes']}\n"
                "الملف مستقل ويعمل دون إنترنت، ويتضمن البحث والفلاتر."
            )
            await send_result(message, output, caption)
        await status.delete()
        await finish_success(message, context)
    except Exception as error:
        await _handle_processing_error(message, status, "colornote_to_html", error)


def deck_slides_keyboard() -> InlineKeyboardMarkup:
    row = [
        InlineKeyboardButton(f"{count} شريحة", callback_data=f"deck:slides:{count}")
        for count in SLIDE_COUNT_CHOICES
    ]
    return InlineKeyboardMarkup([row, [InlineKeyboardButton("❌ إلغاء", callback_data="job:cancel")]])


def deck_review_keyboard() -> InlineKeyboardMarkup:
    return InlineKeyboardMarkup(
        [
            [InlineKeyboardButton("✅ موافق - ابنِ العرض", callback_data="deck:approve")],
            [InlineKeyboardButton("🔄 أعد إنشاء المخطط", callback_data="deck:retry")],
            [InlineKeyboardButton("❌ إلغاء", callback_data="job:cancel")],
        ]
    )


async def start_deck_outline(message, context, organisation: str = "") -> None:
    """Download the document, extract it, and ask the model for an outline."""
    job = context.user_data.get("job")
    if not job or job.get("service") != "pptx_deck":
        await message.reply_text("ابدأ الخدمة من جديد عبر /services.")
        return

    organisation = (organisation or "").strip()
    if organisation in {"-", "/skip", "تخطي"}:
        organisation = ""
    if len(organisation) > 80:
        await message.reply_text("اسم الجهة طويل جدًا. أرسل اسمًا أقصر.")
        return
    job["organisation"] = organisation
    job["stage"] = "building_outline"

    meta = job.get("source")
    slide_count = min(int(job.get("slide_count", DEFAULT_SLIDE_COUNT)), MAX_DECK_SLIDES)
    status = await message.reply_text("⏳ جارٍ استخراج النص من المستند...")
    try:
        with tempfile.TemporaryDirectory(prefix="telegram_deck_") as temp_dir:
            source = Path(temp_dir) / f"input{meta['suffix']}"
            await download_meta(context, meta, source)
            async with PROCESSING_SEMAPHORE:
                await status.edit_text("⏳ جارٍ إعداد مخطط الشرائح...")
                outline = await run_blocking(
                    build_deck_outline, source, slide_count, organisation
                )
        job["outline"] = outline_to_dict(outline)
        job["stage"] = "reviewing_deck_outline"
        await status.delete()
        await message.reply_text(
            outline_preview(outline) + "\n\nراجع المخطط ثم اختر:",
            reply_markup=deck_review_keyboard(),
        )
    except Exception as error:
        await _handle_processing_error(message, status, "pptx_deck", error)
        clear_job(context)


async def process_deck_render(message, context) -> None:
    """Build the .pptx from the outline the user approved."""
    job = context.user_data.get("job")
    if not job or not job.get("outline"):
        await message.reply_text("انتهت صلاحية المخطط. ابدأ الخدمة من جديد عبر /services.")
        clear_job(context)
        return

    organisation = job.get("organisation", "")
    status = await message.reply_text("⏳ جارٍ بناء العرض التقديمي...")
    try:
        outline = outline_from_dict(job["outline"])
        with tempfile.TemporaryDirectory(prefix="telegram_deck_") as temp_dir:
            output = Path(temp_dir) / "presentation.pptx"
            async with PROCESSING_SEMAPHORE:
                await run_blocking(render_outline, outline, output, organisation)
            caption = f"✅ تم إنشاء العرض التقديمي ({len(outline.slides)} شريحة)."
            await send_result(message, output, caption)
        await status.delete()
        await finish_success(message, context)
    except Exception as error:
        await _handle_processing_error(message, status, "pptx_deck", error)


async def on_deck_action(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    query = update.callback_query
    await query.answer()
    message = query.message
    job = context.user_data.get("job")
    if not job or job.get("service") != "pptx_deck":
        await message.reply_text("ابدأ الخدمة من جديد عبر /services.")
        return

    action = query.data.split(":", 1)[1]

    if action.startswith("slides:"):
        if job.get("stage") != "choosing_deck_slides":
            await message.reply_text("أكمل الخطوة المطلوبة أولًا أو أرسل /cancel.")
            return
        job["slide_count"] = min(int(action.split(":")[1]), MAX_DECK_SLIDES)
        job["stage"] = "waiting_deck_org"
        await message.reply_text(
            "أرسل اسم الجهة كما تريده على الغلاف، أو اضغط «بدون اسم».",
            reply_markup=InlineKeyboardMarkup(
                [
                    [InlineKeyboardButton("بدون اسم", callback_data="deck:noorg")],
                    [InlineKeyboardButton("❌ إلغاء", callback_data="job:cancel")],
                ]
            ),
        )
        return

    if action == "noorg":
        if job.get("stage") != "waiting_deck_org":
            await message.reply_text("أكمل الخطوة المطلوبة أولًا أو أرسل /cancel.")
            return
        await start_deck_outline(message, context, organisation="")
        return

    if action == "retry":
        if job.get("stage") != "reviewing_deck_outline":
            await message.reply_text("أكمل الخطوة المطلوبة أولًا أو أرسل /cancel.")
            return
        await start_deck_outline(message, context, job.get("organisation", ""))
        return

    if action == "approve":
        if job.get("stage") != "reviewing_deck_outline":
            await message.reply_text("أكمل الخطوة المطلوبة أولًا أو أرسل /cancel.")
            return
        job["stage"] = "rendering_deck"
        await process_deck_render(message, context)
        return


async def on_file(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    message = update.effective_message
    job = context.user_data.get("job")
    if not job:
        await message.reply_text("اختر خدمة أولًا من /services ثم أرسل الملف.")
        return
    meta = attachment_meta(message)
    if not meta:
        await message.reply_text("لم أتمكن من قراءة الملف المرسل.")
        return
    try:
        service_id = job["service"]
        stage = job["stage"]
        if stage == "collecting":
            await collect_file(message, context, job, meta)
            return
        if stage == "waiting_logo":
            validate_attachment(meta, IMAGE_EXTENSIONS)
            await process_watermark(message, context, logo_meta=meta)
            return
        if stage != "waiting_file":
            await message.reply_text("أكمل الخطوة المطلوبة أولًا أو أرسل /cancel.")
            return

        allowed_by_service = {
            "format_word": {".docx"},
            "remove_background": IMAGE_EXTENSIONS,
            "split_pdf": {".pdf"},
            "manage_pdf_pages": {".pdf"},
            "pdf_password": {".pdf"},
            "optimize_image": IMAGE_EXTENSIONS,
            "watermark_pdf": {".pdf"},
            "ocr_image": IMAGE_EXTENSIONS,
            "excel_to_vcf": {".xlsx", ".xls"},
            "clean_excel": {".xlsx"},
            "split_excel": {".xlsx"},
            "pptx_deck": {".pdf", ".docx"},
        }
        validate_attachment(meta, allowed_by_service.get(service_id, set()))
        if service_id == "pptx_deck":
            if not gemini_key_present():
                raise ServiceInputError(
                    "خدمة العروض التقديمية غير مهيأة على الخادم. راجع مسؤول البوت."
                )
            job.update({"source": meta, "stage": "choosing_deck_slides"})
            await message.reply_text(
                "كم شريحة تريد في العرض؟",
                reply_markup=deck_slides_keyboard(),
            )
        elif service_id == "split_pdf":
            job.update({"source": meta, "stage": "waiting_pages"})
            await message.reply_text("أرسل الصفحات المطلوبة، مثال: 1-3,5,8", reply_markup=job_keyboard())
        elif service_id == "manage_pdf_pages":
            job.update({"source": meta, "stage": "choosing_pdf_page_operation"})
            await message.reply_text(
                "اختر عملية إدارة الصفحات:",
                reply_markup=InlineKeyboardMarkup(
                    [
                        [
                            InlineKeyboardButton(
                                "🔄 تدوير صفحات",
                                callback_data="pdfpages:rotate",
                            )
                        ],
                        [
                            InlineKeyboardButton(
                                "🗑️ حذف صفحات",
                                callback_data="pdfpages:delete",
                            )
                        ],
                        [
                            InlineKeyboardButton(
                                "🔢 إعادة ترتيب الصفحات",
                                callback_data="pdfpages:reorder",
                            )
                        ],
                        [InlineKeyboardButton("❌ إلغاء", callback_data="job:cancel")],
                    ]
                ),
            )
        elif service_id == "pdf_password":
            job.update({"source": meta, "stage": "choosing_pdf_password_action"})
            await message.reply_text(
                "اختر العملية. فك الحماية يتطلب كلمة المرور الصحيحة:",
                reply_markup=InlineKeyboardMarkup(
                    [
                        [
                            InlineKeyboardButton(
                                "🔒 حماية الملف",
                                callback_data="pdfsecurity:protect",
                            )
                        ],
                        [
                            InlineKeyboardButton(
                                "🔓 فك الحماية",
                                callback_data="pdfsecurity:unprotect",
                            )
                        ],
                        [InlineKeyboardButton("❌ إلغاء", callback_data="job:cancel")],
                    ]
                ),
            )
        elif service_id == "optimize_image":
            job.update({"source": meta, "stage": "choosing_image_operation"})
            await message.reply_text(
                "اختر العملية:",
                reply_markup=InlineKeyboardMarkup(
                    [
                        [InlineKeyboardButton("ضغط وتحويل إلى JPG", callback_data="imgopt:jpeg")],
                        [InlineKeyboardButton("تحويل إلى PNG", callback_data="imgopt:png")],
                        [InlineKeyboardButton("تحويل إلى WEBP", callback_data="imgopt:webp")],
                        [InlineKeyboardButton("تصغير 50٪", callback_data="imgopt:half")],
                        [InlineKeyboardButton("❌ إلغاء", callback_data="job:cancel")],
                    ]
                ),
            )
        elif service_id == "watermark_pdf":
            job.update({"source": meta, "stage": "choosing_watermark"})
            await message.reply_text(
                "اختر نوع العلامة المائية:",
                reply_markup=InlineKeyboardMarkup(
                    [
                        [InlineKeyboardButton("📝 نص", callback_data="watermark:text")],
                        [InlineKeyboardButton("🖼️ صورة شعار", callback_data="watermark:logo")],
                        [InlineKeyboardButton("❌ إلغاء", callback_data="job:cancel")],
                    ]
                ),
            )
        elif service_id == "excel_to_vcf":
            if meta["size"] > EXCEL_CONTACTS_MAX_FILE_SIZE:
                raise ServiceInputError(
                    "حجم ملف Excel أكبر من الحد المسموح وهو 10 ميغابايت."
                )
            job.update({"source": meta, "stage": "waiting_country_code"})
            await message.reply_text(
                "أرسل رمز الدولة للأرقام المحلية من 1 إلى 4 أرقام، مثال: 963 لسوريا.\n"
                "الأرقام المكتوبة بصيغة دولية لن تتغير.",
                reply_markup=InlineKeyboardMarkup(
                    [
                        [
                            InlineKeyboardButton(
                                "🇸🇾 استخدام رمز سوريا 963",
                                callback_data="excel:country:963",
                            )
                        ],
                        [InlineKeyboardButton("❌ إلغاء", callback_data="job:cancel")],
                    ]
                ),
            )
        elif service_id == "split_excel":
            job.update({"source": meta, "stage": "waiting_excel_split_column"})
            await message.reply_text(
                "أرسل اسم العمود الذي تريد التقسيم حسبه كما يظهر في الملف، "
                "أو أرسل حرف العمود مثل B. يتم تقسيم الورقة النشطة.",
                reply_markup=job_keyboard(),
            )
        else:
            await process_single_file(message, context, job, meta)
    except Exception as error:
        await _handle_processing_error(message, None, job.get("service", "unknown"), error)


async def on_text(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    message = update.effective_message
    raw_text = message.text
    text = raw_text.strip()
    job = context.user_data.get("job")
    if not job:
        category_id = context.user_data.get("browse_category")
        normalized = text.translate(str.maketrans("٠١٢٣٤٥٦٧٨٩", "0123456789"))
        if category_id and normalized.isdigit():
            service = service_by_category_number(category_id, int(normalized))
            if service:
                await present_service(message, service, context)
                return
        if category_id:
            await message.reply_text(
                "أرسل رقم خدمة صحيحًا من الصنف الحالي، أو عد إلى الأصناف.",
                reply_markup=build_services_keyboard(category_id),
            )
        else:
            await message.reply_text(
                "اختر صنف الخدمات أولًا.",
                reply_markup=build_categories_keyboard(),
            )
        return

    if job["service"] == "create_qr" and job["stage"] == "waiting_text":
        await process_qr(message, context, text)
    elif job["service"] == "text_editor" and job["stage"] == "waiting_text":
        await process_text_editor(message, context, raw_text)
    elif job["service"] == "split_pdf" and job["stage"] == "waiting_pages":
        await process_split(message, context, job, text)
    elif job["service"] == "watermark_pdf" and job["stage"] == "waiting_watermark_text":
        await process_watermark(message, context, text=text)
    elif (
        job["service"] == "manage_pdf_pages"
        and job["stage"] == "waiting_pdf_rotate_pages"
    ):
        await process_pdf_page_management(
            message,
            context,
            "rotate",
            text,
            int(job.get("pdf_rotation_angle", 90)),
        )
    elif (
        job["service"] == "manage_pdf_pages"
        and job["stage"] == "waiting_pdf_delete_pages"
    ):
        await process_pdf_page_management(message, context, "delete", text)
    elif (
        job["service"] == "manage_pdf_pages"
        and job["stage"] == "waiting_pdf_reorder"
    ):
        await process_pdf_page_management(message, context, "reorder", text)
    elif job["service"] == "pdf_password" and job["stage"] == "waiting_pdf_password":
        await process_pdf_password(message, context, raw_text)
    elif job["service"] == "excel_to_vcf" and job["stage"] == "waiting_country_code":
        await process_excel_contacts(message, context, text)
    elif (
        job["service"] == "split_excel"
        and job["stage"] == "waiting_excel_split_column"
    ):
        await process_excel_split(message, context, text)
    elif (
        job["service"] == "colornote_to_html"
        and job["stage"] == "waiting_colornote_password"
    ):
        await process_colornote(message, context, raw_text)
    elif job["service"] == "pptx_deck" and job["stage"] == "waiting_deck_org":
        await start_deck_outline(message, context, organisation=text)
    else:
        await message.reply_text("أكمل الخطوة المطلوبة بالأزرار أو أرسل /cancel.")


async def on_job_action(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    query = update.callback_query
    await query.answer()
    action = query.data.split(":", 1)[1]
    if action == "cancel":
        clear_job(context)
        context.user_data.pop("browse_category", None)
        await query.message.reply_text(
            "تم الإلغاء. اختر صنف الخدمات:",
            reply_markup=build_categories_keyboard(),
        )
    elif action == "done":
        await finish_collection(query.message, context)


async def done_command(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    await finish_collection(update.effective_message, context)


async def on_image_option(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    query = update.callback_query
    await query.answer()
    await process_image_option(query.message, context, query.data.split(":", 1)[1])


async def on_watermark_option(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    query = update.callback_query
    await query.answer()
    job = context.user_data.get("job")
    if not job or job.get("service") != "watermark_pdf":
        await query.message.reply_text("ابدأ خدمة العلامة المائية من القائمة أولًا.")
        return
    option = query.data.split(":", 1)[1]
    if option == "text":
        job["stage"] = "waiting_watermark_text"
        await query.message.reply_text("أرسل نص العلامة المائية الآن.", reply_markup=job_keyboard())
    else:
        job["stage"] = "waiting_logo"
        await query.message.reply_text("أرسل صورة الشعار PNG أو JPG أو WEBP.", reply_markup=job_keyboard())


async def on_pdf_pages_action(
    update: Update,
    context: ContextTypes.DEFAULT_TYPE,
) -> None:
    query = update.callback_query
    await query.answer()
    job = context.user_data.get("job")
    if not job or job.get("service") != "manage_pdf_pages" or "source" not in job:
        await query.message.reply_text("ابدأ خدمة إدارة صفحات PDF من القائمة أولًا.")
        return
    parts = query.data.split(":")
    action = parts[1]
    if action == "rotate":
        job["stage"] = "choosing_pdf_rotation_angle"
        await query.message.reply_text(
            "اختر زاوية التدوير:",
            reply_markup=InlineKeyboardMarkup(
                [
                    [
                        InlineKeyboardButton(
                            "↻ 90° مع عقارب الساعة",
                            callback_data="pdfpages:angle:90",
                        )
                    ],
                    [
                        InlineKeyboardButton(
                            "↺ 90° عكس عقارب الساعة",
                            callback_data="pdfpages:angle:-90",
                        )
                    ],
                    [
                        InlineKeyboardButton(
                            "🔃 180°",
                            callback_data="pdfpages:angle:180",
                        )
                    ],
                    [InlineKeyboardButton("❌ إلغاء", callback_data="job:cancel")],
                ]
            ),
        )
    elif action == "angle" and len(parts) == 3:
        try:
            angle = int(parts[2])
        except ValueError:
            await query.message.reply_text("زاوية تدوير غير صحيحة.")
            return
        if angle not in {-90, 90, 180}:
            await query.message.reply_text("زاوية تدوير غير صحيحة.")
            return
        job.update(
            {
                "stage": "waiting_pdf_rotate_pages",
                "pdf_rotation_angle": angle,
            }
        )
        await query.message.reply_text(
            "أرسل أرقام الصفحات مثل 1-3,5 أو أرسل «الكل».",
            reply_markup=job_keyboard(),
        )
    elif action == "delete":
        job["stage"] = "waiting_pdf_delete_pages"
        await query.message.reply_text(
            "أرسل الصفحات التي تريد حذفها، مثال: 2,5-7",
            reply_markup=job_keyboard(),
        )
    elif action == "reorder":
        job["stage"] = "waiting_pdf_reorder"
        await query.message.reply_text(
            "أرسل ترتيب جميع الصفحات مرة واحدة، مثال لملف من 5 صفحات: "
            "3,1,2,5-4",
            reply_markup=job_keyboard(),
        )
    else:
        await query.message.reply_text("خيار إدارة الصفحات غير معروف.")


async def on_pdf_security_action(
    update: Update,
    context: ContextTypes.DEFAULT_TYPE,
) -> None:
    query = update.callback_query
    await query.answer()
    job = context.user_data.get("job")
    if not job or job.get("service") != "pdf_password" or "source" not in job:
        await query.message.reply_text("ابدأ خدمة حماية PDF من القائمة أولًا.")
        return
    action = query.data.split(":", 1)[1]
    if action not in {"protect", "unprotect"}:
        await query.message.reply_text("خيار حماية PDF غير معروف.")
        return
    job.update(
        {
            "stage": "waiting_pdf_password",
            "pdf_password_action": action,
        }
    )
    if action == "protect":
        instruction = "أرسل كلمة مرور جديدة من 4 إلى 128 محرفًا."
    else:
        instruction = "أرسل كلمة المرور الحالية للملف."
    await query.message.reply_text(
        f"{instruction}\nلن تُحفظ كلمة المرور بعد انتهاء العملية.",
        reply_markup=job_keyboard(),
    )


async def on_excel_merge_action(
    update: Update,
    context: ContextTypes.DEFAULT_TYPE,
) -> None:
    query = update.callback_query
    await query.answer()
    await process_excel_merge(
        query.message,
        context,
        query.data.split(":", 1)[1],
    )


async def on_excel_action(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    query = update.callback_query
    await query.answer()
    parts = query.data.split(":")
    if parts[1] == "template":
        if not EXCEL_CONTACTS_TEMPLATE.is_file():
            await query.message.reply_text("قالب Excel غير متاح حاليًا.")
            return
        with EXCEL_CONTACTS_TEMPLATE.open("rb") as stream:
            await query.message.reply_document(
                document=stream,
                filename="contact-template.xlsx",
                caption="قالب جاهز بالأعمدة المطلوبة لخدمة Excel إلى VCF.",
            )
        return
    if len(parts) == 3 and parts[1] == "country":
        await process_excel_contacts(query.message, context, parts[2])


async def on_colornote_action(
    update: Update,
    context: ContextTypes.DEFAULT_TYPE,
) -> None:
    query = update.callback_query
    await query.answer()
    if query.data == "colornote:default_password":
        await process_colornote(query.message, context, "0000")


def main() -> None:
    if not BOT_TOKEN:
        raise SystemExit("لم يتم ضبط BOT_TOKEN في متغيرات البيئة.")
    app = Application.builder().token(BOT_TOKEN).build()
    app.add_handler(CommandHandler("start", start))
    app.add_handler(CommandHandler("services", services_command))
    app.add_handler(CommandHandler("cancel", cancel))
    app.add_handler(CommandHandler("done", done_command))
    app.add_handler(CallbackQueryHandler(on_category_selected, pattern=r"^cat:"))
    app.add_handler(CallbackQueryHandler(on_service_selected, pattern=r"^svc:"))
    app.add_handler(CallbackQueryHandler(on_back, pattern=r"^back$"))
    app.add_handler(CallbackQueryHandler(on_job_action, pattern=r"^job:"))
    app.add_handler(CallbackQueryHandler(on_image_option, pattern=r"^imgopt:"))
    app.add_handler(CallbackQueryHandler(on_watermark_option, pattern=r"^watermark:"))
    app.add_handler(CallbackQueryHandler(on_pdf_pages_action, pattern=r"^pdfpages:"))
    app.add_handler(
        CallbackQueryHandler(on_pdf_security_action, pattern=r"^pdfsecurity:")
    )
    app.add_handler(
        CallbackQueryHandler(on_excel_merge_action, pattern=r"^excelmerge:")
    )
    app.add_handler(CallbackQueryHandler(on_excel_action, pattern=r"^excel:"))
    app.add_handler(CallbackQueryHandler(on_colornote_action, pattern=r"^colornote:"))
    app.add_handler(CallbackQueryHandler(on_deck_action, pattern=r"^deck:"))
    app.add_handler(MessageHandler(filters.Document.ALL | filters.PHOTO, on_file))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, on_text))
    logger.info("البوت يعمل الآن...")
    app.run_polling(allowed_updates=Update.ALL_TYPES)


if __name__ == "__main__":
    main()
