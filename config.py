import os

BOT_TOKEN = os.environ.get("BOT_TOKEN", "").strip()
MAX_UPLOAD_MB = max(1, int(os.environ.get("MAX_UPLOAD_MB", "20")))
MAX_OUTPUT_MB = max(1, int(os.environ.get("MAX_OUTPUT_MB", "45")))
MAX_MULTI_FILES = max(2, int(os.environ.get("MAX_MULTI_FILES", "10")))
MAX_TOTAL_UPLOAD_MB = max(1, int(os.environ.get("MAX_TOTAL_UPLOAD_MB", "100")))
MAX_IMAGE_PIXELS = max(1_000_000, int(os.environ.get("MAX_IMAGE_PIXELS", "40000000")))
MAX_PDF_PAGES = max(1, int(os.environ.get("MAX_PDF_PAGES", "500")))
MAX_CONCURRENT_JOBS = max(1, int(os.environ.get("MAX_CONCURRENT_JOBS", "2")))
PROCESSING_TIMEOUT_SECONDS = max(
    30, int(os.environ.get("PROCESSING_TIMEOUT_SECONDS", "300"))
)
OCR_LANG = os.environ.get("OCR_LANG", "ara+eng").strip() or "ara+eng"

SERVICE_CATEGORIES = [
    {
        "id": "platforms",
        "title": "المنصات الإلكترونية",
        "emoji": "🌐",
    },
    {
        "id": "documents",
        "title": "المستندات وPDF",
        "emoji": "📁",
    },
    {
        "id": "images",
        "title": "الصور واستخراج النص",
        "emoji": "🖼️",
    },
    {
        "id": "data",
        "title": "البيانات وجهات الاتصال",
        "emoji": "🗃️",
    },
    {
        "id": "tools",
        "title": "أدوات أخرى",
        "emoji": "🧰",
    },
]

SERVICES = [
    {
        "id": "reportnest",
        "category": "platforms",
        "kind": "link",
        "title": "منصة إعداد العقود",
        "emoji": "📄",
        "desc": "منصة إعداد العقود وإنشائها وإدارتها إلكترونيًا.",
        "url": "https://reportnest.moid.gov.sy/login",
    },
    {
        "id": "cardnest",
        "category": "platforms",
        "kind": "link",
        "title": "منصة حفظ الكروت الشخصية",
        "emoji": "🪪",
        "desc": "منصة حفظ بيانات البطاقات الشخصية وأرشفتها والرجوع إليها.",
        "url": "https://cardnest.moid.gov.sy/login",
    },
    {
        "id": "format_word",
        "category": "documents",
        "kind": "direct",
        "title": "تنسيق جداول Word",
        "emoji": "📊",
        "desc": (
            "يرتّب جميع الجداول داخل ملف DOCX، ويوسّط النص أفقيًا وعموديًا، "
            "ويكرّر صف العنوان ويمنع انقسام صف الجدول بين صفحتين."
        ),
    },
    {
        "id": "remove_background",
        "category": "images",
        "kind": "direct",
        "title": "إزالة خلفية صورة",
        "emoji": "🖼️",
        "desc": "يزيل خلفية صورة PNG أو JPG أو WEBP ويعيدها بخلفية شفافة بصيغة PNG.",
    },
    {
        "id": "merge_pdf",
        "category": "documents",
        "kind": "direct",
        "title": "دمج ملفات PDF",
        "emoji": "📚",
        "desc": "يستقبل عدة ملفات PDF ويرجعها ملفًا واحدًا حسب ترتيب الإرسال.",
    },
    {
        "id": "split_pdf",
        "category": "documents",
        "kind": "direct",
        "title": "استخراج صفحات PDF",
        "emoji": "✂️",
        "desc": "يستخرج صفحات أو نطاقات محددة مثل: 1-3,5,8.",
    },
    {
        "id": "images_to_pdf",
        "category": "documents",
        "kind": "direct",
        "title": "تحويل الصور إلى PDF",
        "emoji": "🗂️",
        "desc": "يجمع عدة صور في ملف PDF واحد حسب ترتيب الإرسال.",
    },
    {
        "id": "optimize_image",
        "category": "images",
        "kind": "direct",
        "title": "ضغط وتحويل صورة",
        "emoji": "🗜️",
        "desc": "يضغط الصورة أو يحولها إلى JPG أو PNG أو WEBP أو يصغرها 50٪.",
    },
    {
        "id": "watermark_pdf",
        "category": "documents",
        "kind": "direct",
        "title": "إضافة علامة مائية إلى PDF",
        "emoji": "🔏",
        "desc": "يضيف نصًا عربيًا أو صورة شعار شفافة إلى جميع صفحات PDF.",
    },
    {
        "id": "create_qr",
        "category": "tools",
        "kind": "direct",
        "title": "إنشاء QR Code",
        "emoji": "🔳",
        "desc": "ينشئ رمز QR عالي الجودة من رابط أو نص.",
    },
    {
        "id": "ocr_image",
        "category": "images",
        "kind": "direct",
        "title": "استخراج النص من صورة",
        "emoji": "🔍",
        "desc": "يستخرج النص العربي والإنجليزي من صورة ويرجعه في ملف TXT.",
    },
    {
        "id": "excel_to_vcf",
        "category": "data",
        "kind": "direct",
        "title": "تحويل Excel إلى VCF",
        "emoji": "👥",
        "desc": (
            "ينظّف أرقام جهات الاتصال من ملف XLSX أو XLS، ويوحّد رمز الدولة "
            "ويدمج المكررات، ثم ينتج ملف VCF وتقارير مراجعة داخل حزمة ZIP."
        ),
    },
    {
        "id": "colornote_to_html",
        "category": "data",
        "kind": "direct",
        "title": "تحويل نسخ ColorNote إلى HTML",
        "emoji": "🗒️",
        "desc": (
            "يفك نسخة ColorNote واحدة أو عدة نسخ ويدمج الملاحظات المكررة، "
            "ثم يعيد أرشيف HTML مستقلًا قابلًا للبحث والتصفح دون إنترنت."
        ),
    },
]

WELCOME_MESSAGE = "👋 أهلًا بك في بوت الخدمات. اختر صنف الخدمات الذي تريده."
