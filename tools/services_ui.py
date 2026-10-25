"""A local test page for every direct service the bot offers.

It calls the same service functions `bot.py` hands to `run_blocking`, so what
comes back here is what a Telegram user gets - without Telegram, without a
token, and without the getUpdates conflict a second bot instance causes.

The bot's own `process_*` handlers are not reusable: each one is bound to a
Telegram message and a conversation stage. This page therefore repeats their
argument plumbing and nothing else - the work itself still happens in
services/, so a bug found here is a bug the bot has.

Each run gets its own directory under one session-wide temporary directory,
which is removed on exit. Results stay downloadable until then.

The page follows the visual identity of the business-card platform (ministry
green and gold, Tajawal). The logo and header texture are served from
tools/services_ui_assets/.

    python tools/services_ui.py             # http://127.0.0.1:8000
    python tools/services_ui.py --port 9000

docker-compose.yml also runs it as the `services-ui` service, published on the
host's 127.0.0.1:8765 by default. It has no authentication: keep it on
loopback, or behind a VPN or an authenticating proxy.
"""

from __future__ import annotations

import argparse
import json
import os
import sys
import time
import traceback
import uuid
from dataclasses import dataclass
from email.parser import BytesParser
from email.policy import default as email_policy
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from tempfile import TemporaryDirectory
from urllib.parse import unquote, urlparse

PROJECT_DIR = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(PROJECT_DIR))

# Only meaningful inside the container; Docker reads them, the services do not.
_CONTAINER_ONLY = {"BOT_MEMORY_LIMIT", "BOT_CPU_LIMIT"}


def _is_container_path(value: str) -> bool:
    """True for a POSIX path such as /app/models on a Windows host.

    .env is written for the container. Its paths mean nothing on Windows and
    would send model caches and the like to the root of the current drive.
    """
    return os.name == "nt" and value.startswith("/")


def load_env() -> None:
    """Read .env into os.environ without overriding what the shell already set."""
    path = PROJECT_DIR / ".env"
    if not path.is_file():
        return
    for line in path.read_text(encoding="utf-8").splitlines():
        line = line.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        key, value = line.split("=", 1)
        key, value = key.strip(), value.strip()
        if not key or key in os.environ:
            continue
        if key in _CONTAINER_ONLY or _is_container_path(value):
            continue
        os.environ[key] = value


# config and the services read the environment at import time.
load_env()

import config  # noqa: E402
from services.background_remover import remove_background  # noqa: E402
from services.colornote import convert_colornote_backups  # noqa: E402
from services.excel_contacts import convert_excel_contacts  # noqa: E402
from services.excel_tools import (  # noqa: E402
    clean_excel_workbook,
    merge_excel_workbooks,
    split_excel_workbook,
)
from services.file_services import (  # noqa: E402
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
from services.ocr_services import ocr_image  # noqa: E402
from services.pptx_deck import create_presentation  # noqa: E402
from services.text_editor import (  # noqa: E402
    MAX_INPUT_CHARS,
    api_key_present,
    polish_text,
)
from services.voice_note import (  # noqa: E402
    MAX_DURATION_SECONDS as VOICE_MAX_SECONDS,
    SUPPORTED_SUFFIXES as VOICE_SUFFIXES,
    voice_to_formal,
)
from services.word_formatter import format_word_tables  # noqa: E402
from services.word_to_excel import word_files_to_excel  # noqa: E402

SESSION_DIR = TemporaryDirectory(prefix="services_ui_")
RESULTS: dict[str, Path] = {}

ASSETS_DIR = Path(__file__).resolve().parent / "services_ui_assets"
ASSET_TYPES = {".svg": "image/svg+xml", ".jpg": "image/jpeg", ".png": "image/png"}

IMAGE_ACCEPT = ".png,.jpg,.jpeg,.webp,.bmp,.tif,.tiff"
AUDIO_ACCEPT = ",".join(sorted(VOICE_SUFFIXES))
MAX_BODY_BYTES = max(1, config.MAX_TOTAL_UPLOAD_MB) * 1024 * 1024


# --- running a service -------------------------------------------------------


@dataclass
class Job:
    work: Path
    files: list[tuple[str, Path]]
    params: dict[str, str]

    def text(self, key: str, default: str = "") -> str:
        return (self.params.get(key) or default).strip()

    def number(self, key: str, default: int) -> int:
        raw = self.text(key)
        return int(raw) if raw.isdigit() else default

    def out(self, name: str) -> Path:
        return self.work / name

    def source(self) -> Path:
        if not self.files:
            raise ValueError("أرسل ملفاً أولاً.")
        return self.files[0][1]


def result(
    download: Path | None = None,
    summary: dict | None = None,
    text: str | None = None,
) -> dict:
    return {"download": download, "summary": summary, "text": text}


def _run_format_word(job: Job) -> dict:
    out = job.out("formatted.docx")
    format_word_tables(job.source(), out)
    return result(download=out)


def _run_remove_background(job: Job) -> dict:
    out = job.out("no_background.png")
    remove_background(job.source(), out)
    return result(download=out)


def _run_merge_pdf(job: Job) -> dict:
    out = job.out("merged.pdf")
    merge_pdfs([path for _, path in job.files], out)
    return result(download=out, summary={"الملفات المدمجة": len(job.files)})


def _run_split_pdf(job: Job) -> dict:
    out = job.out("pages.pdf")
    pages = extract_pdf_pages(job.source(), out, job.text("pages"))
    return result(download=out, summary={"الصفحات المستخرجة": pages})


def _run_manage_pdf_pages(job: Job) -> dict:
    out = job.out("managed.pdf")
    summary = manage_pdf_pages(
        job.source(),
        out,
        job.text("operation", "rotate"),
        job.text("pages"),
        job.number("angle", 90),
    )
    return result(download=out, summary=summary)


def _run_pdf_password(job: Job) -> dict:
    password = job.text("password")
    if job.text("mode", "protect") == "protect":
        out = job.out("protected.pdf")
        pages = protect_pdf(job.source(), out, password)
    else:
        out = job.out("unlocked.pdf")
        pages = unprotect_pdf(job.source(), out, password)
    return result(download=out, summary={"عدد الصفحات": pages})


def _run_images_to_pdf(job: Job) -> dict:
    out = job.out("images.pdf")
    images_to_pdf([path for _, path in job.files], out)
    return result(download=out, summary={"الصور المحوّلة": len(job.files)})


def _run_optimize_image(job: Job) -> dict:
    operation = job.text("operation", "jpeg")
    suffix = {"jpeg": "jpg", "half": "jpg", "png": "png", "webp": "webp"}[operation]
    out = job.out(f"optimized.{suffix}")
    before = job.source().stat().st_size
    optimize_image(job.source(), out, operation)
    return result(
        download=out,
        summary={
            "الحجم قبل (كيلوبايت)": round(before / 1024),
            "الحجم بعد (كيلوبايت)": round(out.stat().st_size / 1024),
        },
    )


def _run_watermark_pdf(job: Job) -> dict:
    # One upload field carries both the PDF and the optional logo image.
    source = next((p for _, p in job.files if p.suffix.lower() == ".pdf"), None)
    logo = next((p for _, p in job.files if p.suffix.lower() != ".pdf"), None)
    if source is None:
        raise ValueError("أرسل ملف PDF.")
    text = job.text("text") or None
    if text is None and logo is None:
        raise ValueError("أرسل نص العلامة المائية أو صورة شعار.")
    out = job.out("watermarked.pdf")
    watermark_pdf(source, out, text, logo)
    return result(download=out)


def _run_create_qr(job: Job) -> dict:
    out = job.out("qr.png")
    create_qr_code(job.text("data"), out)
    return result(download=out)


def _run_text_editor(job: Job) -> dict:
    edited = polish_text(job.text("text"))
    return result(text=edited.text, summary=_notes_summary(edited.notes))


def _run_voice_to_formal(job: Job) -> dict:
    # Same two-part reply the bot sends: what was heard, then the formal text,
    # so a bad result can be traced to transcription or to formalising.
    voice = voice_to_formal(job.source())
    body = (
        f"🎙️ كما سُمع:\n{voice.transcript}\n\n✅ بعد التحويل إلى الفصحى:\n"
        f"{voice.text}"
    )
    return result(text=body, summary=_notes_summary(voice.notes))


def _notes_summary(notes: list[str]) -> dict:
    return {"أهم التصحيحات": " • ".join(notes) if notes else "لا شيء"}


def _run_ocr_image(job: Job) -> dict:
    out = job.out("extracted.txt")
    ocr_image(job.source(), out, config.OCR_LANG)
    return result(download=out, text=out.read_text(encoding="utf-8"))


def _run_excel_to_vcf(job: Job) -> dict:
    out = job.out("contacts.zip")
    summary = convert_excel_contacts(
        job.source(),
        out,
        job.text("country_code", "963"),
        job.files[0][0],
        job.out("contacts.vcf"),
    )
    return result(
        download=out,
        summary={
            "الإجمالي": summary["totalRows"],
            "المقبولة": summary["validCount"],
            "المدمجة": summary["duplicateCount"],
            "المرفوضة": summary["invalidCount"],
        },
    )


def _run_clean_excel(job: Job) -> dict:
    out = job.out("cleaned_workbook.xlsx")
    return result(download=out, summary=clean_excel_workbook(job.source(), out))


def _run_split_excel(job: Job) -> dict:
    out = job.out("split_workbooks.zip")
    summary = split_excel_workbook(job.source(), out, job.text("column"))
    return result(download=out, summary=summary)


def _run_merge_excel(job: Job) -> dict:
    out = job.out("merged_workbook.xlsx")
    summary = merge_excel_workbooks(job.files, out, job.text("mode", "rows"))
    return result(download=out, summary=summary)


def _run_pptx_deck(job: Job) -> dict:
    out = job.out("presentation.pptx")
    create_presentation(
        job.source(), out, job.number("slides", 12), job.text("organisation")
    )
    return result(download=out)


def _run_colornote(job: Job) -> dict:
    out = job.out("ColorNote_notes.html")
    summary = convert_colornote_backups(job.files, out, job.text("password", "0000"))
    return result(
        download=out,
        summary={
            "الملفات": summary["files"],
            "الملاحظات": summary["total"],
            "المكررات المحذوفة": summary["duplicates_removed"],
            "الملاحظات المتغيرة": summary["changed_notes"],
        },
    )


def _run_word_to_excel(job: Job) -> dict:
    out = job.out("word_contents.xlsx")
    summary = word_files_to_excel(job.files, out)
    return result(
        download=out,
        summary={
            "الملفات": summary["files"],
            "عدد الأحرف": f"{summary['characters']:,}",
            "ملفات بلا نص": summary["empty"],
            "ملفات أُكملت في أعمدة إضافية": summary["split"],
        },
    )


# What the page renders for each service: the upload field and extra inputs.
FORMS: dict[str, dict] = {
    "format_word": {"files": {"accept": ".docx", "label": "ملف Word"}},
    "remove_background": {"files": {"accept": IMAGE_ACCEPT, "label": "صورة"}},
    "merge_pdf": {
        "files": {
            "accept": ".pdf",
            "multiple": True,
            "label": "ملفات PDF (اثنان فأكثر)",
        }
    },
    "split_pdf": {
        "files": {"accept": ".pdf", "label": "ملف PDF"},
        "fields": [
            {
                "key": "pages",
                "label": "الصفحات",
                "placeholder": "1-3,5,8",
                "required": True,
            }
        ],
    },
    "manage_pdf_pages": {
        "files": {"accept": ".pdf", "label": "ملف PDF"},
        "fields": [
            {
                "key": "operation",
                "label": "العملية",
                "type": "select",
                "options": [
                    ["rotate", "تدوير"],
                    ["delete", "حذف"],
                    ["reorder", "إعادة ترتيب"],
                ],
            },
            {
                "key": "pages",
                "label": "الصفحات",
                "placeholder": "all أو 1-3,5",
                "required": True,
            },
            {
                "key": "angle",
                "label": "زاوية التدوير",
                "type": "select",
                "options": [["90", "90"], ["180", "180"], ["270", "270"]],
            },
        ],
    },
    "pdf_password": {
        "files": {"accept": ".pdf", "label": "ملف PDF"},
        "fields": [
            {
                "key": "mode",
                "label": "العملية",
                "type": "select",
                "options": [["protect", "حماية"], ["unprotect", "فك الحماية"]],
            },
            {
                "key": "password",
                "label": "كلمة المرور",
                "type": "password",
                "required": True,
            },
        ],
    },
    "images_to_pdf": {
        "files": {"accept": IMAGE_ACCEPT, "multiple": True, "label": "صور"}
    },
    "optimize_image": {
        "files": {"accept": IMAGE_ACCEPT, "label": "صورة"},
        "fields": [
            {
                "key": "operation",
                "label": "الخيار",
                "type": "select",
                "options": [
                    ["jpeg", "ضغط JPG"],
                    ["png", "تحويل PNG"],
                    ["webp", "تحويل WEBP"],
                    ["half", "تصغير للنصف"],
                ],
            }
        ],
    },
    "watermark_pdf": {
        "files": {
            "accept": ".pdf," + IMAGE_ACCEPT,
            "multiple": True,
            "label": "ملف PDF (وصورة شعار اختيارياً)",
        },
        "fields": [
            {
                "key": "text",
                "label": "نص العلامة المائية",
                "placeholder": "اتركه فارغاً عند استخدام شعار",
            }
        ],
    },
    "create_qr": {
        "fields": [{"key": "data", "label": "الرابط أو النص", "required": True}]
    },
    "text_editor": {
        "fields": [
            {
                "key": "text",
                "label": f"النص (حتى {MAX_INPUT_CHARS} حرف)",
                "type": "textarea",
                "required": True,
            }
        ]
    },
    "voice_to_formal": {
        "files": {
            "accept": AUDIO_ACCEPT,
            "label": f"تسجيل صوتي (حتى {VOICE_MAX_SECONDS} ثانية)",
        }
    },
    "ocr_image": {"files": {"accept": IMAGE_ACCEPT, "label": "صورة"}},
    "excel_to_vcf": {
        "files": {"accept": ".xlsx,.xls", "label": "ملف Excel"},
        "fields": [
            {
                "key": "country_code",
                "label": "رمز الدولة",
                "value": "963",
                "required": True,
            }
        ],
    },
    "clean_excel": {"files": {"accept": ".xlsx", "label": "ملف Excel"}},
    "split_excel": {
        "files": {"accept": ".xlsx", "label": "ملف Excel"},
        "fields": [
            {
                "key": "column",
                "label": "العمود",
                "placeholder": "A أو اسم العمود",
                "required": True,
            }
        ],
    },
    "merge_excel": {
        "files": {
            "accept": ".xlsx",
            "multiple": True,
            "label": "ملفات Excel (اثنان فأكثر)",
        },
        "fields": [
            {
                "key": "mode",
                "label": "الطريقة",
                "type": "select",
                "options": [["rows", "صفوف متطابقة"], ["sheets", "أوراق مستقلة"]],
            }
        ],
    },
    "pptx_deck": {
        "files": {"accept": ".pdf,.docx", "label": "ملف PDF أو DOCX"},
        "fields": [
            {"key": "slides", "label": "عدد الشرائح", "value": "12"},
            {"key": "organisation", "label": "اسم الجهة", "placeholder": "اختياري"},
        ],
    },
    "colornote_to_html": {
        "files": {
            "accept": ".backup,.doc,.zip",
            "multiple": True,
            "label": "نسخ ColorNote",
        },
        "fields": [{"key": "password", "label": "كلمة المرور", "value": "0000"}],
    },
    "word_to_excel": {
        "files": {
            "accept": ".docx",
            "multiple": True,
            "label": "ملفات Word بصيغة DOCX (كل ملف صف في Excel)",
        }
    },
}

RUNNERS = {
    "format_word": _run_format_word,
    "remove_background": _run_remove_background,
    "merge_pdf": _run_merge_pdf,
    "split_pdf": _run_split_pdf,
    "manage_pdf_pages": _run_manage_pdf_pages,
    "pdf_password": _run_pdf_password,
    "images_to_pdf": _run_images_to_pdf,
    "optimize_image": _run_optimize_image,
    "watermark_pdf": _run_watermark_pdf,
    "create_qr": _run_create_qr,
    "text_editor": _run_text_editor,
    "voice_to_formal": _run_voice_to_formal,
    "ocr_image": _run_ocr_image,
    "excel_to_vcf": _run_excel_to_vcf,
    "clean_excel": _run_clean_excel,
    "split_excel": _run_split_excel,
    "merge_excel": _run_merge_excel,
    "pptx_deck": _run_pptx_deck,
    "colornote_to_html": _run_colornote,
    "word_to_excel": _run_word_to_excel,
}

NEEDS_GEMINI = {"text_editor", "voice_to_formal", "pptx_deck"}


def service_catalogue() -> list[dict]:
    """config.SERVICES plus what the page needs to draw each form."""
    catalogue = []
    for service in config.SERVICES:
        entry = {
            "id": service["id"],
            "title": service["title"],
            "emoji": service.get("emoji", ""),
            "desc": service.get("desc", ""),
            "category": service["category"],
            "kind": service["kind"],
            "url": service.get("url", ""),
            "needs_gemini": service["id"] in NEEDS_GEMINI,
        }
        entry.update(FORMS.get(service["id"], {}))
        if service["kind"] == "direct" and service["id"] not in RUNNERS:
            entry["unsupported"] = True
        catalogue.append(entry)
    return catalogue


def tesseract_available() -> bool:
    try:
        import pytesseract

        pytesseract.get_tesseract_version()
        return True
    except Exception:
        return False


def parse_multipart(content_type: str, body: bytes) -> tuple[dict, list]:
    """Split a multipart/form-data body into text fields and uploaded files."""
    raw = (
        b"Content-Type: "
        + content_type.encode("utf-8")
        + b"\r\nMIME-Version: 1.0\r\n\r\n"
        + body
    )
    message = BytesParser(policy=email_policy).parsebytes(raw)
    params: dict[str, str] = {}
    files: list[tuple[str, bytes]] = []
    for part in message.iter_parts():
        name = part.get_param("name", header="content-disposition")
        filename = part.get_filename()
        payload = part.get_payload(decode=True) or b""
        if filename:
            files.append((filename, payload))
        elif name:
            params[name] = payload.decode("utf-8", "replace")
    return params, files


# --- the page ----------------------------------------------------------------

PAGE = r"""<!doctype html>
<html lang="ar" dir="rtl">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>منصة الخدمات الرقمية</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="stylesheet" href="https://fonts.googleapis.com/css2?family=Tajawal:wght@400;500;700;800;900&family=Cairo:wght@400;600;700&display=swap">
<style>
  /* Tokens taken from the business-card platform's design system. */
  :root {
    --green-deep:#04302a; --green-header:#0a5346; --green-mid:#0e6d5c;
    --green-soft:#1a8f79; --green-wash:#eef5f2;
    --gold:#c7a66b; --gold-hover:#dbc08f; --gold-ink:#8a6a2f;
    --gold-border:rgba(199,166,107,.42); --gold-wash:rgba(199,166,107,.12);
    --bg:#f4f2ec; --bg-tint:#eeebe2; --surface:#fff; --surface-2:#faf9f5;
    --surface-sunken:#f6f4ee;
    --text:#14231f; --text-muted:#5b6f69; --text-soft:#83958f; --text-invert:#f6f2e7;
    --border:rgba(10,83,70,.12); --border-strong:rgba(10,83,70,.20);
    --border-hard:rgba(10,83,70,.32);
    --danger:#b8362a; --danger-ink:#8f2a20; --danger-wash:#fdeeeb;
    --danger-border:rgba(184,54,42,.28);
    --success:#14684a; --success-wash:#eaf7f0; --success-border:rgba(20,104,74,.26);
    --warning:#8a6413; --warning-wash:#fdf4e0; --warning-border:rgba(199,152,42,.35);
    --r-sm:9px; --r-md:12px; --r-lg:16px; --r-xl:20px; --r-pill:999px;
    --shadow-xs:0 1px 2px rgba(10,47,40,.06);
    --shadow-sm:0 1px 3px rgba(10,47,40,.07), 0 4px 12px -6px rgba(10,47,40,.10);
    --shadow-md:0 2px 6px rgba(10,47,40,.07), 0 12px 28px -14px rgba(10,47,40,.20);
    --font:"Tajawal","Cairo",system-ui,-apple-system,"Segoe UI",sans-serif;
    --ease:cubic-bezier(.4,0,.2,1); --fast:140ms;
    --header-h:4rem; --sidebar-w:288px;
    --focus-ring:0 0 0 3px rgba(199,166,107,.40), 0 0 0 1px var(--green-mid);
  }
  *, *::before, *::after { box-sizing:border-box; }
  html, body { margin:0; min-height:100%; }
  body {
    font-family:var(--font); color:var(--text); font-size:15px; line-height:1.7;
    background-color:var(--bg);
    background-image:
      radial-gradient(900px 420px at 88% -8%, rgba(199,166,107,.10), transparent 62%),
      radial-gradient(760px 380px at 4% 0%, rgba(10,83,70,.07), transparent 58%);
    background-attachment:fixed;
    padding-top:var(--header-h);
    -webkit-font-smoothing:antialiased;
  }
  ::selection { background:rgba(199,166,107,.22); color:var(--green-deep); }
  :focus-visible { outline:none; box-shadow:var(--focus-ring); border-radius:var(--r-sm); }
  * { scrollbar-width:thin; scrollbar-color:rgba(10,83,70,.28) transparent; }

  /* Top bar */
  .site-header {
    position:fixed; inset-inline:0; top:0; z-index:50; height:var(--header-h);
    display:flex; align-items:center; padding:0 24px; overflow:hidden;
    background-color:var(--green-header);
    background-image:url(assets/topbar-bg.jpg); background-size:460px;
    background-position:center; background-blend-mode:soft-light;
    border-bottom:1px solid rgba(199,166,107,.28);
    box-shadow:0 1px 0 rgba(255,255,255,.05) inset, 0 10px 30px -18px rgba(4,48,42,.9);
  }
  .header-overlay {
    position:absolute; inset:0; z-index:0; pointer-events:none;
    background:linear-gradient(180deg, rgba(4,48,42,.74), rgba(4,48,42,.86));
  }
  .header-inner {
    position:relative; z-index:1; width:100%; height:100%;
    display:flex; align-items:center; justify-content:space-between; gap:16px;
  }
  .header-logo { height:44px; flex-shrink:0; display:flex; align-items:center; }
  .header-logo img { height:100%; max-height:44px; width:auto; display:block; }
  .header-badges { display:flex; gap:8px; align-items:center; flex-wrap:nowrap; min-width:0; }
  .header-badge {
    display:inline-flex; align-items:center; gap:7px; white-space:nowrap;
    background:rgba(199,166,107,.14); border:1px solid var(--gold-border);
    color:#ecd9b4; font-size:12px; font-weight:700; padding:6px 14px;
    border-radius:var(--r-pill); backdrop-filter:blur(4px);
  }
  .header-badge::before {
    content:""; width:6px; height:6px; border-radius:50%; background:var(--gold);
    flex-shrink:0;
  }
  .header-badge.ok::before { background:#5fd3a2; }
  .header-badge.bad::before { background:#f08a7f; }

  /* Shell */
  .app-shell { display:flex; align-items:stretch; }
  .app-sidebar {
    position:sticky; top:var(--header-h); align-self:flex-start;
    width:var(--sidebar-w); flex:0 0 var(--sidebar-w);
    height:calc(100vh - var(--header-h)); overflow-y:auto;
    padding:22px 14px; color:var(--text-invert);
    background:
      linear-gradient(180deg, rgba(199,166,107,.07), transparent 220px),
      linear-gradient(165deg, var(--green-header), var(--green-deep));
    border-inline-start:1px solid rgba(199,166,107,.22);
    scrollbar-color:rgba(199,166,107,.35) transparent;
  }
  .sidebar-head {
    display:flex; align-items:center; gap:12px;
    padding:4px 6px 16px; margin-bottom:6px;
    border-bottom:1px solid rgba(199,166,107,.20);
  }
  .sidebar-avatar {
    width:42px; height:42px; border-radius:14px; flex-shrink:0;
    display:flex; align-items:center; justify-content:center; font-size:20px;
    background:linear-gradient(140deg, var(--gold-hover), var(--gold));
    box-shadow:0 4px 12px -4px rgba(0,0,0,.45);
  }
  .sidebar-title { font-weight:800; font-size:14.5px; line-height:1.4; }
  .sidebar-sub { font-size:11.5px; font-weight:700; color:var(--gold-hover); }
  .nav-group {
    margin:18px 8px 6px; font-size:11.5px; font-weight:800;
    color:var(--gold-hover); letter-spacing:.02em;
  }
  .nav-link {
    position:relative; display:flex; align-items:center; gap:11px; width:100%;
    padding:10px 13px; margin-bottom:3px; border:0; border-radius:var(--r-md);
    background:transparent; color:rgba(240,235,221,.84);
    font:inherit; font-weight:700; font-size:14px; text-align:start; cursor:pointer;
    transition:background var(--fast) var(--ease), color var(--fast) var(--ease);
  }
  .nav-link:hover { background:rgba(255,255,255,.08); color:#fff; }
  .nav-link.active { background:rgba(199,166,107,.16); color:#fff; }
  .nav-link.active::before {
    content:""; position:absolute; inset-inline-start:0; top:9px; bottom:9px;
    width:3px; border-radius:var(--r-pill);
    background:linear-gradient(180deg, var(--gold-hover), var(--gold));
  }
  .nav-icon { flex:0 0 22px; text-align:center; font-size:16px; }
  .nav-label { min-width:0; overflow:hidden; text-overflow:ellipsis; white-space:nowrap; }
  .nav-new {
    margin-inline-start:auto; font-size:10.5px; font-weight:800; color:var(--green-deep);
    background:var(--gold); border-radius:var(--r-pill); padding:1px 8px;
  }

  .app-main { flex:1 1 auto; min-width:0; padding:26px clamp(16px, 3vw, 34px) 40px; }
  .container { max-width:980px; margin:0 auto; }

  /* Hero */
  .hero {
    position:relative; overflow:hidden; isolation:isolate; color:#f7f3e8;
    background-color:var(--green-header);
    background-image:url(assets/topbar-bg.jpg); background-size:460px;
    background-position:center; background-blend-mode:soft-light;
    border:1px solid rgba(199,166,107,.30); border-radius:var(--r-xl);
    padding:28px 32px; margin-bottom:22px; box-shadow:var(--shadow-md);
  }
  .hero .header-overlay {
    z-index:-1;
    background:linear-gradient(135deg, rgba(4,48,42,.88), rgba(10,83,70,.74) 55%, rgba(4,48,42,.90));
  }
  .hero::after {
    content:""; position:absolute; inset-inline:0; top:0; height:2px; opacity:.7;
    background:linear-gradient(90deg, transparent, var(--gold), transparent);
  }
  .hero h1 { margin:0; font-size:clamp(22px, 3.2vw, 30px); font-weight:800; line-height:1.32; color:#fffaf0; }
  .hero-accent-line {
    width:56px; height:3px; margin:12px 0 13px; border-radius:var(--r-pill);
    background:linear-gradient(90deg, var(--gold), rgba(199,166,107,0));
  }
  .hero p { margin:0; max-width:74ch; color:rgba(247,243,232,.82); font-size:14.5px; line-height:1.85; }
  .hero .badge { margin-top:14px; background:rgba(199,166,107,.16); color:#ecd9b4; }

  /* Cards */
  .card {
    background:var(--surface); border:1px solid var(--border); border-radius:var(--r-lg);
    padding:24px; margin-bottom:18px; box-shadow:var(--shadow-sm);
  }
  .card h2 {
    margin:0 0 18px; font-size:17.5px; font-weight:800; color:var(--green-header);
    display:flex; align-items:center; gap:10px;
  }
  .card h2::before {
    content:""; width:3px; height:19px; flex-shrink:0; border-radius:var(--r-pill);
    background:linear-gradient(180deg, var(--gold-hover), var(--gold));
  }

  /* Forms */
  form { display:grid; gap:16px; }
  label { display:grid; gap:6px; font-weight:700; font-size:13.5px; color:var(--text); }
  .req { color:var(--danger); }
  input, textarea, select {
    width:100%; min-width:0; border:1px solid var(--border-strong); border-radius:var(--r-md);
    padding:11px 14px; font:inherit; font-size:14px; font-weight:500;
    background:var(--surface-2); color:var(--text); outline:none;
    transition:border-color var(--fast) var(--ease), box-shadow var(--fast) var(--ease);
  }
  input::placeholder, textarea::placeholder { color:var(--text-soft); font-weight:400; }
  input:hover, textarea:hover, select:hover { border-color:var(--border-hard); background-color:var(--surface); }
  input:focus, textarea:focus, select:focus {
    background-color:var(--surface); border-color:var(--green-mid);
    box-shadow:0 0 0 3px rgba(14,109,92,.12);
  }
  textarea { resize:vertical; min-height:180px; line-height:1.75; }
  select {
    appearance:none; cursor:pointer; padding-inline-end:38px;
    background-image:url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='12' height='8' viewBox='0 0 12 8' fill='none'%3E%3Cpath d='M1 1.5 6 6.5l5-5' stroke='%235b6f69' stroke-width='1.8' stroke-linecap='round' stroke-linejoin='round'/%3E%3C/svg%3E");
    background-position:left 14px center; background-size:12px 8px; background-repeat:no-repeat;
  }
  input[type="file"] {
    padding:9px 12px; font-size:13px; font-weight:600; cursor:pointer;
    background-color:var(--surface-sunken); border-style:dashed;
  }
  input[type="file"]::file-selector-button {
    font:inherit; font-weight:700; font-size:13px; margin-inline-end:12px;
    padding:7px 14px; border:1px solid var(--border-strong); border-radius:var(--r-sm);
    background:var(--surface); color:var(--green-header); cursor:pointer;
  }
  input[type="file"]::file-selector-button:hover { background:var(--gold-wash); border-color:var(--gold-border); }
  .file-list { display:flex; flex-wrap:wrap; gap:6px; margin-top:2px; }
  .file-list:empty { display:none; }
  .file-chip {
    font-size:12px; font-weight:700; color:var(--green-header);
    background:var(--green-wash); border:1px solid var(--border);
    border-radius:var(--r-pill); padding:2px 10px;
  }
  .helper { color:var(--text-muted); font-size:13px; font-weight:500; }

  /* Buttons */
  .actions { display:flex; gap:10px; flex-wrap:wrap; margin-top:4px; }
  .btn {
    display:inline-flex; align-items:center; justify-content:center; gap:8px;
    min-height:44px; padding:11px 22px; border-radius:var(--r-md);
    font:inherit; font-size:14.5px; font-weight:800; line-height:1.2;
    text-decoration:none; cursor:pointer;
    transition:box-shadow var(--fast) var(--ease), transform var(--fast) var(--ease), filter var(--fast) var(--ease);
  }
  .btn:active:not(:disabled) { transform:translateY(1px); }
  .btn:disabled { opacity:.55; cursor:progress; }
  .btn-gold {
    border:1px solid rgba(138,106,47,.45); color:#2b1f08;
    background:linear-gradient(135deg, #cdaa6b, #b8934f);
    box-shadow:0 1px 2px rgba(90,66,20,.18), 0 8px 18px -10px rgba(138,106,47,.7);
  }
  .btn-gold:hover:not(:disabled) { background:linear-gradient(135deg, #dbbc84, #c7a15c); }
  .btn-green {
    border:1px solid rgba(4,48,42,.45); color:var(--text-invert);
    background:linear-gradient(135deg, var(--green-mid), var(--green-header));
    box-shadow:0 1px 2px rgba(4,48,42,.18), 0 8px 18px -10px rgba(10,83,70,.7);
  }
  .btn-green:hover { background:linear-gradient(135deg, var(--green-soft), var(--green-mid)); }
  .spinner {
    width:16px; height:16px; border-radius:50%; display:none;
    border:2px solid rgba(43,31,8,.25); border-top-color:#2b1f08;
    animation:spin .8s linear infinite;
  }
  .btn:disabled .spinner { display:inline-block; }
  @keyframes spin { to { transform:rotate(360deg); } }

  /* Feedback */
  .status-box {
    padding:12px 15px; border-radius:var(--r-md); border:1px solid var(--border-strong);
    border-inline-start:3px solid var(--green-mid); background:var(--green-wash);
    color:var(--text-muted); font-size:14px; font-weight:600; line-height:1.75;
  }
  .status-box + .status-box { margin-top:10px; }
  .status-box.success { color:var(--success); background:var(--success-wash);
    border-color:var(--success-border); border-inline-start-color:var(--success); }
  .status-box.warning { color:var(--warning); background:var(--warning-wash);
    border-color:var(--warning-border); border-inline-start-color:#d9a441; }
  .status-box.error { color:var(--danger-ink); background:var(--danger-wash);
    border-color:var(--danger-border); border-inline-start-color:var(--danger); }
  .notices:empty { display:none; }
  .notices { margin-bottom:18px; }
  .badge {
    display:inline-flex; align-items:center; gap:6px; padding:4px 11px;
    border-radius:var(--r-pill); border:1px solid var(--gold-border);
    background:var(--gold-wash); color:var(--gold-ink);
    font-size:12px; font-weight:800; white-space:nowrap;
  }
  .badge::before { content:""; width:6px; height:6px; border-radius:50%; background:currentColor; }
  .result-head { display:flex; align-items:center; justify-content:space-between; gap:12px; flex-wrap:wrap; }
  .result-head h2 { margin:0; }
  .empty {
    padding:30px 20px; border:1px dashed var(--border-strong); border-radius:var(--r-md);
    background:var(--surface-sunken); color:var(--text-muted);
    font-weight:700; text-align:center;
  }

  /* Tables */
  .table-wrap { margin-top:16px; border:1px solid var(--border); border-radius:var(--r-lg); overflow:hidden; }
  table { width:100%; border-collapse:separate; border-spacing:0; }
  th, td { padding:11px 14px; text-align:start; font-size:13.5px; border-bottom:1px solid var(--border); overflow-wrap:anywhere; }
  th {
    background:linear-gradient(180deg, var(--green-header), #094a3e); color:#ecf3f0;
    font-weight:700; font-size:12.5px; border-bottom:1px solid rgba(199,166,107,.35);
  }
  tr:last-child td { border-bottom:0; }
  tbody tr:nth-child(even) td { background:var(--surface-2); }
  td:last-child { font-weight:700; color:var(--green-header); }

  pre {
    white-space:pre-wrap; word-break:break-word; margin:16px 0 0;
    background:var(--surface-sunken); border:1px solid var(--border-strong);
    border-radius:var(--r-md); padding:14px 16px;
    font:500 14px/1.85 var(--font);
  }
  pre.trace { direction:ltr; text-align:left; font:12px/1.6 ui-monospace, Consolas, monospace; color:var(--text-muted); }
  .result-actions { margin-top:18px; }

  @media (max-width: 900px) {
    .app-shell { flex-direction:column; }
    /* Stacked: the service list becomes one swipeable row of chips. */
    .app-sidebar {
      position:static; align-self:stretch; width:auto; flex:none; height:auto;
      padding:12px 16px; overflow:hidden;
    }
    .sidebar-head { display:none; }
    #nav { display:flex; gap:8px; overflow-x:auto; padding-bottom:4px; scroll-snap-type:x proximity; }
    .nav-group { display:none; }
    .nav-link {
      flex:0 0 auto; width:auto; margin:0; padding:8px 14px; scroll-snap-align:start;
      border:1px solid rgba(199,166,107,.30); border-radius:var(--r-pill);
    }
    .nav-link.active::before { display:none; }
    .nav-link.active { border-color:var(--gold); }
    .header-badges .header-badge:not(:first-child) { display:none; }
  }
  @media (max-width: 520px) {
    .site-header { padding:0 16px; }
    .header-logo { height:34px; }
    .header-badges { display:none; }
    .hero { padding:22px 20px; }
    .card { padding:18px 16px; }
  }
  @media (prefers-reduced-motion: reduce) { * { transition:none !important; animation:none !important; } }
</style>
</head>
<body>
<header class="site-header">
  <div class="header-overlay" aria-hidden="true"></div>
  <div class="header-inner">
    <div class="header-logo"><img src="assets/ministry-logo-header.svg" alt="وزارة الصناعة والتجارة الداخلية"></div>
    <div class="header-badges">
      <span class="header-badge">بيئة اختبار محلية</span>
      <span class="header-badge" id="c-model"></span>
      <span class="header-badge" id="c-key"></span>
      <span class="header-badge" id="c-ocr"></span>
    </div>
  </div>
</header>
<div class="app-shell">
  <aside class="app-sidebar" aria-label="الخدمات">
    <div class="sidebar-head">
      <div class="sidebar-avatar" aria-hidden="true">🤖</div>
      <div>
        <div class="sidebar-title">الخدمات الرقمية</div>
        <div class="sidebar-sub" id="c-count"></div>
      </div>
    </div>
    <nav id="nav"></nav>
  </aside>
  <main class="app-main">
    <div class="container">
      <section class="hero">
        <div class="header-overlay" aria-hidden="true"></div>
        <h1 id="hero-title">منصة الخدمات الرقمية</h1>
        <div class="hero-accent-line"></div>
        <p id="hero-desc">أدوات لمعالجة المستندات والصور والبيانات في مكان واحد. اختر الخدمة من القائمة، وارفع ملفاتك، واحصل على النتيجة مباشرة.</p>
      </section>
      <div class="notices" id="notices"></div>
      <section class="card" id="panel">
        <h2>ابدأ من هنا</h2>
        <div class="empty">اختر خدمة من القائمة الجانبية.</div>
      </section>
      <section class="card" id="out" hidden></section>
    </div>
  </main>
</div>
<script>
const $ = (id) => document.getElementById(id);
let SERVICES = [], INFO = {}, current = null;

const CATEGORIES = [
  ["platforms", "المنصات الإلكترونية"], ["documents", "المستندات وPDF"],
  ["images", "الصور واستخراج النص"], ["data", "البيانات وجهات الاتصال"],
  ["tools", "أدوات أخرى"],
];
const NEW_SERVICES = new Set(["word_to_excel"]);

function el(tag, props, ...children) {
  const node = document.createElement(tag);
  Object.assign(node, props || {});
  for (const child of children) if (child != null) node.append(child);
  return node;
}

function badge(id, text, ok) {
  const b = $(id);
  b.textContent = text;
  b.className = "header-badge " + (ok ? "ok" : "bad");
}

async function boot() {
  const res = await fetch("api/services");
  const data = await res.json();
  SERVICES = data.services; INFO = data.info;
  badge("c-model", "النموذج: " + INFO.gemini_model, true);
  badge("c-key", "مفتاح Gemini: " + (INFO.gemini_key ? "موجود" : "غير موجود"), INFO.gemini_key);
  badge("c-ocr", "Tesseract: " + (INFO.tesseract ? "مثبّت" : "غير مثبّت"), INFO.tesseract);
  const runnable = SERVICES.filter((s) => s.kind === "direct" && !s.unsupported).length;
  $("c-count").textContent = runnable + " خدمة قابلة للتجربة";
  drawNav();
  const fromHash = decodeURIComponent(location.hash.slice(1));
  if (SERVICES.some((s) => s.id === fromHash)) select(fromHash);
}

function drawNav() {
  const nav = $("nav"); nav.innerHTML = "";
  for (const [cid, title] of CATEGORIES) {
    const items = SERVICES.filter((s) => s.category === cid);
    if (!items.length) continue;
    nav.append(el("div", { className: "nav-group", textContent: title }));
    for (const s of items) {
      const b = el("button", { type: "button", className: "nav-link" },
        el("span", { className: "nav-icon", textContent: s.emoji || "•" }),
        el("span", { className: "nav-label", textContent: s.title }),
        NEW_SERVICES.has(s.id) ? el("span", { className: "nav-new", textContent: "جديد" }) : null);
      b.dataset.id = s.id;
      b.onclick = () => select(s.id);
      nav.append(b);
    }
  }
}

function select(id) {
  current = SERVICES.find((s) => s.id === id);
  history.replaceState(null, "", "#" + id);
  for (const b of document.querySelectorAll(".nav-link")) {
    b.classList.toggle("active", b.dataset.id === id);
    if (b.dataset.id === id) b.scrollIntoView({ block: "nearest", inline: "nearest" });
  }
  $("out").hidden = true;
  $("hero-title").textContent = (current.emoji ? current.emoji + " " : "") + current.title;
  $("hero-desc").textContent = current.desc || "";
  drawForm();
  if (window.innerWidth <= 900) $("panel").scrollIntoView({ behavior: "smooth" });
}

function notice(kind, text) {
  $("notices").append(el("div", { className: "status-box " + kind, textContent: text }));
}

function drawForm() {
  const s = current, p = $("panel");
  $("notices").innerHTML = "";
  p.innerHTML = "";

  if (s.kind === "link") {
    p.append(el("h2", { textContent: "رابط خارجي" }),
      el("div", { className: "status-box", textContent: "هذه الخدمة رابط خارجي، لا معالجة فيها لتُختبر هنا." }),
      el("div", { className: "actions", style: "margin-top:18px" },
        el("a", { className: "btn btn-green", href: s.url || "#", target: "_blank", rel: "noopener", textContent: "فتح المنصة" })));
    return;
  }
  if (s.unsupported) {
    p.append(el("h2", { textContent: "غير متاحة" }),
      el("div", { className: "status-box error", textContent: "لا يوجد مُشغّل لهذه الخدمة في هذه الصفحة." }));
    return;
  }
  if (s.needs_gemini && !INFO.gemini_key)
    notice("warning", "تحتاج هذه الخدمة مفتاح GEMINI_API_KEY في ملف .env");
  if (s.id === "ocr_image" && !INFO.tesseract)
    notice("warning", "Tesseract غير مثبّت على هذا الجهاز، وستفشل الخدمة.");

  p.append(el("h2", { textContent: "مدخلات الخدمة" }));
  const form = el("form", { id: "form" });
  if (s.files) {
    const inp = el("input", { type: "file", name: "files", accept: s.files.accept || "", required: true });
    if (s.files.multiple) inp.multiple = true;
    const list = el("div", { className: "file-list" });
    inp.onchange = () => {
      list.innerHTML = "";
      for (const f of inp.files) list.append(el("span", { className: "file-chip", textContent: f.name }));
    };
    const label = el("label", {}, labelText(s.files.label, true), inp, list);
    if (s.files.multiple)
      label.append(el("span", { className: "helper", textContent: "يمكنك اختيار عدة ملفات معاً، وتُعالَج بترتيب الاختيار." }));
    form.append(label);
  }
  for (const f of (s.fields || [])) {
    let input;
    if (f.type === "select") {
      input = el("select");
      for (const [value, text] of f.options) input.append(el("option", { value, textContent: text }));
    } else if (f.type === "textarea") {
      input = el("textarea");
    } else {
      input = el("input", { type: f.type === "password" ? "password" : "text" });
    }
    input.name = f.key;
    if (f.value) input.value = f.value;
    if (f.placeholder) input.placeholder = f.placeholder;
    if (f.required) input.required = true;
    form.append(el("label", {}, labelText(f.label, f.required), input));
  }
  const go = el("button", { type: "submit", className: "btn btn-gold" },
    el("span", { className: "spinner" }), el("span", { className: "go-label", textContent: "شغّل الخدمة" }));
  form.append(el("div", { className: "actions" }, go));
  form.onsubmit = run;
  p.append(form);
}

function labelText(text, required) {
  const span = el("span", { textContent: text + " " });
  if (required) span.append(el("span", { className: "req", textContent: "*" }));
  return span;
}

async function run(event) {
  event.preventDefault();
  const form = $("form"), go = form.querySelector("button[type=submit]");
  const goLabel = go.querySelector(".go-label");
  const body = new FormData(form);
  body.append("service", current.id);
  go.disabled = true; goLabel.textContent = "جارٍ التنفيذ...";
  const out = $("out");
  out.hidden = false;
  out.innerHTML = "";
  out.append(el("h2", { textContent: "النتيجة" }),
    el("div", { className: "empty", textContent: "جارٍ التنفيذ، قد يستغرق الأمر بعض الوقت." }));
  try {
    const res = await fetch("api/run", { method: "POST", body });
    show(await res.json());
  } catch (err) {
    out.innerHTML = "";
    out.append(el("h2", { textContent: "النتيجة" }),
      el("div", { className: "status-box error", textContent: "تعذّر الاتصال بالخادم: " + err }));
  } finally {
    go.disabled = false; goLabel.textContent = "شغّل الخدمة";
  }
}

function show(data) {
  const out = $("out");
  out.innerHTML = "";
  const head = el("div", { className: "result-head" }, el("h2", { textContent: "النتيجة" }));
  if (data.ms != null) head.append(el("span", { className: "badge", textContent: data.ms.toLocaleString("ar") + " م.ث" }));
  out.append(head);
  out.append(data.ok
    ? el("div", { className: "status-box success", style: "margin-top:16px", textContent: "✓ نجحت الخدمة." })
    : el("div", { className: "status-box error", style: "margin-top:16px", textContent: "✖ " + data.error }));

  if (data.summary && Object.keys(data.summary).length) {
    const table = el("table");
    table.append(el("thead", {}, el("tr", {}, el("th", { textContent: "البند" }), el("th", { textContent: "القيمة" }))));
    const tbody = el("tbody");
    for (const [k, v] of Object.entries(data.summary))
      tbody.append(el("tr", {}, el("td", { textContent: k }), el("td", { textContent: String(v) })));
    table.append(tbody);
    out.append(el("div", { className: "table-wrap" }, table));
  }
  if (data.text) out.append(el("pre", { textContent: data.text }));
  if (data.download) {
    out.append(el("div", { className: "actions result-actions" },
      el("a", { className: "btn btn-green", href: data.download.replace(/^\//, ""), download: data.filename, textContent: "⬇ تنزيل " + data.filename })));
  }
  if (data.trace) out.append(el("pre", { className: "trace", textContent: data.trace }));
  out.scrollIntoView({ behavior: "smooth", block: "nearest" });
}

boot();
</script>
</body>
</html>
"""


# --- HTTP --------------------------------------------------------------------


class Handler(BaseHTTPRequestHandler):
    def log_message(self, fmt, *args):
        sys.stderr.write("  %s\n" % (fmt % args))

    def _send(
        self, status: int, body: bytes, content_type: str, extra: dict | None = None
    ) -> None:
        self.send_response(status)
        self.send_header("Content-Type", content_type)
        self.send_header("Content-Length", str(len(body)))
        for key, value in (extra or {}).items():
            self.send_header(key, value)
        self.end_headers()
        self.wfile.write(body)

    def _json(self, payload: dict, status: int = 200) -> None:
        body = json.dumps(payload, ensure_ascii=False).encode("utf-8")
        self._send(status, body, "application/json; charset=utf-8")

    def do_GET(self) -> None:
        path = urlparse(self.path).path
        if path in ("/", "/index.html"):
            self._send(200, PAGE.encode("utf-8"), "text/html; charset=utf-8")
            return
        if path == "/api/services":
            self._json(
                {
                    "services": service_catalogue(),
                    "info": {
                        "gemini_model": os.environ.get("GEMINI_MODEL", "غير محدد"),
                        "gemini_key": api_key_present(),
                        "tesseract": tesseract_available(),
                        "ocr_lang": config.OCR_LANG,
                    },
                }
            )
            return
        if path.startswith("/assets/"):
            self._asset(unquote(path.split("/assets/", 1)[1]))
            return
        if path.startswith("/download/"):
            self._download(unquote(path.split("/download/", 1)[1]))
            return
        self._json({"error": "not found"}, 404)

    def _asset(self, name: str) -> None:
        # Only plain file names from the assets folder; no path components.
        target = ASSETS_DIR / Path(name).name
        content_type = ASSET_TYPES.get(target.suffix.lower())
        if content_type is None or not target.is_file():
            self._json({"error": "not found"}, 404)
            return
        self._send(
            200,
            target.read_bytes(),
            content_type,
            {"Cache-Control": "max-age=86400"},
        )

    def _download(self, token: str) -> None:
        target = RESULTS.get(token)
        if target is None or not target.is_file():
            self._json({"error": "النتيجة لم تعد متاحة."}, 404)
            return
        self._send(
            200,
            target.read_bytes(),
            "application/octet-stream",
            {"Content-Disposition": f'attachment; filename="{target.name}"'},
        )

    def do_POST(self) -> None:
        if urlparse(self.path).path != "/api/run":
            self._json({"error": "not found"}, 404)
            return
        length = int(self.headers.get("Content-Length") or 0)
        if length > MAX_BODY_BYTES:
            self._json({"ok": False, "error": "حجم الرفع أكبر من الحد المسموح."}, 413)
            return
        params, uploads = parse_multipart(
            self.headers.get("Content-Type", ""), self.rfile.read(length)
        )
        service_id = params.pop("service", "")
        runner = RUNNERS.get(service_id)
        if runner is None:
            self._json({"ok": False, "error": f"خدمة غير معروفة: {service_id}"}, 400)
            return

        work = Path(SESSION_DIR.name) / uuid.uuid4().hex
        work.mkdir(parents=True)
        files = []
        for index, (filename, payload) in enumerate(uploads, start=1):
            if not payload:
                continue
            safe = Path(filename).name or f"upload_{index}"
            target = work / f"{index:02d}_{safe}"
            target.write_bytes(payload)
            files.append((safe, target))

        started = time.perf_counter()
        try:
            produced = runner(Job(work, files, params))
        except Exception as error:
            self._json(
                {
                    "ok": False,
                    "error": str(error) or error.__class__.__name__,
                    "trace": traceback.format_exc(limit=4),
                    "ms": round((time.perf_counter() - started) * 1000),
                }
            )
            return

        payload = {
            "ok": True,
            "ms": round((time.perf_counter() - started) * 1000),
            "summary": produced.get("summary"),
            "text": produced.get("text"),
        }
        download = produced.get("download")
        if download is not None:
            token = uuid.uuid4().hex
            RESULTS[token] = download
            payload["download"] = f"/download/{token}"
            payload["filename"] = download.name
        self._json(payload)


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--host", default="127.0.0.1")
    parser.add_argument("--port", type=int, default=8000)
    args = parser.parse_args()

    runnable = sum(1 for s in config.SERVICES if s["id"] in RUNNERS)
    print(f"Services     : {runnable} runnable, {len(config.SERVICES)} listed")
    print(f"Gemini model : {os.environ.get('GEMINI_MODEL', 'unset')}")
    print(f"Gemini key   : {'present' if api_key_present() else 'MISSING'}")
    print(f"Tesseract    : {'present' if tesseract_available() else 'MISSING'}")
    print(f"Open         : http://{args.host}:{args.port}\n")

    server = ThreadingHTTPServer((args.host, args.port), Handler)
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\nstopped")
    finally:
        server.server_close()
        SESSION_DIR.cleanup()


if __name__ == "__main__":
    main()
