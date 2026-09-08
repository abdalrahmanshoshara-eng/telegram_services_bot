"""Unformatted Arabic text -> corrected, formatted Arabic text, via Gemini.

The model does the language work (spelling, hamza, punctuation choice); the
mechanical cleanup in `tidy` runs on both sides of the call, so the result is
never worse than a deterministic pass even when the model leaves something
behind. `_validate` refuses a response that dropped or summarised the input,
so a truncated answer can never reach the user as if it were their own text.

Only `_generate` talks to the provider - replace it to change providers.
"""
from __future__ import annotations

import json
import os
import re
import time
from dataclasses import dataclass, field
from typing import Any

from google import genai
from google.genai import errors, types

MODEL = os.environ.get("GEMINI_MODEL", "gemini-3.8-flash")

API_KEY_VARS = ("GEMINI_API_KEY", "GOOGLE_API_KEY")

# A Telegram text message caps at 4096 characters, so this is never the binding
# limit in practice; it guards cost if the handler is reused elsewhere.
MAX_INPUT_CHARS = 5000
MIN_INPUT_CHARS = 2

# Above this the reply goes out as a .txt file instead of a message.
MAX_MESSAGE_CHARS = 3500

MAX_NOTES = 6

# Gemini answers a demand spike with 503, which clears in seconds. Retrying
# immediately just hits the same spike, so attempts are spaced out.
RETRY_DELAYS = (1.5, 4.0)

TATWEEL = "ـ"
ARABIC_RANGE = "؀-ۿݐ-ݿ"


class TextEditError(ValueError):
    """Raised when the text cannot be corrected.

    Subclasses ValueError to match the bot's ServiceInputError convention:
    the message is safe to show the user directly.
    """


def api_key_present() -> bool:
    return any(os.environ.get(v) for v in API_KEY_VARS)


@dataclass
class EditResult:
    text: str
    notes: list[str] = field(default_factory=list)


# --------------------------------------------------------------------------
# Mechanical cleanup (pure, provider-independent)
# --------------------------------------------------------------------------
_URL_RE = re.compile(r"(?:https?://|www\.)\S+|\S+@[\w.-]+\.\w+")

_INVISIBLE = {
    " ": " ",   # no-break space
    "​": "",    # zero-width space
    "‌": "",    # zero-width non-joiner
    "‍": "",    # zero-width joiner
    "‎": "",    # LTR mark
    "‏": "",    # RTL mark
    "﻿": "",    # BOM
}

# Closing punctuation: a space before any of these is always wrong.
_CLOSING = r"،؛؟!:\.,\?"

# Closers that must stay glued to the punctuation before them, so a full stop
# inside a quotation does not become `الوقت. "`.
_CLOSERS = "\"'»”’\\)\\]\\}"


def is_mostly_arabic(text: str) -> bool:
    """True when Arabic letters outnumber Latin ones."""
    arabic = len(re.findall(f"[{ARABIC_RANGE}]", text))
    latin = len(re.findall(r"[A-Za-z]", text))
    return arabic > latin


def _mask_urls(text: str) -> tuple[str, list[str]]:
    """Swap URLs and emails for placeholders so punctuation rules skip them."""
    found: list[str] = []

    def take(match: re.Match[str]) -> str:
        found.append(match.group(0))
        return f"\x00{len(found) - 1}\x00"

    return _URL_RE.sub(take, text), found


def _unmask_urls(text: str, found: list[str]) -> str:
    for index, original in enumerate(found):
        text = text.replace(f"\x00{index}\x00", original)
    return text


def tidy(text: str, *, arabic_punctuation: bool | None = None) -> str:
    """Deterministic whitespace and punctuation cleanup.

    `arabic_punctuation` swaps Latin `,` `;` `?` for `،` `؛` `؟`; it defaults
    to whether the text is mostly Arabic. Digits (1,000) and URLs are spared.
    """
    if not text:
        return ""
    if arabic_punctuation is None:
        arabic_punctuation = is_mostly_arabic(text)

    for source, replacement in _INVISIBLE.items():
        text = text.replace(source, replacement)
    text = text.replace(TATWEEL, "")
    text = text.replace("\r\n", "\n").replace("\r", "\n")

    text, urls = _mask_urls(text)

    if arabic_punctuation:
        # Not between digits, so "1,000" and "12;30" survive.
        text = re.sub(r"(?<!\d),(?!\d)", "،", text)
        text = re.sub(r"(?<!\d);(?!\d)", "؛", text)
        text = re.sub(f"(?<=[{ARABIC_RANGE}])[ \t]*\\?", "؟", text)

    # Repeated punctuation, keeping a real ellipsis intact.
    text = re.sub(r"\.{4,}", "...", text)
    text = re.sub(r"([،؛؟!:])\1+", r"\1", text)
    text = re.sub(r"،[ \t]*،+", "،", text)

    text = re.sub(r"[ \t]+", " ", text)
    text = re.sub(f"[ \t]+([{_CLOSING}])", r"\1", text)
    # Exactly one space after punctuation when a word follows.
    text = re.sub(f"([{_CLOSING}])(?=[^\\s\\d{_CLOSING}{_CLOSERS}])", r"\1 ", text)
    text = re.sub(r"\([ \t]+", "(", text)
    text = re.sub(r"[ \t]+\)", ")", text)

    text = _unmask_urls(text, urls)

    text = "\n".join(line.strip() for line in text.split("\n"))
    text = re.sub(r"\n{3,}", "\n\n", text)
    return text.strip()


# --------------------------------------------------------------------------
# Schema and prompt
# --------------------------------------------------------------------------
RESPONSE_SCHEMA: dict[str, Any] = {
    "type": "object",
    "additionalProperties": False,
    "required": ["corrected_text"],
    "properties": {
        "corrected_text": {
            "type": "string",
            "description": "النص كاملاً بعد التصحيح والتنسيق.",
        },
        "notes": {
            "type": "array",
            "maxItems": MAX_NOTES,
            "items": {"type": "string"},
            "description": "أهم التصحيحات، سطر قصير لكل تصحيح.",
        },
    },
}

SYSTEM = """أنت مدقق لغوي عربي محترف. مهمتك إخراج النص بالعربية الفصحى الصحيحة.

حوّل إلى الفصحى:
- كل تعبير عامي يُكتب بما يقابله في الفصحى، ولا تترك أي كلمة عامية.
- أمثلة: «مبسوطين» ← «مسرورين»، «مو» ← «ليس/غير»، «ليش» ← «لماذا»،
  «هيك» ← «هكذا»، «هلق» ← «الآن»، «بدي» ← «أريد»، «شو» ← «ما»،
  «كتير» ← «كثيراً»، «عم يفعل» ← «يفعل»، «رح يفعل» ← «سيفعل»،
  «ما حدا» ← «لا أحد»، «إجا» ← «جاء»، «لحد» ← «حتى»، «هاد» ← «هذا».
- طبّق التحويل على النص كله بالتساوي؛ الخلط بين العامية والفصحى في النص
  الواحد أسوأ من ترك النص كما هو.
- الكلام المنقول داخل علامات الاقتباس يُحوَّل أيضاً.

صحّح:
- الأخطاء الإملائية، والهمزات (أ إ آ ء ئ ؤ)، والألف المقصورة والياء (ى/ي).
- التاء المربوطة والمفتوحة (ة/ت)، والضاد والظاء، والسين والصاد.
- الإعراب الظاهر: التنوين، وجمع المذكر السالم، وأدوات النصب والجزم.
- علامات الترقيم: استخدم الفاصلة العربية «،» والفاصلة المنقوطة «؛» وعلامة الاستفهام «؟».
- المسافات: مسافة واحدة بين الكلمات، ولا مسافة قبل علامة الترقيم، ومسافة واحدة بعدها.
- التطويل والمسافات الزائدة وتكرار علامات الترقيم.
- الحروف المكتوبة بلوحة مفاتيح خاطئة والرموز الدخيلة داخل الكلمات العربية.

نسّق:
- وزّع النص على فقرات منطقية، وافصل بينها بسطر فارغ واحد.
- احتفظ بالقوائم والعناوين إن وُجدت، ووحّد شكلها.

لا تفعل:
- لا تغيّر المعنى، ولا تحذف أي معلومة، ولا تُضف معلومة جديدة.
- لا تلخّص النص، ولا تُطِل العبارة، ولا تتفاصح بكلمات غريبة؛ فصحى واضحة مباشرة.
- لا تترجم، واحتفظ بالكلمات الأجنبية والأرقام والروابط والبريد الإلكتروني كما هي.
- لا تشكّل الحروف إلا حيث يلزم لرفع اللبس.

اذكر في notes أهم ما غيّرته، وأفرد سطراً لتحويل العامية.

أعد النص كاملاً في corrected_text، ولا تعلّق عليه خارج الحقول المطلوبة."""


def build_prompt(text: str) -> str:
    return f"صحّح النص التالي ونسّقه:\n\n{text}"


# --------------------------------------------------------------------------
# Validation
# --------------------------------------------------------------------------
def _validate(data: Any, original: str) -> EditResult:
    if not isinstance(data, dict):
        raise TextEditError("رد غير متوقع من النموذج.")

    corrected = data.get("corrected_text")
    if not isinstance(corrected, str) or not corrected.strip():
        raise TextEditError("لم يُعد النموذج نصاً مصححاً.")

    corrected = tidy(corrected, arabic_punctuation=is_mostly_arabic(original))

    # A far shorter answer means the model summarised or was cut off. The user
    # asked for a correction, so handing back part of their text would be wrong.
    if len(original) >= 200 and len(corrected) < len(original) * 0.6:
        raise TextEditError("النص الناتج ناقص. أعد المحاولة أو أرسل النص على أجزاء.")

    notes: list[str] = []
    raw_notes = data.get("notes")
    if isinstance(raw_notes, list):
        for note in raw_notes[:MAX_NOTES]:
            if isinstance(note, str) and note.strip():
                notes.append(" ".join(note.split()))

    return EditResult(text=corrected, notes=notes)


# --------------------------------------------------------------------------
# Provider call
# --------------------------------------------------------------------------
def _generate(client: genai.Client, prompt: str, model: str) -> str:
    """One schema-constrained request. Returns the raw JSON text."""
    config = types.GenerateContentConfig(
        system_instruction=SYSTEM,
        response_mime_type="application/json",
        response_json_schema=RESPONSE_SCHEMA,
        temperature=0.0,
        max_output_tokens=8000,
    )
    try:
        resp = client.models.generate_content(model=model, contents=prompt, config=config)
    except errors.ClientError as e:
        raise TextEditError(f"طلب غير صالح إلى Gemini: {e}") from e
    except errors.ServerError as e:
        raise TextEditError(f"خدمة Gemini غير متاحة حالياً: {e}") from e
    except errors.APIError as e:
        raise TextEditError(f"تعذّر الاتصال بـ Gemini: {e}") from e

    feedback = getattr(resp, "prompt_feedback", None)
    if feedback is not None and getattr(feedback, "block_reason", None):
        raise TextEditError("تعذّر تدقيق هذا النص.")

    candidates = getattr(resp, "candidates", None) or []
    if candidates:
        reason = str(getattr(candidates[0], "finish_reason", "") or "")
        if "MAX_TOKENS" in reason:
            raise TextEditError("النص طويل جداً. أرسله على أجزاء أقصر.")
        if "SAFETY" in reason or "PROHIBITED" in reason or "BLOCKLIST" in reason:
            raise TextEditError("تعذّر تدقيق هذا النص.")

    text = (resp.text or "").strip()
    if not text:
        raise TextEditError("لم يُعد النموذج أي محتوى.")
    return text


def polish_text(
    text: str,
    *,
    client: genai.Client | None = None,
    model: str | None = None,
) -> EditResult:
    """Correct and format Arabic text. One Gemini call, with a single retry."""
    if not isinstance(text, str) or not text.strip():
        raise TextEditError("أرسل نصاً لتدقيقه.")

    source = tidy(text)
    if len(source) < MIN_INPUT_CHARS:
        raise TextEditError("النص قصير جداً.")
    if len(source) > MAX_INPUT_CHARS:
        raise TextEditError(
            f"النص أطول من {MAX_INPUT_CHARS} حرف. أرسله على أجزاء أقصر."
        )

    if client is None:
        if not api_key_present():
            raise TextEditError(
                "مفتاح Gemini غير محدد. عيّن GEMINI_API_KEY في متغيرات البيئة."
            )
        client = genai.Client()

    model = model or MODEL
    prompt = build_prompt(source)

    last: Exception | None = None
    for attempt in range(len(RETRY_DELAYS) + 1):
        try:
            return _validate(json.loads(_generate(client, prompt, model)), source)
        except json.JSONDecodeError as e:
            last = e
        except TextEditError as e:
            # A refusal or an over-long input will not fix itself on retry.
            if "تعذّر تدقيق" in str(e) or "طويل جداً" in str(e):
                raise
            last = e
        if attempt < len(RETRY_DELAYS):
            time.sleep(RETRY_DELAYS[attempt])
    raise TextEditError(f"فشل تدقيق النص: {last}")
