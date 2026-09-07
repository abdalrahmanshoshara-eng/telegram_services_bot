"""Document text -> a validated slide outline, via Gemini.

The renderer only ever sees a validated Outline, so a malformed model response
can never reach python-pptx.

This is the only module in the project that talks to an LLM. To swap providers,
replace `_generate` - nothing else needs to change.
"""
from __future__ import annotations

import json
import os
from dataclasses import dataclass, field
from typing import Any, Literal

from google import genai
from google.genai import errors, types

from .extract import ExtractedDoc

# Stable Flash model: fast and cheap, and the deck is one call. For a more
# capable outline set GEMINI_MODEL=gemini-3.1-pro-preview.
MODEL = os.environ.get("GEMINI_MODEL", "gemini-3.8-flash")

API_KEY_VARS = ("GEMINI_API_KEY", "GOOGLE_API_KEY")

LAYOUTS = ("cover", "section", "bullets", "two_col", "stat", "chart", "closing")
Layout = Literal["cover", "section", "bullets", "two_col", "stat", "chart", "closing"]


class OutlineError(ValueError):
    """Raised when the model cannot produce a usable outline.

    Subclasses ValueError to match the bot's ServiceInputError convention:
    the message is safe to show the user directly.
    """


def api_key_present() -> bool:
    return any(os.environ.get(v) for v in API_KEY_VARS)


# --------------------------------------------------------------------------
# Schema
#
# Gemini's response_json_schema accepts type, properties, required,
# additionalProperties, enum, items, minItems, maxItems and description - which
# is everything this schema uses, so it goes across unchanged.
# --------------------------------------------------------------------------
SLIDE_SCHEMA: dict[str, Any] = {
    "type": "object",
    "additionalProperties": False,
    "required": ["layout", "title"],
    "properties": {
        "layout": {"type": "string", "enum": list(LAYOUTS)},
        "title": {"type": "string"},
        "subtitle": {"type": "string"},
        "bullets": {
            "type": "array",
            "items": {"type": "string"},
            "maxItems": 6,
            "description": "3-6 short points. One idea each, no trailing punctuation.",
        },
        "columns": {
            "type": "array",
            "minItems": 2,
            "maxItems": 2,
            "description": "two_col only. First entry is the RIGHT column (RTL reads right first).",
            "items": {
                "type": "object",
                "additionalProperties": False,
                "required": ["heading", "bullets"],
                "properties": {
                    "heading": {"type": "string"},
                    "bullets": {"type": "array", "items": {"type": "string"}, "maxItems": 4},
                    "governorate": {
                        "type": "string",
                        "description": "Arabic governorate name, if this column is about one.",
                    },
                },
            },
        },
        "stat": {
            "type": "object",
            "additionalProperties": False,
            "required": ["value", "caption"],
            "properties": {
                "value": {"type": "string", "description": "the figure itself, e.g. 64% or 1,250"},
                "caption": {"type": "string"},
            },
        },
        "chart": {
            "type": "object",
            "additionalProperties": False,
            "required": ["kind", "categories", "series"],
            "properties": {
                "kind": {"type": "string", "enum": ["bar", "column", "line", "pie"]},
                "categories": {"type": "array", "items": {"type": "string"}},
                "series": {
                    "type": "array",
                    "maxItems": 6,
                    "items": {
                        "type": "object",
                        "additionalProperties": False,
                        "required": ["name", "values"],
                        "properties": {
                            "name": {"type": "string"},
                            "values": {"type": "array", "items": {"type": "number"}},
                        },
                    },
                },
            },
        },
        "notes": {"type": "string", "description": "Speaker notes, 1-3 sentences."},
    },
}

OUTLINE_SCHEMA: dict[str, Any] = {
    "type": "object",
    "additionalProperties": False,
    "required": ["title", "slides"],
    "properties": {
        "title": {"type": "string"},
        "subtitle": {"type": "string"},
        "slides": {"type": "array", "minItems": 3, "items": SLIDE_SCHEMA},
    },
}

SYSTEM = """\
أنت مصمم عروض تقديمية رسمية للجهات الحكومية السورية.

مهمتك: تحويل مستند إلى مخطط عرض تقديمي بالعربية الفصحى الرسمية.

قواعد إلزامية:
- كل النصوص بالعربية الفصحى، بأسلوب رسمي مختصر ومباشر.
- العناوين قصيرة: ست كلمات كحد أقصى.
- لا تستخدم النقطتين (:) ولا الشرطة (-) في العناوين؛ اكتب العنوان جملة واحدة.
- تجنّب بدء النقطة برقم؛ ضع الرقم داخل الجملة لا في أولها.
- كل نقطة سطر واحد، فكرة واحدة، بدون نقطة في نهايتها.
- من ثلاث إلى ست نقاط في الشريحة الواحدة، لا أكثر.
- لا تخترع أرقاماً أو حقائق غير موجودة في المستند.
- استخدم التخطيط المناسب لكل شريحة:
  cover  : الشريحة الأولى دائماً.
  section: فاصل بين محاور المستند الرئيسية.
  bullets: المحتوى الاعتيادي.
  two_col: عند المقارنة بين طرفين أو محافظتين.
  stat   : عند وجود رقم واحد مهم يستحق شريحة كاملة.
  chart  : عند وجود بيانات جدولية قابلة للرسم.
  closing: الشريحة الأخيرة دائماً.
- إذا ذكر المستند محافظة سورية، ضع اسمها في حقل governorate ليظهر رمزها.
- استخدم stat وchart حين تسمح البيانات؛ العرض المكوّن من نقاط فقط ضعيف.
"""


# --------------------------------------------------------------------------
# Validated result
# --------------------------------------------------------------------------
@dataclass
class Slide:
    layout: Layout
    title: str
    subtitle: str = ""
    bullets: list[str] = field(default_factory=list)
    columns: list[dict] = field(default_factory=list)
    stat: dict | None = None
    chart: dict | None = None
    notes: str = ""


@dataclass
class Outline:
    title: str
    subtitle: str
    slides: list[Slide]


def _validate(data: dict) -> Outline:
    if not isinstance(data, dict) or "slides" not in data:
        raise OutlineError("النموذج لم يُعد مخططاً صالحاً.")

    slides: list[Slide] = []
    for raw in data["slides"]:
        layout = raw.get("layout")
        if layout not in LAYOUTS:
            layout = "bullets"
        s = Slide(
            layout=layout,
            title=str(raw.get("title", "")).strip(),
            subtitle=str(raw.get("subtitle", "")).strip(),
            bullets=[str(b).strip() for b in raw.get("bullets", []) if str(b).strip()],
            columns=raw.get("columns") or [],
            stat=raw.get("stat"),
            chart=raw.get("chart"),
            notes=str(raw.get("notes", "")).strip(),
        )
        # Demote a layout whose payload is missing rather than render an empty slide.
        if s.layout == "stat" and not s.stat:
            s.layout = "bullets"
        if s.layout == "chart" and not (s.chart and s.chart.get("series")):
            s.layout = "bullets"
        if s.layout == "two_col" and len(s.columns) != 2:
            s.layout = "bullets"
        if s.layout == "bullets" and not s.bullets:
            continue
        slides.append(s)

    if not slides:
        raise OutlineError("لم يتم إنشاء أي شرائح من هذا المستند.")

    # Guarantee the deck opens and closes correctly regardless of the model.
    if slides[0].layout != "cover":
        slides.insert(0, Slide(layout="cover",
                               title=str(data.get("title", "عرض تقديمي")),
                               subtitle=str(data.get("subtitle", ""))))
    if slides[-1].layout != "closing":
        slides.append(Slide(layout="closing", title="شكراً لكم"))

    return Outline(
        title=str(data.get("title", "عرض تقديمي")).strip(),
        subtitle=str(data.get("subtitle", "")).strip(),
        slides=slides,
    )


# --------------------------------------------------------------------------
# Provider call
# --------------------------------------------------------------------------
def _generate(client: genai.Client, prompt: str, model: str) -> str:
    """One schema-constrained request. Returns the raw JSON text."""
    config = types.GenerateContentConfig(
        system_instruction=SYSTEM,
        response_mime_type="application/json",
        response_json_schema=OUTLINE_SCHEMA,
        max_output_tokens=16000,
    )
    try:
        resp = client.models.generate_content(model=model, contents=prompt, config=config)
    except errors.ClientError as e:
        raise OutlineError(f"طلب غير صالح إلى Gemini: {e}") from e
    except errors.ServerError as e:
        raise OutlineError(f"خدمة Gemini غير متاحة حالياً: {e}") from e
    except errors.APIError as e:
        raise OutlineError(f"تعذّر الاتصال بـ Gemini: {e}") from e

    feedback = getattr(resp, "prompt_feedback", None)
    if feedback is not None and getattr(feedback, "block_reason", None):
        raise OutlineError("تعذّر إنشاء العرض من هذا المحتوى.")

    candidates = getattr(resp, "candidates", None) or []
    if candidates:
        reason = str(getattr(candidates[0], "finish_reason", "") or "")
        if "MAX_TOKENS" in reason:
            raise OutlineError("المستند طويل جداً؛ جرّب عدد شرائح أقل.")
        if "SAFETY" in reason or "PROHIBITED" in reason or "BLOCKLIST" in reason:
            raise OutlineError("تعذّر إنشاء العرض من هذا المحتوى.")

    text = (resp.text or "").strip()
    if not text:
        raise OutlineError("لم يُعد النموذج أي محتوى.")
    return text


def build_outline(
    doc: ExtractedDoc,
    *,
    slide_count: int = 12,
    organisation: str = "",
    client: genai.Client | None = None,
    model: str | None = None,
) -> Outline:
    """One Gemini call, schema-constrained, with a single repair retry."""
    if client is None:
        if not api_key_present():
            raise OutlineError(
                "مفتاح Gemini غير محدد. عيّن GEMINI_API_KEY في متغيرات البيئة."
            )
        client = genai.Client()

    model = model or MODEL
    hint = doc.outline_hint()
    parts = [
        f"عدد الشرائح المطلوب: {slide_count} تقريباً (بما فيها الغلاف والخاتمة).",
        f"الجهة: {organisation}" if organisation else "",
        f"\nعناوين المستند:\n{hint}" if hint else "",
        f"\nنص المستند:\n{doc.to_prompt_text()}",
    ]
    prompt = "\n".join(p for p in parts if p)

    last: Exception | None = None
    for _ in range(2):
        try:
            return _validate(json.loads(_generate(client, prompt, model)))
        except json.JSONDecodeError as e:
            last = e
        except OutlineError as e:
            # A refusal or an over-long document will not fix itself on retry.
            if "تعذّر إنشاء العرض" in str(e) or "طويل جداً" in str(e):
                raise
            last = e
    raise OutlineError(f"فشل إنشاء المخطط: {last}")
