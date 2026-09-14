"""Voice note -> verbatim Arabic transcript -> formal Arabic, via Gemini.

The two steps are deliberately separate. Transcription is told to be literal
and to leave dialect alone; turning that into formal Arabic is left to
`text_editor.polish_text`, which already owns the language rules. Asking one
call to do both invites it to paraphrase while transcribing and then
paraphrase again while formalising, and the meaning drifts.

Keeping the transcript also makes a mishearing visible: a wrong name comes
back as fluent, correct Arabic, so the user needs the raw text to catch it.

Only `_transcribe_gemini` talks to the provider, so swapping in another
transcriber (a hosted Whisper, say) is a change confined to that function.
"""
from __future__ import annotations

import json
import os
import time
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

from google import genai
from google.genai import errors, types

from .text_editor import (
    RETRY_DELAYS,
    THINKING_BUDGET,
    TextEditError,
    _is_quota_error,
    api_key_present,
    is_mostly_arabic,
    polish_text,
    quota_message,
    tidy,
)

MODEL = os.environ.get("GEMINI_MODEL", "gemini-3.5-flash")

# Telegram caps bot downloads at 20MB, so this is not usually the binding
# limit; it guards the provider call if the module is reused elsewhere.
MAX_AUDIO_MB = int(os.environ.get("MAX_VOICE_MB", "20"))

# Audio costs roughly 32 tokens per second, so ten minutes is still small for
# the model. The real reason to cap it is PROCESSING_TIMEOUT_SECONDS.
MAX_DURATION_SECONDS = int(os.environ.get("MAX_VOICE_SECONDS", "300"))

MIN_TRANSCRIPT_CHARS = 2

# Telegram voice notes are .oga (Ogg/Opus). The rest are here so an uploaded
# audio file works too, since the same handler receives both.
MIME_BY_SUFFIX = {
    ".oga": "audio/ogg",
    ".ogg": "audio/ogg",
    ".opus": "audio/ogg",
    ".mp3": "audio/mp3",
    ".m4a": "audio/aac",
    ".aac": "audio/aac",
    ".wav": "audio/wav",
    ".flac": "audio/flac",
    ".aiff": "audio/aiff",
}

SUPPORTED_SUFFIXES = frozenset(MIME_BY_SUFFIX)


class VoiceNoteError(ValueError):
    """Raised when the audio cannot be turned into text.

    Subclasses ValueError to match the bot's ServiceInputError convention:
    the message is safe to show the user directly.
    """


class VoiceQuotaError(VoiceNoteError):
    """The provider quota is spent; retrying inside one request cannot help."""


@dataclass
class VoiceResult:
    text: str
    transcript: str
    notes: list[str] = field(default_factory=list)


TRANSCRIBE_SCHEMA: dict[str, Any] = {
    "type": "object",
    "additionalProperties": False,
    "required": ["transcript", "has_speech"],
    "properties": {
        "transcript": {
            "type": "string",
            "description": "النص المنطوق حرفياً كما سُمع، بالحروف العربية.",
        },
        "has_speech": {
            "type": "boolean",
            "description": "false إذا كان المقطع صامتاً أو ضجيجاً بلا كلام.",
        },
    },
}

TRANSCRIBE_SYSTEM = """أنت نظام تحويل كلام إلى نص. مهمتك النسخ الحرفي فقط.

انسخ:
- كل ما يُنطق حرفياً بالحروف العربية، كلمة بكلمة.
- اترك العامية كما هي تماماً: «بدي» تبقى «بدي»، و«هلق» تبقى «هلق».
- اكتب الأرقام والأسماء والكلمات الأجنبية كما نُطقت.

لا تفعل:
- لا تصحح لغةً ولا نحواً ولا إملاءً، ولا تحوّل إلى الفصحى.
- لا تلخّص، ولا تشرح، ولا تُضف كلمةً لم تُنطق، ولا تحذف كلمةً نُطقت.
- لا تكرر عبارةً لم تتكرر في الصوت.
- لا تترجم.

إن لم يكن في المقطع كلام مفهوم، اجعل has_speech قيمتها false
واترك transcript فارغاً. لا تخترع كلاماً لتملأ الفراغ."""


def _mime_for(path: Path, mime_type: str | None) -> str:
    if mime_type:
        return mime_type
    suffix = path.suffix.lower()
    if suffix not in MIME_BY_SUFFIX:
        expected = "، ".join(sorted(s.upper().lstrip(".") for s in SUPPORTED_SUFFIXES))
        raise VoiceNoteError(f"صيغة الصوت غير مدعومة. الصيغ المقبولة: {expected}.")
    return MIME_BY_SUFFIX[suffix]


def _check_audio(path: Path, duration: int | None) -> None:
    if not path.is_file():
        raise VoiceNoteError("لم أتمكن من قراءة الملف الصوتي.")
    size = path.stat().st_size
    if size == 0:
        raise VoiceNoteError("الملف الصوتي فارغ.")
    if size > MAX_AUDIO_MB * 1024 * 1024:
        raise VoiceNoteError(
            f"حجم الصوت أكبر من الحد المسموح ({MAX_AUDIO_MB} ميغابايت)."
        )
    if duration is not None and duration > MAX_DURATION_SECONDS:
        limit = max(1, MAX_DURATION_SECONDS // 60)
        raise VoiceNoteError(
            f"التسجيل أطول من {limit} دقائق. أرسله على مقاطع أقصر."
        )


# --------------------------------------------------------------------------
# Provider call
# --------------------------------------------------------------------------
def _transcribe_gemini(
    audio: bytes, mime: str, client: genai.Client | None, model: str
) -> str:
    """One schema-constrained request. Returns the verbatim transcript."""
    if client is None:
        client = genai.Client()
    config = types.GenerateContentConfig(
        system_instruction=TRANSCRIBE_SYSTEM,
        response_mime_type="application/json",
        response_json_schema=TRANSCRIBE_SCHEMA,
        temperature=0.0,
        max_output_tokens=8000,
        thinking_config=types.ThinkingConfig(thinking_budget=THINKING_BUDGET),
    )
    # The instruction goes BEFORE the audio. With the audio first, the model
    # reads the trailing text as more of the recording and echoes it into the
    # transcript, so every result ended with "انسخ هذا التسجيل حرفياً."
    contents = [
        "انسخ هذا التسجيل حرفياً.",
        types.Part.from_bytes(data=audio, mime_type=mime),
    ]
    try:
        resp = client.models.generate_content(
            model=model, contents=contents, config=config
        )
    except errors.ClientError as e:
        if _is_quota_error(e):
            raise VoiceQuotaError(quota_message(e)) from e
        raise VoiceNoteError(f"طلب غير صالح إلى Gemini: {e}") from e
    except errors.ServerError as e:
        raise VoiceNoteError(f"خدمة Gemini غير متاحة حالياً: {e}") from e
    except errors.APIError as e:
        raise VoiceNoteError(f"تعذّر الاتصال بـ Gemini: {e}") from e

    feedback = getattr(resp, "prompt_feedback", None)
    if feedback is not None and getattr(feedback, "block_reason", None):
        raise VoiceNoteError("تعذّر تحويل هذا التسجيل إلى نص.")

    candidates = getattr(resp, "candidates", None) or []
    if candidates:
        reason = str(getattr(candidates[0], "finish_reason", "") or "")
        if "MAX_TOKENS" in reason:
            raise VoiceNoteError("التسجيل طويل جداً. أرسله على مقاطع أقصر.")
        if "SAFETY" in reason or "PROHIBITED" in reason or "BLOCKLIST" in reason:
            raise VoiceNoteError("تعذّر تحويل هذا التسجيل إلى نص.")

    raw = (resp.text or "").strip()
    if not raw:
        raise VoiceNoteError("لم يُعد النموذج أي نص.")
    return _validate_transcript(raw)


def _validate_transcript(raw: str) -> str:
    try:
        data = json.loads(raw)
    except json.JSONDecodeError as e:
        raise VoiceNoteError(f"رد غير مفهوم من النموذج: {e}") from e
    if not isinstance(data, dict):
        raise VoiceNoteError("رد غير متوقع من النموذج.")

    no_speech = "لم أسمع كلاماً واضحاً في التسجيل. أعد التسجيل في مكان أهدأ."
    if data.get("has_speech") is False:
        raise VoiceNoteError(no_speech)

    transcript = data.get("transcript")
    if not isinstance(transcript, str) or not transcript.strip():
        raise VoiceNoteError(no_speech)

    transcript = tidy(transcript, arabic_punctuation=is_mostly_arabic(transcript))
    if len(transcript) < MIN_TRANSCRIPT_CHARS:
        raise VoiceNoteError("التسجيل قصير جداً.")
    return transcript


def transcribe(
    path: str | Path,
    *,
    mime_type: str | None = None,
    duration: int | None = None,
    client: genai.Client | None = None,
    model: str | None = None,
) -> str:
    """Audio file -> verbatim Arabic transcript. One call, with a single retry."""
    path = Path(path)
    _check_audio(path, duration)
    mime = _mime_for(path, mime_type)

    if client is None and not api_key_present():
        raise VoiceNoteError(
            "مفتاح Gemini غير محدد. عيّن GEMINI_API_KEY في متغيرات البيئة."
        )

    audio = path.read_bytes()
    model = model or MODEL

    last: Exception | None = None
    for attempt in range(len(RETRY_DELAYS) + 1):
        try:
            return _transcribe_gemini(audio, mime, client, model)
        except VoiceQuotaError:
            # Nothing to wait out inside one request; the cap is upstream.
            raise
        except VoiceNoteError as e:
            # A refusal, a silent clip or an over-long one will not fix itself.
            text = str(e)
            if "تعذّر تحويل" in text or "طويل جداً" in text or "لم أسمع" in text:
                raise
            last = e
        if attempt < len(RETRY_DELAYS):
            time.sleep(RETRY_DELAYS[attempt])
    raise VoiceNoteError(f"فشل تحويل الصوت إلى نص: {last}")


def voice_to_formal(
    path: str | Path,
    *,
    mime_type: str | None = None,
    duration: int | None = None,
    client: genai.Client | None = None,
    model: str | None = None,
) -> VoiceResult:
    """The whole pipeline: audio -> verbatim transcript -> formal Arabic."""
    transcript = transcribe(
        path, mime_type=mime_type, duration=duration, client=client, model=model
    )
    try:
        edited = polish_text(transcript, client=client, model=model)
    except TextEditError as e:
        # The transcript is the user's words and is worth something on its own,
        # so a failure in the second step says so instead of losing it.
        raise VoiceNoteError(
            f"{e}\n\nالنص المنسوخ كما سُمع:\n{transcript}"
        ) from e
    return VoiceResult(text=edited.text, transcript=transcript, notes=edited.notes)
