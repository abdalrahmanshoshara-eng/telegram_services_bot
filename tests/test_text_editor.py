import json

import pytest

import services.text_editor as text_editor
from services.text_editor import (
    MAX_INPUT_CHARS,
    EditResult,
    TextEditError,
    _validate,
    build_prompt,
    is_mostly_arabic,
    polish_text,
    tidy,
)


@pytest.fixture(autouse=True)
def no_retry_backoff(monkeypatch):
    """Keep the retry path exercised without the real seconds-long waits."""
    monkeypatch.setattr(text_editor, "RETRY_DELAYS", (0, 0))


# --------------------------------------------------------------------------
# tidy
# --------------------------------------------------------------------------
def test_tidy_collapses_spaces_and_removes_tatweel():
    assert tidy("هذا   نص  فيه ـــتطويل") == "هذا نص فيه تطويل"


def test_tidy_fixes_space_before_punctuation():
    assert tidy("مرحبا ، كيف حالك ؟") == "مرحبا، كيف حالك؟"


def test_tidy_adds_space_after_punctuation():
    assert tidy("مرحبا،كيف حالك") == "مرحبا، كيف حالك"


def test_tidy_keeps_punctuation_glued_to_a_closing_quote():
    """Regression: a full stop inside a quotation grew a space before the quote."""
    assert tidy('قال: "انتهى الوقت."') == 'قال: "انتهى الوقت."'
    assert tidy("قال: (انتهى.) ثم مشى") == "قال: (انتهى.) ثم مشى"
    assert tidy("«لا أحد يهتم!»") == "«لا أحد يهتم!»"


def test_tidy_still_spaces_punctuation_before_a_word():
    assert tidy("انتهى.ثم مشى") == "انتهى. ثم مشى"


def test_tidy_collapses_repeated_punctuation():
    assert tidy("ماذا؟؟؟") == "ماذا؟"
    assert tidy("انتظر....") == "انتظر..."


def test_tidy_converts_latin_punctuation_in_arabic_text():
    assert tidy("نعم, وأيضا; ثم ماذا?") == "نعم، وأيضا؛ ثم ماذا؟"


def test_tidy_keeps_latin_punctuation_in_english_text():
    assert tidy("yes, and also; then what?") == "yes, and also; then what?"


def test_tidy_spares_digit_separators():
    assert "1,000" in tidy("المبلغ 1,000 ليرة")
    assert "12;30" in tidy("الوقت 12;30 صباحا")


def test_tidy_spares_urls_and_emails():
    text = "زر https://a.com/x?y=1,2 أو راسلنا info@mei.gov.sy"
    result = tidy(text)
    assert "https://a.com/x?y=1,2" in result
    assert "info@mei.gov.sy" in result


def test_tidy_strips_invisible_characters():
    assert tidy("نص​مع﻿رموز") == "نصمعرموز"


def test_tidy_limits_blank_lines_to_one():
    assert tidy("فقرة\n\n\n\nفقرة أخرى") == "فقرة\n\nفقرة أخرى"


def test_tidy_handles_empty_input():
    assert tidy("") == ""
    assert tidy("   ") == ""


def test_is_mostly_arabic():
    assert is_mostly_arabic("نص عربي")
    assert not is_mostly_arabic("english text")


# --------------------------------------------------------------------------
# validation
# --------------------------------------------------------------------------
def test_validate_returns_tidied_text_and_notes():
    result = _validate(
        {"corrected_text": "مرحبا ، كيف حالك ؟", "notes": ["  فاصلة   زائدة "]},
        "مرحبا, كيف حالك?",
    )
    assert isinstance(result, EditResult)
    assert result.text == "مرحبا، كيف حالك؟"
    assert result.notes == ["فاصلة زائدة"]


def test_validate_rejects_missing_text():
    with pytest.raises(TextEditError):
        _validate({"notes": []}, "نص")
    with pytest.raises(TextEditError):
        _validate({"corrected_text": "   "}, "نص")
    with pytest.raises(TextEditError):
        _validate("not a dict", "نص")


def test_validate_rejects_summarised_answer():
    original = "كلمة " * 100
    with pytest.raises(TextEditError, match="ناقص"):
        _validate({"corrected_text": "كلمة واحدة"}, original)


def test_validate_allows_short_answer_for_short_input():
    result = _validate({"corrected_text": "نعم"}, "نعم")
    assert result.text == "نعم"


def test_validate_caps_notes():
    result = _validate(
        {"corrected_text": "نص", "notes": [f"ملاحظة {i}" for i in range(20)]},
        "نص",
    )
    assert len(result.notes) <= 6


def test_validate_ignores_malformed_notes():
    result = _validate({"corrected_text": "نص", "notes": ["صحيح", "", 5, None]}, "نص")
    assert result.notes == ["صحيح"]


def test_build_prompt_includes_text():
    assert "نصي هنا" in build_prompt("نصي هنا")


# --------------------------------------------------------------------------
# polish_text, with a stub client instead of the provider
# --------------------------------------------------------------------------
class _Response:
    def __init__(self, text):
        self.text = text
        self.candidates = []
        self.prompt_feedback = None


class _Models:
    def __init__(self, payloads):
        self._payloads = list(payloads)
        self.calls = 0

    def generate_content(self, *, model, contents, config):
        self.calls += 1
        return _Response(self._payloads.pop(0))


class _Client:
    def __init__(self, *payloads):
        self.models = _Models(payloads)


def test_polish_text_returns_corrected_text():
    client = _Client(json.dumps({"corrected_text": "مرحبا، كيف حالك؟"}))
    result = polish_text("مرحبا ,كيف حالك ?", client=client)
    assert result.text == "مرحبا، كيف حالك؟"
    assert client.models.calls == 1


def test_polish_text_retries_once_on_invalid_json():
    client = _Client("not json", json.dumps({"corrected_text": "نص سليم"}))
    result = polish_text("نص", client=client)
    assert result.text == "نص سليم"
    assert client.models.calls == 2


def test_polish_text_gives_up_after_exhausting_retries():
    attempts = len(text_editor.RETRY_DELAYS) + 1
    client = _Client(*(["not json"] * attempts))
    with pytest.raises(TextEditError):
        polish_text("نص", client=client)
    assert client.models.calls == attempts


def test_polish_text_rejects_empty_input():
    with pytest.raises(TextEditError, match="أرسل نصاً"):
        polish_text("   ", client=_Client())


def test_polish_text_rejects_over_long_input():
    with pytest.raises(TextEditError, match="أطول"):
        polish_text("ا" * (MAX_INPUT_CHARS + 1), client=_Client())


def test_polish_text_does_not_call_provider_for_invalid_input():
    client = _Client()
    with pytest.raises(TextEditError):
        polish_text("", client=client)
    assert client.models.calls == 0
