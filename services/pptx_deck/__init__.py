"""PDF/DOCX -> presentation in the Syrian visual identity.

Two entry points, both synchronous so bot.py can hand them to run_blocking:

    outline = build_deck_outline(source, slide_count, organisation)
    outline_preview(outline)                      # text shown for approval
    render_outline(outline, destination, organisation)

`create_presentation` runs both in one call when no preview is wanted.

Identity assets: https://syrian.zone/syid
"""
from __future__ import annotations

from pathlib import Path

from .extract import ExtractionError, ScannedDocumentError, extract
from .outline import (
    MODEL,
    Outline,
    OutlineError,
    Slide,
    api_key_present,
    build_outline,
)
from .render import render

__all__ = [
    "build_deck_outline",
    "render_outline",
    "outline_preview",
    "create_presentation",
    "api_key_present",
    "outline_to_dict",
    "outline_from_dict",
    "DeckError",
    "ExtractionError",
    "ScannedDocumentError",
    "OutlineError",
    "Outline",
    "Slide",
    "MODEL",
]

# Both error types already subclass ValueError; this is the single name bot.py
# can catch for anything this service raises.
DeckError = ValueError

SLIDE_COUNT_CHOICES = (8, 12, 20)
DEFAULT_SLIDE_COUNT = 12

_LAYOUT_LABELS = {
    "cover": "غلاف",
    "section": "فاصل",
    "bullets": "نقاط",
    "two_col": "عمودان",
    "stat": "رقم",
    "chart": "رسم بياني",
    "closing": "خاتمة",
}


def build_deck_outline(
    source: str | Path,
    slide_count: int = DEFAULT_SLIDE_COUNT,
    organisation: str = "",
) -> Outline:
    """Extract the document and ask the model for a slide outline."""
    doc = extract(source)
    return build_outline(doc, slide_count=slide_count, organisation=organisation)


def render_outline(
    outline: Outline,
    destination: str | Path,
    organisation: str = "",
) -> None:
    """Build the .pptx from an already-approved outline."""
    render(outline, destination, organisation=organisation)


def create_presentation(
    source: str | Path,
    destination: str | Path,
    slide_count: int = DEFAULT_SLIDE_COUNT,
    organisation: str = "",
) -> None:
    """Full pipeline in one call, for use without the approval step."""
    outline = build_deck_outline(source, slide_count, organisation)
    render_outline(outline, destination, organisation)


def outline_preview(outline: Outline, limit: int = 25) -> str:
    """Human-readable summary of an outline, for the approval message."""
    lines = [f"📋 {outline.title}"]
    if outline.subtitle:
        lines.append(outline.subtitle)
    lines.append("")
    for i, slide in enumerate(outline.slides[:limit], start=1):
        label = _LAYOUT_LABELS.get(slide.layout, slide.layout)
        lines.append(f"{i}. [{label}] {slide.title}".rstrip())
        for bullet in slide.bullets[:3]:
            lines.append(f"    • {bullet}")
    remaining = len(outline.slides) - limit
    if remaining > 0:
        lines.append(f"… و{remaining} شريحة أخرى")
    lines.append("")
    lines.append(f"المجموع: {len(outline.slides)} شريحة")
    return "\n".join(lines)


# --------------------------------------------------------------------------
# Serialisation
#
# An approved outline has to survive between two Telegram updates. Storing the
# dataclass in user_data would work in-process, but plain dicts keep the job
# state JSON-serialisable and safe to log.
# --------------------------------------------------------------------------
def outline_to_dict(outline: Outline) -> dict:
    return {
        "title": outline.title,
        "subtitle": outline.subtitle,
        "slides": [vars(s) for s in outline.slides],
    }


def outline_from_dict(data: dict) -> Outline:
    return Outline(
        title=data.get("title", ""),
        subtitle=data.get("subtitle", ""),
        slides=[Slide(**s) for s in data.get("slides", [])],
    )
