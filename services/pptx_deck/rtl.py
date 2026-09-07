"""Right-to-left and Arabic-font plumbing that python-pptx does not expose.

Two things python-pptx will not do for you:

1. `run.font.name` writes only <a:latin>. Arabic glyphs resolve from the
   *complex script* face, <a:cs>, so an Arabic run keeps rendering in the
   theme font until you set that element yourself.
2. There is no paragraph-direction API. RTL paragraphs need rtl="1" and an
   explicit algn on <a:pPr>.
"""
from __future__ import annotations

from pptx.dml.color import RGBColor
from pptx.enum.text import MSO_ANCHOR
from pptx.oxml.ns import qn
from pptx.util import Emu, Length, Pt

import re as _re

from . import brand


def _set_typeface(rPr, tag: str, family: str) -> None:
    el = rPr.find(qn(tag))
    if el is None:
        el = rPr.makeelement(qn(tag), {})
        rPr.append(el)
    el.set("typeface", family)


def style_run(
    run,
    *,
    family: str = brand.FONT_REGULAR,
    size: Length | None = None,
    color: str | None = None,
    bold: bool = False,
) -> None:
    """Apply an Arabic-safe font to one run.

    Sets <a:latin>, <a:ea> and <a:cs> to the same family so mixed
    Arabic/Latin/digits in a single run stay visually consistent.
    """
    run.font.size = size
    run.font.bold = bold
    if color is not None:
        run.font.color.rgb = brand.rgb(color)
    rPr = run._r.get_or_add_rPr()
    for tag in ("a:latin", "a:ea", "a:cs"):
        _set_typeface(rPr, tag, family)


def rtl_paragraph(paragraph, align: str = "r") -> None:
    """Mark a paragraph right-to-left. align: r | l | ctr | just."""
    pPr = paragraph._p.get_or_add_pPr()
    pPr.set("rtl", "1")
    pPr.set("algn", align)


def set_line_spacing(paragraph, multiple: float) -> None:
    paragraph.line_spacing = multiple


def textbox(
    slide,
    left: Length,
    top: Length,
    width: Length,
    height: Length,
    *,
    anchor=MSO_ANCHOR.TOP,
    word_wrap: bool = True,
):
    box = slide.shapes.add_textbox(left, top, width, height)
    tf = box.text_frame
    tf.word_wrap = word_wrap
    tf.vertical_anchor = anchor
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    return box, tf


# A Western digit run sitting inside Arabic text loses its trailing space
# visually (bidi puts the neutral space inside the LTR run). A right-to-left
# mark after the digits restores the gap without changing the reading order.
RLM = "‏"
_ARABIC = _re.compile(r"[؀-ۿ]")
_DIGIT_RUN = _re.compile(r"(\d[\d.,:/%٫٬]*)")


def fix_bidi_digits(text: str) -> str:
    if not text or not _ARABIC.search(text):
        return text
    return _DIGIT_RUN.sub(lambda m: m.group(1) + RLM, text)


def write(
    tf,
    lines: list[str] | str,
    *,
    family: str = brand.FONT_REGULAR,
    size: Length = brand.SIZE_BODY,
    color: str = brand.BODY,
    bold: bool = False,
    align: str = "r",
    line_spacing: float | None = None,
    space_after: Length | None = None,
) -> None:
    """Fill a text frame with RTL paragraphs, one per line."""
    if isinstance(lines, str):
        lines = [lines]
    tf.clear()
    for i, line in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        run = p.add_run()
        run.text = fix_bidi_digits(line)
        style_run(run, family=family, size=size, color=color, bold=bold)
        rtl_paragraph(p, align)
        if line_spacing is not None:
            p.line_spacing = line_spacing
        if space_after is not None:
            p.space_after = space_after


# --------------------------------------------------------------------------
# Text fitting
#
# python-pptx cannot autofit: PowerPoint computes autofit at render time and
# python-pptx does not run PowerPoint. So we estimate and shrink ourselves.
# --------------------------------------------------------------------------

# Mean advance width as a fraction of font size, measured over Arabic text set
# in Hayyakum Allah. Arabic is narrower per character than Latin.
_AR_ADVANCE = 0.46
_LATIN_ADVANCE = 0.52


def _advance(text: str) -> float:
    arabic = sum(1 for c in text if "؀" <= c <= "ۿ")
    if not text:
        return _LATIN_ADVANCE
    frac = arabic / len(text)
    return _AR_ADVANCE * frac + _LATIN_ADVANCE * (1 - frac)


def estimate_lines(text: str, width: Length, size: Length) -> int:
    """How many wrapped lines `text` needs at `size` inside `width`."""
    if not text:
        return 1
    chars_per_line = max(1, int(Emu(width).inches * 72 / (size.pt * _advance(text))))
    # wrap on words, not mid-word
    lines, cur = 1, 0
    for word in text.split():
        need = len(word) + (1 if cur else 0)
        if cur + need > chars_per_line and cur:
            lines += 1
            cur = len(word)
        else:
            cur += need
    return lines


def fit_size(
    text: str,
    width: Length,
    height: Length,
    start: Length,
    *,
    minimum: Length = Pt(12),
    step: Length = Pt(1),
    leading: float = 1.25,
) -> Length:
    """Largest size <= `start` at which `text` fits `width` x `height`."""
    size = start
    while size > minimum:
        needed = estimate_lines(text, width, size) * size.pt * leading
        if needed <= Emu(height).inches * 72:
            return size
        size = Pt(size.pt - step.pt)
    return minimum


def split_to_fit(
    items: list[str],
    width: Length,
    height: Length,
    size: Length,
    *,
    step: Length,
    leading: float = 1.25,
) -> list[list[str]]:
    """Chunk bullet items into groups that each fit one slide body.

    Returns one list per slide. Callers title continuation slides "... (تابع)".
    """
    budget = int(height) // int(step)
    if budget < 1:
        budget = 1
    pages: list[list[str]] = []
    page: list[str] = []
    used = 0
    for item in items:
        cost = estimate_lines(item, width, size)
        if page and used + cost > budget:
            pages.append(page)
            page, used = [], 0
        page.append(item)
        used += cost
    if page:
        pages.append(page)
    return pages


CONTINUED = " (تابع)"


def centre_top(top: Length, available: Length, content: Length) -> Length:
    """Top edge that optically centres `content` inside `available`.

    Never rises above `top`, so a block that overfills still starts at the
    top of the body area rather than being pushed off the slide.
    """
    from pptx.util import Emu
    slack = int(available) - int(content)
    if slack <= 0:
        return Emu(int(top))
    return Emu(int(top) + slack // 2)


def block_height(items: list[str], width: Length, size: Length, step: Length) -> Length:
    """Rendered height of a bullet block, counting wrapped lines."""
    from pptx.util import Emu
    total = sum(max(1, estimate_lines(i, width, size)) for i in items)
    return Emu(total * int(step))
