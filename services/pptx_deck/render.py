"""Outline -> .pptx, in the Syrian visual identity.

Seven layouts, drawn with absolute geometry from brand.py. Nothing here depends
on a PowerPoint template layout, so the deck renders identically everywhere.

python-pptx note: Length subclasses int, and int arithmetic returns a plain int.
Every computed dimension is re-wrapped in Emu() so `.inches` / `.pt` survive.
"""
from __future__ import annotations

import logging
from datetime import date
from pathlib import Path

from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION, XL_TICK_MARK
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR
from pptx.oxml.ns import qn
from pptx.util import Emu, Inches, Pt

from . import brand as B
from . import rtl
from .embed_fonts import EmbedError, embed
from .outline import Outline, Slide

logger = logging.getLogger(__name__)

CHART_KINDS = {
    "bar": XL_CHART_TYPE.BAR_CLUSTERED,
    "column": XL_CHART_TYPE.COLUMN_CLUSTERED,
    "line": XL_CHART_TYPE.LINE_MARKERS,
    "pie": XL_CHART_TYPE.PIE,
}

RIGHT_EDGE = Emu(B.SLIDE_W - B.MARGIN_X)


# --------------------------------------------------------------------------
# Primitives
# --------------------------------------------------------------------------
def _blank(prs: Presentation, ctx: dict):
    """Add a slide and advance the true slide counter."""
    ctx["n"] = ctx.get("n", 0) + 1
    return prs.slides.add_slide(prs.slide_layouts[6])


def _fill(slide, hex_: str) -> None:
    shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, B.SLIDE_W, B.SLIDE_H)
    shape.fill.solid()
    shape.fill.fore_color.rgb = B.rgb(hex_)
    shape.line.fill.background()
    shape.shadow.inherit = False
    # Send to back so everything drawn afterwards sits on top.
    slide.shapes._spTree.remove(shape._element)
    slide.shapes._spTree.insert(2, shape._element)


def _rect(slide, left, top, width, height, hex_: str):
    shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, width, height)
    shape.fill.solid()
    shape.fill.fore_color.rgb = B.rgb(hex_)
    shape.line.fill.background()
    shape.shadow.inherit = False
    return shape


def _rule(slide, right_x, y, width, hex_: str, thickness=B.RULE_THICK):
    """Horizontal rule anchored at its RIGHT edge (RTL)."""
    return _rect(slide, Emu(int(right_x) - int(width)), y, width,
                 Emu(int(thickness)), hex_)


def _picture(slide, path: Path, left, top, width=None, height=None):
    if not path or not Path(path).exists():
        return None
    return slide.shapes.add_picture(str(path), left, top, width=width, height=height)


def _notes(slide, text: str) -> None:
    if text:
        slide.notes_slide.notes_text_frame.text = text


def _footer(slide, ctx: dict) -> None:
    """Hairline + true slide number bottom-left (mirrored, the deck is RTL)."""
    _rect(slide, B.MARGIN_X, B.HAIRLINE_Y, B.CONTENT_W,
          Emu(int(B.HAIRLINE_THICK)), B.WHEAT_DARK)

    _, tf = rtl.textbox(slide, B.MARGIN_X, B.FOOTER_Y, Inches(1.2), Inches(0.3))
    rtl.write(tf, str(ctx["n"]), size=B.SIZE_FOOTER, color=B.MUTED, align="l")

    org = ctx.get("organisation", "")
    if org:
        _, tf = rtl.textbox(slide, Emu(int(RIGHT_EDGE) - int(Inches(5.0))),
                            B.FOOTER_Y, Inches(5.0), Inches(0.3))
        rtl.write(tf, org, size=B.SIZE_FOOTER, color=B.MUTED, align="r")


def _title_block(slide, text: str) -> None:
    """Slide title + the gold rule under it. Shared by bullets/two_col/stat/chart."""
    size = rtl.fit_size(text, B.CONTENT_W, B.TITLE_H, B.SIZE_SLIDE_TITLE, minimum=Pt(18))
    _, tf = rtl.textbox(slide, B.MARGIN_X, B.MARGIN_TOP, B.CONTENT_W, B.TITLE_H)
    rtl.write(tf, text, family=B.FONT_BOLD, size=size, color=B.HEADING, bold=True)
    _rule(slide, RIGHT_EDGE, B.RULE_Y, B.RULE_W, B.WHEAT_DARK)


def _bullet_list(slide, items: list[str], top, width, right_x, *,
                 size=B.SIZE_BODY, marker=B.ACCENT) -> None:
    """Bullets drawn by hand: PowerPoint's auto-bullets are unreliable in RTL."""
    y = int(top)
    for item in items:
        # marker square, optically centred against the first line
        _rect(slide, Emu(int(right_x) - int(B.MARKER_SIZE)),
              Emu(y + int(int(size) * 0.42)),
              B.MARKER_SIZE, B.MARKER_SIZE, marker)

        text_w = Emu(int(width) - int(B.MARKER_SIZE) - int(B.MARKER_GAP))
        left = int(right_x) - int(B.MARKER_SIZE) - int(B.MARKER_GAP) - int(text_w)
        _, tf = rtl.textbox(slide, Emu(left), Emu(y), text_w, B.LINE_STEP)
        rtl.write(tf, item, size=size, color=B.BODY, line_spacing=1.2)

        lines = max(1, rtl.estimate_lines(item, text_w, size))
        y += int(B.LINE_STEP) * lines


# --------------------------------------------------------------------------
# Chart chrome
#
# python-pptx exposes series colours but not axis direction, and its font API
# writes only <a:latin> - so Arabic axis labels fall back to a system font.
# Both are fixed on the XML directly.
# --------------------------------------------------------------------------
def _reverse_category_axis(ch) -> None:
    """Put the first category on the right (or top, for a bar chart)."""
    cat_ax = ch._chartSpace.find(".//" + qn("c:catAx"))
    if cat_ax is None:
        return  # pie has no category axis
    scaling = cat_ax.find(qn("c:scaling"))
    if scaling is None:
        return
    orientation = scaling.find(qn("c:orientation"))
    if orientation is None:
        orientation = scaling.makeelement(qn("c:orientation"), {})
        scaling.insert(0, orientation)
    orientation.set("val", "maxMin")

    # The value axis stays on the left. Arabic convention would put it on the
    # right, but PowerPoint ignores c:crosses / c:axPos on a reversed category
    # axis - verified against both "min" and "max" - so we do not pretend to
    # set it. Category order is what actually carries the reading direction.


def _chart_typefaces(ch, family: str) -> None:
    """Force every text run in the chart onto an Arabic-capable face."""
    for def_rpr in ch._chartSpace.iter(qn("a:defRPr")):
        for tag in ("a:latin", "a:ea", "a:cs"):
            el = def_rpr.find(qn(tag))
            if el is None:
                el = def_rpr.makeelement(qn(tag), {})
                def_rpr.append(el)
            el.set("typeface", family)


def _style_chart(ch, kind, n_series: int) -> None:
    if n_series > 1 or kind == XL_CHART_TYPE.PIE:
        ch.has_legend = True
        ch.legend.position = XL_LEGEND_POSITION.RIGHT
        ch.legend.include_in_layout = False
        ch.legend.font.size = B.SIZE_CHART_LABEL
        ch.legend.font.color.rgb = B.rgb(B.CHART_LABEL)
    else:
        ch.has_legend = False

    for i, series in enumerate(ch.series):
        colour = B.CHART_SERIES[i % len(B.CHART_SERIES)]
        if kind == XL_CHART_TYPE.LINE_MARKERS:
            series.format.line.color.rgb = B.rgb(colour)
            series.format.line.width = Pt(2.25)
        else:
            series.format.fill.solid()
            series.format.fill.fore_color.rgb = B.rgb(colour)

    if kind == XL_CHART_TYPE.PIE:
        _chart_typefaces(ch, B.FONT_REGULAR)
        return

    _reverse_category_axis(ch)

    cat = ch.category_axis
    cat.tick_labels.font.size = B.SIZE_CHART_LABEL
    cat.tick_labels.font.color.rgb = B.rgb(B.CHART_LABEL)
    cat.format.line.color.rgb = B.rgb(B.CHART_AXIS_LINE)
    cat.format.line.width = Pt(0.75)
    cat.has_major_gridlines = False
    cat.major_tick_mark = XL_TICK_MARK.NONE
    cat.minor_tick_mark = XL_TICK_MARK.NONE

    val = ch.value_axis
    val.tick_labels.font.size = B.SIZE_CHART_LABEL
    val.tick_labels.font.color.rgb = B.rgb(B.CHART_LABEL)
    val.format.line.fill.background()
    val.has_major_gridlines = True
    val.major_gridlines.format.line.color.rgb = B.rgb(B.CHART_GRIDLINE)
    val.major_gridlines.format.line.width = Pt(0.75)
    val.major_tick_mark = XL_TICK_MARK.NONE
    val.minor_tick_mark = XL_TICK_MARK.NONE

    try:
        ch.plots[0].gap_width = 60
    except (AttributeError, ValueError):
        pass

    # Must run last: the calls above create the defRPr elements we retypeface.
    _chart_typefaces(ch, B.FONT_REGULAR)


# --------------------------------------------------------------------------
# Layouts
# --------------------------------------------------------------------------
def cover(prs, s: Slide, ctx: dict) -> None:
    slide = _blank(prs, ctx)
    _fill(slide, B.FOREST)

    _picture(slide, B.emblem("gold"),
             Emu(int(RIGHT_EDGE) - int(B.EMBLEM_SIZE)), B.MARGIN_TOP,
             width=B.EMBLEM_SIZE)

    top = Emu(B.MARGIN_TOP + B.EMBLEM_SIZE + Inches(0.9))
    size = rtl.fit_size(s.title, B.CONTENT_W, Inches(1.8), B.SIZE_COVER_TITLE, minimum=Pt(24))
    _, tf = rtl.textbox(slide, B.MARGIN_X, top, B.CONTENT_W, Inches(1.8))
    rtl.write(tf, s.title, family=B.FONT_BOLD, size=size, color=B.WHITE,
              bold=True, line_spacing=1.15)

    y = Emu(top + Inches(1.75))
    _rule(slide, RIGHT_EDGE, y, Inches(2.6), B.WHEAT_DARK)

    if s.subtitle:
        _, tf = rtl.textbox(slide, B.MARGIN_X, Emu(y + Inches(0.32)),
                            B.CONTENT_W, Inches(0.9))
        rtl.write(tf, s.subtitle, family=B.FONT_LIGHT, size=B.SIZE_COVER_SUB,
                  color=B.WHEAT, line_spacing=1.3)

    footer = " • ".join(x for x in (ctx.get("organisation", ""), ctx["date"]) if x)
    _, tf = rtl.textbox(slide, B.MARGIN_X, Emu(B.SLIDE_H - Inches(0.95)),
                        B.CONTENT_W, Inches(0.4))
    rtl.write(tf, footer, size=B.SIZE_SMALL, color=B.WHEAT)

    _notes(slide, s.notes)


def section(prs, s: Slide, ctx: dict) -> None:
    """Number and title share one horizontal band, so they read as one mark."""
    slide = _blank(prs, ctx)
    _fill(slide, B.WHEAT_LIGHT)

    band_h = Inches(2.2)
    band_y = Emu(int((int(B.SLIDE_H) - int(band_h)) / 2))

    ctx["section_no"] = ctx.get("section_no", 0) + 1
    _, tf = rtl.textbox(slide, B.MARGIN_X, band_y, Inches(3.4), band_h,
                        anchor=MSO_ANCHOR.MIDDLE)
    rtl.write(tf, f"{ctx['section_no']:02d}", family=B.FONT_BOLD,
              size=B.SIZE_WATERMARK, color=B.WHEAT, bold=True, align="l")

    title_w = Emu(int(B.CONTENT_W) - int(Inches(3.8)))
    size = rtl.fit_size(s.title, title_w, band_h, B.SIZE_SECTION_TITLE, minimum=Pt(20))
    _, tf = rtl.textbox(slide, Emu(int(RIGHT_EDGE) - int(title_w)), band_y,
                        title_w, band_h, anchor=MSO_ANCHOR.MIDDLE)
    rtl.write(tf, s.title, family=B.FONT_BOLD, size=size, color=B.FOREST,
              bold=True, line_spacing=1.2)

    _rule(slide, RIGHT_EDGE, Emu(int(band_y) + int(band_h) + int(Inches(0.15))),
          Inches(3.2), B.WHEAT_DARK)
    _notes(slide, s.notes)


def bullets(prs, s: Slide, ctx: dict) -> None:
    pages = rtl.split_to_fit(s.bullets, B.CONTENT_W, B.BODY_H, B.SIZE_BODY,
                             step=B.LINE_STEP)
    for i, page in enumerate(pages):
        slide = _blank(prs, ctx)
        _fill(slide, B.SURFACE)
        _title_block(slide, s.title + (rtl.CONTINUED if i else ""))

        height = rtl.block_height(page, B.CONTENT_W, B.SIZE_BODY, B.LINE_STEP)
        top = rtl.centre_top(B.BODY_TOP, B.BODY_H, height)
        _bullet_list(slide, page, top, B.CONTENT_W, RIGHT_EDGE)

        _footer(slide, ctx)
        if i == 0:
            _notes(slide, s.notes)


def two_col(prs, s: Slide, ctx: dict) -> None:
    slide = _blank(prs, ctx)
    _fill(slide, B.SURFACE)
    _title_block(slide, s.title)

    # First column renders on the RIGHT: that is where an RTL reader starts.
    rights = [
        RIGHT_EDGE,
        Emu(int(RIGHT_EDGE) - int(B.COLUMN_W) - int(B.GUTTER)),
    ]
    has_icon = any(c.get("governorate") for c in s.columns)
    head_h = int(Inches(0.6))
    icon_h = int(B.ICON_SIZE) + int(Inches(0.25)) if has_icon else 0

    tallest = max(
        (int(rtl.block_height(c.get("bullets", []), B.COLUMN_W, Pt(16), B.LINE_STEP))
         for c in s.columns),
        default=0,
    )
    top = int(rtl.centre_top(B.BODY_TOP, B.BODY_H, Emu(icon_h + head_h + tallest)))

    for col, right_x in zip(s.columns, rights):
        y = top

        if has_icon:
            icon = B.governorate_icon(col.get("governorate", "")) or None
            if icon:
                # These stamps are portrait, so their rendered width is
                # narrower than ICON_SIZE. Place it, then right-align on the
                # width PowerPoint actually gave it.
                pic = _picture(slide, icon, Emu(int(right_x)), Emu(y),
                               height=B.ICON_SIZE)
                if pic is not None:
                    pic.left = Emu(int(right_x) - int(pic.width))
            y += icon_h

        _, tf = rtl.textbox(slide, Emu(int(right_x) - int(B.COLUMN_W)), Emu(y),
                            B.COLUMN_W, Inches(0.55))
        rtl.write(tf, col.get("heading", ""), family=B.FONT_MEDIUM,
                  size=Pt(20), color=B.FOREST_LIGHT, bold=True)
        y += head_h

        _bullet_list(slide, col.get("bullets", []), Emu(y), B.COLUMN_W, right_x,
                     size=Pt(16), marker=B.WHEAT_DARK)

    _footer(slide, ctx)
    _notes(slide, s.notes)


def stat(prs, s: Slide, ctx: dict) -> None:
    slide = _blank(prs, ctx)
    _fill(slide, B.SURFACE)
    _title_block(slide, s.title)

    value = str(s.stat.get("value", ""))
    caption = str(s.stat.get("caption", ""))

    figure_h = Inches(1.7)
    caption_h = Inches(1.0)
    block = Emu(int(figure_h) + int(Inches(0.45)) + int(caption_h))
    top = int(rtl.centre_top(B.BODY_TOP, B.BODY_H, block))

    size = rtl.fit_size(value, B.CONTENT_W, figure_h, B.SIZE_STAT, minimum=Pt(36))
    _, tf = rtl.textbox(slide, B.MARGIN_X, Emu(top), B.CONTENT_W, figure_h,
                        anchor=MSO_ANCHOR.MIDDLE)
    rtl.write(tf, value, family=B.FONT_BOLD, size=size, color=B.FOREST_LIGHT, bold=True)

    rule_y = top + int(figure_h) + int(Inches(0.18))
    _rule(slide, RIGHT_EDGE, Emu(rule_y), Inches(1.6), B.WHEAT_DARK)

    _, tf = rtl.textbox(slide, B.MARGIN_X, Emu(rule_y + int(Inches(0.27))),
                        B.CONTENT_W, caption_h)
    rtl.write(tf, caption, family=B.FONT_LIGHT, size=Pt(16), color=B.MUTED,
              line_spacing=1.35)

    _footer(slide, ctx)
    _notes(slide, s.notes)


def chart(prs, s: Slide, ctx: dict) -> None:
    slide = _blank(prs, ctx)
    _fill(slide, B.SURFACE)
    _title_block(slide, s.title)

    spec = s.chart
    categories = spec.get("categories", [])
    data = CategoryChartData()
    data.categories = categories
    for series in spec.get("series", []):
        values = list(series.get("values", []))[:len(categories)]
        values += [None] * (len(categories) - len(values))
        data.add_series(series.get("name", ""), values)

    kind = CHART_KINDS.get(spec.get("kind", "column"), XL_CHART_TYPE.COLUMN_CLUSTERED)
    frame = slide.shapes.add_chart(
        kind, B.MARGIN_X, B.BODY_TOP, B.CONTENT_W,
        Emu(int(B.BODY_H) - int(Inches(0.2))), data,
    )
    ch = frame.chart
    ch.has_title = False
    _style_chart(ch, kind, len(spec.get("series", [])))

    _footer(slide, ctx)
    _notes(slide, s.notes)


def closing(prs, s: Slide, ctx: dict) -> None:
    slide = _blank(prs, ctx)
    _fill(slide, B.FOREST)

    emblem = Emu(int(B.EMBLEM_SIZE * 1.4))
    title_h = int(Inches(0.9))
    org_h = int(Inches(0.5)) if ctx.get("organisation") else 0
    block = int(emblem) + int(Inches(0.55)) + title_h + org_h
    top = int((int(B.SLIDE_H) - block) / 2)

    _picture(slide, B.emblem("gold"),
             Emu(int((int(B.SLIDE_W) - int(emblem)) / 2)), Emu(top), width=emblem)

    y = top + int(emblem) + int(Inches(0.55))
    _, tf = rtl.textbox(slide, B.MARGIN_X, Emu(y), B.CONTENT_W, Emu(title_h))
    rtl.write(tf, s.title or "شكراً لكم", family=B.FONT_BOLD,
              size=B.SIZE_SECTION_TITLE, color=B.WHITE, bold=True, align="ctr")

    if org_h:
        _, tf = rtl.textbox(slide, B.MARGIN_X, Emu(y + title_h), B.CONTENT_W, Emu(org_h))
        rtl.write(tf, ctx["organisation"], family=B.FONT_LIGHT,
                  size=B.SIZE_SMALL, color=B.WHEAT, align="ctr")

    _notes(slide, s.notes)


RENDERERS = {
    "cover": cover,
    "section": section,
    "bullets": bullets,
    "two_col": two_col,
    "stat": stat,
    "chart": chart,
    "closing": closing,
}


# --------------------------------------------------------------------------
# Entry point
# --------------------------------------------------------------------------
def render(
    outline: Outline,
    out_path: str | Path,
    *,
    organisation: str = "",
    template: str | Path | None = None,
    embed_fonts: bool = True,
) -> Path:
    """Render an outline to a .pptx file.

    With `embed_fonts` the identity fonts are written into the package, so the
    deck renders correctly on a machine that does not have them installed.
    """
    prs = Presentation(str(template)) if template else Presentation()
    prs.slide_width = B.SLIDE_W
    prs.slide_height = B.SLIDE_H

    ctx = {
        "organisation": organisation,
        "date": date.today().strftime("%Y/%m/%d"),
        "n": 0,
    }

    for s in outline.slides:
        RENDERERS[s.layout](prs, s, ctx)

    out_path = Path(out_path)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    prs.save(str(out_path))

    if embed_fonts:
        try:
            embed(out_path)
        except EmbedError as e:
            # A deck without embedded fonts is still usable where they are
            # installed - never fail the whole job over this.
            logger.warning("font embedding skipped: %s", e)

    return out_path
