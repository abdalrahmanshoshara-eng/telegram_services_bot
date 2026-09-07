"""Syrian visual identity (SYID) design tokens and deck geometry.

Every colour, font name and flag proportion here is taken from the official kit
published at https://syrian.zone/syid (mirror of syrianidentity.sy).
"""
from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path

from pptx.dml.color import RGBColor
from pptx.util import Emu, Inches, Length, Pt

# services/pptx_deck/brand.py -> repo root is three levels up.
ASSETS = Path(__file__).resolve().parent.parent.parent / "assets" / "syid"


def rgb(hex_: str) -> RGBColor:
    return RGBColor.from_string(hex_.lstrip("#").upper())


# --------------------------------------------------------------------------
# Palette - the four official families
# --------------------------------------------------------------------------
FOREST_LIGHT = "#428177"
FOREST = "#054239"
FOREST_DARK = "#002623"

WHEAT_LIGHT = "#edebe0"
WHEAT = "#b9a779"
WHEAT_DARK = "#988561"

UMBER = "#6b1f2a"
UMBER_DARK = "#4a151e"
UMBER_DARKEST = "#260f14"

WHITE = "#ffffff"
GREY = "#3d3a3b"
INK = "#161616"

# Flag colours. These belong to the flag ONLY - never use them as accent,
# chart or UI colours.
FLAG_GREEN = "#007a3d"
FLAG_BLACK = "#161616"
FLAG_WHITE = "#ffffff"
FLAG_RED = "#ce1126"
FLAG_RATIO = 3 / 2          # width : height
FLAG_STAR_POSITIONS = (0.25, 0.50, 0.75)   # fraction of width, centre of each star

# --------------------------------------------------------------------------
# Semantic roles
# --------------------------------------------------------------------------
HEADING = FOREST           # slide titles, cover ground
ACCENT = FOREST_LIGHT      # rules, bullet markers, primary chart series
GOLD = WHEAT_DARK          # emblem, hairlines, dividers - sparingly
BODY = INK
MUTED = GREY
SURFACE = WHITE
SECTION_GROUND = WHEAT_LIGHT
NEGATIVE = UMBER           # reserved: deficits, warnings, "before" figures

CHART_SERIES = [FOREST, FOREST_LIGHT, WHEAT_DARK, WHEAT, GREY, UMBER]

# Chart chrome: everything that is not a data mark stays quiet.
CHART_GRIDLINE = WHEAT_LIGHT
CHART_AXIS_LINE = WHEAT
CHART_LABEL = GREY

# --------------------------------------------------------------------------
# Typography
#
# Qomra is the official identity typeface but is commercially licensed
# (iwantype.com/product/qomra). Hayyakum Allah is the free near-equivalent
# published alongside it and is what we render with.
#
# NOTE: the four weights are NOT one family. Regular and Bold share the family
# name "Hayyakum Allah"; Light and Medium register as their own families. These
# strings are read from the TTF name tables and must be used verbatim.
# --------------------------------------------------------------------------
FONT_LIGHT = "Hayyakum Allah Light"
FONT_REGULAR = "Hayyakum Allah"
FONT_MEDIUM = "Hayyakum Allah Medium"
FONT_BOLD = "Hayyakum Allah"       # with bold=True
FONT_FALLBACK = "IBM Plex Sans Arabic"

SIZE_COVER_TITLE = Pt(40)
SIZE_COVER_SUB = Pt(18)
SIZE_SECTION_TITLE = Pt(32)
SIZE_SLIDE_TITLE = Pt(28)
SIZE_BODY = Pt(18)
SIZE_SMALL = Pt(14)
SIZE_CAPTION = Pt(12)
SIZE_STAT = Pt(72)
SIZE_FOOTER = Pt(10)
SIZE_CHART_LABEL = Pt(12)
SIZE_WATERMARK = Pt(150)

# --------------------------------------------------------------------------
# Geometry - 16:9
# --------------------------------------------------------------------------
SLIDE_W = Inches(13.333)
SLIDE_H = Inches(7.5)

MARGIN_X = Inches(0.9)
MARGIN_TOP = Inches(0.7)
MARGIN_BOTTOM = Inches(0.55)
GUTTER = Inches(0.25)

# NOTE: Length subclasses int, and int arithmetic returns int - not Length.
# Every computed dimension is re-wrapped in Emu() so `.inches` / `.pt` survive.
CONTENT_W = Emu(SLIDE_W - 2 * MARGIN_X)             # 11.533"
COLUMN_W = Emu(int((CONTENT_W - GUTTER) / 2))       # for two_col

TITLE_H = Inches(0.85)
RULE_Y = Emu(MARGIN_TOP + TITLE_H + Inches(0.10))
RULE_W = Inches(2.2)
RULE_THICK = Pt(2.5)

BODY_TOP = Emu(RULE_Y + Inches(0.45))
BODY_H = Emu(SLIDE_H - BODY_TOP - MARGIN_BOTTOM - Inches(0.35))

LINE_STEP = Inches(0.55)                    # bullet-to-bullet leading
MARKER_SIZE = Inches(0.085)                 # gold/accent square before a bullet
MARKER_GAP = Inches(0.22)

FOOTER_Y = Emu(SLIDE_H - Inches(0.55))
HAIRLINE_Y = Emu(SLIDE_H - Inches(0.62))
HAIRLINE_THICK = Pt(0.75)

EMBLEM_SIZE = Inches(1.1)                   # cover / closing eagle
ICON_SIZE = Inches(1.8)                     # governorate icon in a column head


# --------------------------------------------------------------------------
# Assets
# --------------------------------------------------------------------------
@dataclass(frozen=True)
class Governorate:
    name_ar: str
    landmark_ar: str
    slug: str


GOVERNORATES: tuple[Governorate, ...] = (
    Governorate("دمشق", "السيف الدمشقي", "damascus-sword"),
    Governorate("ريف دمشق", "غوطة ريف دمشق", "rif-dimashq-ghouta"),
    Governorate("حلب", "قلعة حلب", "aleppo-citadel"),
    Governorate("حمص", "ساعة حمص", "homs-clock"),
    Governorate("حماة", "نواعير حماة", "hama-norias"),
    Governorate("اللاذقية", "قوس النصر", "latakia-triumph-arch"),
    Governorate("طرطوس", "جزيرة أرواد", "tartous-arwad"),
    Governorate("إدلب", "رويحة إدلب", "idlib-ruweiha"),
    Governorate("دير الزور", "الجسر المعلق", "deir-ez-zor-bridge"),
    Governorate("الرقة", "بوابة بغداد", "raqqa-baghdad-gate"),
    Governorate("الحسكة", "جسر عين ديوار", "hasakah-ain-diwar"),
    Governorate("درعا", "المسجد العمري", "daraa-omari-mosque"),
    Governorate("السويداء", "قنوات السويداء", "sweida-qanawat"),
    Governorate("القنيطرة", "بيت صيدا", "quneitra-beit-saida"),
)

_BY_NAME = {g.name_ar: g for g in GOVERNORATES}


def governorate_icon(name_ar: str, tint: str = "forest") -> Path | None:
    """Path to a rasterised governorate icon, or None if the name isn't one.

    `tint` is "forest" or "gold". Matches on the Arabic governorate name as it
    appears in the identity kit, tolerating a leading "ال".
    """
    g = _BY_NAME.get(name_ar.strip())
    if g is None:
        stripped = name_ar.strip().removeprefix("ال")
        g = next((v for k, v in _BY_NAME.items() if k.removeprefix("ال") == stripped), None)
    if g is None:
        return None
    p = ASSETS / "icons" / f"{g.slug}-{tint}.png"
    return p if p.exists() else None


def emblem(variant: str = "gold") -> Path:
    """Eagle emblem PNG. variant: gold | white | forest."""
    return ASSETS / "logo" / f"eagle-{variant}.png"


def wordmark(on_dark: bool = False) -> Path:
    return ASSETS / "logo" / ("whiteonblack.png" if on_dark else "blackonwhite.png")
