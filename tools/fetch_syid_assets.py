"""Download the Syrian visual-identity kit from syrian.zone/syid and prepare it for python-pptx.

Run once:  python tools/fetch_assets.py

SVG is rasterised to PNG because python-pptx cannot place SVG. PyMuPDF does the
rasterising, so this adds no dependency beyond what extract.py already needs.
"""
from __future__ import annotations

import io
import re
import sys
import urllib.request
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

import pymupdf

from svg_inline import inline_styles

BASE = "https://syrian.zone/syid-assets"
UA = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/128.0 Safari/537.36"

ROOT = Path(__file__).resolve().parent.parent
ASSETS = ROOT / "assets" / "syid"

FONT_WEIGHTS = ["Light", "Regular", "Medium", "Bold"]

MATERIALS = [
    "materials/logo.ai.svg",
    "materials/blackonwhite.png",
    "materials/whiteonblack.png",
    "materials/syrian-flag-proportions.svg",
]

# governorate -> landmark, as published on syrian.zone/syid
GOVERNORATES = {
    "damascus-sword.svg": ("دمشق", "السيف الدمشقي"),
    "rif-dimashq-ghouta.svg": ("ريف دمشق", "غوطة ريف دمشق"),
    "aleppo-citadel.svg": ("حلب", "قلعة حلب"),
    "homs-clock.svg": ("حمص", "ساعة حمص"),
    "hama-norias.svg": ("حماة", "نواعير حماة"),
    "latakia-triumph-arch.svg": ("اللاذقية", "قوس النصر"),
    "tartous-arwad.svg": ("طرطوس", "جزيرة أرواد"),
    "idlib-ruweiha.svg": ("إدلب", "رويحة إدلب"),
    "deir-ez-zor-bridge.svg": ("دير الزور", "الجسر المعلق"),
    "raqqa-baghdad-gate.svg": ("الرقة", "بوابة بغداد"),
    "hasakah-ain-diwar.svg": ("الحسكة", "جسر عين ديوار"),
    "daraa-omari-mosque.svg": ("درعا", "المسجد العمري"),
    "sweida-qanawat.svg": ("السويداء", "قنوات السويداء"),
    "quneitra-beit-saida.svg": ("القنيطرة", "بيت صيدا"),
}

# Source fills as published. The official site recolours by replacing these.
ICON_SOURCE_HEX = "#04018c"   # governorate icons
LOGO_SOURCE_HEX = "#988561"   # eagle emblem (Golden Wheat dark)


def get(path: str) -> bytes:
    req = urllib.request.Request(f"{BASE}/{path}", headers={"User-Agent": UA})
    with urllib.request.urlopen(req, timeout=60) as r:
        return r.read()


def write(rel: Path, data: bytes) -> Path:
    rel.parent.mkdir(parents=True, exist_ok=True)
    rel.write_bytes(data)
    print(f"  {rel.relative_to(ROOT)}  ({len(data):,} bytes)")
    return rel


def rasterise(
    svg: bytes,
    out: Path,
    dpi: int = 300,
    swap: tuple[str, str] | None = None,
) -> Path:
    """SVG bytes -> PNG on disk. `swap` is (from_hex, to_hex), applied case-insensitively."""
    text = inline_styles(svg.decode("utf-8"))
    if swap:
        src, dst = swap
        text = re.sub(re.escape(src), dst, text, flags=re.I)
    svg = text.encode("utf-8")
    doc = pymupdf.open(stream=io.BytesIO(svg), filetype="svg")
    pix = doc[0].get_pixmap(dpi=dpi, alpha=True)
    out.parent.mkdir(parents=True, exist_ok=True)
    pix.save(out)
    print(f"  {out.relative_to(ROOT)}  ({out.stat().st_size:,} bytes, {pix.width}x{pix.height})")
    return out


def main() -> int:
    print("fonts (Hayyakum Allah, 4 weights)")
    for w in FONT_WEIGHTS:
        write(ASSETS / "fonts" / f"HayyakumAllah-{w}.ttf",
              get(f"fonts/HayyakumAllah/TTF/HayyakumAllah-{w}.ttf"))

    print("\nlogo + wordmarks")
    logo_svg = get("materials/logo.ai.svg")
    write(ASSETS / "logo" / "logo.svg", logo_svg)
    rasterise(logo_svg, ASSETS / "logo" / "eagle-gold.png")
    rasterise(logo_svg, ASSETS / "logo" / "eagle-white.png", swap=(LOGO_SOURCE_HEX, "#ffffff"))
    rasterise(logo_svg, ASSETS / "logo" / "eagle-forest.png", swap=(LOGO_SOURCE_HEX, "#054239"))
    for m in MATERIALS[1:]:
        data = get(m)
        name = Path(m).name
        if name.endswith(".svg"):
            write(ASSETS / "logo" / name, data)
            rasterise(data, ASSETS / "logo" / name.replace(".svg", ".png"))
        else:
            write(ASSETS / "logo" / name, data)

    print("\ngovernorate icons (rasterised in Forest + gold)")
    for f in GOVERNORATES:
        svg = get(f"icons/governorates/with-frame/{f}")
        stem = f[:-4]
        write(ASSETS / "icons" / f, svg)
        rasterise(svg, ASSETS / "icons" / f"{stem}-forest.png", dpi=200,
                  swap=(ICON_SOURCE_HEX, "#054239"))
        rasterise(svg, ASSETS / "icons" / f"{stem}-gold.png", dpi=200,
                  swap=(ICON_SOURCE_HEX, "#988561"))

    print("\ndone ->", ASSETS)
    return 0


if __name__ == "__main__":
    sys.exit(main())
