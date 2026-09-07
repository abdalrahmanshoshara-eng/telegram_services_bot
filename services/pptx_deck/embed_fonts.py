"""Embed the identity fonts into a .pptx so it renders correctly anywhere.

PowerPoint's own "Embed fonts in the file" needs PowerPoint. This does the same
thing by writing the OOXML parts directly, so it works on a Linux bot host.

The font's OS/2 fsType must permit embedding - Hayyakum Allah is 0x0008
("editable"), which allows it. `embed()` re-checks and refuses otherwise.

Structure written into the package:
  ppt/fonts/fontN.fntdata          the raw TTF
  ppt/_rels/presentation.xml.rels  relationship per font
  ppt/presentation.xml             <p:embeddedFontLst> + embedTrueTypeFonts="1"
  [Content_Types].xml              Default for the fntdata extension
"""
from __future__ import annotations

import re
import shutil
import struct
import zipfile
from pathlib import Path

from . import brand as B

FONT_DIR = B.ASSETS / "fonts"

# family name -> {slot: filename}. Slot is the OOXML child element name.
# Light and Medium register as their own families, so each is a "regular".
FAMILIES: dict[str, dict[str, str]] = {
    "Hayyakum Allah": {
        "regular": "HayyakumAllah-Regular.ttf",
        "bold": "HayyakumAllah-Bold.ttf",
    },
    "Hayyakum Allah Light": {"regular": "HayyakumAllah-Light.ttf"},
    "Hayyakum Allah Medium": {"regular": "HayyakumAllah-Medium.ttf"},
}

FNTDATA_CT = "application/x-fontdata"
FONT_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/font"

# Schema order of CT_Presentation children. embeddedFontLst must land after
# these and before anything else, or PowerPoint reports a repair error.
_BEFORE_FONTLST = ("p:smartTags", "p:notesSz", "p:sldSz", "p:sldIdLst",
                   "p:handoutMasterIdLst", "p:notesMasterIdLst", "p:sldMasterIdLst")


class EmbedError(Exception):
    pass


def _fstype(path: Path) -> int:
    d = path.read_bytes()
    for i in range(struct.unpack(">H", d[4:6])[0]):
        rec = d[12 + 16 * i: 12 + 16 * i + 16]
        if rec[0:4] == b"OS/2":
            off = struct.unpack(">I", rec[8:12])[0]
            return struct.unpack(">H", d[off + 8:off + 10])[0]
    raise EmbedError(f"no OS/2 table in {path.name}")


def _embeddable(path: Path) -> bool:
    """fsType bits 0-3: 0 installable, 2 restricted, 4 preview/print, 8 editable."""
    return (_fstype(path) & 0x000F) != 0x0002


def _next_rel_id(rels_xml: str) -> int:
    ids = [int(m) for m in re.findall(r'Id="rId(\d+)"', rels_xml)]
    return (max(ids) + 1) if ids else 1


def embed(pptx_path: str | Path, font_dir: str | Path | None = None) -> Path:
    """Rewrite `pptx_path` in place with the identity fonts embedded."""
    pptx_path = Path(pptx_path)
    fonts = Path(font_dir) if font_dir else FONT_DIR

    plan: list[tuple[str, str, Path]] = []   # (family, slot, path)
    for family, slots in FAMILIES.items():
        for slot, filename in slots.items():
            p = fonts / filename
            if not p.exists():
                raise EmbedError(f"font not found: {p}")
            if not _embeddable(p):
                raise EmbedError(f"{filename} forbids embedding (fsType restricted)")
            plan.append((family, slot, p))

    with zipfile.ZipFile(pptx_path) as z:
        names = z.namelist()
        if any(n.startswith("ppt/fonts/") for n in names):
            return pptx_path                      # already embedded
        items = {n: z.read(n) for n in names}

    rels_name = "ppt/_rels/presentation.xml.rels"
    pres_name = "ppt/presentation.xml"
    ct_name = "[Content_Types].xml"
    for required in (rels_name, pres_name, ct_name):
        if required not in items:
            raise EmbedError(f"malformed pptx: {required} missing")

    rels = items[rels_name].decode("utf-8")
    rid = _next_rel_id(rels)

    # ---- font parts + relationships ------------------------------------
    added_rels: list[str] = []
    font_lst: dict[str, list[str]] = {}
    for i, (family, slot, path) in enumerate(plan, start=1):
        part = f"ppt/fonts/font{i}.fntdata"
        items[part] = path.read_bytes()
        added_rels.append(
            f'<Relationship Id="rId{rid}" Type="{FONT_REL}" Target="fonts/font{i}.fntdata"/>'
        )
        font_lst.setdefault(family, []).append(f'<p:{slot} r:id="rId{rid}"/>')
        rid += 1

    rels = rels.replace("</Relationships>", "".join(added_rels) + "</Relationships>")
    items[rels_name] = rels.encode("utf-8")

    # ---- content types --------------------------------------------------
    ct = items[ct_name].decode("utf-8")
    if "fntdata" not in ct:
        ct = ct.replace(
            "<Types ",
            "<Types ", 1,
        ).replace(
            "</Types>",
            f'<Default Extension="fntdata" ContentType="{FNTDATA_CT}"/></Types>',
        )
    items[ct_name] = ct.encode("utf-8")

    # ---- presentation.xml ----------------------------------------------
    pres = items[pres_name].decode("utf-8")

    blocks = []
    for family, children in font_lst.items():
        blocks.append(
            f'<p:embeddedFont><p:font typeface="{family}" pitchFamily="34" charset="0"/>'
            + "".join(children)
            + "</p:embeddedFont>"
        )
    fontlst = "<p:embeddedFontLst>" + "".join(blocks) + "</p:embeddedFontLst>"

    inserted = False
    for tag in _BEFORE_FONTLST:
        m = re.search(rf"</{tag}>", pres)
        if m:
            pres = pres[: m.end()] + fontlst + pres[m.end():]
            inserted = True
            break
        m = re.search(rf"<{tag}\b[^>]*/>", pres)
        if m:
            pres = pres[: m.end()] + fontlst + pres[m.end():]
            inserted = True
            break
    if not inserted:
        raise EmbedError("could not locate an insertion point in presentation.xml")

    # Tell PowerPoint the embedded fonts are there to be used.
    if "embedTrueTypeFonts" not in pres:
        pres = re.sub(r"(<p:presentation\b)", r'\1 embedTrueTypeFonts="1"', pres, count=1)

    items[pres_name] = pres.encode("utf-8")

    # ---- rewrite the package -------------------------------------------
    tmp = pptx_path.with_suffix(".embedding.tmp")
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as z:
        for name, data in items.items():
            z.writestr(name, data)
    shutil.move(str(tmp), str(pptx_path))
    return pptx_path
