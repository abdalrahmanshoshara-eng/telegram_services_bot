"""Small DOCX rewrites that make LibreOffice lay a document out the way Word does.

Each fix was found by converting real Arabic documents with both programs and
comparing the pages. None of them changes what the reader sees in Word; they
only remove differences in how LibreOffice reads the same file.

* Complex-script twins. Word draws Arabic in a run that is not marked w:rtl or
  w:cs with the run's ordinary size, bold and italic. LibreOffice always uses
  the complex-script ones (szCs, bCs, iCs). Google Docs exports set only
  w:sz, so headings came out at the default 11pt. Missing twins are added;
  existing values are never changed.
* Content controls. Word often wraps a table of contents in a content
  control, and LibreOffice prints the field codes of the first paragraph
  inside one. Paragraph-level controls are unwrapped; their content stays.
* Table-of-contents fields. The TOC field wrapper is removed and its saved
  entries kept as text. The page-number fields inside them stay, so
  LibreOffice numbers them for its own pages. TOC entries also lose the blue
  "Hyperlink" style, which Word does not show there.
* Floating pictures. LibreOffice measures a picture's horizontal offset from
  the right when its paragraph is right-to-left, so a logo placed top-left
  lands top-right. A paragraph holding only floating pictures is marked
  left-to-right, which changes no text.
"""
from __future__ import annotations

import zipfile
from pathlib import Path

from lxml import etree

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
WP = "{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}"
TWINS = (("sz", "szCs"), ("b", "bCs"), ("i", "iCs"))
# Elements that come after w:bidi inside w:pPr, in schema order.
_AFTER_BIDI = (
    "adjustRightInd", "snapToGrid", "spacing", "ind", "contextualSpacing",
    "mirrorIndents", "suppressOverlap", "jc", "textDirection", "textAlignment",
    "textboxTightWrap", "outlineLvl", "divId", "cnfStyle", "rPr", "sectPr",
    "pPrChange",
)
_PARTS = ("word/document.xml", "word/footnotes.xml", "word/endnotes.xml")


def _on(element) -> bool:
    return element is not None and element.get(W + "val", "1") not in ("0", "false", "off")


def _in_text_box(element) -> bool:
    return any(a.tag == W + "txbxContent" for a in element.iterancestors())


def add_complex_script_twins(root) -> int:
    added = 0
    for rpr in root.iter(W + "rPr"):
        if rpr.getparent() is None or rpr.getparent().tag != W + "r":
            continue  # runs only; styles and defaults are left alone
        if _on(rpr.find(W + "rtl")) or _on(rpr.find(W + "cs")):
            continue
        for plain, twin_name in TWINS:
            source = rpr.find(W + plain)
            if source is None or rpr.find(W + twin_name) is not None:
                continue
            twin = etree.Element(W + twin_name)
            if source.get(W + "val") is not None:
                twin.set(W + "val", source.get(W + "val"))
            source.addnext(twin)
            added += 1
    return added


def unwrap_content_controls(root) -> int:
    unwrapped = 0
    for sdt in list(root.iter(W + "sdt")):
        content = sdt.find(W + "sdtContent")
        if content is None or (content.find(W + "p") is None and content.find(W + "tbl") is None):
            continue  # inline controls inside a paragraph are left alone
        parent = sdt.getparent()
        index = parent.index(sdt)
        for offset, child in enumerate(list(content)):
            parent.insert(index + offset, child)
        parent.remove(sdt)
        unwrapped += 1
    return unwrapped


def unwrap_toc_fields(root) -> int:
    # A run with formatting but no content between a field's code and its
    # result also makes LibreOffice print the code; Word ignores such runs.
    for run in list(root.iter(W + "r")):
        if len(run) == 1 and run[0].tag == W + "rPr" and run.getparent() is not None:
            run.getparent().remove(run)

    stack: list[dict] = []
    removed = 0
    for element in list(root.iter(W + "fldChar", W + "instrText")):
        run = element.getparent()
        if element.tag == W + "instrText":
            if stack:
                stack[-1]["code"] += element.text or ""
                stack[-1]["runs"].append(run)
            continue
        kind = element.get(W + "fldCharType")
        if kind == "begin":
            stack.append({"code": "", "runs": [run]})
        elif kind == "separate" and stack:
            stack[-1]["runs"].append(run)
        elif kind == "end" and stack:
            field = stack.pop()
            if field["code"].strip().upper().startswith("TOC"):
                for field_run in field["runs"] + [run]:
                    if field_run.getparent() is not None:
                        field_run.getparent().remove(field_run)
                removed += 1

    for style in root.iter(W + "pStyle"):
        if not (style.get(W + "val") or "").upper().startswith("TOC"):
            continue
        paragraph = style.getparent().getparent()
        for run_style in list(paragraph.iter(W + "rStyle")):
            if run_style.get(W + "val") == "Hyperlink":
                run_style.getparent().remove(run_style)
    return removed


def unmirror_floating_pictures(root) -> int:
    changed = 0
    for paragraph in root.iter(W + "p"):
        bidi = paragraph.find(f"{W}pPr/{W}bidi")
        if bidi is not None and not _on(bidi):
            continue  # already explicitly left-to-right
        if next(paragraph.iter(WP + "anchor"), None) is None or _in_text_box(paragraph):
            continue
        # Text in floating text boxes keeps its own direction; only the
        # paragraph's own text counts.
        own_text = "".join(
            t.text or "" for t in paragraph.iter(W + "t") if not _in_text_box(t)
        )
        if own_text.strip():
            continue
        ppr = paragraph.find(W + "pPr")
        if ppr is None:
            ppr = etree.Element(W + "pPr")
            paragraph.insert(0, ppr)
        if bidi is None:
            # Written explicitly: right-to-left is often the document default.
            bidi = etree.Element(W + "bidi")
            follower = next(
                (ppr.find(W + name) for name in _AFTER_BIDI if ppr.find(W + name) is not None),
                None,
            )
            if follower is not None:
                follower.addprevious(bidi)
            else:
                ppr.append(bidi)
        bidi.set(W + "val", "0")
        changed += 1
    return changed


def fix_part(xml: bytes) -> tuple[bytes, int]:
    root = etree.fromstring(xml)
    changes = (
        unwrap_content_controls(root)
        + unwrap_toc_fields(root)
        + unmirror_floating_pictures(root)
        + add_complex_script_twins(root)
    )
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True), changes


def prepare_docx(source: Path, destination: Path) -> int:
    """Write a fixed copy of `source` to `destination`; return the change count."""
    total = 0
    with zipfile.ZipFile(source) as archive, zipfile.ZipFile(
        destination, "w", zipfile.ZIP_DEFLATED
    ) as output:
        for item in archive.infolist():
            data = archive.read(item)
            name = item.filename
            if name in _PARTS or (
                name.startswith(("word/header", "word/footer")) and name.endswith(".xml")
            ):
                data, changes = fix_part(data)
                total += changes
            output.writestr(item, data)
    return total
