"""Pull text and heading structure out of a PDF or DOCX.

DOCX heading levels are kept because they seed section breaks for free - the
model does not have to invent a structure the document already states. PDFs
carry no heading metadata, so headings there are inferred from font size.

pdfplumber (MIT) is used rather than PyMuPDF, which is AGPL-3.0 and would put
a copyleft obligation on a network-facing bot.
"""
from __future__ import annotations

import re
from dataclasses import dataclass, field
from pathlib import Path

import pdfplumber
from docx import Document

# Below this, the file is almost certainly a scan with no text layer.
MIN_CHARS = 200

# pdfminer is slow on very long documents, and the model cannot use that much
# text anyway. Read at most this many pages.
MAX_PDF_PAGES = 120


class ExtractionError(ValueError):
    """Raised when an uploaded document cannot be turned into text.

    Subclasses ValueError to match the bot's ServiceInputError convention:
    the message is safe to show the user directly.
    """


class ScannedDocumentError(ExtractionError):
    """PDF carries images but no extractable text."""


@dataclass
class Block:
    text: str
    level: int = 0          # 0 = body, 1..6 = heading depth


@dataclass
class ExtractedDoc:
    blocks: list[Block] = field(default_factory=list)
    source: str = ""
    page_count: int = 0

    @property
    def text(self) -> str:
        return "\n".join(b.text for b in self.blocks)

    @property
    def headings(self) -> list[Block]:
        return [b for b in self.blocks if b.level > 0]

    def outline_hint(self, limit: int = 40) -> str:
        """Indented heading tree, handed to the model as structural guidance."""
        return "\n".join(
            f"{'  ' * (b.level - 1)}- {b.text}" for b in self.headings[:limit]
        )

    def to_prompt_text(self, max_chars: int = 120_000) -> str:
        parts = []
        for b in self.blocks:
            parts.append(f"{'#' * b.level} {b.text}" if b.level else b.text)
        return "\n".join(parts)[:max_chars]


_WS = re.compile(r"[ \t ]+")
_BLANK = re.compile(r"\n{3,}")


def _clean(s: str) -> str:
    return _BLANK.sub("\n\n", _WS.sub(" ", s)).strip()


def _mode(values: list[float]) -> float:
    counts: dict[float, int] = {}
    for v in values:
        k = round(v, 1)
        counts[k] = counts.get(k, 0) + 1
    return max(counts, key=counts.get) if counts else 0.0


def _lines_from_page(page) -> list[tuple[str, float, bool]]:
    """Group a page's characters into lines of (text, max size, any bold)."""
    lines: dict[float, list[dict]] = {}
    for char in page.chars:
        # Round the baseline so characters on the same visual line group together.
        lines.setdefault(round(char["top"], 1), []).append(char)

    out: list[tuple[str, float, bool]] = []
    for key in sorted(lines):
        chars = sorted(lines[key], key=lambda c: c["x0"])
        text = _clean("".join(c["text"] for c in chars))
        if not text:
            continue
        size = max((c.get("size") or 0.0) for c in chars)
        bold = any("bold" in str(c.get("fontname", "")).lower() for c in chars)
        out.append((text, size, bold))
    return out


def from_pdf(path: Path) -> ExtractedDoc:
    blocks: list[Block] = []
    page_lines: list[list[tuple[str, float, bool]]] = []
    sizes: list[float] = []

    try:
        with pdfplumber.open(str(path)) as pdf:
            page_count = len(pdf.pages)
            for page in pdf.pages[:MAX_PDF_PAGES]:
                lines = _lines_from_page(page)
                page_lines.append(lines)
                sizes.extend(size for _, size, _ in lines if size)
    except Exception as error:                       # corrupt or encrypted PDF
        raise ExtractionError(
            "تعذّرت قراءة ملف PDF. تأكد من أنه غير محمي بكلمة مرور وغير تالف."
        ) from error

    body_size = _mode(sizes) if sizes else 0.0

    for lines in page_lines:
        for text, size, bold in lines:
            level = 0
            if body_size and size > body_size * 1.35:
                level = 1
            elif body_size and (size > body_size * 1.12 or (bold and len(text) < 90)):
                level = 2
            blocks.append(Block(text, level))

    out = ExtractedDoc(blocks=blocks, source=path.name, page_count=page_count)
    if len(out.text.strip()) < MIN_CHARS:
        raise ScannedDocumentError(
            "الملف لا يحتوي على نص قابل للاستخراج (يبدو أنه ملف ممسوح ضوئياً)."
        )
    return out


_HEADING_STYLE = re.compile(r"^Heading (\d)$", re.I)


def from_docx(path: Path) -> ExtractedDoc:
    try:
        doc = Document(str(path))
    except Exception as error:
        raise ExtractionError("تعذّرت قراءة ملف Word. تأكد من سلامة الملف.") from error

    blocks: list[Block] = []
    for para in doc.paragraphs:
        text = _clean(para.text)
        if not text:
            continue
        level = 0
        style = (para.style.name or "") if para.style is not None else ""
        m = _HEADING_STYLE.match(style)
        if m:
            level = int(m.group(1))
        elif style.lower() in {"title", "subtitle"}:
            level = 1
        blocks.append(Block(text, level))

    # Tables carry the numbers in ministry documents - keep them as rows.
    for table in doc.tables:
        for row in table.rows:
            cells = [_clean(c.text) for c in row.cells]
            if any(cells):
                blocks.append(Block(" | ".join(cells), 0))

    out = ExtractedDoc(blocks=blocks, source=path.name)
    if len(out.text.strip()) < MIN_CHARS:
        raise ExtractionError("الملف فارغ أو لا يحتوي على نص كافٍ لإنشاء عرض تقديمي.")
    return out


def extract(path: str | Path) -> ExtractedDoc:
    path = Path(path)
    suffix = path.suffix.lower()
    if suffix == ".pdf":
        return from_pdf(path)
    if suffix == ".docx":
        return from_docx(path)
    if suffix == ".doc":
        raise ExtractionError("صيغة .doc غير مدعومة، الرجاء تحويل الملف إلى .docx أو .pdf")
    raise ExtractionError(f"صيغة غير مدعومة: {suffix or 'unknown'}")
