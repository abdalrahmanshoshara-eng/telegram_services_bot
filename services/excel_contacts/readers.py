"""Read contacts from CSV, TXT, VCF, DOCX and PDF files.

Excel workbooks are read in `processor.py`. Every reader here returns the same
thing: a list of row dicts keyed by the contact columns in `domain.py`, each
carrying a `_row` number that points back into the source file, plus a short
label for the summary report. From there all formats share one pipeline:
phone normalisation, duplicate merging, the VCF and the review reports.

Two ways of finding contacts are used:

* Tables. A CSV file, a Word table or a PDF table is used column by column
  when one of its first rows is a recognisable header - Arabic or English, and
  the Google Contacts and Outlook export headers too.
* Free text. Anything else is read line by line: a phone number and an email
  are picked out of each line, and what is left is the name. A line holding
  only a name is kept for the next line, so "الاسم: أحمد" followed by
  "الهاتف: 0933..." still makes one contact.
"""

from __future__ import annotations

import csv
import io
import quopri
import re
import unicodedata
import zipfile

from .domain import EMAIL, FULL_NAME, PHONE, is_empty

MAX_TEXT_CHARS = 5_000_000
MAX_PDF_PAGES = 200
MAX_ZIP_PARTS = 2_000
MAX_ZIP_UNCOMPRESSED = 100 * 1024 * 1024
HEADER_SEARCH_ROWS = 10


class ContactsReadError(ValueError):
    """Expected rejection; the message is safe to show to the user."""


# --------------------------------------------------------------------------
# Header recognition
# --------------------------------------------------------------------------

def _key(value) -> str:
    """Comparable form of a header cell: NFKC, no tatweel, plain alef, lower."""
    text = unicodedata.normalize("NFKC", str(value or "")).lstrip("﻿")
    text = text.replace("ـ", "")
    text = re.sub("[أإآ]", "ا", text)
    text = re.sub(r"[\s_]+", " ", text).strip().lower()
    return text.strip(" :*")


_FULL = {
    "الاسم الكامل", "الاسم", "اسم", "الاسم الثلاثي", "الاسم الرباعي",
    "اسم الشخص", "اسم جهة الاتصال", "الاسم والكنية", "name", "full name",
    "fullname", "display name", "contact name", "contact",
}
_FIRST = {"first name", "given name", "الاسم الاول"}
_MIDDLE = {"middle name", "additional name", "اسم الاب"}
_LAST = {
    "last name", "family name", "surname", "الكنية", "اسم العائلة",
    "العائلة", "الشهرة", "اللقب",
}
_PHONE = {
    "رقم التواصل", "رقم الهاتف", "الهاتف", "هاتف", "الجوال", "جوال",
    "الموبايل", "موبايل", "رقم الموبايل", "رقم الجوال", "الرقم", "رقم",
    "الهاتف المحمول", "المحمول", "رقم المحمول", "واتساب", "رقم الواتساب",
    "phone", "phone number", "mobile", "mobile phone", "mobile number",
    "telephone", "tel", "cell", "cell phone", "home phone", "business phone",
    "work phone", "primary phone", "other phone", "whatsapp",
}
_EMAIL = {
    "البريد الالكتروني", "البريد", "الايميل", "ايميل", "بريد الكتروني",
    "email", "e-mail", "email address", "e-mail address", "mail",
}
_GOOGLE_PHONE = re.compile(r"^phone \d+ - value$")
_GOOGLE_EMAIL = re.compile(r"^e-mail \d+ - value$")
_OUTLOOK_EMAIL = re.compile(r"^e-mail \d+ address$")


def _role(cell) -> str | None:
    key = _key(cell)
    if not key:
        return None
    if key in _FULL:
        return "full"
    if key in _FIRST:
        return "first"
    if key in _MIDDLE:
        return "middle"
    if key in _LAST:
        return "last"
    if key in _PHONE or _GOOGLE_PHONE.match(key):
        return "phone"
    if key in _EMAIL or _GOOGLE_EMAIL.match(key) or _OUTLOOK_EMAIL.match(key):
        return "email"
    return None


def _header_roles(row: list) -> list[str | None] | None:
    """Roles for a header row, or None when the row is not a contacts header."""
    roles = [_role(cell) for cell in row]
    has_name = any(r in {"full", "first", "last"} for r in roles)
    if has_name and "phone" in roles:
        return roles
    return None


def _cell(row: list, index: int) -> str:
    if index >= len(row) or is_empty(row[index]):
        return ""
    return " ".join(str(row[index]).split())


def _split_multi(value: str) -> list[str]:
    """Google exports put several numbers in one cell, separated by ':::'."""
    return [part.strip() for part in re.split(r":::|\n", value) if part.strip()]


def rows_from_table(table: list[list], first_line: int = 1) -> list[dict] | None:
    """Contacts from a table whose header is among its first rows, else None.

    `first_line` is the source line or row number of `table[0]`, so reported
    row numbers match what the user sees in the file.
    """
    for offset, row in enumerate(table[:HEADER_SEARCH_ROWS]):
        roles = _header_roles(row)
        if roles is None:
            continue
        rows: list[dict] = []
        for position, cells in enumerate(table[offset + 1 :], start=offset + 1):
            if all(is_empty(value) for value in cells):
                continue
            if _header_roles(cells) is not None:
                continue  # the header repeated, e.g. on every PDF page
            column = lambda role: [  # noqa: E731
                _cell(cells, i) for i, r in enumerate(roles) if r == role
            ]
            # "الاسم" + "اسم الأب" + "الكنية" make one name; a part already
            # inside the full name is not repeated.
            full = next((v for v in column("full") if v), "")
            parts = [full] if full else [v for v in column("first") if v]
            for role in ("middle", "last"):
                parts += [v for v in column(role) if v and v not in full]
            name = " ".join(parts)
            phones = [p for value in column("phone") for p in _split_multi(value)]
            emails = [e for value in column("email") for e in _split_multi(value)]
            email = emails[0] if emails else ""
            for phone in phones or [""]:
                rows.append(
                    {
                        FULL_NAME: name,
                        PHONE: phone,
                        EMAIL: email,
                        "_row": first_line + position,
                    }
                )
        return rows
    return None


# --------------------------------------------------------------------------
# Free text
# --------------------------------------------------------------------------

_DIGIT = "0-9٠-٩۰-۹"
_PHONE_TOKEN = re.compile(
    rf"(?<![{_DIGIT}])(?:\+|00)?[{_DIGIT}](?:[{_DIGIT}]|[ \-.()/](?=[ \-.()/]?[{_DIGIT}])){{5,24}}[{_DIGIT}](?![{_DIGIT}])"
)
_EMAIL_TOKEN = re.compile(r"[^\s@<>()\[\],;:|\"']+@[^\s@<>()\[\],;:|\"']+\.[A-Za-z]{2,}")
_LABEL = re.compile(
    r"(?i)(?:الاسم الكامل|الاسم|اسم|رقم التواصل|رقم الهاتف|رقم الموبايل|رقم الجوال|"
    r"الهاتف|هاتف|الموبايل|موبايل|الجوال|جوال|الرقم|رقم|البريد الإلكتروني|"
    r"البريد الالكتروني|البريد|الإيميل|الايميل|ايميل|name|phone|mobile|tel|"
    r"e-?mail)\s*[:：]"
)
# A list number in front of the name: "1.", "2)", "3 -" or just "4 ".
_NUMBERING = re.compile(rf"^\s*[{_DIGIT}]{{1,4}}(?:\s*[.)\-:،]|\s)\s*")
_SEPARATORS = " \t|,;:،؛-–—_/\\•·*\"'()[]{}<>"
_HAS_LETTER = re.compile(r"[^\W\d_]")


def _digits(token: str) -> int:
    return sum(ch.isdigit() for ch in token)


def _clean_name(text: str) -> str:
    text = _LABEL.sub(" ", text)
    text = _NUMBERING.sub("", text)
    text = " ".join(text.split()).strip(_SEPARATORS)
    text = re.sub(r"\s*[|,;،؛]\s*", " ", text)
    text = " ".join(text.split()).strip(_SEPARATORS)
    return text if _HAS_LETTER.search(text) else ""


def rows_from_lines(lines: list[tuple[int, str]]) -> list[dict]:
    """Contacts from (line number, text) pairs of unstructured text."""
    rows: list[dict] = []
    pending: tuple[int, str] | None = None  # a name still waiting for its number
    phone_line = -1  # line of the last number found
    for number, line in lines:
        line = unicodedata.normalize("NFKC", line)
        emails = _EMAIL_TOKEN.findall(line)
        rest = _EMAIL_TOKEN.sub(" ", line)
        phones = [
            match.group().strip()
            for match in _PHONE_TOKEN.finditer(rest)
            if 7 <= _digits(match.group()) <= 15
        ]
        for phone in phones:
            rest = rest.replace(phone, " ", 1)
        name = _clean_name(rest)

        if not phones:
            if (
                emails
                and not name
                and rows
                and phone_line == number - 1
                and not rows[-1][EMAIL]
            ):
                rows[-1][EMAIL] = emails[0]  # "البريد: x@y.z" under its contact
            elif name and not emails:
                pending = (number, name)
            continue

        row_number = number
        if not name and pending is not None:
            row_number, name = pending
        pending = None
        phone_line = number
        for phone in phones:
            rows.append(
                {
                    FULL_NAME: name,
                    PHONE: phone,
                    EMAIL: emails[0] if emails else "",
                    "_row": row_number,
                }
            )
    return rows


# --------------------------------------------------------------------------
# Formats
# --------------------------------------------------------------------------

def decode_text(data: bytes) -> str:
    """UTF-8 (with or without BOM), UTF-16 with BOM, else Windows Arabic."""
    if data.startswith((b"\xff\xfe", b"\xfe\xff")):
        text = data.decode("utf-16", errors="replace")
    else:
        try:
            text = data.decode("utf-8-sig")
        except UnicodeDecodeError:
            # Excel on Arabic Windows saves "CSV" in this code page.
            text = data.decode("cp1256", errors="replace")
    if "\x00" in text:
        raise ContactsReadError("الملف ليس ملفًا نصيًا صالحًا.")
    if len(text) > MAX_TEXT_CHARS:
        raise ContactsReadError("حجم النص في الملف يتجاوز الحد المسموح.")
    return text


def _csv_table(text: str) -> list[list[str]]:
    sample = text[:20_000]
    try:
        dialect = csv.Sniffer().sniff(sample, delimiters=",;\t|؛")
        delimiter = dialect.delimiter
    except csv.Error:
        delimiter = ","
    return [row for row in csv.reader(io.StringIO(text), delimiter=delimiter)]


def _text_lines(text: str) -> list[tuple[int, str]]:
    return [(i, line) for i, line in enumerate(text.splitlines(), start=1)]


def read_delimited(data: bytes, kind: str) -> tuple[list[dict], str]:
    """CSV and TXT: a header table if there is one, otherwise free text."""
    text = decode_text(data)
    try:
        rows = rows_from_table(_csv_table(text))
    except csv.Error:
        rows = None
    if rows is not None:
        return rows, f"جدول {kind}"
    return rows_from_lines(_text_lines(text)), f"نص {kind}"


# --- VCF -------------------------------------------------------------------

def _unfold_vcard(text: str) -> list[str]:
    lines: list[str] = []
    for raw in text.splitlines():
        if raw[:1] in {" ", "\t"} and lines:
            lines[-1] += raw[1:]  # RFC 6350 folding
        elif lines and lines[-1].endswith("=") and "QUOTED-PRINTABLE" in lines[-1].upper():
            lines[-1] = lines[-1][:-1] + raw  # quoted-printable soft break
        else:
            lines.append(raw)
    return lines


def _vcard_value(params: list[str], value: str) -> str:
    upper = [p.upper() for p in params]
    if "ENCODING=QUOTED-PRINTABLE" in upper or "QUOTED-PRINTABLE" in upper:
        charset = next(
            (p.split("=", 1)[1] for p in params if p.upper().startswith("CHARSET=")),
            "utf-8",
        )
        try:
            value = quopri.decodestring(value.encode("latin-1", "replace")).decode(
                charset, errors="replace"
            )
        except LookupError:
            value = quopri.decodestring(value.encode("latin-1", "replace")).decode(
                "utf-8", errors="replace"
            )
    value = re.sub(r"\\([,;\\])", r"\1", value)
    return value.replace("\\n", " ").replace("\\N", " ").strip()


def read_vcf(data: bytes) -> tuple[list[dict], str]:
    text = decode_text(data)
    if "BEGIN:VCARD" not in text.upper():
        raise ContactsReadError("الملف لا يحتوي على جهات اتصال بصيغة VCF.")
    rows: list[dict] = []
    card: dict | None = None
    index = 0
    for line in _unfold_vcard(text):
        if ":" not in line:
            continue
        head, value = line.split(":", 1)
        parts = head.split(";")
        prop = parts[0].split(".")[-1].upper()  # "item1.TEL" -> "TEL"
        if prop == "BEGIN" and value.strip().upper() == "VCARD":
            index += 1
            card = {"fn": "", "n": "", "tels": [], "emails": []}
        elif card is None:
            continue
        elif prop == "END":
            name = card["fn"] or card["n"]
            for phone in card["tels"] or [""]:
                rows.append(
                    {
                        FULL_NAME: name,
                        PHONE: phone,
                        EMAIL: card["emails"][0] if card["emails"] else "",
                        "_row": index,
                    }
                )
            card = None
        elif prop == "FN":
            card["fn"] = _vcard_value(parts[1:], value)
        elif prop == "N":
            # Family;Given;Additional;Prefix;Suffix -> "Given Additional Family"
            fields = _vcard_value(parts[1:], value.replace("\\;", "\x00")).split(";")
            fields = [f.replace("\x00", ";").strip() for f in fields] + [""] * 5
            card["n"] = " ".join(f for f in (fields[1], fields[2], fields[0]) if f)
        elif prop == "TEL":
            phone = _vcard_value(parts[1:], value)
            if phone:
                card["tels"].append(phone)
        elif prop == "EMAIL":
            email = _vcard_value(parts[1:], value)
            if email:
                card["emails"].append(email)
    return rows, f"ملف VCF ({index} جهة)"


# --- DOCX ------------------------------------------------------------------

def _check_zip(data: bytes, required: str, label: str) -> None:
    try:
        with zipfile.ZipFile(io.BytesIO(data)) as archive:
            members = archive.infolist()
            if len(members) > MAX_ZIP_PARTS:
                raise ContactsReadError(f"يحتوي ملف {label} على عدد مفرط من الأجزاء.")
            if sum(m.file_size for m in members) > MAX_ZIP_UNCOMPRESSED:
                raise ContactsReadError(
                    f"حجم ملف {label} بعد فك الضغط يتجاوز الحد المسموح."
                )
            names = {m.filename.lower() for m in members}
            if "encryptedpackage" in names:
                raise ContactsReadError(f"ملفات {label} المحمية بكلمة مرور غير مدعومة.")
            if required not in names:
                raise ContactsReadError(f"محتوى الملف لا يطابق صيغة {label}.")
    except zipfile.BadZipFile as error:
        raise ContactsReadError(f"ملف {label} تالف أو غير صالح.") from error


def read_docx(data: bytes) -> tuple[list[dict], str]:
    from docx import Document
    from docx.oxml.ns import qn
    from docx.table import Table
    from docx.text.paragraph import Paragraph

    _check_zip(data, "word/document.xml", "Word")
    try:
        document = Document(io.BytesIO(data))
    except Exception as error:
        raise ContactsReadError("تعذر فتح ملف Word. تأكد من سلامته.") from error

    tables: list[list[list[str]]] = []
    lines: list[tuple[int, str]] = []
    for child in document.element.body.iterchildren():
        if child.tag == qn("w:p"):
            lines.append((len(lines) + 1, Paragraph(child, document).text))
        elif child.tag == qn("w:tbl"):
            table = []
            for row in Table(child, document).rows:
                cells, previous = [], None
                for cell in row.cells:
                    if previous is not None and cell._tc is previous:
                        continue  # a merged cell is returned once per grid column
                    previous = cell._tc
                    cells.append(" ".join(cell.text.split()))
                table.append(cells)
                lines.append((len(lines) + 1, " | ".join(cells)))
            tables.append(table)

    table_rows: list[dict] = []
    for number, table in enumerate(tables, start=1):
        rows = rows_from_table(table)
        if rows:
            for row in rows:
                row["_row"] = f"جدول {number} - صف {row['_row']}"
            table_rows.extend(rows)
    if table_rows:
        return table_rows, "جداول Word"
    return rows_from_lines(lines), "نص Word"


# --- PDF -------------------------------------------------------------------

_ARABIC = re.compile("[؀-ۿݐ-ݿﭐ-﷿ﹰ-﻿]")
_LATIN = re.compile("[A-Za-z]")
# PDF fonts often map Arabic glyphs to their Persian look-alikes.
_PERSIAN = str.maketrans({"ی": "ي", "ھ": "ه", "ک": "ك"})
_MIRROR = str.maketrans("()[]{}<>", ")(][}{><")
_NUMBER_PARTS = re.compile(r"^(\D*)(\d(?:.*\d)?)(\D*)$")


def _letters(text: str) -> str:
    return unicodedata.normalize("NFKC", text).translate(_PERSIAN)


def _logical_token(token: str) -> str:
    if _ARABIC.search(token):
        # Reverse before NFKC: a lam-alef ligature is one glyph on the page and
        # must expand to "لا" after the word is turned around, not before.
        return _letters(token[::-1].translate(_MIRROR))
    token = _letters(token)
    match = _NUMBER_PARTS.match(token)
    if match and not _LATIN.search(token) and "@" not in token:
        # "963+" -> "+963", ".1" -> "1.", "(2" -> "2)": digits keep their order.
        lead, core, trail = match.groups()
        return trail[::-1].translate(_MIRROR) + core + lead[::-1].translate(_MIRROR)
    return token


def _logical(text: str, rtl: bool | None = None) -> str:
    """Put a line of PDF text that holds Arabic back in reading order.

    pdfplumber lays characters out left to right as they appear on the page,
    so in a right-to-left line every Arabic word arrives with its letters
    reversed and the words arrive in reverse order. Numbers and Latin words
    keep their letters in order, and a run of Latin words keeps its order
    too. The phone number "0933 123 456", which the page shows as
    "456 123 0933", comes back as written.

    `rtl` is for a table cell: a cell holding only a number has no Arabic of
    its own, but it still sits in a right-to-left table.
    """
    if rtl is None:
        rtl = bool(_ARABIC.search(text))
    if not rtl:
        return _letters(text)
    groups: list[tuple[bool, list[str]]] = []
    for token in text.split():
        latin = bool(_LATIN.search(token)) and not _ARABIC.search(token)
        if latin and groups and groups[-1][0]:
            groups[-1][1].append(token)
        else:
            groups.append((latin, [token]))
    words: list[str] = []
    for latin, tokens in reversed(groups):
        words.extend(
            [_letters(t) for t in tokens] if latin else [_logical_token(t) for t in tokens]
        )
    return " ".join(words)


def read_pdf(data: bytes) -> tuple[list[dict], str]:
    import pdfplumber

    try:
        pdf = pdfplumber.open(io.BytesIO(data))
    except Exception as error:
        raise ContactsReadError(
            "تعذرت قراءة ملف PDF. تأكد من أنه غير محمي بكلمة مرور وغير تالف."
        ) from error
    with pdf:
        if len(pdf.pages) > MAX_PDF_PAGES:
            raise ContactsReadError(
                f"عدد صفحات PDF يتجاوز الحد ({MAX_PDF_PAGES} صفحة)."
            )
        table_rows: list[dict] = []
        lines: list[tuple[int, str]] = []
        try:
            for page_number, page in enumerate(pdf.pages, start=1):
                for table in page.extract_tables():
                    # Column order does not matter: columns are matched by header.
                    rtl = any(
                        _ARABIC.search(cell or "") for row in table for cell in row
                    )
                    cleaned = [
                        [_logical(cell or "", rtl) for cell in row] for row in table
                    ]
                    rows = rows_from_table(cleaned)
                    for row in rows or []:
                        row["_row"] = f"صفحة {page_number} - صف {row['_row']}"
                    table_rows.extend(rows or [])
                text = page.extract_text() or ""
                for line in text.splitlines():
                    lines.append((len(lines) + 1, _logical(line)))
        except ContactsReadError:
            raise
        except Exception as error:
            raise ContactsReadError("تعذرت قراءة محتوى ملف PDF.") from error
    if table_rows:
        return table_rows, "جداول PDF"
    if not any(text.strip() for _, text in lines):
        raise ContactsReadError(
            "ملف PDF لا يحتوي على نص قابل للقراءة، ويبدو أنه ممسوح ضوئيًا."
        )
    return rows_from_lines(lines), "نص PDF"


READERS = {
    ".csv": lambda data: read_delimited(data, "CSV"),
    ".txt": lambda data: read_delimited(data, "TXT"),
    ".vcf": read_vcf,
    ".docx": read_docx,
    ".pdf": read_pdf,
}


def read_contacts(data: bytes, extension: str) -> tuple[list[dict], str]:
    reader = READERS.get(extension)
    if reader is None:
        raise ContactsReadError("صيغة الملف غير مدعومة.")
    rows, label = reader(data)
    if not rows:
        raise ContactsReadError(
            "لم أجد أي جهة اتصال في الملف. تأكد أن كل جهة تحتوي اسمًا ورقم هاتف."
        )
    return rows, label
