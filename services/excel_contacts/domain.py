"""Pure contact normalization and classification rules."""

import math
import re
import unicodedata

FULL_NAME = "الاسم الكامل"
PHONE = "رقم التواصل"
EMAIL = "البريد الالكتروني"
EXPECTED_COLUMNS = [FULL_NAME, PHONE, EMAIL]

VALID_COLUMNS = EXPECTED_COLUMNS
DUPLICATE_COLUMNS = [
    "صف Excel المكرر",
    "الاسم الكامل المكرر",
    PHONE,
    "البريد الالكتروني المكرر",
    "صف Excel المحتفظ به",
    "الاسم المحتفظ به",
    "ملاحظة",
]
INVALID_COLUMNS = [
    "صف Excel",
    FULL_NAME,
    "رقم التواصل الأصلي",
    "البريد الالكتروني الأصلي",
    "سبب الاستبعاد",
]

ARABIC_DIGITS = str.maketrans("٠١٢٣٤٥٦٧٨٩۰۱۲۳۴۵۶۷۸۹", "01234567890123456789")
EMAIL_PATTERN = re.compile(r"^[^@\s]+@[^@\s]+\.[^@\s]+$")


def is_empty(value) -> bool:
    return value is None or str(value).strip() == ""


def normalize_header(value) -> str:
    return re.sub(
        r"\s+",
        " ",
        unicodedata.normalize("NFKC", str(value or "")).lstrip("\ufeff").strip(),
    )


def value_to_text(value) -> str:
    if is_empty(value):
        return ""
    if isinstance(value, float):
        if not math.isfinite(value):
            return ""
        if value.is_integer():
            return str(int(value))
        return f"{value:.12f}".rstrip("0").rstrip(".")
    return str(value).strip()


def normalize_country_code(value) -> str:
    code = re.sub(r"\D", "", str(value or ""))
    if not re.fullmatch(r"\d{1,4}", code):
        raise ValueError("رمز الدولة يجب أن يتكون من 1 إلى 4 أرقام.")
    return code


def clean_phone_text(value) -> str:
    text = unicodedata.normalize("NFKC", value_to_text(value)).translate(ARABIC_DIGITS)
    text = re.sub(r"[^0-9+]", "", text.strip())
    if "+" in text:
        text = ("+" if text.startswith("+") else "") + text.replace("+", "")
    if text.startswith("00"):
        text = f"+{text[2:]}"
    return text


def normalize_phone(value, country_code="963") -> str:
    default_country_code = normalize_country_code(country_code)
    raw = clean_phone_text(value)
    if not raw:
        raise ValueError("رقم التواصل فارغ")

    if raw.startswith("+"):
        digits = raw[1:]
        if not re.fullmatch(r"\d{8,15}", digits):
            raise ValueError(f"رقم دولي غير صالح: {raw}")
        return f"+{digits}"

    if not raw.isdigit():
        raise ValueError(f"تعذر فهم الرقم: {raw}")
    if raw.startswith(default_country_code):
        if not 8 <= len(raw) <= 15:
            raise ValueError(f"رقم غير صالح: {raw}")
        return f"+{raw}"
    if raw.startswith("0"):
        local = raw[1:]
        if not 7 <= len(local) <= 12:
            raise ValueError(f"رقم محلي غير صالح: {raw}")
        return f"+{default_country_code}{local}"
    if default_country_code == "963" and re.fullmatch(r"9\d{8}", raw):
        return f"+{default_country_code}{raw}"
    raise ValueError(
        f"رقم غير معروف الصيغة: {raw}. استخدم صيغة دولية مثل "
        f"+{default_country_code}9XXXXXXXX"
    )


def clean_email(value) -> str:
    if is_empty(value):
        return ""
    email = unicodedata.normalize("NFKC", str(value)).strip().lower()
    if not EMAIL_PATTERN.fullmatch(email):
        raise ValueError("بريد إلكتروني غير صالح")
    return email


def process_contact_rows(rows: list[dict], country_code="963") -> dict:
    valid_rows: list[dict] = []
    duplicate_rows: list[dict] = []
    invalid_rows: list[dict] = []
    seen_phones: dict[str, dict] = {}

    for index, row in enumerate(rows):
        excel_row = index + 2
        name = "" if is_empty(row.get(FULL_NAME)) else str(row[FULL_NAME]).strip()
        raw_phone = row.get(PHONE)
        raw_email = row.get(EMAIL)
        reasons: list[str] = []

        if not name:
            reasons.append("الاسم الكامل فارغ")
        try:
            normalized_phone = normalize_phone(raw_phone, country_code)
        except ValueError as error:
            normalized_phone = ""
            reasons.append(str(error))
        try:
            normalized_email = clean_email(raw_email)
        except ValueError as error:
            normalized_email = ""
            reasons.append(str(error))

        if reasons:
            invalid_rows.append(
                {
                    "صف Excel": excel_row,
                    FULL_NAME: name,
                    "رقم التواصل الأصلي": value_to_text(raw_phone),
                    "البريد الالكتروني الأصلي": (
                        "" if is_empty(raw_email) else str(raw_email).strip()
                    ),
                    "سبب الاستبعاد": " | ".join(reasons),
                }
            )
            continue

        if normalized_phone in seen_phones:
            first = seen_phones[normalized_phone]
            duplicate_rows.append(
                {
                    "صف Excel المكرر": excel_row,
                    "الاسم الكامل المكرر": name,
                    PHONE: normalized_phone,
                    "البريد الالكتروني المكرر": normalized_email,
                    "صف Excel المحتفظ به": first["excel_row"],
                    "الاسم المحتفظ به": first["name"],
                    "ملاحظة": "تم الاحتفاظ بأول ظهور وتجاهل هذا التكرار",
                }
            )
            continue

        seen_phones[normalized_phone] = {"excel_row": excel_row, "name": name}
        valid_rows.append(
            {FULL_NAME: name, PHONE: normalized_phone, EMAIL: normalized_email}
        )

    return {
        "valid_rows": valid_rows,
        "duplicate_rows": duplicate_rows,
        "invalid_rows": invalid_rows,
        "summary": {
            "totalRows": len(rows),
            "validCount": len(valid_rows),
            "duplicateCount": len(duplicate_rows),
            "invalidCount": len(invalid_rows),
        },
    }


def escape_vcard(value) -> str:
    return (
        str(value)
        .replace("\\", "\\\\")
        .replace("\r\n", "\\n")
        .replace("\n", "\\n")
        .replace(";", "\\;")
        .replace(",", "\\,")
    )


def create_vcard(contacts: list[dict], category="WhatsApp Excel Import") -> str:
    lines: list[str] = []
    for contact in contacts:
        name = escape_vcard(contact[FULL_NAME])
        email = escape_vcard(contact.get(EMAIL, ""))
        lines.extend(
            [
                "BEGIN:VCARD",
                "VERSION:3.0",
                f"FN;CHARSET=UTF-8:{name}",
                f"N;CHARSET=UTF-8:;{name};;;",
                f"TEL;TYPE=CELL:{contact[PHONE]}",
            ]
        )
        if email:
            lines.append(f"EMAIL;TYPE=INTERNET:{email}")
        lines.extend([f"CATEGORIES:{category}", "END:VCARD"])
    return "\r\n".join(lines) + "\r\n"


def safe_spreadsheet_value(value):
    """Prevent exported user strings from being interpreted as formulas."""
    if isinstance(value, str) and value.startswith(("=", "+", "-", "@")):
        return f"'{value}"
    return value
