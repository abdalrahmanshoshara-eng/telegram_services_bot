"""Pure contact normalization and classification rules."""

import math
import re
import unicodedata

FULL_NAME = "الاسم الكامل"
PHONE = "رقم التواصل"
EMAIL = "البريد الالكتروني"
# Only the name and the phone number are needed to build a contact; every other
# column is used when the workbook happens to carry it and silently skipped when
# it does not.
REQUIRED_COLUMNS = [FULL_NAME, PHONE]
OPTIONAL_COLUMNS = [EMAIL]
EXPECTED_COLUMNS = REQUIRED_COLUMNS + OPTIONAL_COLUMNS

EXTRA_PHONES = "أرقام إضافية"
VALID_COLUMNS = EXPECTED_COLUMNS + [EXTRA_PHONES]
# "Location" is a worksheet row, a line of text, a vCard or a table row, and
# carries the file name when several files are processed together.
LOCATION = "الموقع"
DUPLICATE_LOCATION = "موقع المكرر"
KEPT_LOCATION = "موقع المحتفظ به"
DUPLICATE_COLUMNS = [
    DUPLICATE_LOCATION,
    "الاسم الكامل المكرر",
    PHONE,
    "البريد الالكتروني المكرر",
    KEPT_LOCATION,
    "الاسم المحتفظ به",
    "ملاحظة",
]
INVALID_COLUMNS = [
    LOCATION,
    FULL_NAME,
    "رقم التواصل الأصلي",
    "البريد الالكتروني الأصلي",
    "سبب الاستبعاد",
]
MAX_PHONES_PER_CONTACT = 10

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


def _records(rows: list[dict]) -> list[dict]:
    """Group rows into contact records.

    A reader emits one row per phone number, so a vCard, a CSV line or a text
    line holding two numbers arrives as two rows with the same `_row` and
    `_source`. Rows without `_row` (Excel) are each their own record.
    """
    records: list[dict] = []
    for index, row in enumerate(rows):
        location = row.get("_row", index + 2)
        key = (row.get("_source"), location) if "_row" in row else None
        name = "" if is_empty(row.get(FULL_NAME)) else str(row[FULL_NAME]).strip()
        if key is not None and records and records[-1]["key"] == key:
            record = records[-1]
        else:
            record = {"key": key, "location": location, "name": name,
                      "phones": [], "email": None}
            records.append(record)
        record["phones"].append(row.get(PHONE))
        if record["email"] is None and not is_empty(row.get(EMAIL)):
            record["email"] = row.get(EMAIL)
    return records


def _invalid(location, name, phone, email, reasons: list[str]) -> dict:
    return {
        LOCATION: location,
        FULL_NAME: name,
        "رقم التواصل الأصلي": value_to_text(phone),
        "البريد الالكتروني الأصلي": "" if is_empty(email) else str(email).strip(),
        "سبب الاستبعاد": " | ".join(reasons),
    }


def process_contact_rows(rows: list[dict], country_code="963") -> dict:
    """Validate, merge and classify contacts.

    One contact keeps all of its numbers. Two records that share a number are
    the same person, wherever they came from: the later one is merged into the
    first, and any numbers it adds are kept on that first contact.
    """
    contacts: list[dict] = []
    duplicate_rows: list[dict] = []
    invalid_rows: list[dict] = []
    owner_of: dict[str, dict] = {}  # normalized phone -> contact

    records = _records(rows)
    for record in records:
        location, name, raw_email = record["location"], record["name"], record["email"]
        reasons: list[str] = []
        if not name:
            reasons.append("الاسم الكامل فارغ")
        try:
            email = clean_email(raw_email)
        except ValueError as error:
            email = ""
            reasons.append(str(error))

        phones: list[str] = []
        bad_phones: list[tuple[object, str]] = []
        for raw_phone in record["phones"]:
            try:
                phone = normalize_phone(raw_phone, country_code)
            except ValueError as error:
                bad_phones.append((raw_phone, str(error)))
                continue
            if phone not in phones:
                phones.append(phone)

        if reasons or not phones:
            if not phones:
                reasons += [reason for _, reason in bad_phones] or ["رقم التواصل فارغ"]
            raw = " | ".join(value_to_text(p) for p in record["phones"] if not is_empty(p))
            invalid_rows.append(_invalid(location, name, raw, raw_email, reasons))
            continue
        # A contact with one good number and one broken one keeps the good one;
        # the broken one is still listed for review.
        for raw_phone, reason in bad_phones:
            invalid_rows.append(
                _invalid(location, name, raw_phone, raw_email,
                         [f"{reason} (تم حفظ الجهة بأرقامها الأخرى)"])
            )

        target = next((owner_of[p] for p in phones if p in owner_of), None)
        if target is None:
            contact = {FULL_NAME: name, "phones": phones[:MAX_PHONES_PER_CONTACT],
                       EMAIL: email, "location": location}
            contacts.append(contact)
            for phone in contact["phones"]:
                owner_of.setdefault(phone, contact)
            continue

        added = [p for p in phones if p not in target["phones"]]
        room = MAX_PHONES_PER_CONTACT - len(target["phones"])
        added = added[: max(0, room)]
        target["phones"].extend(added)
        for phone in added:
            owner_of.setdefault(phone, target)
        if not target[EMAIL] and email:
            target[EMAIL] = email
        duplicate_rows.append(
            {
                DUPLICATE_LOCATION: location,
                "الاسم الكامل المكرر": name,
                PHONE: "، ".join(phones),
                "البريد الالكتروني المكرر": email,
                KEPT_LOCATION: target["location"],
                "الاسم المحتفظ به": target[FULL_NAME],
                "ملاحظة": (
                    "تم دمجه مع الجهة المحتفظ بها وإضافة أرقامه الجديدة إليها"
                    if added
                    else "تم الاحتفاظ بأول ظهور وتجاهل هذا التكرار"
                ),
            }
        )

    valid_rows = [
        {
            FULL_NAME: contact[FULL_NAME],
            PHONE: contact["phones"][0],
            EMAIL: contact[EMAIL],
            EXTRA_PHONES: "، ".join(contact["phones"][1:]),
            "phones": list(contact["phones"]),
        }
        for contact in contacts
    ]
    return {
        "valid_rows": valid_rows,
        "duplicate_rows": duplicate_rows,
        "invalid_rows": invalid_rows,
        "summary": {
            "totalRows": len(records),
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
            ]
        )
        for phone in contact.get("phones") or [contact[PHONE]]:
            lines.append(f"TEL;TYPE=CELL:{phone}")
        if email:
            lines.append(f"EMAIL;TYPE=INTERNET:{email}")
        lines.extend([f"CATEGORIES:{category}", "END:VCARD"])
    return "\r\n".join(lines) + "\r\n"


def safe_spreadsheet_value(value):
    """Prevent exported user strings from being interpreted as formulas."""
    if isinstance(value, str) and value.startswith(("=", "+", "-", "@")):
        return f"'{value}"
    return value
