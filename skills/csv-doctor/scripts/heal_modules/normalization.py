from __future__ import annotations

import re
from datetime import datetime, timedelta, timezone

from heal_modules.shared import (
    COL,
    HEADERS,
    N_COLS,
    SMART_QUOTES,
    Change,
    CleanRow,
    SemanticPlan,
    _strip_nulls,
)


def _clean_cell_text(value: object) -> tuple[str, list[str]]:
    """Strip BOM, null bytes, line breaks, smart quotes. Returns (cleaned_value, reasons)."""
    new_val = _strip_nulls(value)
    reasons = []

    if "\ufeff" in new_val:
        new_val = new_val.replace("\ufeff", "")
        reasons.append("BOM byte-order mark stripped")
    if "\x00" in new_val:
        new_val = new_val.replace("\x00", "")
        reasons.append("Null byte removed")
    if "\n" in new_val or "\r" in new_val:
        new_val = new_val.replace("\r\n", " ").replace("\n", " ").replace("\r", " ")
        reasons.append("Embedded line break replaced with space")

    had_smart = any(s in new_val for s in SMART_QUOTES)
    for smart, straight in SMART_QUOTES.items():
        new_val = new_val.replace(smart, straight)
    if had_smart:
        reasons.append("Smart/curly quotes normalised to straight quotes")

    new_val = new_val.strip()
    return new_val, reasons


def _col_name(i: int) -> str:
    return HEADERS[i] if i < N_COLS else f"[col {i + 1}]"


def clean_row(row: list[str], row_num: int) -> tuple[list[str], list[Change]]:
    cleaned, changes = [], []
    for i, cell in enumerate(row):
        new_val, reasons = _clean_cell_text(cell)
        if reasons:
            orig_display = (
                cell
                .replace("\ufeff", "[BOM]")
                .replace("\x00", "[NULL]")
                .strip()
            )
            changes.append(Change(
                row_num, _col_name(i), orig_display, new_val, "Fixed", "; ".join(reasons)
            ))
        cleaned.append(new_val)
    return cleaned, changes


def parse_amount_like(value: str) -> float | None:
    if not value.strip():
        return None
    normalised, changed, _ = normalise_amount(value)
    candidate = normalised if changed or normalised != value else value.strip()
    try:
        return float(candidate)
    except ValueError:
        return None


def extract_currency_from_text(value: str) -> tuple[str | None, str | None]:
    raw = value.strip()
    if not raw:
        return None, None

    code_match = re.search(r"\b(USD|EUR|GBP|INR|CAD|AUD|JPY)\b", raw, flags=re.IGNORECASE)
    symbol_match = next((symbol for symbol in {"$": "USD", "€": "EUR", "£": "GBP", "₹": "INR", "¥": "JPY"} if symbol in raw), None)
    currency = None
    if code_match:
        currency = code_match.group(1).upper()
    elif symbol_match:
        currency = {"$": "USD", "€": "EUR", "£": "GBP", "₹": "INR", "¥": "JPY"}[symbol_match]

    amount_candidate = raw
    if code_match:
        amount_candidate = re.sub(r"\b(USD|EUR|GBP|INR|CAD|AUD|JPY)\b", "", amount_candidate, flags=re.IGNORECASE)
    if symbol_match:
        amount_candidate = amount_candidate.replace(symbol_match, "")
    amount_candidate = " ".join(amount_candidate.split()).strip()

    if currency and parse_amount_like(amount_candidate) is not None:
        return amount_candidate, currency
    return None, currency


EXCEL_EPOCH  = datetime(1899, 12, 30)
MONTH_NAMES  = {
    "january":1,"february":2,"march":3,"april":4,"may":5,"june":6,
    "july":7,"august":8,"september":9,"october":10,"november":11,"december":12,
    "jan":1,"feb":2,"mar":3,"apr":4,"jun":6,"jul":7,"aug":8,
    "sep":9,"oct":10,"nov":11,"dec":12,
}

def _fmt(dt: datetime) -> str:
    return dt.strftime("%Y-%m-%d")

def normalise_date(value: str) -> tuple[str, bool, str]:
    v = value.strip()
    if not v or re.match(r"^\d{4}-\d{2}-\d{2}$", v):
        return v, False, ""

    # ISO 8601 with time: 2023-01-18T00:00:00Z
    m = re.match(r"^(\d{4}-\d{2}-\d{2})T", v)
    if m:
        return m.group(1), True, "ISO 8601 datetime truncated to date-only"

    # DD/MM/YYYY or MM/DD/YYYY — try day-first, fall back to month-first
    m = re.match(r"^(\d{1,2})/(\d{1,2})/(\d{4})$", v)
    if m:
        a, b, year = int(m.group(1)), int(m.group(2)), int(m.group(3))
        for day, month, fmt in [(a, b, "DD/MM/YYYY"), (b, a, "MM/DD/YYYY")]:
            try:
                return _fmt(datetime(year, month, day)), True, \
                       f"{fmt} normalised to ISO YYYY-MM-DD (day-first assumed for ambiguous dates)"
            except ValueError:
                continue
        return v, False, ""

    # YYYY/MM/DD
    m = re.match(r"^(\d{4})/(\d{1,2})/(\d{1,2})$", v)
    if m:
        year, month, day = int(m.group(1)), int(m.group(2)), int(m.group(3))
        try:
            return _fmt(datetime(year, month, day)), True, "Slash-separated ISO-style date normalised to YYYY-MM-DD"
        except ValueError:
            return v, False, ""

    # DD-MM-YYYY or MM-DD-YYYY — prefer day-first, fall back to month-first
    m = re.match(r"^(\d{1,2})-(\d{1,2})-(\d{4})$", v)
    if m:
        a, b, year = int(m.group(1)), int(m.group(2)), int(m.group(3))
        for day, month, fmt in [(a, b, "DD-MM-YYYY"), (b, a, "MM-DD-YYYY")]:
            try:
                return _fmt(datetime(year, month, day)), True, \
                       f"{fmt} normalised to ISO YYYY-MM-DD (day-first assumed for ambiguous dates)"
            except ValueError:
                continue
        return v, False, ""

    # DD-MM-YY or MM-DD-YY (two-digit year, hyphens) — prefer day-first, fall back to month-first
    m = re.match(r"^(\d{2})-(\d{2})-(\d{2})$", v)
    if m:
        a, b, yr = int(m.group(1)), int(m.group(2)), int(m.group(3))
        year = 2000 + yr if yr < 50 else 1900 + yr
        for day, month, fmt in [(a, b, "DD-MM-YY"), (b, a, "MM-DD-YY")]:
            try:
                return _fmt(datetime(year, month, day)), True, \
                       f"{fmt} normalised to ISO YYYY-MM-DD (day-first assumed; 20xx for year < 50)"
            except ValueError:
                continue
        return v, False, ""

    # Month DD YYYY or Month D YYYY
    m = re.match(r"^([A-Za-z]+)\s+(\d{1,2})\s+(\d{4})$", v)
    if m:
        month_num = MONTH_NAMES.get(m.group(1).lower())
        if month_num:
            try:
                return _fmt(datetime(int(m.group(3)), month_num, int(m.group(2)))), True, \
                       "Written-out month name normalised to ISO YYYY-MM-DD"
            except ValueError:
                pass
        return v, False, ""

    # Unix timestamp (10 digits)
    if re.match(r"^\d{10}$", v):
        try:
            dt = datetime.fromtimestamp(int(v), tz=timezone.utc)
            return _fmt(dt), True, "Unix timestamp (UTC) converted to YYYY-MM-DD"
        except (ValueError, OSError):
            return v, False, ""

    # Excel serial date (5-digit integer in plausible range)
    if re.match(r"^\d{5}$", v) and 40_000 <= int(v) <= 55_000:
        return _fmt(EXCEL_EPOCH + timedelta(days=int(v))), True, \
               "Excel serial date (Windows epoch 1899-12-30) converted to YYYY-MM-DD"

    return v, False, ""


_AMOUNT_NULL = {"n/a", "tbd", "-", "na", "nil", "none", ""}

def normalise_amount(value: str) -> tuple[str, bool, str]:
    v = value.strip()
    orig = v
    if v.lower() in _AMOUNT_NULL:
        result = ""
        return result, (result != orig), "Non-numeric placeholder left blank (N/A / TBD)"

    # Strip currency symbols and trailing ISO codes
    v = re.sub(r"[€£¥₹$]", "", v)
    v = re.sub(r"\s*(USD|EUR|GBP|INR|CAD|AUD)\s*$", "", v, flags=re.IGNORECASE).strip()

    # Negative accounting notation: (500) → -500
    m = re.match(r"^\(([0-9,. ]+)\)$", v)
    if m:
        v = "-" + m.group(1)

    # European decimal: 1.200,00  (period = thousands, comma = decimal)
    if "," in v and "." in v:
        if v.index(".") < v.index(","):
            v = v.replace(".", "").replace(",", ".")
            desc = "European decimal format (1.200,00) converted"
        else:
            v = v.replace(",", "")
            desc = "US thousands separator removed"
    elif "," in v:
        if re.search(r",\d{2}$", v):
            v = v.replace(",", ".")
            desc = "Comma decimal separator converted to period"
        else:
            v = v.replace(",", "")
            desc = "Thousands-separator comma removed"
    else:
        desc = "Amount normalised to 2 decimal places"

    try:
        result = f"{float(v):.2f}"
        return result, (result != orig), desc
    except ValueError:
        return orig, False, ""


CURRENCY_MAP = {
    "usd": "USD", "us dollar": "USD", "u.s. dollar": "USD", "dollar": "USD", "$": "USD",
    "eur": "EUR", "euro": "EUR", "€": "EUR",
    "gbp": "GBP", "pound": "GBP", "sterling": "GBP", "£": "GBP",
    "inr": "INR", "rupee": "INR", "indian rupee": "INR", "₹": "INR",
    "cad": "CAD", "canadian dollar": "CAD",
    "aud": "AUD", "australian dollar": "AUD",
}

def normalise_currency(value: str) -> tuple[str, bool, str]:
    v = value.strip()
    if not v:
        return v, False, ""
    cleaned = v.replace("₹", "").replace("€", "").replace("$", "").replace("£", "").strip()
    for lookup in (cleaned.lower(), v.lower()):
        if lookup in CURRENCY_MAP:
            result = CURRENCY_MAP[lookup]
            return result, result != v, f"Currency '{v}' standardised to ISO 3-letter code"
    if re.match(r"^[A-Z]{3}$", cleaned):
        return cleaned, cleaned != v, "Currency uppercased to ISO format"
    return v, False, ""


def split_amount_currency_fields(row: list[str], row_num: int) -> tuple[list[str], list[Change]]:
    return split_amount_currency_fields_dynamic(
        row,
        row_num,
        HEADERS,
        COL["Amount"],
        COL["Currency"],
    )


def normalise_name(value: str) -> tuple[str, bool, str]:
    v = " ".join(value.split())   # collapse multiple spaces
    if not v:
        return v, False, ""
    if "," in v:
        parts = [p.strip() for p in v.split(",", 1)]
        if len(parts) == 2 and parts[1]:
            v = f"{parts[1]} {parts[0]}"
    result = v.title()
    reasons = []
    if " ".join(value.split()) != value:
        reasons.append("extra whitespace collapsed")
    if "," in value:
        reasons.append("Last, First → First Last")
    if result != " ".join(value.split()).title():
        pass
    if result != value:
        if not reasons:
            reasons.append("name title-cased")
    return result, result != value, "Name normalised: " + "; ".join(reasons) if reasons else "Name title-cased"


STATUS_MAP = {
    "approved":       "Approved",
    "approve":        "Approved",
    "rejected":       "Rejected",
    "reject":         "Rejected",
    "pending":        "Pending",
    "pending review": "Pending",
}

def normalise_status(value: str) -> tuple[str, bool, str]:
    v = value.strip()
    result = STATUS_MAP.get(v.lower(), v.title() if v else v)
    reason  = f"Status '{v}' standardised to canonical form" if result != v else ""
    return result, result != v, reason


def apply_normalisations(row: list[str], row_num: int) -> tuple[list[str], list[Change]]:
    row     = row.copy()
    changes = []

    row, split_changes = split_amount_currency_fields(row, row_num)
    changes.extend(split_changes)

    def maybe_fix(idx, fn, col_label):
        orig = row[idx]
        new, changed, reason = fn(orig)
        if changed:
            row[idx] = new
            changes.append(Change(row_num, col_label, orig, new, "Fixed", reason))

    maybe_fix(COL["Date"],          normalise_date,     "Date")
    maybe_fix(COL["Amount"],        normalise_amount,   "Amount")
    maybe_fix(COL["Currency"],      normalise_currency, "Currency")
    maybe_fix(COL["Employee Name"], normalise_name,     "Employee Name")
    maybe_fix(COL["Status"],        normalise_status,   "Status")

    # Department — title-case + collapse whitespace
    orig = row[COL["Department"]]
    fixed_dept = " ".join(orig.split()).title()
    if fixed_dept != orig:
        row[COL["Department"]] = fixed_dept
        changes.append(Change(row_num, "Department", orig, fixed_dept,
                               "Fixed", "Department title-cased"))

    # Category — title-case
    orig = row[COL["Category"]]
    fixed_cat = orig.strip().title()
    if fixed_cat != orig:
        row[COL["Category"]] = fixed_cat
        changes.append(Change(row_num, "Category", orig, fixed_cat,
                               "Fixed", "Category title-cased"))

    return row, changes


# ══════════════════════════════════════════════════════════════════════════
# STEP 6 — needs_review decision
# ══════════════════════════════════════════════════════════════════════════

def needs_review(row: list[str], was_padded: bool) -> bool:
    """Row needs human review if: amount blank, date unparseable, or padded."""
    if not row[COL["Amount"]]:
        return True
    date_val = row[COL["Date"]]
    if date_val and not re.match(r"^\d{4}-\d{2}-\d{2}$", date_val):
        return True
    return was_padded


def forward_fill_merged_cell_gaps_generic(
    clean_data: list[CleanRow], changelog: list[Change], headers: list[str], fill_columns: list[int]
) -> None:
    def apply_gap_fill(gap_indexes: list[int], last_value: str, col_idx: int) -> None:
        if not last_value or not gap_indexes or len(gap_indexes) > 5:
            return
        for gap_idx in gap_indexes:
            gap_entry = clean_data[gap_idx]
            if gap_entry.row[col_idx].strip():
                continue
            gap_entry.row[col_idx] = last_value
            gap_entry.was_modified = True
            gap_entry.needs_review = True
            changelog.append(
                Change(
                    gap_entry.row_num,
                    headers[col_idx],
                    "",
                    last_value,
                    "Fixed",
                    "Blank categorical cell forward-filled to repair a merged-cell style export gap",
                )
            )

    for col_idx in fill_columns:
        last_value = ""
        gap_indexes: list[int] = []

        for idx, entry in enumerate(clean_data):
            cell_value = entry.row[col_idx].strip()
            other_filled = sum(1 for j, cell in enumerate(entry.row) if j != col_idx and cell.strip())

            if cell_value:
                if last_value and gap_indexes and len(gap_indexes) <= 5:
                    apply_gap_fill(gap_indexes, last_value, col_idx)
                last_value = cell_value
                gap_indexes = []
                continue

            if last_value and other_filled >= 2:
                gap_indexes.append(idx)
            else:
                gap_indexes = []

        apply_gap_fill(gap_indexes, last_value, col_idx)


def forward_fill_merged_cell_gaps(clean_data: list[CleanRow], changelog: list[Change]) -> None:
    forward_fill_merged_cell_gaps_generic(
        clean_data,
        changelog,
        HEADERS,
        [COL["Department"], COL["Category"], COL["Status"], COL["Currency"]],
    )

def split_amount_currency_fields_dynamic(
    row: list[str], row_num: int, headers: list[str], amount_idx: int | None, currency_idx: int | None
) -> tuple[list[str], list[Change]]:
    if amount_idx is None or currency_idx is None:
        return row, []

    row = row.copy()
    changes: list[Change] = []
    amount_label = headers[amount_idx]
    currency_label = headers[currency_idx]

    amount_val = row[amount_idx]
    currency_val = row[currency_idx]

    extracted_amount, extracted_currency = extract_currency_from_text(amount_val)
    if extracted_currency and (not currency_val.strip()):
        if extracted_amount and extracted_amount != amount_val:
            row[amount_idx] = extracted_amount
            changes.append(
                Change(
                    row_num,
                    amount_label,
                    amount_val,
                    extracted_amount,
                    "Fixed",
                    "Currency marker removed from amount-like field so the numeric value can be parsed cleanly",
                )
            )
        row[currency_idx] = extracted_currency
        changes.append(
            Change(
                row_num,
                currency_label,
                currency_val,
                extracted_currency,
                "Fixed",
                "Currency recovered from amount-like field",
            )
        )
        amount_val = row[amount_idx]
        currency_val = row[currency_idx]

    if not row[amount_idx].strip() and currency_val.strip():
        extracted_amount, extracted_currency = extract_currency_from_text(currency_val)
        if extracted_amount:
            original_currency = row[currency_idx]
            row[amount_idx] = extracted_amount
            changes.append(
                Change(
                    row_num,
                    amount_label,
                    "",
                    extracted_amount,
                    "Fixed",
                    "Amount recovered from currency-like field",
                )
            )
            if extracted_currency:
                row[currency_idx] = extracted_currency
                changes.append(
                    Change(
                        row_num,
                        currency_label,
                        original_currency,
                        extracted_currency,
                        "Fixed",
                        "Currency standardised after recovering combined amount/currency text",
                    )
                )

    return row, changes


def apply_semantic_normalisations(
    row: list[str], row_num: int, headers: list[str], semantic_plan: SemanticPlan
) -> tuple[list[str], list[Change]]:
    row = row.copy()
    changes: list[Change] = []

    row, split_changes = split_amount_currency_fields_dynamic(
        row, row_num, headers, semantic_plan.amount_idx, semantic_plan.currency_idx
    )
    changes.extend(split_changes)

    for idx, role in semantic_plan.roles_by_index.items():
        original = row[idx]
        if role == "date":
            new_value, changed, reason = normalise_date(original)
        elif role == "amount":
            new_value, changed, reason = normalise_amount(original)
        elif role == "identifier":
            new_value = " ".join(original.split()) if original.strip() else ""
            changed = new_value != original
            reason = "Identifier spacing normalised"
        elif role == "measurement":
            new_value = " ".join(original.split()) if original.strip() else ""
            changed = new_value != original
            reason = "Measurement text spacing normalised"
        elif role == "currency":
            new_value, changed, reason = normalise_currency(original)
        elif role == "name":
            new_value, changed, reason = normalise_name(original)
        elif role == "status":
            new_value, changed, reason = normalise_status(original)
        elif role in {"department", "category"}:
            new_value = " ".join(original.split()).title() if original.strip() else ""
            changed = new_value != original
            reason = f"{role.title()} title-cased"
        else:
            continue

        if changed:
            row[idx] = new_value
            changes.append(Change(row_num, headers[idx], original, new_value, "Fixed", reason))

    return row, changes

def needs_review_semantic(row: list[str], structure_changed: bool, semantic_plan: SemanticPlan) -> bool:
    if structure_changed:
        return True
    if semantic_plan.amount_idx is not None and not row[semantic_plan.amount_idx].strip():
        return True
    if semantic_plan.date_idx is not None:
        date_value = row[semantic_plan.date_idx].strip()
        if date_value and not re.match(r"^\d{4}-\d{2}-\d{2}$", date_value):
            return True
    review_tokens = {"nan", "null", "n/a", "na", "not applicable", "none", "inf"}
    return any(cell.strip().lower() in review_tokens for cell in row if cell.strip())


def needs_review_generic(row: list[str], structure_changed: bool) -> bool:
    if structure_changed:
        return True
    review_tokens = {"nan", "null", "n/a", "na", "not applicable", "none", "inf"}
    return any(cell.strip().lower() in review_tokens for cell in row if cell.strip())
