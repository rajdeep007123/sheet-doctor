from __future__ import annotations

import re
from collections import Counter
from dataclasses import dataclass
from datetime import datetime

HEADERS = [
    "Employee Name",
    "Department",
    "Date",
    "Amount",
    "Currency",
    "Category",
    "Status",
    "Notes",
]
N_COLS = len(HEADERS)
COL = {name: i for i, name in enumerate(HEADERS)}

SPARSE_THRESHOLD_SCHEMA = 0.50
SPARSE_THRESHOLD_GENERIC = 0.25

WRITE_ONLY_THRESHOLD = 5_000
LARGE_FILE_SKIP_EXTRAS = 10_000

VALID_SEMANTIC_ROLES = (
    "identifier",
    "name",
    "date",
    "amount",
    "measurement",
    "currency",
    "status",
    "department",
    "category",
    "notes",
)

FORMULA_RE = re.compile(r"^\s*=")
TOTAL_LABEL_RE = re.compile(r"\b(grand\s+total|subtotal|sub-total|total|sum)\b", re.IGNORECASE)
NOTES_ROW_RE = re.compile(r"\b(approved|manager|note|comment|memo|generated|report|expense|expenses)\b", re.IGNORECASE)
CURRENCY_SYMBOL_MAP = {"$": "USD", "€": "EUR", "£": "GBP", "₹": "INR", "¥": "JPY"}

STATUS_MAP = {
    "approved": "Approved",
    "approve": "Approved",
    "rejected": "Rejected",
    "reject": "Rejected",
    "pending": "Pending",
    "pending review": "Pending",
}
STATUS_VALUE_HINTS = set(STATUS_MAP.keys())

ROLE_HEADER_HINTS = {
    "identifier": ("id", "code", "study id", "study_id", "pat_id", "patient id", "subject id", "record id"),
    "name": ("name", "person", "contact"),
    "date": ("date", "dated", "txn date", "transaction", "invoice date", "posted", "dob", "dofb"),
    "amount": ("amount", "cost", "price", "value", "expense", "spend", "salary", "pay", "total"),
    "measurement": ("bp", "hr", "gfr", "glucose", "weight", "height", "score", "rate", "result", "reading", "pre", "post"),
    "currency": ("currency", "curr", "fx", "ccy"),
    "status": ("status", "state", "approval", "approved", "decision"),
    "department": ("department", "dept", "division", "team", "unit", "function", "ward", "location", "clinic"),
    "category": ("category", "type", "class", "group", "bucket", "expense type", "race", "sex", "ethnicity", "hispanic", "diagnosis", "sediment"),
    "notes": ("notes", "note", "comment", "comments", "description", "details", "memo", "remarks"),
}

SMART_QUOTES = {
    "\u201c": '"',
    "\u201d": '"',
    "\u2018": "'",
    "\u2019": "'",
}

GENERIC_QUARANTINE_REASONS = {
    "EMPTY": "Completely empty row",
    "WHITESPACE": "Row is all whitespace",
    "STRUCTURAL_HEADER": "Structural row (TOTAL/subtotal/header repeat)",
    "STRUCTURAL_TOTAL": "Structural row (TOTAL/subtotal/header repeat)",
    "CALCULATED_SUBTOTAL": "Calculated subtotal row",
    "NOTES_ROW": "Appears to be a notes row",
    "FORMULA": "Excel formula found, not data",
    "SPARSE": f"Less than {int(SPARSE_THRESHOLD_SCHEMA * 100)}% columns filled",
}

ASSUMPTIONS = [
    "Rows that still contain Excel formulas as text (for example '=SUM(...)') are quarantined because they are not stable data values",
    "Ambiguous DD/MM vs MM/DD dates: day-first assumed; fallback to month-first if day-first is impossible",
    "DD-MM-YY / MM-DD-YY two-digit years: day-first assumed; fallback to month-first if impossible; 2000–2049 for YY < 50, 1950–1999 for YY ≥ 50",
    "Unix timestamps: interpreted as UTC; converted to YYYY-MM-DD",
    "Excel serial dates: Windows epoch (1899-12-30); range 40,000–55,000 treated as dates",
    "European decimal 1.200,00: detected when period precedes comma; converted to 1200.00",
    "Combined values like '$1,200 USD' are split so Amount keeps the numeric value and Currency keeps the ISO code",
    "Blank / N/A / TBD amounts: cleared to empty string and flagged needs_review=TRUE",
    '"INR ₹" and similar symbol+code combos: symbol stripped, ISO code kept',
    '"Eng" department abbreviation: kept as-is (expanding abbreviations requires a lookup table)',
    "Short blank runs in categorical columns are forward-filled when they look like merged-cell export gaps; filled rows are flagged for review",
    "Rows with long single-cell prose are treated as notes/metadata rather than transactional data",
    "Rows labelled TOTAL/Subtotal/SUM are quarantined when the amount matches the running total closely enough to look calculated",
    "Near-duplicate rows (same Name/Amount/Currency/Category, date within 2 days): both kept, both flagged",
    "Exact duplicates: first occurrence kept; subsequent occurrences removed and logged",
    "Short rows (< 8 columns): padded with empty strings; flagged needs_review=TRUE",
    f"Metadata export row (sparse, 1/8 fields filled): quarantined as 'Less than {int(SPARSE_THRESHOLD_SCHEMA * 100)}% columns filled'",
]

GENERIC_ASSUMPTIONS = [
    "Rows that still contain Excel formulas as text are quarantined because they are not stable data values",
    "Delimiter is auto-detected from the file content (comma/semicolon/tab/pipe)",
    "Rows with overflow columns are repaired by merging overflow into the last column",
    "Rows with missing trailing columns are padded with empty strings",
    "Repeated header rows and subtotal/total structural rows are quarantined",
    "Long one-cell prose rows are treated as notes rather than tabular data",
    "Rows before a detected header are moved to File Metadata entries in the Change Log",
    "BOM/null bytes/line breaks/smart quotes are normalised in text cells",
    "Exact duplicate rows are removed (first occurrence kept)",
]

SEMANTIC_ASSUMPTIONS = GENERIC_ASSUMPTIONS + [
    "When headers are non-standard, likely semantic roles are inferred from the column values and header hints",
    "Date-like columns are normalised to YYYY-MM-DD when confidence is high enough",
    "Amount-like and currency-like columns are normalised even when their headers are not exact schema matches",
    "Status-like, department-like, and category-like columns are title-cased or canonicalised when their semantics are clear enough",
    "Near-duplicate detection uses inferred semantic key columns instead of fixed schema names",
]


@dataclass
class Change:
    original_row_number: int
    column_affected: str
    original_value: str
    new_value: str
    action_taken: str
    reason: str


@dataclass
class CleanRow:
    row: list
    row_num: int
    was_modified: bool
    needs_review: bool


@dataclass
class QuarantineRow:
    row: list
    row_num: int
    reason: str


@dataclass
class SemanticPlan:
    enabled: bool
    roles_by_index: dict[int, str]
    confidence_by_index: dict[int, float]
    label_idx: int
    amount_idx: int | None
    currency_idx: int | None
    date_idx: int | None
    fill_down_indices: list[int]


def _strip_nulls(value: object) -> str:
    if value is None:
        return ""
    return str(value).replace("\x00", "")


def _normalise_header_for_match(value: str) -> str:
    return " ".join(value.strip().lower().split())


def is_schema_specific_header(header_row: list[str]) -> bool:
    if len(header_row) != N_COLS:
        return False
    cleaned = tuple(_normalise_header_for_match(c or "") for c in header_row)
    expected = tuple(_normalise_header_for_match(c) for c in HEADERS)
    return cleaned == expected
