#!/usr/bin/env python3
"""
heal.py — CSV Doctor Healer  (v2)

Reads a messy CSV and writes a 3-sheet Excel workbook:

  Sheet 1 — "Clean Data"    rows that were fixed and are ready to use
  Sheet 2 — "Quarantine"    rows that could not be fixed or are unusable
  Sheet 3 — "Change Log"    one entry per individual change made

Usage:
    python skills/csv-doctor/scripts/heal.py [input.csv [output.xlsx]]

Exit codes:
    0 — completed
    1 — input file not found or unreadable
"""

import csv
import io
import re
import sys
from dataclasses import dataclass
from datetime import datetime, timedelta, timezone
from pathlib import Path

try:
    import openpyxl
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter
except ImportError:
    print("ERROR: openpyxl not installed — run: pip install openpyxl", file=sys.stderr)
    sys.exit(1)

# ── Paths ─────────────────────────────────────────────────────────────────
HERE  = Path(__file__).parent
ROOT  = HERE.parent.parent.parent
INPUT  = ROOT / "sample-data" / "extreme_mess.csv"
OUTPUT = ROOT / "sample-data" / "extreme_mess_healed.xlsx"

# ── Schema ────────────────────────────────────────────────────────────────
HEADERS = ["Employee Name", "Department", "Date", "Amount", "Currency",
           "Category", "Status", "Notes"]
N_COLS  = len(HEADERS)
COL     = {name: i for i, name in enumerate(HEADERS)}


# ══════════════════════════════════════════════════════════════════════════
# DATA CLASSES
# ══════════════════════════════════════════════════════════════════════════

@dataclass
class Change:
    original_row_number: int
    column_affected:     str
    original_value:      str
    new_value:           str
    action_taken:        str   # Fixed | Quarantined | Flagged | Removed
    reason:              str

@dataclass
class CleanRow:
    row:          list
    row_num:      int
    was_modified: bool
    needs_review: bool

@dataclass
class QuarantineRow:
    row:    list
    row_num: int
    reason: str


# ══════════════════════════════════════════════════════════════════════════
# STEP 1 — Read with mixed-encoding tolerance
# ══════════════════════════════════════════════════════════════════════════

def read_mixed_encoding(path: Path) -> list[list[str]]:
    """
    Files with mixed Latin-1 and UTF-8 bytes: split on raw newlines, try
    UTF-8 per line, fall back to Latin-1. Rejoin and re-parse with
    csv.reader so multi-line quoted fields are reconstructed correctly.
    """
    with open(path, "rb") as f:
        raw = f.read()
    lines = []
    for line in raw.split(b"\n"):
        try:
            lines.append(line.decode("utf-8"))
        except UnicodeDecodeError:
            lines.append(line.decode("latin-1"))
    return list(csv.reader(io.StringIO("\n".join(lines))))


# ══════════════════════════════════════════════════════════════════════════
# STEP 2 — Row classification
# ══════════════════════════════════════════════════════════════════════════

QUARANTINE_REASONS = {
    "EMPTY":               "Completely empty row",
    "WHITESPACE":          "Row is all whitespace",
    "STRUCTURAL_HEADER":   "Structural row (TOTAL/subtotal/header repeat)",
    "STRUCTURAL_TOTAL":    "Structural row (TOTAL/subtotal/header repeat)",
    "SPARSE":              "Less than 50% columns filled",
}

def classify_raw_row(row: list[str], header_sig: tuple[str, ...]) -> str:
    stripped = [c.strip() for c in row]

    if all(c == "" for c in stripped):
        # Distinguish truly empty (all "") from whitespace-only
        return "WHITESPACE" if any(c != "" for c in row) else "EMPTY"

    if tuple(c.lower() for c in stripped) == header_sig:
        return "STRUCTURAL_HEADER"

    if stripped[0].upper() == "TOTAL" and all(c == "" for c in stripped[1:3]):
        return "STRUCTURAL_TOTAL"

    non_empty = sum(1 for c in stripped if c)
    if non_empty < N_COLS * 0.5:
        return "SPARSE"

    return "NORMAL"


# ══════════════════════════════════════════════════════════════════════════
# STEP 3 — Alignment fix (with Change log entry)
# ══════════════════════════════════════════════════════════════════════════

def fix_alignment(row: list[str], row_num: int) -> tuple[list[str], Change | None]:
    """Detect and fix three structural column problems."""
    n = len(row)
    if n == N_COLS:
        return row, None

    if n > N_COLS:
        # Shifted right: leading empty ghost column
        if row[0].strip() == "" and row[1].strip() != "":
            fixed = row[1: N_COLS + 1]
            return fixed, Change(
                row_num, "[row structure]",
                f"{n} columns (empty leading ghost col)",
                f"{N_COLS} columns",
                "Fixed", "Shifted-right row: empty leading column stripped"
            )
        # Phantom comma: empty ghost field sits between Status and Notes
        if n == N_COLS + 1 and row[N_COLS - 1].strip() == "" and row[N_COLS].strip() != "":
            fixed = row[: N_COLS - 1] + [row[N_COLS]]
            return fixed, Change(
                row_num, "Notes",
                f"[ghost field] + '{row[N_COLS]}'",
                row[N_COLS],
                "Fixed", "Phantom comma: empty ghost field before Notes removed"
            )
        # Unquoted commas in Notes: merge overflow columns back
        merged = ", ".join(row[N_COLS - 1:]).strip()
        fixed = row[: N_COLS - 1] + [merged]
        return fixed, Change(
            row_num, "Notes",
            f"{n - N_COLS + 1} fragments: {row[N_COLS - 1]!r}...",
            merged,
            "Fixed", "Unquoted commas in Notes field: fragments merged into one"
        )

    # Short row: pad with empty strings
    fixed = row + [""] * (N_COLS - n)
    return fixed, Change(
        row_num, "[row structure]",
        f"{n} columns",
        f"{N_COLS} columns ({N_COLS - n} empty field(s) appended)",
        "Fixed", f"Short row padded with {N_COLS - n} empty field(s)"
    )


# ══════════════════════════════════════════════════════════════════════════
# STEP 4 — Cell-level cleaning (BOM, null, line breaks, smart quotes)
# ══════════════════════════════════════════════════════════════════════════

SMART_QUOTES = {
    "\u201c": '"', "\u201d": '"',   # "" double curly
    "\u2018": "'", "\u2019": "'",   # '' single curly / smart apostrophe
}

def _col_name(i: int) -> str:
    return HEADERS[i] if i < N_COLS else f"[col {i + 1}]"

def clean_row(row: list[str], row_num: int) -> tuple[list[str], list[Change]]:
    cleaned, changes = [], []
    for i, cell in enumerate(row):
        orig = cell
        val  = cell
        reasons = []

        if "\ufeff" in val:
            val = val.replace("\ufeff", "")
            reasons.append("BOM byte-order mark stripped")

        if "\x00" in val:
            val = val.replace("\x00", "")
            reasons.append("Null byte removed")

        if "\n" in val or "\r" in val:
            val = val.replace("\r\n", " ").replace("\n", " ").replace("\r", " ")
            reasons.append("Embedded line break replaced with space")

        had_smart = any(s in val for s in SMART_QUOTES)
        for smart, straight in SMART_QUOTES.items():
            val = val.replace(smart, straight)
        if had_smart:
            reasons.append("Smart/curly quotes normalised to straight quotes")

        val = val.strip()

        if reasons:
            orig_display = (orig
                            .replace("\ufeff", "[BOM]")
                            .replace("\x00", "[NULL]")
                            .strip())
            changes.append(Change(
                row_num, _col_name(i),
                orig_display, val,
                "Fixed", "; ".join(reasons)
            ))
        cleaned.append(val)
    return cleaned, changes


# ══════════════════════════════════════════════════════════════════════════
# STEP 5 — Value normalisation
# ══════════════════════════════════════════════════════════════════════════

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

    # MM-DD-YY (two-digit year, hyphens)
    m = re.match(r"^(\d{2})-(\d{2})-(\d{2})$", v)
    if m:
        mo, day, yr = int(m.group(1)), int(m.group(2)), int(m.group(3))
        year = 2000 + yr if yr < 50 else 1900 + yr
        try:
            return _fmt(datetime(year, mo, day)), True, \
                   f"MM-DD-YY normalised (20xx assumed for year < 50)"
        except ValueError:
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


# ══════════════════════════════════════════════════════════════════════════
# MAIN PROCESSING LOOP
# ══════════════════════════════════════════════════════════════════════════

def process(all_rows: list[list[str]]) -> tuple[list[CleanRow], list[QuarantineRow], list[Change]]:
    header_sig = tuple(c.strip().lower() for c in all_rows[0])

    clean_data: list[CleanRow]       = []
    quarantine: list[QuarantineRow]  = []
    changelog:  list[Change]         = []
    seen_exact: dict[tuple, int]     = {}   # normalized row → original row_num

    # Skip row 0 (actual column headers); start from row 1 (metadata / first data row)
    data_rows = all_rows[1:]

    for i, raw_row in enumerate(data_rows):
        row_num = i + 2   # 1-based; header = row 1

        # ── Classify on raw values ────────────────────────────────────────
        cls = classify_raw_row(raw_row, header_sig)

        if cls != "NORMAL":
            q_reason = QUARANTINE_REASONS[cls]
            # Light-clean the row for display in Quarantine tab
            q_row = [c.strip() for c in raw_row]
            q_row = (q_row + [""] * N_COLS)[:N_COLS]          # normalise length
            row_id = next((c for c in q_row if c), "[empty]")
            quarantine.append(QuarantineRow(q_row, row_num, q_reason))
            changelog.append(Change(
                row_num, "[row]", row_id[:60], "",
                "Quarantined", q_reason
            ))
            continue

        # ── Fix alignment ─────────────────────────────────────────────────
        aligned, align_chg = fix_alignment(raw_row, row_num)
        was_padded = align_chg is not None and "padded" in align_chg.reason.lower()
        if align_chg:
            changelog.append(align_chg)

        # ── Clean cells ───────────────────────────────────────────────────
        cleaned, cell_chgs = clean_row(aligned, row_num)
        changelog.extend(cell_chgs)

        # ── Normalise values ──────────────────────────────────────────────
        fixed, norm_chgs = apply_normalisations(cleaned, row_num)
        changelog.extend(norm_chgs)

        was_modified = bool(align_chg or cell_chgs or norm_chgs)

        # ── Exact-duplicate removal ───────────────────────────────────────
        row_key = tuple(fixed)
        if row_key in seen_exact:
            first_num = seen_exact[row_key]
            changelog.append(Change(
                row_num, "[row]", fixed[0], "",
                "Removed", f"Exact duplicate of row {first_num}"
            ))
            continue
        seen_exact[row_key] = row_num

        clean_data.append(CleanRow(
            row          = fixed,
            row_num      = row_num,
            was_modified = was_modified,
            needs_review = needs_review(fixed, was_padded),
        ))

    # ── Near-duplicate detection (second pass on clean_data) ─────────────
    nd_index: dict[tuple, int] = {}   # key → index in clean_data
    for idx, entry in enumerate(clean_data):
        r = entry.row
        key = (r[COL["Employee Name"]], r[COL["Amount"]],
               r[COL["Currency"]],     r[COL["Category"]])
        if key in nd_index:
            j       = nd_index[key]
            prev    = clean_data[j]
            d1, d2  = prev.row[COL["Date"]], entry.row[COL["Date"]]
            if (d1 and d2
                    and re.match(r"^\d{4}-\d{2}-\d{2}$", d1)
                    and re.match(r"^\d{4}-\d{2}-\d{2}$", d2)):
                try:
                    delta = abs((datetime.strptime(d2, "%Y-%m-%d")
                                 - datetime.strptime(d1, "%Y-%m-%d")).days)
                    if delta <= 2:
                        prev.needs_review  = True
                        entry.needs_review = True
                        for flagged, other_date, other_row_num in [
                            (entry, d1, prev.row_num),
                            (prev,  d2, entry.row_num),
                        ]:
                            changelog.append(Change(
                                flagged.row_num, "[row]",
                                flagged.row[COL["Employee Name"]], "",
                                "Flagged",
                                f"Near-duplicate: same Name/Amount/Currency/Category; "
                                f"date {flagged.row[COL['Date']]} differs by {delta} day(s) "
                                f"from row {other_row_num} ({other_date})"
                            ))
                except ValueError:
                    pass
        else:
            nd_index[key] = idx

    return clean_data, quarantine, changelog


# ══════════════════════════════════════════════════════════════════════════
# EXCEL OUTPUT
# ══════════════════════════════════════════════════════════════════════════

def _header_font() -> Font:
    return Font(bold=True, color="FFFFFF")

def _header_fill(hex_color: str) -> PatternFill:
    return PatternFill("solid", fgColor=hex_color)

def _style_sheet(ws, col_widths: list[int], header_color: str):
    """Apply bold header, color, frozen row, and column widths."""
    fill = _header_fill(header_color)
    font = _header_font()
    for cell in ws[1]:
        cell.font  = font
        cell.fill  = fill
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=False)
    ws.freeze_panes = "A2"
    for i, width in enumerate(col_widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = width

# Accent fills for was_modified / needs_review cells
FILL_MODIFIED = PatternFill("solid", fgColor="FFF2CC")   # soft yellow
FILL_REVIEW   = PatternFill("solid", fgColor="FCE4D6")   # soft orange

def write_workbook(
    clean_data:  list[CleanRow],
    quarantine:  list[QuarantineRow],
    changelog:   list[Change],
    output_path: Path,
) -> None:
    wb = openpyxl.Workbook()

    # ── Sheet 1 — Clean Data ─────────────────────────────────────────────
    ws1 = wb.active
    ws1.title = "Clean Data"
    clean_headers = HEADERS + ["was_modified", "needs_review"]
    ws1.append(clean_headers)
    for entry in clean_data:
        row_out = entry.row + [entry.was_modified, entry.needs_review]
        ws1.append(row_out)
        # Accent modified / review flag cells
        last = ws1.max_row
        mod_cell    = ws1.cell(last, len(HEADERS) + 1)
        review_cell = ws1.cell(last, len(HEADERS) + 2)
        if entry.was_modified:
            mod_cell.fill = FILL_MODIFIED
        if entry.needs_review:
            review_cell.fill = FILL_REVIEW
    _style_sheet(ws1,
                 [22, 18, 13, 11, 11, 16, 12, 40, 14, 13],
                 "4CAF50")   # green

    # ── Sheet 2 — Quarantine ─────────────────────────────────────────────
    ws2 = wb.create_sheet("Quarantine")
    ws2.append(HEADERS + ["quarantine_reason"])
    for q in quarantine:
        ws2.append(q.row + [q.reason])
    _style_sheet(ws2,
                 [22, 18, 13, 11, 11, 16, 12, 35, 35],
                 "E53935")   # red

    # ── Sheet 3 — Change Log ─────────────────────────────────────────────
    ws3 = wb.create_sheet("Change Log")
    log_headers = ["original_row_number", "column_affected",
                   "original_value", "new_value", "action_taken", "reason"]
    ws3.append(log_headers)
    for c in changelog:
        ws3.append([c.original_row_number, c.column_affected,
                    c.original_value, c.new_value, c.action_taken, c.reason])
    _style_sheet(ws3,
                 [22, 20, 35, 35, 14, 50],
                 "1565C0")   # blue

    # Notes column text-wrap in Sheet 1
    notes_col = get_column_letter(COL["Notes"] + 1)
    for cell in ws1[notes_col][1:]:
        cell.alignment = Alignment(wrap_text=True, vertical="top")

    # Reason column text-wrap in Sheet 3
    for cell in ws3["F"][1:]:
        cell.alignment = Alignment(wrap_text=True, vertical="top")

    output_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(output_path)


# ══════════════════════════════════════════════════════════════════════════
# MAIN
# ══════════════════════════════════════════════════════════════════════════

ASSUMPTIONS = [
    "Ambiguous DD/MM vs MM/DD dates: day-first assumed; fallback to month-first if day-first is impossible",
    "MM-DD-YY two-digit years: treated as 2000–2049 for YY < 50, 1950–1999 for YY ≥ 50",
    "Unix timestamps: interpreted as UTC; converted to YYYY-MM-DD",
    "Excel serial dates: Windows epoch (1899-12-30); range 40,000–55,000 treated as dates",
    "European decimal 1.200,00: detected when period precedes comma; converted to 1200.00",
    "Blank / N/A / TBD amounts: cleared to empty string and flagged needs_review=TRUE",
    "\"INR ₹\" and similar symbol+code combos: symbol stripped, ISO code kept",
    "\"Eng\" department abbreviation: kept as-is (expanding abbreviations requires a lookup table)",
    "Near-duplicate rows (same Name/Amount/Currency/Category, date within 2 days): both kept, both flagged",
    "Exact duplicates: first occurrence kept; subsequent occurrences removed and logged",
    "Short rows (< 8 columns): padded with empty strings; flagged needs_review=TRUE",
    "Metadata export row (sparse, 1/8 fields filled): quarantined as 'Less than 50% columns filled'",
]

def main():
    input_path  = Path(sys.argv[1]) if len(sys.argv) > 1 else INPUT
    output_path = Path(sys.argv[2]) if len(sys.argv) > 2 else OUTPUT

    if not input_path.exists():
        print(f"ERROR: File not found: {input_path}", file=sys.stderr)
        sys.exit(1)

    all_rows = read_mixed_encoding(input_path)
    if len(all_rows) < 2:
        print("ERROR: File is empty or has only a header.", file=sys.stderr)
        sys.exit(1)

    total_in = len(all_rows)   # includes the actual header row

    clean_data, quarantine, changelog = process(all_rows)

    write_workbook(clean_data, quarantine, changelog, output_path)

    # ── Counts by action type ─────────────────────────────────────────────
    from collections import Counter
    action_counts = Counter(c.action_taken for c in changelog)

    W = 60
    print()
    print("═" * W)
    print("  CSV Doctor  ·  Heal Report  (Excel output)")
    print("═" * W)
    print(f"  Input file   : {input_path.name}")
    print(f"  Output file  : {output_path.name}")
    print("─" * W)
    print(f"  Rows in      : {total_in}  (incl. column header row)")
    print(f"  Clean Data   : {len(clean_data)} rows")
    print(f"    · was_modified = TRUE  : {sum(1 for r in clean_data if r.was_modified)}")
    print(f"    · needs_review = TRUE  : {sum(1 for r in clean_data if r.needs_review)}")
    print(f"  Quarantine   : {len(quarantine)} rows")
    for reason, rows in sorted(
        {r.reason: sum(1 for q in quarantine if q.reason == r.reason)
         for r in quarantine}.items()
    ):
        print(f"    · {reason:<40} {rows}")
    print(f"  Changes logged: {len(changelog)}")
    print(f"    · Fixed       : {action_counts.get('Fixed', 0)}")
    print(f"    · Quarantined : {action_counts.get('Quarantined', 0)}")
    print(f"    · Removed     : {action_counts.get('Removed', 0)}")
    print(f"    · Flagged     : {action_counts.get('Flagged', 0)}")
    print("─" * W)
    print("  ASSUMPTIONS MADE:")
    for a in ASSUMPTIONS:
        print(f"    · {a}")
    print("═" * W)
    print()


if __name__ == "__main__":
    main()
