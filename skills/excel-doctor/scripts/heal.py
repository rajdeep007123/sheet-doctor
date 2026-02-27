#!/usr/bin/env python3
"""
excel-doctor heal.py

Applies safe, deterministic cleanup passes to an Excel workbook and writes:
1) A healed workbook with normalised sheets
2) A "Change Log" sheet recording every change

Usage:
    python skills/excel-doctor/scripts/heal.py <input.xlsx> [output.xlsx]
"""

from __future__ import annotations

import re
import sys
from collections import Counter
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path

try:
    from openpyxl import load_workbook
except ImportError:
    print("ERROR: openpyxl not installed — run: pip install openpyxl", file=sys.stderr)
    sys.exit(1)


SMART_QUOTES = {
    "\u201c": '"',
    "\u201d": '"',
    "\u2018": "'",
    "\u2019": "'",
}
INTERNAL_SHEETS = {"Change Log"}


@dataclass
class Change:
    sheet: str
    cell_or_range: str
    action: str
    old_value: str
    new_value: str
    reason: str


def is_blank(value) -> bool:
    return value is None or (isinstance(value, str) and value.strip() == "")


def is_sheet_empty(sheet) -> bool:
    for row in sheet.iter_rows(values_only=True):
        if any(not is_blank(value) for value in row):
            return False
    return True


def text(value) -> str:
    if value is None:
        return ""
    return str(value)


def clean_text(value: str) -> tuple[str, list[str]]:
    new_value = value
    reasons = []

    if "\ufeff" in new_value:
        new_value = new_value.replace("\ufeff", "")
        reasons.append("BOM removed")
    if "\x00" in new_value:
        new_value = new_value.replace("\x00", "")
        reasons.append("NULL byte removed")
    if "\r" in new_value or "\n" in new_value:
        new_value = new_value.replace("\r\n", " ").replace("\n", " ").replace("\r", " ")
        reasons.append("line breaks replaced with spaces")

    had_smart_quotes = any(char in new_value for char in SMART_QUOTES)
    for smart, straight in SMART_QUOTES.items():
        new_value = new_value.replace(smart, straight)
    if had_smart_quotes:
        reasons.append("smart quotes normalised")

    collapsed = " ".join(new_value.split())
    if collapsed != new_value:
        reasons.append("extra whitespace collapsed")
    new_value = collapsed

    return new_value, reasons


def normalise_date_text(value: str) -> tuple[str, bool, str]:
    v = value.strip()
    if not v:
        return v, False, ""

    if re.match(r"^\d{4}-\d{2}-\d{2}$", v):
        return v, False, ""

    if re.match(r"^\d{4}/\d{2}/\d{2}$", v):
        try:
            dt = datetime.strptime(v, "%Y/%m/%d")
            return dt.strftime("%Y-%m-%d"), True, "YYYY/MM/DD normalised to YYYY-MM-DD"
        except ValueError:
            return v, False, ""

    m = re.match(r"^(\d{1,2})/(\d{1,2})/(\d{4})$", v)
    if m:
        first, second, year = int(m.group(1)), int(m.group(2)), int(m.group(3))
        for day, month, label in [
            (first, second, "DD/MM/YYYY"),
            (second, first, "MM/DD/YYYY"),
        ]:
            try:
                dt = datetime(year, month, day)
                return (
                    dt.strftime("%Y-%m-%d"),
                    True,
                    f"{label} normalised to YYYY-MM-DD (day-first preferred when ambiguous)",
                )
            except ValueError:
                continue
        return v, False, ""

    m = re.match(r"^(\d{1,2})-(\d{1,2})-(\d{4})$", v)
    if m:
        first, second, year = int(m.group(1)), int(m.group(2)), int(m.group(3))
        for day, month, label in [
            (first, second, "DD-MM-YYYY"),
            (second, first, "MM-DD-YYYY"),
        ]:
            try:
                dt = datetime(year, month, day)
                return (
                    dt.strftime("%Y-%m-%d"),
                    True,
                    f"{label} normalised to YYYY-MM-DD (day-first preferred when ambiguous)",
                )
            except ValueError:
                continue
        return v, False, ""

    m = re.match(r"^(\d{2})-(\d{2})-(\d{2})$", v)
    if m:
        month, day, year_short = int(m.group(1)), int(m.group(2)), int(m.group(3))
        year = 2000 + year_short if year_short < 50 else 1900 + year_short
        try:
            dt = datetime(year, month, day)
            return dt.strftime("%Y-%m-%d"), True, "MM-DD-YY normalised to YYYY-MM-DD"
        except ValueError:
            return v, False, ""

    m = re.match(r"^(\d{1,2})/(\d{1,2})/(\d{2})$", v)
    if m:
        first, second, year_short = int(m.group(1)), int(m.group(2)), int(m.group(3))
        year = 2000 + year_short if year_short < 50 else 1900 + year_short
        for day, month, label in [
            (first, second, "DD/MM/YY"),
            (second, first, "MM/DD/YY"),
        ]:
            try:
                dt = datetime(year, month, day)
                return (
                    dt.strftime("%Y-%m-%d"),
                    True,
                    f"{label} normalised to YYYY-MM-DD (day-first preferred when ambiguous)",
                )
            except ValueError:
                continue
        return v, False, ""

    for fmt, label in [
        ("%B %d %Y", "Month D YYYY"),
        ("%b %d %Y", "Mon D YYYY"),
        ("%B %d, %Y", "Month D, YYYY"),
        ("%b %d, %Y", "Mon D, YYYY"),
    ]:
        try:
            dt = datetime.strptime(v, fmt)
            return dt.strftime("%Y-%m-%d"), True, f"{label} normalised to YYYY-MM-DD"
        except ValueError:
            continue

    return v, False, ""


def normalise_header(raw_header, index: int) -> str:
    base = text(raw_header).strip()
    base = " ".join(base.split())
    if not base:
        return f"column_{index}"
    return base


def add_change(
    changes: list[Change],
    sheet_name: str,
    target: str,
    action: str,
    old_value,
    new_value,
    reason: str,
) -> None:
    changes.append(
        Change(
            sheet=sheet_name,
            cell_or_range=target,
            action=action,
            old_value=text(old_value),
            new_value=text(new_value),
            reason=reason,
        )
    )


def heal_sheet(sheet, changes: list[Change], stats: Counter) -> None:
    sheet_name = sheet.title
    if is_sheet_empty(sheet):
        stats["empty_sheets_skipped"] += 1
        return

    # 1) Unmerge ranges and fill all cells with the anchor value.
    merged_ranges = list(sheet.merged_cells.ranges)
    for merged in merged_ranges:
        min_col, min_row, max_col, max_row = merged.bounds
        range_ref = str(merged)
        anchor_value = sheet.cell(row=min_row, column=min_col).value

        sheet.unmerge_cells(range_ref)
        stats["merged_ranges_unmerged"] += 1
        add_change(
            changes,
            sheet_name,
            range_ref,
            "Fixed",
            "[merged]",
            "[unmerged]",
            "Merged range unmerged for tabular compatibility",
        )

        for row in range(min_row, max_row + 1):
            for col in range(min_col, max_col + 1):
                if row == min_row and col == min_col:
                    continue
                cell = sheet.cell(row=row, column=col)
                old = cell.value
                if old != anchor_value:
                    cell.value = anchor_value
                    stats["cells_filled_from_merged_anchor"] += 1
                    add_change(
                        changes,
                        sheet_name,
                        cell.coordinate,
                        "Fixed",
                        old,
                        anchor_value,
                        f"Filled from merged anchor cell {sheet.cell(min_row, min_col).coordinate}",
                    )

    max_col = sheet.max_column
    if max_col == 0:
        return

    # 2) Standardise and deduplicate headers.
    seen_headers = Counter()
    for col in range(1, max_col + 1):
        cell = sheet.cell(row=1, column=col)
        old = cell.value
        header = normalise_header(old, col)

        key = header.lower()
        seen_headers[key] += 1
        if seen_headers[key] > 1:
            deduped = f"{header}_{seen_headers[key]}"
            stats["headers_deduplicated"] += 1
            reason = "Duplicate header renamed with numeric suffix"
        else:
            deduped = header
            reason = "Header standardised"

        if text(old) != deduped:
            cell.value = deduped
            stats["headers_standardised"] += 1
            add_change(
                changes,
                sheet_name,
                cell.coordinate,
                "Fixed",
                old,
                deduped,
                reason,
            )

    # 3) Clean text + normalise date strings.
    for row in range(2, sheet.max_row + 1):
        for col in range(1, max_col + 1):
            cell = sheet.cell(row=row, column=col)
            value = cell.value

            if value is None:
                continue
            if cell.data_type == "f":
                continue
            if not isinstance(value, str):
                continue

            cleaned, reasons = clean_text(value)
            date_value, date_changed, date_reason = normalise_date_text(cleaned)
            final = date_value if date_changed else cleaned

            if date_changed:
                reasons.append(date_reason)

            if final != value:
                cell.value = final
                if reasons:
                    stats["text_cells_cleaned"] += 1
                    if date_changed:
                        stats["dates_normalised"] += 1
                    add_change(
                        changes,
                        sheet_name,
                        cell.coordinate,
                        "Fixed",
                        value,
                        final,
                        "; ".join(reasons),
                    )

    # 4) Remove fully empty rows from bottom-up (skip header row 1).
    for row in range(sheet.max_row, 1, -1):
        row_values = [sheet.cell(row=row, column=col).value for col in range(1, max_col + 1)]
        if all(is_blank(v) for v in row_values):
            sheet.delete_rows(row, 1)
            stats["empty_rows_removed"] += 1
            add_change(
                changes,
                sheet_name,
                f"{row}:{row}",
                "Removed",
                "[empty row]",
                "",
                "Fully empty row removed",
            )


def write_change_log_sheet(workbook, changes: list[Change]) -> None:
    if "Change Log" in workbook.sheetnames:
        workbook.remove(workbook["Change Log"])

    ws = workbook.create_sheet("Change Log")
    ws.append(["sheet", "cell_or_range", "action", "old_value", "new_value", "reason"])
    for change in changes:
        ws.append(
            [
                change.sheet,
                change.cell_or_range,
                change.action,
                change.old_value,
                change.new_value,
                change.reason,
            ]
        )


def main():
    if len(sys.argv) < 2:
        print(
            "ERROR: No file path provided. Usage: "
            "python skills/excel-doctor/scripts/heal.py <input.xlsx> [output.xlsx]",
            file=sys.stderr,
        )
        sys.exit(1)

    input_path = Path(sys.argv[1])
    if not input_path.exists():
        print(f"ERROR: File not found: {input_path}", file=sys.stderr)
        sys.exit(1)

    if input_path.suffix.lower() not in (".xlsx", ".xlsm"):
        print(f"ERROR: Expected an .xlsx/.xlsm file, got: {input_path.suffix}", file=sys.stderr)
        sys.exit(1)

    output_path = (
        Path(sys.argv[2])
        if len(sys.argv) > 2
        else input_path.with_name(f"{input_path.stem}_healed.xlsx")
    )

    try:
        workbook = load_workbook(input_path)
    except Exception as exc:
        print(f"ERROR: Could not read workbook: {exc}", file=sys.stderr)
        sys.exit(1)

    changes: list[Change] = []
    stats = Counter()

    for sheet in workbook.worksheets:
        if sheet.title in INTERNAL_SHEETS:
            continue
        stats["sheets_processed"] += 1
        heal_sheet(sheet, changes, stats)

    write_change_log_sheet(workbook, changes)
    workbook.save(output_path)

    print()
    print("═" * 62)
    print("  Excel Doctor  ·  Heal Report")
    print("═" * 62)
    print(f"  Input file   : {input_path.name}")
    print(f"  Output file  : {output_path.name}")
    print("─" * 62)
    print(f"  Sheets processed                : {stats['sheets_processed']}")
    print(f"  Changes logged                  : {len(changes)}")
    print(f"    · Headers standardised        : {stats['headers_standardised']}")
    print(f"    · Headers deduplicated        : {stats['headers_deduplicated']}")
    print(f"    · Merged ranges unmerged      : {stats['merged_ranges_unmerged']}")
    print(f"    · Cells filled from merges    : {stats['cells_filled_from_merged_anchor']}")
    print(f"    · Text cells cleaned          : {stats['text_cells_cleaned']}")
    print(f"    · Dates normalised            : {stats['dates_normalised']}")
    print(f"    · Empty rows removed          : {stats['empty_rows_removed']}")
    print("─" * 62)
    print("  ASSUMPTIONS:")
    print("    · Ambiguous DD/MM vs MM/DD dates prefer day-first when both parse")
    print("    · MM-DD-YY dates assume 20xx for YY<50, otherwise 19xx")
    print("    · Formula cells are preserved unchanged")
    print("═" * 62)
    print()


if __name__ == "__main__":
    main()
