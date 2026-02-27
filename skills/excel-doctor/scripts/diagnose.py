#!/usr/bin/env python3
"""
excel-doctor diagnose.py

Analyses an Excel workbook for structural and data-quality issues and
outputs a JSON health report to stdout.

Usage:
    python skills/excel-doctor/scripts/diagnose.py <path-to-xlsx>
"""

from __future__ import annotations

import json
import re
import sys
from collections import Counter
from datetime import date, datetime, time
from pathlib import Path

try:
    from openpyxl import load_workbook
except ImportError:
    print(
        json.dumps({"error": "openpyxl not installed — run: pip install openpyxl"}),
        file=sys.stdout,
    )
    sys.exit(1)


ERROR_VALUES = {"#REF!", "#VALUE!", "#DIV/0!", "#NAME?", "#NULL!", "#N/A", "#NUM!"}
INTERNAL_SHEETS = {"Change Log"}
STRUCTURAL_ROW_RE = re.compile(r"^\s*(grand\s+total|subtotal|total)\b", re.IGNORECASE)
DATE_LIKE_RE = re.compile(
    r".*\d{2,4}[-/]\d{1,2}[-/]\d{1,4}.*|.*\d{1,2}\s+\w{3,9}\s+\d{2,4}.*"
)
DATE_PATTERNS = [
    (r"^\d{4}-\d{2}-\d{2}$", "YYYY-MM-DD"),
    (r"^\d{2}/\d{2}/\d{4}$", "DD/MM/YYYY or MM/DD/YYYY"),
    (r"^\d{2}-\d{2}-\d{4}$", "DD-MM-YYYY or MM-DD-YYYY"),
    (r"^\d{2}/\d{2}/\d{2}$", "DD/MM/YY or MM/DD/YY"),
    (r"^\d{2}-\d{2}-\d{2}$", "DD-MM-YY or MM-DD-YY"),
    (r"^\d{1,2}\s+\w+\s+\d{4}$", "D Month YYYY"),
    (r"^\w+\s+\d{1,2},?\s+\d{4}$", "Month D, YYYY"),
    (r"^\d{8}$", "YYYYMMDD"),
    (r"^\d{4}/\d{2}/\d{2}$", "YYYY/MM/DD"),
    (r"^\d{1,2}/\d{1,2}/\d{4}$", "M/D/YYYY or D/M/YYYY"),
]


def is_blank(value) -> bool:
    return value is None or (isinstance(value, str) and value.strip() == "")


def to_text(value) -> str:
    if value is None:
        return ""
    if isinstance(value, datetime):
        return value.strftime("%Y-%m-%d %H:%M:%S")
    if isinstance(value, date):
        return value.strftime("%Y-%m-%d")
    if isinstance(value, time):
        return value.strftime("%H:%M:%S")
    return str(value).strip()


def classify_value(value) -> str | None:
    if is_blank(value):
        return None
    if isinstance(value, bool):
        return "boolean"
    if isinstance(value, (int, float)):
        return "number"
    if isinstance(value, (datetime, date, time)):
        return "date"
    text = to_text(value)
    if text.upper() in ERROR_VALUES:
        return "error"
    return "text"


def is_sheet_empty(sheet) -> bool:
    for row in sheet.iter_rows(values_only=True):
        if any(not is_blank(value) for value in row):
            return False
    return True


def headers_for_sheet(sheet) -> list[str]:
    headers = []
    for col_idx in range(1, sheet.max_column + 1):
        raw = sheet.cell(row=1, column=col_idx).value
        cleaned = to_text(raw)
        headers.append(cleaned if cleaned else f"column_{col_idx}")
    return headers


def header_whitespace(headers: list[str]) -> list[str]:
    return [header for header in headers if header != header.strip()]


def duplicate_headers(headers: list[str]) -> list[str]:
    canonical_to_name = {}
    counts = Counter()
    for header in headers:
        key = header.strip().lower()
        if not key:
            continue
        counts[key] += 1
        canonical_to_name.setdefault(key, header.strip())
    return [canonical_to_name[key] for key, n in counts.items() if n > 1]


def detect_date_format(value: str) -> str | None:
    for pattern, label in DATE_PATTERNS:
        if re.match(pattern, value):
            return label
    return None


def detect_mixed_date_formats(text_values: list[str]) -> dict:
    format_examples: dict[str, str] = {}
    looks_like_date = 0

    for value in text_values:
        if DATE_LIKE_RE.match(value):
            looks_like_date += 1
        label = detect_date_format(value)
        if label and label not in format_examples:
            format_examples[label] = value

    if looks_like_date < 2 or len(format_examples) <= 1:
        return {}

    return {
        "formats_found": list(format_examples.keys()),
        "examples": format_examples,
    }


def scan_formula_errors(values_sheet) -> list[dict]:
    errors = []
    for row in values_sheet.iter_rows():
        for cell in row:
            if is_blank(cell.value):
                continue
            if cell.data_type == "e":
                errors.append({"cell": cell.coordinate, "value": to_text(cell.value)})
                continue
            if isinstance(cell.value, str) and cell.value.strip().upper() in ERROR_VALUES:
                errors.append({"cell": cell.coordinate, "value": cell.value.strip()})
    return errors


def scan_formula_cache_misses(formula_sheet, values_sheet) -> list[dict]:
    misses = []
    for row in formula_sheet.iter_rows():
        for formula_cell in row:
            value = formula_cell.value
            if not isinstance(value, str) or not value.startswith("="):
                continue

            cached_cell = values_sheet[formula_cell.coordinate]
            if cached_cell.value is None:
                misses.append(
                    {
                        "cell": formula_cell.coordinate,
                        "formula": value,
                        "reason": "No cached value. Open/recalculate in Excel and save.",
                    }
                )
    return misses


def scan_empty_rows(sheet) -> list[int]:
    empty = []
    for row_idx in range(2, sheet.max_row + 1):
        row_values = [
            sheet.cell(row=row_idx, column=c).value for c in range(1, sheet.max_column + 1)
        ]
        if all(is_blank(value) for value in row_values):
            empty.append(row_idx)
    return empty


def scan_structural_rows(sheet) -> list[dict]:
    structural = []
    for row_idx in range(2, sheet.max_row + 1):
        row_values = [
            sheet.cell(row=row_idx, column=c).value for c in range(1, sheet.max_column + 1)
        ]
        non_empty = [to_text(v) for v in row_values if not is_blank(v)]
        if not non_empty:
            continue

        first = non_empty[0]
        if STRUCTURAL_ROW_RE.match(first) and len(non_empty) <= 3:
            structural.append(
                {
                    "row": row_idx,
                    "label": first,
                }
            )
    return structural


def scan_columns(sheet, headers: list[str]) -> tuple[dict, list[str], dict, dict, dict]:
    mixed_types = {}
    empty_columns = []
    single_value_columns = {}
    date_formats = {}
    high_null_columns = {}

    data_row_count = max(sheet.max_row - 1, 0)

    for col_idx, header in enumerate(headers, start=1):
        types_seen = set()
        type_examples = {}
        non_empty_text = []
        canonical_values = []
        empty_cells = 0

        for row_idx in range(2, sheet.max_row + 1):
            value = sheet.cell(row=row_idx, column=col_idx).value
            value_type = classify_value(value)
            if not value_type:
                empty_cells += 1
                continue

            types_seen.add(value_type)
            type_examples.setdefault(value_type, to_text(value))

            text = to_text(value)
            canonical_values.append(text)
            if value_type == "text":
                non_empty_text.append(text)

        if not canonical_values:
            empty_columns.append(header)
            continue

        unique_values = {value for value in canonical_values if value != ""}
        if len(unique_values) == 1:
            single_value_columns[header] = next(iter(unique_values))

        if len(types_seen) > 1:
            mixed_types[header] = {
                "types": sorted(types_seen),
                "examples": {t: type_examples[t] for t in sorted(type_examples)},
            }

        date_mix = detect_mixed_date_formats(non_empty_text)
        if date_mix:
            date_formats[header] = date_mix

        if data_row_count >= 5:
            null_ratio = round(empty_cells / data_row_count, 2)
            if 0.8 <= null_ratio < 1.0:
                high_null_columns[header] = {
                    "null_ratio": null_ratio,
                    "empty_cells": empty_cells,
                    "data_rows": data_row_count,
                }

    return mixed_types, empty_columns, single_value_columns, date_formats, high_null_columns


def count_issue_events(report: dict) -> dict:
    sheets = report["sheets"]
    counts = {
        "empty_sheets": len(sheets.get("empty", [])),
        "hidden_sheets": len(sheets.get("hidden", [])),
        "merged_ranges": sum(len(v) for v in report.get("merged_cells", {}).values()),
        "formula_errors": sum(len(v) for v in report.get("formula_errors", {}).values()),
        "formula_cache_misses": sum(
            len(v) for v in report.get("formula_cache_misses", {}).values()
        ),
        "mixed_type_columns": sum(len(v) for v in report.get("mixed_types", {}).values()),
        "empty_rows": sum(v["count"] for v in report.get("empty_rows", {}).values()),
        "empty_columns": sum(len(v) for v in report.get("empty_columns", {}).values()),
        "duplicate_headers": sum(len(v) for v in report.get("duplicate_headers", {}).values()),
        "whitespace_headers": sum(len(v) for v in report.get("whitespace_headers", {}).values()),
        "date_format_columns": sum(len(v) for v in report.get("date_formats", {}).values()),
        "single_value_columns": sum(
            len(v) for v in report.get("single_value_columns", {}).values()
        ),
        "structural_rows": sum(len(v) for v in report.get("structural_rows", {}).values()),
        "high_null_columns": sum(len(v) for v in report.get("high_null_columns", {}).values()),
    }
    return counts


def build_summary(report: dict) -> dict:
    counts = count_issue_events(report)

    critical = 0
    high = 0
    medium = 0

    if counts["formula_errors"] > 0:
        critical += 1
    if counts["formula_cache_misses"] > 0:
        high += 1
    if counts["duplicate_headers"] > 0:
        high += 1
    if counts["mixed_type_columns"] > 0:
        high += 1
    if counts["date_format_columns"] > 0:
        high += 1
    if counts["merged_ranges"] > 0:
        high += 1

    if counts["empty_sheets"] > 0:
        medium += 1
    if counts["hidden_sheets"] > 0:
        medium += 1
    if counts["empty_rows"] > 0:
        medium += 1
    if counts["empty_columns"] > 0:
        medium += 1
    if counts["single_value_columns"] > 0:
        medium += 1
    if counts["whitespace_headers"] > 0:
        medium += 1
    if counts["structural_rows"] > 0:
        medium += 1
    if counts["high_null_columns"] > 0:
        medium += 1

    if critical > 0 or high >= 3:
        verdict = "CRITICAL"
    elif high > 0 or medium >= 2:
        verdict = "NEEDS ATTENTION"
    else:
        verdict = "HEALTHY"

    issue_categories_triggered = sum(1 for _, v in counts.items() if v > 0)
    issue_count = sum(counts.values())

    return {
        "verdict": verdict,
        "issue_count": issue_count,
        "issue_categories_triggered": issue_categories_triggered,
        "severity_breakdown": {
            "critical_categories": critical,
            "high_categories": high,
            "medium_categories": medium,
        },
    }


def main():
    if len(sys.argv) < 2:
        print(
            json.dumps({"error": "No file path provided. Usage: diagnose.py <file.xlsx>"}),
            file=sys.stdout,
        )
        sys.exit(1)

    file_path = Path(sys.argv[1])
    if not file_path.exists():
        print(json.dumps({"error": f"File not found: {file_path}"}), file=sys.stdout)
        sys.exit(1)

    if file_path.suffix.lower() not in (".xlsx", ".xlsm"):
        print(
            json.dumps({"error": f"Expected an .xlsx/.xlsm file, got: {file_path.suffix}"}),
            file=sys.stdout,
        )
        sys.exit(1)

    try:
        workbook_values = load_workbook(file_path, data_only=True)
        workbook_formulas = load_workbook(file_path, data_only=False)
    except Exception as exc:
        print(json.dumps({"error": f"Could not read workbook: {exc}"}), file=sys.stdout)
        sys.exit(1)

    sheets_info = {
        "all": [sheet.title for sheet in workbook_values.worksheets],
        "count": len(workbook_values.worksheets),
        "empty": [],
        "hidden": [],
    }
    merged_cells = {}
    formula_errors = {}
    formula_cache_misses = {}
    mixed_types = {}
    empty_rows = {}
    empty_columns = {}
    duplicate_headers_report = {}
    whitespace_headers_report = {}
    structural_rows_report = {}
    date_formats = {}
    single_value_columns = {}
    high_null_columns = {}

    for values_sheet in workbook_values.worksheets:
        sheet_name = values_sheet.title
        formula_sheet = workbook_formulas[sheet_name]

        sheet_empty = is_sheet_empty(values_sheet)
        if sheet_empty:
            sheets_info["empty"].append(sheet_name)

        if values_sheet.sheet_state != "visible":
            sheets_info["hidden"].append({"name": sheet_name, "state": values_sheet.sheet_state})

        if sheet_name in INTERNAL_SHEETS:
            # Generated metadata tabs are included in inventory but excluded
            # from quality findings.
            continue

        merged = [str(rng) for rng in formula_sheet.merged_cells.ranges]
        if merged:
            merged_cells[sheet_name] = merged

        errors = scan_formula_errors(values_sheet)
        if errors:
            formula_errors[sheet_name] = errors

        cache_misses = scan_formula_cache_misses(formula_sheet, values_sheet)
        if cache_misses:
            formula_cache_misses[sheet_name] = cache_misses

        if sheet_empty:
            # Empty sheets are already reported in inventory; skip data-profile checks.
            continue

        headers = headers_for_sheet(values_sheet)
        whitespace_headers = header_whitespace(headers)
        if whitespace_headers:
            whitespace_headers_report[sheet_name] = whitespace_headers

        duplicates = duplicate_headers(headers)
        if duplicates:
            duplicate_headers_report[sheet_name] = duplicates

        blank_rows = scan_empty_rows(values_sheet)
        if blank_rows:
            empty_rows[sheet_name] = {"count": len(blank_rows), "rows": blank_rows}

        structural_rows = scan_structural_rows(values_sheet)
        if structural_rows:
            structural_rows_report[sheet_name] = structural_rows

        mix, empties, singles, dates, high_nulls = scan_columns(values_sheet, headers)
        if mix:
            mixed_types[sheet_name] = mix
        if empties:
            empty_columns[sheet_name] = empties
        if singles:
            single_value_columns[sheet_name] = singles
        if dates:
            date_formats[sheet_name] = dates
        if high_nulls:
            high_null_columns[sheet_name] = high_nulls

    report = {
        "file": file_path.name,
        "sheets": sheets_info,
        "merged_cells": merged_cells,
        "formula_errors": formula_errors,
        "formula_cache_misses": formula_cache_misses,
        "mixed_types": mixed_types,
        "empty_rows": empty_rows,
        "empty_columns": empty_columns,
        "duplicate_headers": duplicate_headers_report,
        "whitespace_headers": whitespace_headers_report,
        "structural_rows": structural_rows_report,
        "date_formats": date_formats,
        "single_value_columns": single_value_columns,
        "high_null_columns": high_null_columns,
    }
    report["summary"] = build_summary(report)
    report["issue_counts"] = count_issue_events(report)

    print(json.dumps(report, indent=2, ensure_ascii=False))
    sys.exit(0)


if __name__ == "__main__":
    main()
