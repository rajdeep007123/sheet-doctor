#!/usr/bin/env python3
"""
csv-doctor diagnose.py
Part of sheet-doctor — https://github.com/razzo007/sheet-doctor

Analyses a CSV file for common data quality problems and outputs a JSON
health report to stdout. Designed to be run by Claude Code's csv-doctor skill.

Usage:
    python diagnose.py <path-to-csv>

Exit codes:
    0 — script ran successfully (issues may still have been found)
    1 — script failed (file not found, completely unreadable, etc.)
"""

import sys
import json
import io
import csv
import re
from pathlib import Path
from collections import Counter

# loader.py lives in the same directory as this script.
sys.path.insert(0, str(Path(__file__).parent))
from loader import load_file


def check_column_alignment(raw_rows: list[list[str]]) -> dict:
    if not raw_rows:
        return {"expected": 0, "misaligned_rows": []}

    expected = len(raw_rows[0])
    misaligned = []

    for i, row in enumerate(raw_rows[1:], start=2):  # row 2 is first data row
        count = len(row)
        if count != expected:
            misaligned.append({"row": i, "count": count})

    return {
        "expected": expected,
        "misaligned_rows": misaligned[:50],  # cap report at 50
    }


def check_empty_rows(raw_rows: list[list[str]]) -> dict:
    empty = []
    for i, row in enumerate(raw_rows, start=1):
        if all(cell.strip() == "" for cell in row):
            empty.append(i)
    return {"count": len(empty), "rows": empty}


def check_duplicate_headers(raw_rows: list[list[str]]) -> dict:
    if not raw_rows:
        return {"duplicate_columns": [], "repeated_header_rows": []}

    headers = [h.strip() for h in raw_rows[0]]

    # Duplicate column names within the header
    counts = Counter(headers)
    duplicate_columns = [h for h, n in counts.items() if n > 1 and h != ""]

    # Header row appearing again inside the data
    header_signature = tuple(h.strip().lower() for h in raw_rows[0])
    repeated_at = []
    for i, row in enumerate(raw_rows[1:], start=2):
        if tuple(c.strip().lower() for c in row) == header_signature:
            repeated_at.append(i)

    return {
        "duplicate_columns": duplicate_columns,
        "repeated_header_rows": repeated_at,
    }


def check_whitespace_headers(raw_rows: list[list[str]]) -> list[str]:
    if not raw_rows:
        return []
    return [h for h in raw_rows[0] if h != h.strip()]


def check_date_formats(df) -> dict:
    """Check for mixed date formats across columns. Accepts a pandas DataFrame."""
    # Common date patterns with human-readable labels
    date_patterns = [
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

    results = {}

    for col in df.columns:
        series = df[col].dropna().astype(str).str.strip()
        if series.empty:
            continue

        # Quick heuristic: does this column look date-like at all?
        sample = series.head(20)
        looks_like_date = sample.str.match(
            r".*\d{2,4}[-/]\d{1,2}[-/]\d{1,4}.*|.*\d{1,2}\s+\w{3,9}\s+\d{2,4}.*"
        ).sum()
        if looks_like_date < 1:
            continue

        format_hits: dict[str, list[str]] = {}
        for value in series:
            for pattern, label in date_patterns:
                if re.match(pattern, value.strip()):
                    format_hits.setdefault(label, []).append(value)
                    break

        if len(format_hits) > 1:
            results[col] = {
                "formats_found": list(format_hits.keys()),
                "examples": {fmt: vals[0] for fmt, vals in format_hits.items()},
            }

    return results


def check_columns_quality(df) -> dict:
    """Check for empty and single-value columns. Accepts a pandas DataFrame."""
    empty_cols = []
    single_val_cols = {}

    for col in df.columns:
        series = df[col].dropna().astype(str).str.strip().replace("", None).dropna()
        if series.empty:
            empty_cols.append(col)
        elif series.nunique() == 1:
            single_val_cols[col] = series.iloc[0]

    return {
        "empty_columns": empty_cols,
        "single_value_columns": single_val_cols,
    }


def build_summary(report: dict) -> dict:
    issues = 0

    enc = report.get("encoding", {})
    if not enc.get("is_utf8") and enc.get("detected", "unknown") != "unknown":
        issues += 1
    if enc.get("suspicious_chars"):
        issues += 1

    misaligned = report.get("column_count", {}).get("misaligned_rows", [])
    if misaligned:
        issues += 1

    if report.get("date_formats"):
        issues += len(report["date_formats"])

    if report.get("empty_rows", {}).get("count", 0) > 0:
        issues += 1

    dup = report.get("duplicate_headers", {})
    if dup.get("duplicate_columns") or dup.get("repeated_header_rows"):
        issues += 1

    if report.get("whitespace_headers"):
        issues += 1

    col_quality = report.get("column_quality", {})
    if col_quality.get("empty_columns"):
        issues += 1
    if col_quality.get("single_value_columns"):
        issues += 1

    if issues == 0:
        verdict = "HEALTHY"
    elif issues <= 3:
        verdict = "NEEDS ATTENTION"
    else:
        verdict = "CRITICAL"

    return {"verdict": verdict, "issue_count": issues}


def main():
    if len(sys.argv) < 2:
        print(
            json.dumps({"error": "No file path provided. Usage: diagnose.py <file.csv>"}),
            file=sys.stdout,
        )
        sys.exit(1)

    file_path = Path(sys.argv[1])

    if not file_path.exists():
        print(
            json.dumps({"error": f"File not found: {file_path}"}),
            file=sys.stdout,
        )
        sys.exit(1)

    if file_path.suffix.lower() not in (".csv", ".tsv", ".txt"):
        print(
            json.dumps(
                {"error": f"Expected a CSV/TSV/TXT file, got: {file_path.suffix}"}
            ),
            file=sys.stdout,
        )
        sys.exit(1)

    # ── Load the file via the universal loader ─────────────────────────────
    try:
        loaded = load_file(file_path)
    except Exception as e:
        print(json.dumps({"error": f"Could not read file: {e}"}), file=sys.stdout)
        sys.exit(1)

    encoding_info = loaded["encoding_info"]
    delimiter     = loaded["delimiter"]
    raw_text      = loaded["raw_text"]
    df            = loaded["dataframe"]

    # ── Parse raw rows for structural checks ──────────────────────────────
    try:
        raw_rows = list(csv.reader(io.StringIO(raw_text), delimiter=delimiter))
    except Exception as e:
        print(json.dumps({"error": f"Could not parse file: {e}"}), file=sys.stdout)
        sys.exit(1)

    # ── Run all checks ─────────────────────────────────────────────────────
    report = {
        "file":              file_path.name,
        "total_rows":        len(raw_rows),
        "dialect":           {"delimiter": delimiter},
        "encoding":          encoding_info,
        "column_count":      check_column_alignment(raw_rows),
        "date_formats":      check_date_formats(df),
        "empty_rows":        check_empty_rows(raw_rows),
        "duplicate_headers": check_duplicate_headers(raw_rows),
        "whitespace_headers": check_whitespace_headers(raw_rows),
        "column_quality":    check_columns_quality(df),
    }

    if loaded["warnings"]:
        report["loader_warnings"] = loaded["warnings"]

    report["summary"] = build_summary(report)

    print(json.dumps(report, indent=2, ensure_ascii=False))
    sys.exit(0)


if __name__ == "__main__":
    main()
