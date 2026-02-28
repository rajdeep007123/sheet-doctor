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

SCRIPT_DIR = Path(__file__).resolve().parent
ROOT_DIR = SCRIPT_DIR.parents[2]
sys.path.insert(0, str(ROOT_DIR))
sys.path.insert(0, str(SCRIPT_DIR))

from sheet_doctor import __version__ as TOOL_VERSION
from sheet_doctor.contracts import build_contract, build_run_summary
from loader import load_file
from column_detector import analyse_dataframe
from issue_taxonomy import build_issue, infer_healing_mode


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

    column_semantics = report.get("column_semantics", {})
    semantic_summary = column_semantics.get("summary", {})
    if semantic_summary.get("issue_counts"):
        issues += 1
    unknown_columns = semantic_summary.get("detected_types", {}).get("unknown", 0)
    if unknown_columns:
        issues += 1

    if issues == 0:
        verdict = "HEALTHY"
    elif issues <= 3:
        verdict = "NEEDS ATTENTION"
    else:
        verdict = "CRITICAL"

    return {"verdict": verdict, "issue_count": issues}


def non_null_rows(column_stats: dict, total_rows: int) -> int:
    return max(0, total_rows - int(column_stats.get("null_count", 0)))


def build_normalized_issues(report: dict, healing_mode: str) -> list[dict]:
    issues: list[dict] = []
    row_accounting = report.get("row_accounting", {})
    total_rows = row_accounting.get("raw_data_rows_total", report.get("total_rows", 0) - 1)
    columns = report.get("column_semantics", {}).get("columns", {})

    encoding = report.get("encoding", {})
    if not encoding.get("is_utf8") and encoding.get("detected", "unknown") != "unknown":
        issues.append(
            build_issue(
                issue_id="encoding_non_utf8",
                plain_english=f"The file uses {encoding.get('detected')} instead of UTF-8, so some systems may misread special characters.",
                columns=["file-wide"],
                rows_affected=total_rows,
                healing_mode=healing_mode,
            )
        )
    if encoding.get("suspicious_chars"):
        issues.append(
            build_issue(
                issue_id="encoding_suspicious_chars",
                plain_english="Some characters look corrupted, which may change names, notes, or other text values.",
                columns=["file-wide"],
                rows_affected=len(encoding.get("suspicious_chars", [])),
                healing_mode=healing_mode,
            )
        )

    misaligned_rows = report.get("column_count", {}).get("misaligned_rows", [])
    if misaligned_rows:
        issues.append(
            build_issue(
                issue_id="structural_misaligned_rows",
                plain_english="Some rows have too many or too few columns, which will break imports and shift values into the wrong fields.",
                columns=["row structure"],
                rows_affected=len(misaligned_rows),
                healing_mode=healing_mode,
                details={"rows": misaligned_rows},
            )
        )

    repeated_headers = report.get("duplicate_headers", {}).get("repeated_header_rows", [])
    if repeated_headers:
        issues.append(
            build_issue(
                issue_id="structural_repeated_header_rows",
                plain_english="The header row appears again inside the data, which will be treated as a broken data row during import.",
                columns=["row structure"],
                rows_affected=len(repeated_headers),
                healing_mode=healing_mode,
                details={"rows": repeated_headers},
            )
        )

    empty_rows = report.get("empty_rows", {})
    if empty_rows.get("count", 0):
        issues.append(
            build_issue(
                issue_id="structural_empty_rows",
                plain_english="Completely blank rows are mixed into the file and can interfere with import counts and analysis.",
                columns=["row structure"],
                rows_affected=empty_rows["count"],
                healing_mode=healing_mode,
                details={"rows": empty_rows.get("rows", [])},
            )
        )

    whitespace_headers = report.get("whitespace_headers", [])
    if whitespace_headers:
        issues.append(
            build_issue(
                issue_id="header_whitespace",
                plain_english="Some headers contain leading or trailing spaces, which can make column matching fail silently.",
                columns=whitespace_headers,
                rows_affected=1,
                healing_mode=healing_mode,
            )
        )

    column_quality = report.get("column_quality", {})
    if column_quality.get("empty_columns"):
        issues.append(
            build_issue(
                issue_id="quality_empty_columns",
                plain_english="Some columns are completely empty and add noise without carrying any usable information.",
                columns=column_quality["empty_columns"],
                rows_affected=total_rows,
                healing_mode=healing_mode,
            )
        )
    if column_quality.get("single_value_columns"):
        issues.append(
            build_issue(
                issue_id="quality_single_value_columns",
                plain_english="Some columns only contain one repeated value, which may indicate a fill-down or export problem.",
                columns=list(column_quality["single_value_columns"].keys()),
                rows_affected=total_rows,
                healing_mode=healing_mode,
            )
        )

    for column_name, info in report.get("date_formats", {}).items():
        issues.append(
            build_issue(
                issue_id="date_mixed_formats",
                plain_english=f"The {column_name} column mixes multiple date formats, so the same date may be interpreted differently by different tools.",
                columns=[column_name],
                rows_affected=non_null_rows(columns.get(column_name, {}), total_rows),
                healing_mode=healing_mode,
                details={"formats_found": info.get("formats_found", [])},
            )
        )

    for column_name, column_stats in columns.items():
        detected_type = column_stats.get("detected_type", "unknown")
        for suspected_issue in column_stats.get("suspected_issues", []):
            issue_id = None
            plain_english = ""
            if suspected_issue == "Mixed date formats detected":
                continue
            if suspected_issue == "Inconsistent capitalisation":
                if detected_type in {"date", "plain number", "currency/amount", "percentage", "URL", "phone number", "ID/code"}:
                    continue
                issue_id = "semantic_inconsistent_capitalisation"
                plain_english = (
                    f"The {column_name} column uses inconsistent capitalisation, which can split what should be one category into several versions."
                )
            elif suspected_issue.startswith("Trailing/leading whitespace in "):
                issue_id = "semantic_trim_whitespace"
                plain_english = f"The {column_name} column contains extra spaces at the start or end of values."
            elif suspected_issue == "Possible duplicates with slight differences":
                issue_id = "semantic_near_duplicates"
                plain_english = (
                    f"The {column_name} column appears to contain near-duplicate values that differ only slightly, such as spacing or casing changes."
                )
            elif suspected_issue == "Values suspiciously all the same":
                issue_id = "semantic_constant_values"
                plain_english = (
                    f"The {column_name} column is almost entirely the same value, which may mean the export lost useful variation."
                )
            elif suspected_issue == "Outliers detected (values outside 3 standard deviations)":
                issue_id = "semantic_outliers"
                plain_english = (
                    f"The {column_name} column contains outlier values that look unusually large or small compared with the rest of the file."
                )
            elif suspected_issue == "Possible PII detected (emails/phones/names)":
                issue_id = "semantic_pii"
                plain_english = (
                    f"The {column_name} column appears to contain personally identifiable information that may need extra care before sharing."
                )

            if issue_id:
                issues.append(
                    build_issue(
                        issue_id=issue_id,
                        plain_english=plain_english,
                        columns=[column_name],
                        rows_affected=non_null_rows(column_stats, total_rows),
                        healing_mode=healing_mode,
                    )
                )

    return sorted(
        issues,
        key=lambda issue: (
            {"critical": 0, "warning": 1, "info": 2}[issue["severity"]],
            -issue["rows_affected"],
            issue["columns"],
        ),
    )


def build_report(file_path: Path) -> dict:
    if not file_path.exists():
        raise FileNotFoundError(f"File not found: {file_path}")

    if file_path.suffix.lower() not in (".csv", ".tsv", ".txt"):
        raise ValueError(f"Expected a CSV/TSV/TXT file, got: {file_path.suffix}")

    loaded = load_file(file_path)
    encoding_info = loaded["encoding_info"]
    delimiter = loaded["delimiter"]
    raw_text = loaded["raw_text"]
    df = loaded["dataframe"]
    column_semantics = analyse_dataframe(df)
    healing_mode = infer_healing_mode(list(df.columns), column_semantics)

    raw_rows = list(csv.reader(io.StringIO(raw_text), delimiter=delimiter))

    report = {
        "contract": build_contract("csv_doctor.diagnose"),
        "schema_version": build_contract("csv_doctor.diagnose")["version"],
        "tool_version": TOOL_VERSION,
        "file": file_path.name,
        "detected_format": loaded["detected_format"],
        "detected_encoding": loaded["detected_encoding"],
        "total_rows": len(raw_rows),
        "healing_mode_candidate": healing_mode,
        "dialect": {"delimiter": delimiter},
        "encoding": encoding_info,
        "column_count": check_column_alignment(raw_rows),
        "date_formats": check_date_formats(df),
        "empty_rows": check_empty_rows(raw_rows),
        "duplicate_headers": check_duplicate_headers(raw_rows),
        "whitespace_headers": check_whitespace_headers(raw_rows),
        "column_quality": check_columns_quality(df),
        "column_semantics": column_semantics,
        "row_accounting": loaded.get("row_accounting"),
    }

    if loaded["warnings"]:
        report["loader_warnings"] = loaded["warnings"]
    if loaded.get("degraded_mode"):
        report["degraded_mode"] = loaded["degraded_mode"]

    report["issues"] = build_normalized_issues(report, healing_mode=healing_mode)
    report["summary"] = build_summary(report)
    report["run_summary"] = build_run_summary(
        tool="csv-doctor",
        script="diagnose.py",
        input_path=file_path,
        warnings=loaded.get("warnings", []),
        metrics={
            "detected_format": loaded["detected_format"],
            "healing_mode_candidate": healing_mode,
            "raw_rows_total": report["row_accounting"].get("raw_rows_total", len(raw_rows)),
            "parsed_rows_total": report["row_accounting"].get("parsed_rows_total", len(df)),
            "issues_found": report["summary"]["issue_count"],
            "verdict": report["summary"]["verdict"],
            "degraded_mode_active": bool(loaded.get("degraded_mode", {}).get("active", False)),
        },
    )
    return report


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

    try:
        report = build_report(file_path)
    except Exception as e:
        print(json.dumps({"error": f"Could not read file: {e}"}), file=sys.stdout)
        sys.exit(1)

    print(json.dumps(report, indent=2, ensure_ascii=False))
    sys.exit(0)


if __name__ == "__main__":
    main()
