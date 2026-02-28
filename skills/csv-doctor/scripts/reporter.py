#!/usr/bin/env python3
"""
csv-doctor reporter.py

Builds a plain-text and JSON health report from diagnose.py and
column_detector.py output.

Usage:
    python reporter.py <path-to-file> [output.txt] [output.json]
"""

from __future__ import annotations

import json
import math
import sys
from datetime import datetime
from pathlib import Path
from typing import Any

sys.path.insert(0, str(Path(__file__).parent))
from diagnose import build_report as build_diagnose_report
from heal import ASSUMPTIONS, GENERIC_ASSUMPTIONS, SEMANTIC_ASSUMPTIONS


SEVERITY_HEADINGS = {
    "critical": "🚨 Critical (will break imports)",
    "warning": "⚠️  Warning (will cause analysis errors)",
    "info": "ℹ️  Info (cosmetic, worth fixing)",
}

SCORE_LABELS = [
    (90, "Excellent — minor cleanup only"),
    (70, "Good — a few issues to address"),
    (50, "Fair — significant cleaning needed"),
    (30, "Poor — major surgery required"),
    (0, "Critical — severe data quality issues"),
]


def calc_health_score(diagnose_report: dict[str, Any], column_report: dict[str, Any]) -> dict[str, Any]:
    encoding = diagnose_report.get("encoding", {})
    encoding_issue_count = 0
    if not encoding.get("is_utf8") and encoding.get("detected", "unknown") != "unknown":
        encoding_issue_count += 1
    if encoding.get("suspicious_chars"):
        encoding_issue_count += 1
    encoding_deduction = min(20, encoding_issue_count * 5)

    structural_issue_count = 0
    if diagnose_report.get("column_count", {}).get("misaligned_rows"):
        structural_issue_count += 1
    if diagnose_report.get("empty_rows", {}).get("count", 0):
        structural_issue_count += 1
    duplicate_headers = diagnose_report.get("duplicate_headers", {})
    if duplicate_headers.get("duplicate_columns"):
        structural_issue_count += 1
    if duplicate_headers.get("repeated_header_rows"):
        structural_issue_count += 1
    if diagnose_report.get("whitespace_headers"):
        structural_issue_count += 1
    if diagnose_report.get("column_quality", {}).get("empty_columns"):
        structural_issue_count += 1
    if diagnose_report.get("column_quality", {}).get("single_value_columns"):
        structural_issue_count += 1
    structural_deduction = min(30, structural_issue_count * 10)

    date_deduction = min(20, len(diagnose_report.get("date_formats", {})) * 5)

    semantic_columns = column_report.get("columns", {})
    total_rows = column_report.get("summary", {}).get("total_rows", 0)
    total_columns = max(1, column_report.get("summary", {}).get("total_columns", 0))
    total_nulls = sum(column.get("null_count", 0) for column in semantic_columns.values())
    total_cells = max(1, total_rows * total_columns)
    missing_percentage = (total_nulls / total_cells) * 100
    missing_deduction = min(15, math.ceil(missing_percentage / 5))

    duplicate_deduction = 0
    if duplicate_headers.get("duplicate_columns") or duplicate_headers.get("repeated_header_rows"):
        duplicate_deduction += 5
    semantic_issue_counts = column_report.get("summary", {}).get("issue_counts", {})
    if semantic_issue_counts.get("Possible duplicates with slight differences"):
        duplicate_deduction += 5
    duplicate_deduction = min(10, duplicate_deduction)

    data_quality_deduction = min(15, len(semantic_issue_counts) * 2)

    deductions = {
        "encoding": encoding_deduction,
        "structural": structural_deduction,
        "date_chaos": date_deduction,
        "missing_data": missing_deduction,
        "duplicates": duplicate_deduction,
        "data_quality": data_quality_deduction,
    }
    score = max(0, 100 - sum(deductions.values()))

    label = next(text for threshold, text in SCORE_LABELS if score >= threshold)
    return {
        "score": score,
        "label": label,
        "deductions": deductions,
        "metrics": {
            "overall_null_percentage": round(missing_percentage, 2),
            "encoding_issue_count": encoding_issue_count,
            "structural_issue_count": structural_issue_count,
            "date_columns_affected": len(diagnose_report.get("date_formats", {})),
            "semantic_issue_types": len(semantic_issue_counts),
        },
    }


def build_column_breakdown(column_report: dict[str, Any]) -> list[dict[str, Any]]:
    rows = []
    for column_name, column_stats in column_report.get("columns", {}).items():
        top_issues = column_stats.get("suspected_issues") or ["No major issues detected"]
        rows.append(
            {
                "column": column_name,
                "detected_type": column_stats.get("detected_type", "unknown"),
                "null_percentage": column_stats.get("null_percentage", 0.0),
                "top_issues": top_issues,
            }
        )
    return rows


def build_actions(
    diagnose_report: dict[str, Any],
    column_report: dict[str, Any],
    healing_mode: str,
) -> list[str]:
    actions: list[str] = []
    columns = column_report.get("columns", {})
    issues = diagnose_report.get("issues", [])

    misaligned_count = len(diagnose_report.get("column_count", {}).get("misaligned_rows", []))
    if misaligned_count:
        actions.append(
            f"Repair {misaligned_count} structurally broken row(s) with missing or extra columns (auto-fixable)"
        )

    repeated_headers = len(diagnose_report.get("duplicate_headers", {}).get("repeated_header_rows", []))
    empty_rows = diagnose_report.get("empty_rows", {}).get("count", 0)
    if repeated_headers or empty_rows:
        actions.append(
            f"Remove or quarantine {repeated_headers + empty_rows} structural row(s) such as blank lines or repeated headers (auto-fixable)"
        )

    for column_name, info in diagnose_report.get("date_formats", {}).items():
        affected_rows = next(
            (issue["rows_affected"] for issue in issues if issue["id"] == "date_mixed_formats" and issue["columns"] == [column_name]),
            0,
        )
        actions.append(
            f"Normalize mixed date formats in {column_name} ({len(info.get('formats_found', []))} formats across about {affected_rows} populated row(s)) "
            f"({'auto-fixable' if healing_mode == 'schema-specific' and column_name == 'Date' else 'manual review needed'})"
        )

    duplicate_like_columns = [
        issue["columns"][0]
        for issue in issues
        if issue["id"] == "semantic_near_duplicates"
    ]
    if duplicate_like_columns:
        actions.append(
            f"Review near-duplicate values in {', '.join(dict.fromkeys(duplicate_like_columns))} to decide which versions should be merged or kept"
        )

    outlier_columns = [
        issue["columns"][0]
        for issue in issues
        if issue["id"] == "semantic_outliers"
    ]
    if outlier_columns:
        actions.append(
            f"Manually check outlier values in {', '.join(dict.fromkeys(outlier_columns))} before trusting totals or downstream analysis"
        )

    encoding = diagnose_report.get("encoding", {})
    if not encoding.get("is_utf8") or encoding.get("suspicious_chars"):
        actions.append(
            f"Re-decode and normalise text values from {encoding.get('detected', 'the detected')} encoding so names and notes are readable everywhere (auto-fixable)"
        )

    return actions[:6]


def build_assumptions(healing_mode: str) -> list[str]:
    if healing_mode == "schema-specific":
        return ASSUMPTIONS
    if healing_mode == "semantic":
        return SEMANTIC_ASSUMPTIONS
    return GENERIC_ASSUMPTIONS


def format_issue(issue: dict[str, Any]) -> str:
    columns = ", ".join(issue["columns"])
    auto_fix = "✅" if issue["auto_fixable"] else "❌"
    return (
        f"- {issue['plain_english']}\n"
        f"  Columns: {columns}\n"
        f"  Rows affected: {issue['rows_affected']}\n"
        f"  Auto-fixable: {auto_fix}"
    )


def render_text_report(report_json: dict[str, Any]) -> str:
    overview = report_json["file_overview"]
    score = report_json["health_score"]
    issues = report_json["issues"]
    column_breakdown = report_json["column_breakdown"]
    actions = report_json["recommended_actions"]
    assumptions = report_json["assumptions"]

    lines = [
        "SECTION 1 — FILE OVERVIEW",
        f"📄 File: {overview['file']}",
        f"📊 Size: {overview['rows']} rows × {overview['columns']} columns",
        f"🧾 Parsed cleanly: {overview['parsed_rows']} rows",
        f"🧯 Malformed rows: {overview['malformed_rows']} | Skipped by parser: {overview['dropped_rows']}",
        f"💾 Format: {overview['format']}",
        f"🔤 Encoding: {overview['encoding']}",
        f"⏱  Scanned: {overview['scanned_at']}",
        "",
        "SECTION 2 — HEALTH SCORE",
        f"🩺 Health Score: {score['score']}/100 ({score['label']})",
        f"  • Encoding problems: -{score['deductions']['encoding']}",
        f"  • Structural problems: -{score['deductions']['structural']}",
        f"  • Date chaos: -{score['deductions']['date_chaos']}",
        f"  • Missing data: -{score['deductions']['missing_data']}",
        f"  • Duplicates: -{score['deductions']['duplicates']}",
        f"  • Data quality issues: -{score['deductions']['data_quality']}",
        "",
        "SECTION 3 — ISSUES FOUND",
    ]

    if report_json.get("pii_warning"):
        lines.insert(-1, report_json["pii_warning"])
        lines.insert(-1, "")

    for severity in ("critical", "warning", "info"):
        lines.append(SEVERITY_HEADINGS[severity])
        severity_items = issues[severity]
        if not severity_items:
            lines.append("- None")
        else:
            for item in severity_items:
                lines.append(format_issue(item))
        lines.append("")

    lines.extend(
        [
            "SECTION 4 — COLUMN BREAKDOWN",
        ]
    )
    for item in column_breakdown:
        lines.append(
            f"{item['column']} | {item['detected_type']} | {item['null_percentage']}% null | {'; '.join(item['top_issues'])}"
        )

    lines.extend(
        [
            "",
            "SECTION 5 — RECOMMENDED ACTIONS",
        ]
    )
    if not actions:
        lines.append("1. No urgent action required.")
    else:
        for index, action in enumerate(actions, start=1):
            lines.append(f"{index}. {action}")

    lines.extend(
        [
            "",
            "SECTION 6 — ASSUMPTIONS",
        ]
    )
    for assumption in assumptions:
        lines.append(f"• {assumption}")

    return "\n".join(lines).strip() + "\n"


def build_report(file_path: Path) -> dict[str, Any]:
    diagnose_report = build_diagnose_report(file_path)
    column_report = diagnose_report.get("column_semantics", {})
    semantic_columns = list(column_report.get("columns", {}).keys())
    healing_mode = diagnose_report.get("healing_mode_candidate", "generic")

    score = calc_health_score(diagnose_report, column_report)
    issues = diagnose_report.get("issues", [])
    issues_by_severity = {
        severity: [issue for issue in issues if issue["severity"] == severity]
        for severity in ("critical", "warning", "info")
    }
    column_breakdown = build_column_breakdown(column_report)
    actions = build_actions(diagnose_report, column_report, healing_mode=healing_mode)
    assumptions = build_assumptions(healing_mode)
    pii_warning = None
    if any(
        issue.get("plain_english", "").startswith("The ") and "personally identifiable information" in issue.get("plain_english", "").lower()
        for issue in issues_by_severity["info"] + issues_by_severity["warning"]
    ):
        pii_warning = "⚠️ This file appears to contain PII. Handle according to your data protection policy."

    row_accounting = diagnose_report.get("row_accounting") or {}
    report_json = {
        "file_overview": {
            "file": file_path.name,
            "rows": row_accounting.get("raw_data_rows_total", column_report.get("summary", {}).get("total_rows", 0)),
            "parsed_rows": row_accounting.get("parsed_rows_total", column_report.get("summary", {}).get("total_rows", 0)),
            "malformed_rows": row_accounting.get("malformed_rows_total", 0),
            "dropped_rows": row_accounting.get("dropped_rows_total", 0),
            "columns": diagnose_report.get("column_count", {}).get("expected", column_report.get("summary", {}).get("total_columns", 0)),
            "format": diagnose_report.get("detected_format", "unknown"),
            "encoding": diagnose_report.get("detected_encoding", "unknown"),
            "scanned_at": datetime.now().isoformat(timespec="seconds"),
        },
        "health_score": score,
        "issues": issues_by_severity,
        "column_breakdown": column_breakdown,
        "recommended_actions": actions,
        "assumptions": assumptions,
        "pii_warning": pii_warning,
        "source_reports": {
            "diagnose": diagnose_report,
            "column_detector": column_report,
        },
    }
    report_json["text_report"] = render_text_report(report_json)
    return report_json


def default_output_paths(input_path: Path) -> tuple[Path, Path]:
    base = input_path.with_name(f"{input_path.stem}_report")
    return base.with_suffix(".txt"), base.with_suffix(".json")


def main() -> int:
    if len(sys.argv) < 2:
        print(json.dumps({"error": "Usage: reporter.py <file> [output.txt] [output.json]"}))
        return 1

    input_path = Path(sys.argv[1])
    if not input_path.exists():
        print(json.dumps({"error": f"File not found: {input_path}"}))
        return 1

    txt_path, json_path = default_output_paths(input_path)
    if len(sys.argv) >= 3:
        txt_path = Path(sys.argv[2])
    if len(sys.argv) >= 4:
        json_path = Path(sys.argv[3])

    try:
        report_json = build_report(input_path)
    except Exception as exc:
        print(json.dumps({"error": str(exc)}))
        return 1

    txt_path.parent.mkdir(parents=True, exist_ok=True)
    json_path.parent.mkdir(parents=True, exist_ok=True)
    txt_path.write_text(report_json["text_report"], encoding="utf-8")
    json_path.write_text(json.dumps(report_json, indent=2, ensure_ascii=False), encoding="utf-8")

    print(report_json["text_report"], end="")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
