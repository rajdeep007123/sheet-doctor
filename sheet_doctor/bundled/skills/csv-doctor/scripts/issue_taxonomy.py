#!/usr/bin/env python3
"""
Shared csv-doctor issue taxonomy.

This keeps severity and auto-fixability logic in one place so diagnose.py
and reporter.py do not drift.
"""

from __future__ import annotations

from typing import Any


ISSUE_DEFINITIONS = {
    "encoding_non_utf8": {"severity": "warning"},
    "encoding_suspicious_chars": {"severity": "warning"},
    "structural_misaligned_rows": {"severity": "critical"},
    "structural_repeated_header_rows": {"severity": "critical"},
    "structural_empty_rows": {"severity": "warning"},
    "header_whitespace": {"severity": "info"},
    "quality_empty_columns": {"severity": "warning"},
    "quality_single_value_columns": {"severity": "info"},
    "date_mixed_formats": {"severity": "warning"},
    "semantic_inconsistent_capitalisation": {"severity": "warning"},
    "semantic_trim_whitespace": {"severity": "info"},
    "semantic_near_duplicates": {"severity": "warning"},
    "semantic_constant_values": {"severity": "info"},
    "semantic_outliers": {"severity": "warning"},
    "semantic_pii": {"severity": "info"},
}


def infer_healing_mode(headers: list[str], column_semantics: dict[str, Any] | None = None) -> str:
    normalized = tuple(" ".join((header or "").strip().lower().split()) for header in headers)
    expected = (
        "employee name",
        "department",
        "date",
        "amount",
        "currency",
        "category",
        "status",
        "notes",
    )
    if normalized == expected:
        return "schema-specific"

    if column_semantics:
        detected_types = column_semantics.get("summary", {}).get("detected_types", {})
        signals = 0
        if detected_types.get("date", 0) >= 1:
            signals += 1
        if detected_types.get("currency/amount", 0) + detected_types.get("plain number", 0) >= 1:
            signals += 1
        if (
            detected_types.get("name", 0) >= 1
            or detected_types.get("currency code", 0) >= 1
            or detected_types.get("categorical", 0) >= 2
        ):
            signals += 1
        if signals >= 3:
            return "semantic"

    return "generic"


def is_auto_fixable(issue_id: str, columns: list[str], healing_mode: str) -> bool:
    if issue_id in {
        "encoding_non_utf8",
        "encoding_suspicious_chars",
        "structural_misaligned_rows",
        "structural_repeated_header_rows",
        "structural_empty_rows",
        "header_whitespace",
        "semantic_trim_whitespace",
    }:
        return True

    if issue_id in {"quality_empty_columns", "quality_single_value_columns", "semantic_constant_values", "semantic_outliers", "semantic_pii"}:
        return False

    if issue_id == "date_mixed_formats":
        return healing_mode in {"schema-specific", "semantic"} and len(columns) == 1

    if issue_id == "semantic_inconsistent_capitalisation":
        return healing_mode in {"schema-specific", "semantic"}

    if issue_id == "semantic_near_duplicates":
        return healing_mode == "schema-specific" and all(
            column in {"Employee Name", "Amount", "Currency", "Status"}
            for column in columns
        )

    return False


def build_issue(
    *,
    issue_id: str,
    plain_english: str,
    columns: list[str],
    rows_affected: int,
    healing_mode: str,
    details: dict[str, Any] | None = None,
) -> dict[str, Any]:
    definition = ISSUE_DEFINITIONS[issue_id]
    return {
        "id": issue_id,
        "severity": definition["severity"],
        "plain_english": plain_english,
        "columns": columns,
        "rows_affected": rows_affected,
        "auto_fixable": is_auto_fixable(issue_id, columns, healing_mode),
        "details": details or {},
    }
