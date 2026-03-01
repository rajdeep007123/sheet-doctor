from __future__ import annotations

from pathlib import Path

from sheet_doctor import __version__ as TOOL_VERSION
from sheet_doctor.contracts import build_contract, build_run_summary

def build_structured_summary(
    result: dict,
    *,
    input_path: Path,
    output_path: Path,
    sheet_name: str | None = None,
    consolidate_sheets: bool | None = None,
    header_row_override: int | None = None,
    role_overrides: dict[int, str] | None = None,
    plan_confirmed: bool = False,
) -> dict:
    contract = build_contract("csv_doctor.heal_summary")
    clean_rows = len(result["clean_data"])
    quarantine_rows = len(result["quarantine"])
    needs_review_rows = sum(1 for row in result["clean_data"] if row.needs_review)
    modified_rows = sum(1 for row in result["clean_data"] if row.was_modified)
    applied_role_overrides = {
        str(idx + 1): role for idx, role in sorted((role_overrides or {}).items())
    }
    return {
        "contract": contract,
        "schema_version": contract["version"],
        "tool_version": TOOL_VERSION,
        "mode": result["mode"],
        "delimiter": result["delimiter"],
        "input_file": str(input_path),
        "output_file": str(output_path),
        "rows": {
            "total_including_header": result["total_in"],
            "clean_rows": clean_rows,
            "quarantine_rows": quarantine_rows,
            "needs_review_rows": needs_review_rows,
            "modified_rows": modified_rows,
        },
        "changes": {
            "logged": len(result["changelog"]),
            "action_counts": dict(result["action_counts"]),
            "quarantine_reason_counts": result["quarantine_reason_counts"],
        },
        "workbook_plan": {
            "sheet_name": sheet_name,
            "consolidate_sheets": bool(consolidate_sheets),
            "header_row_override": header_row_override,
            "role_overrides": applied_role_overrides,
            "plan_confirmed": plan_confirmed,
        },
        "assumptions": list(result["assumptions"]),
        "run_summary": build_run_summary(
            tool="csv-doctor",
            script="heal.py",
            input_path=input_path,
            output_path=output_path,
            metrics={
                "mode": result["mode"],
                "total_including_header": result["total_in"],
                "clean_rows": clean_rows,
                "quarantine_rows": quarantine_rows,
                "needs_review_rows": needs_review_rows,
                "modified_rows": modified_rows,
                "changes_logged": len(result["changelog"]),
                "plan_confirmed": plan_confirmed,
                "role_override_count": len(applied_role_overrides),
            },
        ),
    }
