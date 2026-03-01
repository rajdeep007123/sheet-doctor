#!/usr/bin/env python3
from __future__ import annotations

import argparse
import json
import os
import sys
import tempfile
from pathlib import Path

SCRIPT_DIR = Path(__file__).resolve().parent
ROOT_DIR = SCRIPT_DIR.parents[2]
sys.path.insert(0, str(ROOT_DIR))
sys.path.insert(0, str(SCRIPT_DIR))

from heal_modules.preprocessing import preprocess_rows, read_file
from heal_modules.semantic import (
    ASSUMPTIONS,
    GENERIC_ASSUMPTIONS,
    SEMANTIC_ASSUMPTIONS,
    VALID_SEMANTIC_ROLES,
    execute_healing_pipeline,
    inspect_healing_plan,
    process_generic,
    process_schema_specific,
)
from heal_modules.shared import COL, HEADERS, Change, CleanRow, QuarantineRow, is_schema_specific_header
from heal_modules.summary import build_structured_summary
from heal_modules.workbook import (
    WRITE_ONLY_THRESHOLD,
    _write_workbook_fast_impl,
    _write_workbook_standard_impl,
)

HERE = SCRIPT_DIR
ROOT = HERE.parent.parent.parent
INPUT = ROOT / "sample-data" / "extreme_mess.csv"
OUTPUT = ROOT / "sample-data" / "extreme_mess_healed.xlsx"


def parse_args(argv: list[str]) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Heal messy tabular files into a 3-sheet Excel workbook.")
    parser.add_argument("input", nargs="?", default=str(INPUT), help="Input file path")
    parser.add_argument("output", nargs="?", default=None, help="Output .xlsx path (default: <input_stem>_healed.xlsx next to input file)")
    parser.add_argument(
        "--header-row",
        dest="header_row",
        type=int,
        help="Optional 1-based row number to force as the detected header row",
    )
    parser.add_argument(
        "--json-summary",
        dest="json_summary",
        help="Optional path to write a structured JSON healing summary for UI/backend use",
    )
    parser.add_argument(
        "--role-override",
        dest="role_overrides",
        action="append",
        default=[],
        help="Optional semantic override in the form <column_index>=<role> using 1-based indexes; role can also be 'ignore'",
    )
    parser.add_argument(
        "--confirm-plan",
        dest="confirm_plan",
        action="store_true",
        help="Persist that the current workbook header/role plan was explicitly confirmed by the user",
    )
    group = parser.add_mutually_exclusive_group()
    group.add_argument("--sheet", dest="sheet_name", help="Workbook sheet name to heal")
    group.add_argument(
        "--all-sheets",
        dest="all_sheets",
        action="store_true",
        help="Consolidate all sheets before healing (only when columns are compatible)",
    )
    return parser.parse_args(argv)



def parse_role_overrides(raw_overrides: list[str]) -> dict[int, str]:
    overrides: dict[int, str] = {}
    allowed = set(VALID_SEMANTIC_ROLES) | {"ignore"}
    for raw in raw_overrides:
        if "=" not in raw:
            raise ValueError(f"Invalid role override '{raw}'. Expected <column_index>=<role>.")
        index_text, role_text = raw.split("=", 1)
        try:
            index = int(index_text.strip())
        except ValueError as exc:
            raise ValueError(f"Invalid role override column index '{index_text}'.") from exc
        role = role_text.strip().lower()
        if index < 1:
            raise ValueError(f"Role override column index must be >= 1: '{raw}'")
        if role not in allowed:
            raise ValueError(
                f"Invalid role override role '{role}'. Expected one of: {', '.join(sorted(allowed))}"
            )
        overrides[index - 1] = role
    return overrides



def write_workbook(
    clean_data: list[CleanRow],
    quarantine: list[QuarantineRow],
    changelog: list[Change],
    output_path: Path,
    headers: list[str] | None = None,
) -> None:
    headers = headers or HEADERS
    writer = _write_workbook_fast_impl if len(clean_data) > WRITE_ONLY_THRESHOLD else _write_workbook_standard_impl
    output_path.parent.mkdir(parents=True, exist_ok=True)

    fd, tmp_name = tempfile.mkstemp(
        prefix=f".{output_path.stem}.",
        suffix=output_path.suffix,
        dir=str(output_path.parent),
    )
    os.close(fd)
    temp_path = Path(tmp_name)
    try:
        writer(clean_data, quarantine, changelog, temp_path, headers)
        os.replace(temp_path, output_path)
    finally:
        if temp_path.exists():
            temp_path.unlink()



def execute_healing(
    input_path: Path,
    *,
    sheet_name: str | None = None,
    consolidate_sheets: bool | None = None,
    header_row_override: int | None = None,
    role_overrides: dict[int, str] | None = None,
) -> dict:
    return execute_healing_pipeline(
        input_path,
        sheet_name=sheet_name,
        consolidate_sheets=consolidate_sheets,
        header_row_override=header_row_override,
        role_overrides=role_overrides,
    )



def main() -> None:
    args = parse_args(sys.argv[1:])
    input_path = Path(args.input)
    if args.output is not None:
        output_path = Path(args.output)
    else:
        output_path = input_path.parent / f"{input_path.stem}_healed.xlsx"
    role_overrides = parse_role_overrides(args.role_overrides)

    try:
        result = execute_healing(
            input_path,
            sheet_name=args.sheet_name,
            consolidate_sheets=True if args.all_sheets else None,
            header_row_override=args.header_row,
            role_overrides=role_overrides,
        )
    except (FileNotFoundError, ValueError, ImportError) as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        sys.exit(1)

    write_workbook(
        result["clean_data"],
        result["quarantine"],
        result["changelog"],
        output_path,
        headers=result["headers"],
    )
    if args.json_summary:
        summary_path = Path(args.json_summary)
        summary_path.parent.mkdir(parents=True, exist_ok=True)
        summary_path.write_text(
            json.dumps(
                build_structured_summary(
                    result,
                    input_path=input_path,
                    output_path=output_path,
                    sheet_name=args.sheet_name,
                    consolidate_sheets=True if args.all_sheets else None,
                    header_row_override=args.header_row,
                    role_overrides=role_overrides,
                    plan_confirmed=args.confirm_plan,
                ),
                indent=2,
                ensure_ascii=False,
            ),
            encoding="utf-8",
        )

    width = 60
    print()
    print("═" * width)
    print("  CSV Doctor  ·  Heal Report  (Excel output)")
    print("═" * width)
    print(f"  Input file   : {input_path.name}")
    print(f"  Output file  : {output_path.name}")
    print(f"  Mode         : {result['mode']}")
    print(f"  Delimiter    : {result['delimiter']!r}")
    print("─" * width)
    print(f"  Rows in      : {result['total_in']}  (incl. column header row)")
    print(f"  Clean Data   : {len(result['clean_data'])} rows")
    print(f"    · was_modified = TRUE  : {sum(1 for r in result['clean_data'] if r.was_modified)}")
    print(f"    · needs_review = TRUE  : {sum(1 for r in result['clean_data'] if r.needs_review)}")
    print(f"  Quarantine   : {len(result['quarantine'])} rows")
    for reason, rows in result["quarantine_reason_counts"].items():
        print(f"    · {reason:<40} {rows}")
    print(f"  Changes logged: {len(result['changelog'])}")
    print(f"    · Fixed       : {result['action_counts'].get('Fixed', 0)}")
    print(f"    · Quarantined : {result['action_counts'].get('Quarantined', 0)}")
    print(f"    · Removed     : {result['action_counts'].get('Removed', 0)}")
    print(f"    · Flagged     : {result['action_counts'].get('Flagged', 0)}")
    print("─" * width)
    print("  ASSUMPTIONS MADE:")
    for assumption in result["assumptions"]:
        print(f"    · {assumption}")
    print("═" * width)
    print()


if __name__ == "__main__":
    main()
