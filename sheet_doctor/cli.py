from __future__ import annotations

import argparse
import importlib.util
import json
import sys
from pathlib import Path
from typing import Any


PACKAGE_DIR = Path(__file__).resolve().parent
PROJECT_ROOT = PACKAGE_DIR.parent
SOURCE_SKILLS_DIR = PROJECT_ROOT / "skills"
BUNDLED_SKILLS_DIR = PACKAGE_DIR / "bundled" / "skills"

TABULAR_FORMATS = {".csv", ".tsv", ".txt", ".json", ".jsonl"}
MODERN_WORKBOOK_FORMATS = {".xlsx", ".xlsm"}
TABULAR_WORKBOOK_FORMATS = {".xls", ".ods"}
ALL_SUPPORTED_FORMATS = TABULAR_FORMATS | MODERN_WORKBOOK_FORMATS | TABULAR_WORKBOOK_FORMATS

_MODULE_CACHE: dict[str, Any] = {}


def available_script_root() -> Path:
    if SOURCE_SKILLS_DIR.exists():
        return SOURCE_SKILLS_DIR
    return BUNDLED_SKILLS_DIR


def script_path(skill_dir: str, script_name: str) -> Path:
    path = available_script_root() / skill_dir / "scripts" / script_name
    if not path.exists():
        raise FileNotFoundError(f"Bundled runtime script not found: {path}")
    return path


def load_script_module(skill_dir: str, script_name: str):
    path = script_path(skill_dir, script_name)
    cache_key = f"{skill_dir}:{script_name}:{path}"
    if cache_key in _MODULE_CACHE:
        return _MODULE_CACHE[cache_key]
    module_name = (
        "sheet_doctor_runtime_"
        + skill_dir.replace("-", "_")
        + "_"
        + script_name.replace(".py", "").replace("-", "_")
    )
    spec = importlib.util.spec_from_file_location(module_name, path)
    module = importlib.util.module_from_spec(spec)
    assert spec and spec.loader
    sys.modules[module_name] = module
    spec.loader.exec_module(module)
    _MODULE_CACHE[cache_key] = module
    return module


def stderr(message: str) -> None:
    print(f"ERROR: {message}", file=sys.stderr)


def write_output(payload: str, output_path: Path | None) -> None:
    if output_path is not None:
        output_path.parent.mkdir(parents=True, exist_ok=True)
        output_path.write_text(payload, encoding="utf-8")
    else:
        print(payload, end="" if payload.endswith("\n") else "\n")


def choose_backend(
    *,
    input_path: Path,
    mode: str,
    sheet_name: str | None,
    all_sheets: bool,
) -> str:
    suffix = input_path.suffix.lower()
    if suffix not in ALL_SUPPORTED_FORMATS:
        raise ValueError(
            f"Unsupported file type '{suffix or '[missing extension]'}'. "
            f"Supported: {', '.join(sorted(ALL_SUPPORTED_FORMATS))}"
        )

    if suffix in MODERN_WORKBOOK_FORMATS:
        if mode == "tabular":
            return "csv"
        if mode == "workbook":
            if sheet_name or all_sheets:
                raise ValueError(
                    "--sheet/--all-sheets apply to tabular rescue only for .xlsx/.xlsm inputs. "
                    "Use --mode tabular if you want sheet-level rescue."
                )
            return "excel"
        if sheet_name or all_sheets:
            return "csv"
        return "excel"

    if suffix in TABULAR_WORKBOOK_FORMATS:
        if mode == "workbook":
            raise ValueError(
                f"Workbook-native mode is not supported for {suffix}. "
                "Use --mode tabular or convert the workbook to .xlsx first."
            )
        return "csv"

    if mode == "workbook":
        raise ValueError("Workbook-native mode is only supported for .xlsx and .xlsm files.")
    return "csv"


def render_tabular_diagnose_text(report: dict[str, Any]) -> str:
    summary = report.get("summary", {})
    row_accounting = report.get("row_accounting", {})
    lines = [
        "sheet-doctor diagnose",
        f"File: {report.get('file', '[unknown]')}",
        f"Format: {report.get('detected_format', '[unknown]')}",
        f"Encoding: {report.get('detected_encoding', '[unknown]')}",
        f"Verdict: {summary.get('verdict', '[unknown]')}",
        f"Issues: {summary.get('issue_count', 0)}",
    ]
    if row_accounting:
        lines.extend(
            [
                f"Raw rows: {row_accounting.get('raw_rows_total', 0)}",
                f"Parsed rows: {row_accounting.get('parsed_rows_total', 0)}",
                f"Dropped rows: {row_accounting.get('dropped_rows_total', 0)}",
            ]
        )
    if report.get("sheet_name"):
        lines.append(f"Sheet: {report['sheet_name']}")
    if report.get("sheet_names"):
        lines.append(f"Workbook sheets: {', '.join(report['sheet_names'])}")
    return "\n".join(lines) + "\n"


def render_workbook_report_text(report: dict[str, Any]) -> str:
    summary = report.get("summary", {})
    triage = report.get("workbook_triage", {})
    residual = report.get("residual_risk", {})
    warnings = report.get("manual_review_warnings", [])
    lines = [
        "sheet-doctor workbook report",
        f"File: {report.get('file', '[unknown]')}",
        f"Type: {report.get('file_type', '[unknown]')}",
        f"Verdict: {summary.get('verdict', '[unknown]')}",
        f"Issues: {summary.get('issue_count', 0)}",
        f"Triage: {triage.get('classification', '[unknown]')}",
        f"Reason: {triage.get('reason', '[none]')}",
        f"Next action: {triage.get('recommended_next_action', '[none]')}",
        f"Sheets: {report.get('sheets', {}).get('count', 0)}",
    ]

    safe_fixes = residual.get("safe_auto_fix_candidates", [])
    if safe_fixes:
        lines.append("Safe auto-fix candidates:")
        lines.extend(f"- {item['issue']}: {item['count']}" for item in safe_fixes)

    remaining = residual.get("remaining_risks", [])
    if remaining:
        lines.append("Remaining risks:")
        lines.extend(f"- {item['issue']}: {item['count']}" for item in remaining)

    manual = residual.get("manual_review_required", [])
    if manual:
        lines.append("Manual review required:")
        lines.extend(f"- {item['issue']}: {item['count']}" for item in manual)

    if warnings:
        lines.append("Warnings:")
        lines.extend(f"- {warning}" for warning in warnings)

    return "\n".join(lines) + "\n"


def render_excel_heal_summary(summary: dict[str, Any]) -> str:
    triage = summary.get("workbook_triage", {})
    before_after = summary.get("before_after_issue_summary", {}).get("issue_counts", {})
    stats = summary.get("stats", {})
    lines = [
        "sheet-doctor heal",
        f"Input: {summary.get('input_file', '[unknown]')}",
        f"Output: {summary.get('output_file', '[unknown]')}",
        f"Mode: {summary.get('mode', '[unknown]')}",
        f"Sheets processed: {stats.get('sheets_processed', 0)}",
        f"Changes logged: {summary.get('changes_logged', 0)}",
        f"Triage after heal: {triage.get('classification', '[unknown]')}",
        f"Triage reason: {triage.get('reason', '[none]')}",
    ]
    for issue in ["merged_ranges", "duplicate_headers", "empty_rows", "formula_errors", "formula_cache_misses"]:
        values = before_after.get(issue)
        if values is not None:
            lines.append(f"{issue}: {values['before']} -> {values['after']}")
    if summary.get("warnings"):
        lines.append("Warnings:")
        lines.extend(f"- {warning}" for warning in summary["warnings"])
    return "\n".join(lines) + "\n"


def render_csv_heal_summary(result: dict[str, Any], output_path: Path) -> str:
    return (
        "sheet-doctor heal\n"
        f"Input: {result['input_path']}\n"
        f"Output: {output_path}\n"
        f"Mode: {result['mode']}\n"
        f"Rows in: {result['total_in']}\n"
        f"Clean rows: {len(result['clean_data'])}\n"
        f"Quarantine rows: {len(result['quarantine'])}\n"
        f"Changes logged: {len(result['changelog'])}\n"
    )


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(prog="sheet-doctor", description="Local-first spreadsheet diagnosis and cleanup.")
    subparsers = parser.add_subparsers(dest="command", required=True)

    diagnose = subparsers.add_parser("diagnose", help="Diagnose a file and print a health summary or JSON report.")
    diagnose.add_argument("input", help="Input file path")
    diagnose.add_argument("--sheet", dest="sheet_name", help="Workbook sheet name for tabular rescue mode")
    diagnose.add_argument("--all-sheets", dest="all_sheets", action="store_true", help="Consolidate compatible workbook sheets in tabular rescue mode")
    diagnose.add_argument("--mode", choices=["auto", "tabular", "workbook"], default="auto", help="Backend selection policy")
    diagnose.add_argument("--format", choices=["json", "text"], default="json", help="Output format")
    diagnose.add_argument("--output", help="Optional output file path")

    heal = subparsers.add_parser("heal", help="Heal a file and write a cleaned workbook or cleaned workbook-native output.")
    heal.add_argument("input", help="Input file path")
    heal.add_argument("output_positional", nargs="?", default=None, help="Optional output path")
    heal.add_argument("--output", dest="output_flag", help="Optional output path")
    heal.add_argument("--sheet", dest="sheet_name", help="Workbook sheet name for tabular rescue mode")
    heal.add_argument("--all-sheets", dest="all_sheets", action="store_true", help="Consolidate compatible workbook sheets in tabular rescue mode")
    heal.add_argument("--mode", choices=["auto", "tabular", "workbook"], default="auto", help="Backend selection policy")
    heal.add_argument("--json-summary", dest="json_summary", help="Optional JSON summary output path")

    report = subparsers.add_parser("report", help="Generate a human-readable or JSON report.")
    report.add_argument("input", help="Input file path")
    report.add_argument("--sheet", dest="sheet_name", help="Workbook sheet name for tabular rescue mode")
    report.add_argument("--all-sheets", dest="all_sheets", action="store_true", help="Consolidate compatible workbook sheets in tabular rescue mode")
    report.add_argument("--mode", choices=["auto", "tabular", "workbook"], default="auto", help="Backend selection policy")
    report.add_argument("--format", choices=["text", "json"], default="text", help="Output format")
    report.add_argument("--output", help="Optional output file path")

    return parser


def resolve_heal_output(args: argparse.Namespace, input_path: Path, backend: str) -> Path:
    if args.output_positional and args.output_flag:
        raise ValueError("Use either positional output or --output, not both.")
    if args.output_flag:
        return Path(args.output_flag)
    if args.output_positional:
        return Path(args.output_positional)
    if backend == "excel":
        return input_path.with_name(f"{input_path.stem}_healed{input_path.suffix}")
    return input_path.with_name(f"{input_path.stem}_healed.xlsx")


def run_diagnose(args: argparse.Namespace) -> int:
    input_path = Path(args.input)
    if not input_path.exists():
        stderr(f"File not found: {input_path}")
        return 1

    try:
        backend = choose_backend(
            input_path=input_path,
            mode=args.mode,
            sheet_name=args.sheet_name,
            all_sheets=args.all_sheets,
        )
        output_path = Path(args.output) if args.output else None
        if backend == "excel":
            report = load_script_module("excel-doctor", "diagnose.py").build_report(input_path)
            payload = (
                json.dumps(report, indent=2, ensure_ascii=False)
                if args.format == "json"
                else render_workbook_report_text(report)
            )
        else:
            report = load_script_module("csv-doctor", "diagnose.py").build_report(
                input_path,
                sheet_name=args.sheet_name,
                consolidate_sheets=True if args.all_sheets else None,
            )
            payload = (
                json.dumps(report, indent=2, ensure_ascii=False)
                if args.format == "json"
                else render_tabular_diagnose_text(report)
            )
        write_output(payload, output_path)
        return 0
    except Exception as exc:
        stderr(str(exc))
        return 1


def run_report(args: argparse.Namespace) -> int:
    input_path = Path(args.input)
    if not input_path.exists():
        stderr(f"File not found: {input_path}")
        return 1

    try:
        backend = choose_backend(
            input_path=input_path,
            mode=args.mode,
            sheet_name=args.sheet_name,
            all_sheets=args.all_sheets,
        )
        output_path = Path(args.output) if args.output else None
        if backend == "excel":
            report = load_script_module("excel-doctor", "diagnose.py").build_report(input_path)
            payload = (
                json.dumps(report, indent=2, ensure_ascii=False)
                if args.format == "json"
                else render_workbook_report_text(report)
            )
        else:
            report = load_script_module("csv-doctor", "reporter.py").build_report(
                input_path,
                sheet_name=args.sheet_name,
                consolidate_sheets=True if args.all_sheets else None,
            )
            payload = (
                json.dumps(report, indent=2, ensure_ascii=False)
                if args.format == "json"
                else report["text_report"]
            )
        write_output(payload, output_path)
        return 0
    except Exception as exc:
        stderr(str(exc))
        return 1


def run_heal(args: argparse.Namespace) -> int:
    input_path = Path(args.input)
    if not input_path.exists():
        stderr(f"File not found: {input_path}")
        return 1

    try:
        backend = choose_backend(
            input_path=input_path,
            mode=args.mode,
            sheet_name=args.sheet_name,
            all_sheets=args.all_sheets,
        )
        output_path = resolve_heal_output(args, input_path, backend)
        if backend == "excel":
            excel_heal = load_script_module("excel-doctor", "heal.py")
            changes, stats = excel_heal.execute_healing(input_path, output_path)
            summary = excel_heal.build_structured_summary(
                input_path=input_path,
                output_path=output_path,
                changes=changes,
                stats=stats,
            )
            if args.json_summary:
                summary_path = Path(args.json_summary)
                summary_path.parent.mkdir(parents=True, exist_ok=True)
                summary_path.write_text(json.dumps(summary, indent=2, ensure_ascii=False), encoding="utf-8")
            write_output(render_excel_heal_summary(summary), None)
            return 0

        csv_heal = load_script_module("csv-doctor", "heal.py")
        role_overrides: dict[int, str] = {}
        result = csv_heal.execute_healing(
            input_path,
            sheet_name=args.sheet_name,
            consolidate_sheets=True if args.all_sheets else None,
            header_row_override=None,
            role_overrides=role_overrides,
        )
        csv_heal.write_workbook(
            result["clean_data"],
            result["quarantine"],
            result["changelog"],
            output_path,
            headers=result["headers"],
        )
        if args.json_summary:
            summary = csv_heal.build_structured_summary(
                result,
                input_path=input_path,
                output_path=output_path,
                sheet_name=args.sheet_name,
                consolidate_sheets=True if args.all_sheets else None,
                header_row_override=None,
                role_overrides=role_overrides,
                plan_confirmed=False,
            )
            summary_path = Path(args.json_summary)
            summary_path.parent.mkdir(parents=True, exist_ok=True)
            summary_path.write_text(json.dumps(summary, indent=2, ensure_ascii=False), encoding="utf-8")
        write_output(render_csv_heal_summary(result, output_path), None)
        return 0
    except Exception as exc:
        stderr(str(exc))
        return 1


def main(argv: list[str] | None = None) -> int:
    parser = build_parser()
    args = parser.parse_args(argv)
    if args.command == "diagnose":
        return run_diagnose(args)
    if args.command == "heal":
        return run_heal(args)
    if args.command == "report":
        return run_report(args)
    parser.error(f"Unknown command: {args.command}")
    return 2


if __name__ == "__main__":
    raise SystemExit(main())
