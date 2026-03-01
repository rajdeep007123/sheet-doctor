from __future__ import annotations

import argparse
import importlib.util
import json
import os
import sys
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

from sheet_doctor import __version__ as TOOL_VERSION


PACKAGE_DIR = Path(__file__).resolve().parent
PROJECT_ROOT = PACKAGE_DIR.parent
SOURCE_SKILLS_DIR = PROJECT_ROOT / "skills"
BUNDLED_SKILLS_DIR = PACKAGE_DIR / "bundled" / "skills"

TABULAR_FORMATS = {".csv", ".tsv", ".txt", ".json", ".jsonl"}
MODERN_WORKBOOK_FORMATS = {".xlsx", ".xlsm"}
TABULAR_WORKBOOK_FORMATS = {".xls", ".ods"}
ALL_SUPPORTED_FORMATS = TABULAR_FORMATS | MODERN_WORKBOOK_FORMATS | TABULAR_WORKBOOK_FORMATS
SUPPORTED_SCHEMA_SUFFIXES = {".json", ".yml", ".yaml"}

EXIT_SUCCESS = 0
EXIT_COMMAND_ERROR = 1
EXIT_PARSE_FAILED = 2
EXIT_DIAGNOSE_ISSUES = 3
EXIT_HEAL_QUARANTINE = 4
EXIT_VALIDATE_FAILED = 5
EXIT_PARTIAL = 6

_MODULE_CACHE: dict[str, Any] = {}


class CliError(Exception):
    def __init__(self, message: str, code: int = EXIT_COMMAND_ERROR) -> None:
        super().__init__(message)
        self.code = code


class SheetDoctorArgumentParser(argparse.ArgumentParser):
    def error(self, message: str) -> None:
        raise CliError(message, EXIT_COMMAND_ERROR)


def available_script_root() -> Path:
    if SOURCE_SKILLS_DIR.exists():
        return SOURCE_SKILLS_DIR
    return BUNDLED_SKILLS_DIR


def script_path(skill_dir: str, script_name: str) -> Path:
    path = available_script_root() / skill_dir / "scripts" / script_name
    if not path.exists():
        raise CliError(f"Bundled runtime script not found: {path}", EXIT_COMMAND_ERROR)
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


def eprint(message: str) -> None:
    print(message, file=sys.stderr)


def emit_human(message: str, *, quiet: bool = False) -> None:
    if not quiet:
        eprint(message)


def json_dumps(payload: Any) -> str:
    return json.dumps(payload, indent=2, ensure_ascii=False, sort_keys=True)


def timestamp_token() -> str:
    override = os.environ.get("SHEET_DOCTOR_OUTPUT_STAMP")
    if override:
        return override
    return datetime.now(timezone.utc).strftime("%Y%m%dT%H%M%SZ")


def default_output_dir(input_path: Path) -> Path:
    return Path.cwd() / "sheet-doctor-output" / f"{input_path.stem}-{timestamp_token()}"


def determine_output_dir(args: argparse.Namespace, input_path: Path) -> Path:
    if getattr(args, "out_dir", None):
        return Path(args.out_dir)
    return default_output_dir(input_path)


def ensure_parent(path: Path) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)


def write_text(path: Path, payload: str) -> None:
    ensure_parent(path)
    path.write_text(payload, encoding="utf-8")


def write_json(path: Path, payload: Any) -> None:
    write_text(path, json_dumps(payload))


def output_exists(path: Path) -> bool:
    return path.exists()


def remove_generated_at(value: Any) -> Any:
    if isinstance(value, dict):
        result = {}
        for key, item in value.items():
            if key == "generated_at":
                result[key] = "1970-01-01T00:00:00Z"
            else:
                result[key] = remove_generated_at(item)
        return result
    if isinstance(value, list):
        return [remove_generated_at(item) for item in value]
    return value


def normalize_report_for_cli(payload: Any) -> Any:
    payload = remove_generated_at(payload)
    if isinstance(payload, dict):
        overview = payload.get("file_overview")
        if isinstance(overview, dict) and "scanned_at" in overview:
            overview["scanned_at"] = "1970-01-01T00:00:00"
    return payload


def choose_backend(
    *,
    input_path: Path,
    mode: str,
    sheet_name: str | None,
    all_sheets: bool,
) -> str:
    suffix = input_path.suffix.lower()
    if suffix not in ALL_SUPPORTED_FORMATS:
        raise CliError(
            f"Unsupported file type '{suffix or '[missing extension]'}'. "
            f"Supported: {', '.join(sorted(ALL_SUPPORTED_FORMATS))}",
            EXIT_COMMAND_ERROR,
        )

    if suffix in MODERN_WORKBOOK_FORMATS:
        if mode == "tabular":
            return "csv"
        if mode == "workbook":
            if sheet_name or all_sheets:
                raise CliError(
                    "--sheet/--all-sheets apply to tabular rescue only for .xlsx/.xlsm inputs. "
                    "Use --mode tabular if you want sheet-level rescue.",
                    EXIT_COMMAND_ERROR,
                )
            return "excel"
        if sheet_name or all_sheets:
            return "csv"
        return "excel"

    if suffix in TABULAR_WORKBOOK_FORMATS:
        if mode == "workbook":
            raise CliError(
                f"Workbook-native mode is not supported for {suffix}. "
                "Use --mode tabular or convert the workbook to .xlsx first.",
                EXIT_COMMAND_ERROR,
            )
        return "csv"

    if mode == "workbook":
        raise CliError("Workbook-native mode is only supported for .xlsx and .xlsm files.", EXIT_COMMAND_ERROR)
    return "csv"


def classify_backend_exception(exc: Exception) -> int:
    if isinstance(exc, CliError):
        return exc.code
    if isinstance(exc, (ImportError, UnicodeDecodeError)):
        return EXIT_PARSE_FAILED
    if isinstance(exc, FileNotFoundError):
        return EXIT_COMMAND_ERROR
    if isinstance(exc, ValueError):
        return EXIT_PARSE_FAILED
    return EXIT_COMMAND_ERROR


def maybe_default_sheet(input_path: Path, sheet_name: str | None, all_sheets: bool) -> str | None:
    if sheet_name or all_sheets:
        return sheet_name
    if input_path.suffix.lower() not in MODERN_WORKBOOK_FORMATS:
        return sheet_name
    try:
        report = load_script_module("excel-doctor", "diagnose.py").build_report(input_path)
    except Exception:
        return sheet_name
    hidden = {item["name"] for item in report.get("sheets", {}).get("hidden", [])}
    very_hidden = {item["name"] for item in report.get("sheets", {}).get("very_hidden", [])}
    for name in report.get("sheets", {}).get("all", []):
        if name in {"Change Log"}:
            continue
        if name in hidden or name in very_hidden:
            continue
        return name
    return sheet_name


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


def render_validate_text(payload: dict[str, Any]) -> str:
    lines = [
        "sheet-doctor validate",
        f"Input: {payload['input']}",
        f"Valid: {payload['valid']}",
        f"Errors: {payload['error_count']}",
    ]
    if payload["missing_columns"]:
        lines.append("Missing columns: " + ", ".join(payload["missing_columns"]))
    if payload["type_mismatches"]:
        lines.append("Type mismatches:")
        for item in payload["type_mismatches"]:
            lines.append(f"- {item['column']}: expected {item['expected']}, got {item['detected']}")
    return "\n".join(lines) + "\n"


EXPLAIN_RULES = {
    "encoding_non_utf8": {
        "description": "The file decoded as a non-UTF-8 encoding.",
        "evidence": "Detected encoding metadata says the file is not UTF-8.",
        "auto_fixable": True,
        "disable_hint": "Force a known encoding at ingest time if you do not want auto-detection.",
    },
    "encoding_suspicious_chars": {
        "description": "The decoded text contains characters that look corrupted.",
        "evidence": "Suspicious replacement or mojibake-like characters appeared during scan.",
        "auto_fixable": True,
        "disable_hint": "Force the correct encoding if the detector chose the wrong one.",
    },
    "structural_misaligned_rows": {
        "description": "Some rows have too many or too few columns for the detected header.",
        "evidence": "Row length does not match the expected column count.",
        "auto_fixable": True,
        "disable_hint": "Disable structural cleanup only if you want to inspect raw broken rows manually.",
    },
    "structural_repeated_header_rows": {
        "description": "The header row appears again inside the data body.",
        "evidence": "A data row exactly matches the normalized header signature.",
        "auto_fixable": True,
        "disable_hint": "Keep it only if your source intentionally repeats headers mid-file.",
    },
    "date_mixed_formats": {
        "description": "A date-like column contains multiple text date formats.",
        "evidence": "More than one recognized date pattern was found in the same column.",
        "auto_fixable": True,
        "disable_hint": "Disable date normalization if preserving raw source formatting matters more than consistency.",
    },
    "semantic_near_duplicates": {
        "description": "Values look like slight variations of the same thing.",
        "evidence": "Canonicalized versions of values collide while raw values still differ.",
        "auto_fixable": False,
        "disable_hint": "Reduce or skip near-duplicate detection if your domain intentionally uses many close variants.",
    },
    "formula_errors": {
        "description": "Workbook formula cells currently evaluate to Excel error values.",
        "evidence": "Cells with #REF!, #VALUE!, #DIV/0!, or similar errors were found.",
        "auto_fixable": False,
        "disable_hint": "excel-doctor preserves formulas; fix the workbook logic in Excel.",
    },
    "formula_cache_misses": {
        "description": "Workbook formula cells have no cached computed values.",
        "evidence": "A formula exists but the data-only workbook view returned no cached result.",
        "auto_fixable": False,
        "disable_hint": "Open the workbook in Excel, recalculate, and save if cached values matter.",
    },
}


def build_parser() -> argparse.ArgumentParser:
    parser = SheetDoctorArgumentParser(prog="sheet-doctor", description="Local-first spreadsheet diagnosis and cleanup.")
    subparsers = parser.add_subparsers(dest="command", required=True)

    diagnose = subparsers.add_parser("diagnose", help="Diagnose a file and produce a report.")
    diagnose.add_argument("input", help="Input file path")
    diagnose.add_argument("-o", "--out", dest="out_dir", help="Output directory")
    diagnose.add_argument("--output", help="Explicit report output path")
    diagnose.add_argument("--json", action="store_true", help="Write machine JSON to stdout")
    diagnose.add_argument("--sheet", dest="sheet_name", help="Workbook sheet name for tabular rescue mode")
    diagnose.add_argument("--all-sheets", dest="all_sheets", action="store_true", help="Consolidate compatible workbook sheets in tabular rescue mode")
    diagnose.add_argument("--mode", choices=["auto", "tabular", "workbook"], default="auto", help="Backend selection policy")
    diagnose.add_argument("-q", "--quiet", action="store_true", help="Minimal human logs")
    diagnose.add_argument("-v", "--verbose", action="store_true", help="More human logs")

    heal = subparsers.add_parser("heal", help="Heal a file and write cleaned outputs.")
    heal.add_argument("input", help="Input file path")
    heal.add_argument("output_positional", nargs="?", default=None, help="Optional output path")
    heal.add_argument("-o", "--out", dest="out_dir", help="Output directory")
    heal.add_argument("--output", dest="output_flag", help="Explicit output path")
    heal.add_argument("--format", choices=["xlsx", "csv"], default="xlsx", help="Output format for tabular healing")
    heal.add_argument("--in-place", action="store_true", help="Overwrite the input (CSV/tabular only)")
    heal.add_argument("--sheet", dest="sheet_name", help="Workbook sheet name for tabular rescue mode")
    heal.add_argument("--all-sheets", dest="all_sheets", action="store_true", help="Consolidate compatible workbook sheets in tabular rescue mode")
    heal.add_argument("--mode", choices=["auto", "tabular", "workbook"], default="auto", help="Backend selection policy")
    heal.add_argument("--json", action="store_true", help="Write machine JSON to stdout")
    heal.add_argument("--json-summary", dest="json_summary", help="Explicit JSON summary output path")
    heal.add_argument("--dry-run", action="store_true", help="Run healing logic without writing outputs")
    heal.add_argument("--fail-on-quarantine", action="store_true", help="Return exit code 5 instead of 4 when quarantine has rows")
    heal.add_argument("-q", "--quiet", action="store_true", help="Minimal human logs")
    heal.add_argument("-v", "--verbose", action="store_true", help="More human logs")

    report = subparsers.add_parser("report", help="Generate a human-readable or JSON report.")
    report.add_argument("input", help="Input file path")
    report.add_argument("-o", "--out", dest="out_dir", help="Output directory")
    report.add_argument("--output", help="Explicit report output path")
    report.add_argument("--sheet", dest="sheet_name", help="Workbook sheet name for tabular rescue mode")
    report.add_argument("--all-sheets", dest="all_sheets", action="store_true", help="Consolidate compatible workbook sheets in tabular rescue mode")
    report.add_argument("--mode", choices=["auto", "tabular", "workbook"], default="auto", help="Backend selection policy")
    report.add_argument("--json", action="store_true", help="Write machine JSON to stdout")
    report.add_argument("--format", choices=["text", "json"], default="text", help="Output format when --json is not used")
    report.add_argument("-q", "--quiet", action="store_true", help="Minimal human logs")
    report.add_argument("-v", "--verbose", action="store_true", help="More human logs")

    validate = subparsers.add_parser("validate", help="Validate a dataset against a simple schema.")
    validate.add_argument("input", help="Input file path")
    validate.add_argument("--schema", required=True, help="Schema path (.json supported; .yml/.yaml rejected honestly for now)")
    validate.add_argument("--sheet", dest="sheet_name", help="Workbook sheet name")
    validate.add_argument("--all-sheets", dest="all_sheets", action="store_true", help="Consolidate compatible workbook sheets")
    validate.add_argument("--json", action="store_true", help="Write machine JSON to stdout")
    validate.add_argument("-o", "--out", dest="out_dir", help="Output directory")
    validate.add_argument("--output", help="Explicit validation output path")
    validate.add_argument("-q", "--quiet", action="store_true", help="Minimal human logs")
    validate.add_argument("-v", "--verbose", action="store_true", help="More human logs")

    config = subparsers.add_parser("config", help="Generate or inspect configuration.")
    config_subparsers = config.add_subparsers(dest="config_command", required=True)
    config_init = config_subparsers.add_parser("init", help="Write a starter config file.")
    config_init.add_argument("--path", default="sheet-doctor.yml", help="Config output path")

    explain = subparsers.add_parser("explain", help="Explain a stable rule id.")
    explain.add_argument("rule_id", help="Rule identifier")
    explain.add_argument("--json", action="store_true", help="Write machine JSON to stdout")

    subparsers.add_parser("version", help="Print version")
    return parser


def maybe_emit_json_stdout(payload: Any, enabled: bool) -> None:
    if enabled:
        print(json_dumps(payload))


def safe_output_path(explicit: Path | None, default_path: Path, *, in_place: bool = False) -> Path:
    path = explicit or default_path
    if not in_place and output_exists(path):
        raise CliError(f"Refusing to overwrite existing output: {path}", EXIT_COMMAND_ERROR)
    return path


def diagnose_default_paths(args: argparse.Namespace, input_path: Path) -> tuple[Path, Path]:
    out_dir = determine_output_dir(args, input_path)
    report_path = Path(args.output) if args.output else out_dir / "report.json"
    return out_dir, report_path


def report_default_paths(args: argparse.Namespace, input_path: Path) -> tuple[Path, Path]:
    out_dir = determine_output_dir(args, input_path)
    suffix = ".json" if args.json or args.format == "json" else ".txt"
    report_path = Path(args.output) if args.output else out_dir / f"report{suffix}"
    return out_dir, report_path


def heal_default_paths(args: argparse.Namespace, input_path: Path, backend: str) -> tuple[Path, Path, Path]:
    out_dir = determine_output_dir(args, input_path)
    explicit_output = Path(args.output_flag) if args.output_flag else (Path(args.output_positional) if args.output_positional else None)
    if explicit_output and args.output_flag and args.output_positional:
        raise CliError("Use either positional output or --output, not both.", EXIT_COMMAND_ERROR)

    if args.in_place:
        if backend != "csv":
            raise CliError("--in-place is only supported for tabular CSV-style healing.", EXIT_COMMAND_ERROR)
        if input_path.suffix.lower() not in TABULAR_FORMATS:
            raise CliError("--in-place is only supported for .csv/.tsv/.txt/.json/.jsonl inputs.", EXIT_COMMAND_ERROR)
        if args.format != "csv":
            raise CliError("--in-place requires --format csv.", EXIT_COMMAND_ERROR)
        output_path = input_path
    elif explicit_output:
        output_path = explicit_output
    else:
        if backend == "excel":
            output_name = f"{input_path.stem}-cleaned{input_path.suffix}"
        elif args.format == "csv":
            output_name = f"{input_path.stem}-clean.csv"
        else:
            output_name = f"{input_path.stem}-cleaned.xlsx"
        output_path = out_dir / output_name

    summary_path = Path(args.json_summary) if args.json_summary else out_dir / "heal-summary.json"
    return out_dir, output_path, summary_path


def exit_code_for_diagnose_report(report: dict[str, Any]) -> int:
    summary = report.get("summary", {})
    if summary.get("issue_count", 0) > 0:
        return EXIT_DIAGNOSE_ISSUES
    return EXIT_SUCCESS


def exit_code_for_report_payload(payload: dict[str, Any]) -> int:
    summary = payload.get("summary", {})
    if summary.get("issue_count", 0) > 0:
        return EXIT_DIAGNOSE_ISSUES
    run_metrics = payload.get("run_summary", {}).get("metrics", {})
    if run_metrics.get("issues_found", 0) > 0:
        return EXIT_DIAGNOSE_ISSUES
    source_diagnose = payload.get("source_reports", {}).get("diagnose", {})
    if source_diagnose.get("summary", {}).get("issue_count", 0) > 0:
        return EXIT_DIAGNOSE_ISSUES
    return EXIT_SUCCESS


def run_diagnose(args: argparse.Namespace) -> int:
    input_path = Path(args.input)
    if not input_path.exists():
        eprint(f"File not found: {input_path}")
        return EXIT_COMMAND_ERROR

    try:
        backend = choose_backend(
            input_path=input_path,
            mode=args.mode,
            sheet_name=args.sheet_name,
            all_sheets=args.all_sheets,
        )
        out_dir, report_path = diagnose_default_paths(args, input_path)
        sheet_name = args.sheet_name
        if backend == "csv":
            sheet_name = maybe_default_sheet(input_path, args.sheet_name, args.all_sheets)
            report = load_script_module("csv-doctor", "diagnose.py").build_report(
                input_path,
                sheet_name=sheet_name,
                consolidate_sheets=True if args.all_sheets else None,
            )
        else:
            report = load_script_module("excel-doctor", "diagnose.py").build_report(input_path)
        report = normalize_report_for_cli(report)
        write_json(report_path, report)
        if args.json:
            maybe_emit_json_stdout(report, True)
        else:
            text = render_workbook_report_text(report) if backend == "excel" else render_tabular_diagnose_text(report)
            emit_human(text.rstrip(), quiet=args.quiet)
            emit_human(f"Report written: {report_path}", quiet=args.quiet)
        return exit_code_for_diagnose_report(report)
    except Exception as exc:
        eprint(str(exc))
        return classify_backend_exception(exc)


def run_report(args: argparse.Namespace) -> int:
    input_path = Path(args.input)
    if not input_path.exists():
        eprint(f"File not found: {input_path}")
        return EXIT_COMMAND_ERROR

    try:
        backend = choose_backend(
            input_path=input_path,
            mode=args.mode,
            sheet_name=args.sheet_name,
            all_sheets=args.all_sheets,
        )
        out_dir, report_path = report_default_paths(args, input_path)
        if backend == "excel":
            report = load_script_module("excel-doctor", "diagnose.py").build_report(input_path)
            normalized = normalize_report_for_cli(report)
            text_payload = render_workbook_report_text(normalized)
            machine_payload = normalized
        else:
            sheet_name = maybe_default_sheet(input_path, args.sheet_name, args.all_sheets)
            report = load_script_module("csv-doctor", "reporter.py").build_report(
                input_path,
                sheet_name=sheet_name,
                consolidate_sheets=True if args.all_sheets else None,
            )
            machine_payload = normalize_report_for_cli(report)
            text_payload = machine_payload["text_report"]

        if args.json or args.format == "json":
            write_json(report_path, machine_payload)
            if args.json:
                maybe_emit_json_stdout(machine_payload, True)
            else:
                emit_human(f"Report written: {report_path}", quiet=args.quiet)
        else:
            write_text(report_path, text_payload)
            emit_human(text_payload.rstrip(), quiet=args.quiet)
            emit_human(f"Report written: {report_path}", quiet=args.quiet)

        return exit_code_for_report_payload(machine_payload)
    except Exception as exc:
        eprint(str(exc))
        return classify_backend_exception(exc)


def write_tabular_csv_outputs(result: dict[str, Any], output_path: Path) -> dict[str, str]:
    clean_path = output_path
    quarantine_path = output_path.with_name(output_path.stem.replace("-clean", "-quarantine") + output_path.suffix)
    changelog_path = output_path.with_name(output_path.stem.replace("-clean", "-changelog") + output_path.suffix)
    ensure_parent(clean_path)
    headers = result["headers"]

    def render_rows(rows: list[list[Any]], extra_header: list[str] | None = None) -> str:
        all_headers = headers + (extra_header or [])
        lines = [",".join(json.dumps(str(value) if value is not None else "")[1:-1] for value in all_headers)]
        for row in rows:
            lines.append(",".join(json.dumps(str(value) if value is not None else "")[1:-1] for value in row))
        return "\n".join(lines) + "\n"

    clean_rows = [
        entry.row + ["TRUE" if entry.was_modified else "FALSE", "TRUE" if entry.needs_review else "FALSE"]
        for entry in result["clean_data"]
    ]
    quarantine_rows = [entry.row + [entry.reason] for entry in result["quarantine"]]
    changelog_rows = [
        [change.action, change.row_number, change.column_name, change.old_value, change.new_value, change.reason]
        for change in result["changelog"]
    ]

    write_text(clean_path, render_rows(clean_rows, ["was_modified", "needs_review"]))
    write_text(quarantine_path, render_rows(quarantine_rows, ["quarantine_reason"]))
    write_text(
        changelog_path,
        "\n".join(
            [
                "action,row_number,column_name,old_value,new_value,reason",
                *[
                    ",".join(json.dumps("" if value is None else str(value))[1:-1] for value in row)
                    for row in changelog_rows
                ],
            ]
        )
        + "\n",
    )
    return {
        "clean": str(clean_path),
        "quarantine": str(quarantine_path),
        "changelog": str(changelog_path),
    }


def run_heal(args: argparse.Namespace) -> int:
    input_path = Path(args.input)
    if not input_path.exists():
        eprint(f"File not found: {input_path}")
        return EXIT_COMMAND_ERROR

    try:
        backend = choose_backend(
            input_path=input_path,
            mode=args.mode,
            sheet_name=args.sheet_name,
            all_sheets=args.all_sheets,
        )
        if backend == "excel" and args.format == "csv":
            raise CliError("Workbook-native healing does not support --format csv.", EXIT_COMMAND_ERROR)

        out_dir, output_path, summary_path = heal_default_paths(args, input_path, backend)
        if not args.in_place and not args.dry_run:
            output_path = safe_output_path(output_path, output_path)
            summary_path = safe_output_path(summary_path, summary_path)

        if backend == "excel":
            excel_heal = load_script_module("excel-doctor", "heal.py")
            changes, stats = excel_heal.execute_healing(input_path, output_path if not args.dry_run else out_dir / "_dry_run.xlsx")
            summary = excel_heal.build_structured_summary(
                input_path=input_path,
                output_path=output_path,
                changes=changes,
                stats=stats,
            )
            summary = normalize_report_for_cli(summary)
            if args.dry_run:
                dry_path = out_dir / "_dry_run.xlsx"
                if dry_path.exists():
                    dry_path.unlink()
            else:
                write_json(summary_path, summary)
            if args.json:
                maybe_emit_json_stdout(summary, True)
            else:
                emit_human(render_excel_heal_summary(summary).rstrip(), quiet=args.quiet)
                if not args.dry_run:
                    emit_human(f"Healed workbook: {output_path}", quiet=args.quiet)
                    emit_human(f"Heal summary: {summary_path}", quiet=args.quiet)
            return EXIT_SUCCESS

        csv_heal = load_script_module("csv-doctor", "heal.py")
        role_overrides: dict[int, str] = {}
        sheet_name = maybe_default_sheet(input_path, args.sheet_name, args.all_sheets)
        result = csv_heal.execute_healing(
            input_path,
            sheet_name=sheet_name,
            consolidate_sheets=True if args.all_sheets else None,
            header_row_override=None,
            role_overrides=role_overrides,
        )
        summary = csv_heal.build_structured_summary(
            result,
            input_path=input_path,
            output_path=output_path,
            sheet_name=sheet_name,
            consolidate_sheets=True if args.all_sheets else None,
            header_row_override=None,
            role_overrides=role_overrides,
            plan_confirmed=False,
        )
        summary = normalize_report_for_cli(summary)
        output_manifest: dict[str, str] | None = None
        if not args.dry_run:
            if args.format == "xlsx":
                csv_heal.write_workbook(
                    result["clean_data"],
                    result["quarantine"],
                    result["changelog"],
                    output_path,
                    headers=result["headers"],
                )
                output_manifest = {"workbook": str(output_path)}
            else:
                output_manifest = write_tabular_csv_outputs(result, output_path)
            write_json(summary_path, summary)
        if args.json:
            payload = dict(summary)
            if output_manifest:
                payload["outputs"] = output_manifest
            maybe_emit_json_stdout(payload, True)
        else:
            emit_human(render_csv_heal_summary(result, output_path).rstrip(), quiet=args.quiet)
            if not args.dry_run:
                emit_human(f"Heal summary: {summary_path}", quiet=args.quiet)
        quarantine_rows = len(result["quarantine"])
        if quarantine_rows > 0:
            return EXIT_VALIDATE_FAILED if args.fail_on_quarantine else EXIT_HEAL_QUARANTINE
        return EXIT_SUCCESS
    except Exception as exc:
        eprint(str(exc))
        return classify_backend_exception(exc)


def load_validate_schema(schema_path: Path) -> dict[str, Any]:
    if not schema_path.exists():
        raise CliError(f"Schema not found: {schema_path}", EXIT_COMMAND_ERROR)
    suffix = schema_path.suffix.lower()
    if suffix not in SUPPORTED_SCHEMA_SUFFIXES:
        raise CliError("Schema must be .json, .yml, or .yaml", EXIT_COMMAND_ERROR)
    if suffix in {".yml", ".yaml"}:
        raise CliError("YAML schemas are not supported yet. Use JSON for now.", EXIT_COMMAND_ERROR)
    try:
        payload = json.loads(schema_path.read_text(encoding="utf-8"))
    except Exception as exc:
        raise CliError(f"Could not read schema: {exc}", EXIT_COMMAND_ERROR) from exc
    if not isinstance(payload, dict):
        raise CliError("Schema root must be a JSON object.", EXIT_COMMAND_ERROR)
    return payload


def infer_validation_sheet_name(input_path: Path, sheet_name: str | None, all_sheets: bool) -> str | None:
    return maybe_default_sheet(input_path, sheet_name, all_sheets)


def run_validate(args: argparse.Namespace) -> int:
    input_path = Path(args.input)
    if not input_path.exists():
        eprint(f"File not found: {input_path}")
        return EXIT_COMMAND_ERROR

    try:
        schema = load_validate_schema(Path(args.schema))
        loader = load_script_module("csv-doctor", "loader.py")
        column_detector = load_script_module("csv-doctor", "column_detector.py")
        sheet_name = infer_validation_sheet_name(input_path, args.sheet_name, args.all_sheets)
        loaded = loader.load_file(
            input_path,
            sheet_name=sheet_name,
            consolidate_sheets=True if args.all_sheets else None,
        )
        df = loaded["dataframe"]
        profile = column_detector.analyse_dataframe(df)
        required = schema.get("required_columns", [])
        types = schema.get("types", {})
        missing_columns = [name for name in required if name not in df.columns]
        type_mismatches = []
        columns = profile.get("columns", {})
        for column_name, expected_type in types.items():
            if column_name not in df.columns:
                continue
            detected = columns.get(column_name, {}).get("detected_type", "unknown")
            if detected != expected_type:
                type_mismatches.append({"column": column_name, "expected": expected_type, "detected": detected})
        payload = {
            "tool": "sheet-doctor",
            "command": "validate",
            "version": TOOL_VERSION,
            "input": str(input_path),
            "valid": not missing_columns and not type_mismatches,
            "missing_columns": missing_columns,
            "type_mismatches": sorted(type_mismatches, key=lambda item: item["column"]),
            "error_count": len(missing_columns) + len(type_mismatches),
            "sheet_name": loaded.get("sheet_name"),
        }
        payload = normalize_report_for_cli(payload)
        if args.output or args.out_dir:
            out_dir = determine_output_dir(args, input_path)
            output_path = Path(args.output) if args.output else out_dir / "validation.json"
            write_json(output_path, payload)
            emit_human(f"Validation report: {output_path}", quiet=args.quiet)
        if args.json:
            maybe_emit_json_stdout(payload, True)
        else:
            emit_human(render_validate_text(payload).rstrip(), quiet=args.quiet)
        return EXIT_SUCCESS if payload["valid"] else EXIT_VALIDATE_FAILED
    except Exception as exc:
        eprint(str(exc))
        return classify_backend_exception(exc)


def run_config_init(args: argparse.Namespace) -> int:
    config_path = Path(args.path)
    if config_path.exists():
        eprint(f"Refusing to overwrite existing config: {config_path}")
        return EXIT_COMMAND_ERROR
    payload = """io:
  max_rows_scan: 20000
  output_format: xlsx

rules:
  enable:
    - DATE_NORMALIZE
    - ENCODING_FIX
    - SCHEMA_ALIGN
    - DEDUP_EXACT
    - DEDUP_NEAR
  disable: []

dedupe:
  near_threshold: 0.92
  key_columns: []

dates:
  dayfirst: false
  prefer_iso: true

quarantine:
  on_unparseable_date: true
  on_column_shift: true
"""
    write_text(config_path, payload)
    emit_human(f"Config written: {config_path}")
    return EXIT_SUCCESS


def run_explain(args: argparse.Namespace) -> int:
    rule = EXPLAIN_RULES.get(args.rule_id)
    if rule is None:
        eprint(f"Unknown rule id: {args.rule_id}")
        return EXIT_COMMAND_ERROR
    payload = {
        "rule_id": args.rule_id,
        "description": rule["description"],
        "evidence": rule["evidence"],
        "auto_fixable": rule["auto_fixable"],
        "disable_hint": rule["disable_hint"],
    }
    if args.json:
        maybe_emit_json_stdout(payload, True)
    else:
        print(
            "\n".join(
                [
                    f"Rule: {args.rule_id}",
                    f"What it does: {payload['description']}",
                    f"What triggers it: {payload['evidence']}",
                    f"Auto-fixable: {'yes' if payload['auto_fixable'] else 'no'}",
                    f"How to avoid/disable it: {payload['disable_hint']}",
                ]
            )
        )
    return EXIT_SUCCESS


def run_version() -> int:
    print(TOOL_VERSION)
    return EXIT_SUCCESS


def main(argv: list[str] | None = None) -> int:
    try:
        parser = build_parser()
        args = parser.parse_args(argv)
        if args.command == "diagnose":
            return run_diagnose(args)
        if args.command == "heal":
            return run_heal(args)
        if args.command == "report":
            return run_report(args)
        if args.command == "validate":
            return run_validate(args)
        if args.command == "config":
            if args.config_command == "init":
                return run_config_init(args)
        if args.command == "explain":
            return run_explain(args)
        if args.command == "version":
            return run_version()
        raise CliError(f"Unknown command: {args.command}", EXIT_COMMAND_ERROR)
    except CliError as exc:
        eprint(str(exc))
        return exc.code


if __name__ == "__main__":
    raise SystemExit(main())
