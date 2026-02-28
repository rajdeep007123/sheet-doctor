#!/usr/bin/env python3
"""
loader.py — Universal file loader for sheet-doctor

Supports: .csv .tsv .txt .xlsx .xls .xlsm .ods .json .jsonl

Public API:
    result = load_file("path/to/file.csv")
    df     = result["dataframe"]

Result dict keys:
    dataframe         — pandas DataFrame (always present)
    detected_format   — "csv", "xlsx", "json", etc.
    detected_encoding — encoding name for text files; None for binary
    encoding_info     — full dict: detected, confidence, is_utf8, suspicious_chars
    delimiter         — delimiter char for text files; None otherwise
    raw_text          — decoded text for text files; None otherwise
    sheet_name        — active sheet name for spreadsheets; None otherwise
    sheet_names       — all available sheet names for spreadsheets; None otherwise
    original_rows     — row count including header row
    original_columns  — column count
    warnings          — list of warning strings
"""

from __future__ import annotations

import csv
import io
import json as _json
import sys
from pathlib import Path
from typing import Optional

# ── Format groups ──────────────────────────────────────────────────────────────
TEXT_FORMATS  = {".csv", ".tsv", ".txt"}
EXCEL_FORMATS = {".xlsx", ".xls", ".xlsm"}
ODS_FORMATS   = {".ods"}
JSON_FORMATS  = {".json"}
JSONL_FORMATS = {".jsonl"}
ALL_FORMATS   = TEXT_FORMATS | EXCEL_FORMATS | ODS_FORMATS | JSON_FORMATS | JSONL_FORMATS


# ══════════════════════════════════════════════════════════════════════════════
# ENCODING DETECTION
# ══════════════════════════════════════════════════════════════════════════════

def _detect_encoding_info(raw: bytes) -> dict:
    """
    Detect encoding from raw bytes.

    Returns dict with: detected, confidence, is_utf8, suspicious_chars.
    Uses chardet when available; falls back to heuristics.
    """
    try:
        import chardet
        result   = chardet.detect(raw)
        detected = result.get("encoding") or "unknown"
        confidence = round(result.get("confidence") or 0.0, 2)
    except ImportError:
        detected   = "unknown"
        confidence = 0.0

    is_utf8 = detected.upper().replace("-", "") in ("UTF8", "ASCII")

    suspicious: list[str] = []
    if not is_utf8:
        for row_idx, line in enumerate(raw.split(b"\n")[:100], start=1):
            try:
                line.decode("utf-8")
            except UnicodeDecodeError as e:
                bad_byte = line[e.start : e.end]
                suspicious.append(
                    f"row {row_idx}: byte {bad_byte!r} at position {e.start}"
                )

    return {
        "detected":        detected,
        "confidence":      confidence,
        "is_utf8":         is_utf8,
        "suspicious_chars": suspicious[:10],
    }


# ══════════════════════════════════════════════════════════════════════════════
# SAFE TEXT READING (mixed-encoding tolerant)
# ══════════════════════════════════════════════════════════════════════════════

def _read_text_safely(raw: bytes, preferred_encoding: str) -> str:
    """
    Decode raw bytes line-by-line.

    Strategy per line:
      1. Try UTF-8
      2. Try preferred_encoding (chardet result)
      3. Try latin-1
      4. CP1252 with replace (never crashes)

    Also strips embedded null bytes so downstream parsers don't choke.
    """
    decoded_lines: list[str] = []
    for raw_line in raw.split(b"\n"):
        decoded: str | None = None
        for enc in ("utf-8", preferred_encoding, "latin-1"):
            if not enc or enc == "unknown":
                continue
            try:
                decoded = raw_line.decode(enc)
                break
            except (LookupError, UnicodeDecodeError):
                continue
        if decoded is None:
            decoded = raw_line.decode("cp1252", errors="replace")
        decoded_lines.append(decoded.replace("\x00", ""))
    return "\n".join(decoded_lines)


# ══════════════════════════════════════════════════════════════════════════════
# DELIMITER DETECTION
# ══════════════════════════════════════════════════════════════════════════════

def _detect_delimiter(text: str) -> str:
    """
    Infer CSV delimiter from sample lines.

    Uses csv.Sniffer first; falls back to scoring each candidate delimiter by
    column-count consistency and column width.
    """
    from collections import Counter

    sample_lines = [l for l in text.splitlines() if l.strip()][:50]
    sample = "\n".join(sample_lines[:25])

    if sample:
        try:
            sniffed = csv.Sniffer().sniff(sample, delimiters=",;\t|")
            return sniffed.delimiter
        except csv.Error:
            pass

    candidates  = [",", ";", "\t", "|"]
    best_delim  = ","
    best_score  = float("-inf")
    best_width  = 0
    sample_text = "\n".join(sample_lines[:120])

    for delim in candidates:
        rows = [
            row
            for row in csv.reader(io.StringIO(sample_text), delimiter=delim)
            if any(cell.strip() for cell in row)
        ]
        if len(rows) < 2:
            continue

        widths      = [len(row) for row in rows]
        width_counts = Counter(widths)
        mode_width, mode_count = width_counts.most_common(1)[0]
        consistency  = mode_count / len(widths)
        header_width = len(rows[0])

        score = (mode_width * 2.0) + (consistency * mode_width)
        if header_width == mode_width:
            score += 1.0
        if mode_width == 1:
            score -= 10.0

        if score > best_score or (score == best_score and mode_width > best_width):
            best_score = score
            best_width = mode_width
            best_delim = delim

    return best_delim


def _sample_delimited_rows(text: str, delimiter: str, limit: int = 50) -> list[list[str]]:
    """Parse a sample of non-empty lines using the given delimiter."""
    sample_lines = [line for line in text.splitlines() if line.strip()][:limit]
    return [
        row
        for row in csv.reader(io.StringIO("\n".join(sample_lines)), delimiter=delimiter)
        if any(cell.strip() for cell in row)
    ]


def _validate_txt_table(text: str, delimiter: str) -> None:
    """
    Reject plain-text .txt files that are not actually delimited/tabular data.

    The loader should not report success for prose or notes files that only
    happen to end in .txt.
    """
    rows = _sample_delimited_rows(text, delimiter)
    if len(rows) < 2:
        raise ValueError(
            ".txt file does not appear to contain delimited/tabular data "
            "(need at least 2 non-empty rows)"
        )

    multi_field_rows = sum(1 for row in rows if len(row) > 1)
    if multi_field_rows < 2:
        raise ValueError(
            ".txt file does not appear to contain delimited/tabular data "
            f"(detected delimiter {delimiter!r} but fewer than 2 rows contain multiple fields)"
        )


def _sheet_selection_message(
    all_sheets: list[str],
    same_columns: bool,
    suffix: str,
) -> str:
    action = "pass sheet_name='...'"
    if same_columns:
        action += " or consolidate_sheets=True"
    return (
        f"Multiple sheets found in {suffix} workbook; {action}. "
        f"Available sheets: {all_sheets}"
    )


# ══════════════════════════════════════════════════════════════════════════════
# INTERACTIVE SHEET SELECTION
# ══════════════════════════════════════════════════════════════════════════════

def _prompt_sheet_choice(
    sheet_names: list[str],
    same_columns: bool,
) -> tuple[Optional[str], bool]:
    """
    Ask the user to pick a sheet or consolidate all sheets.

    Returns:
        (sheet_name, consolidate)
        sheet_name is None when consolidating all sheets.

    Caller must ensure stdin is interactive before using this helper.
    """
    print(f"\nFound {len(sheet_names)} sheets:", file=sys.stderr)
    for i, name in enumerate(sheet_names, 1):
        print(f"  {i}. {name}", file=sys.stderr)

    if same_columns:
        print(
            f"  {len(sheet_names) + 1}. Consolidate all sheets (same columns detected)",
            file=sys.stderr,
        )

    while True:
        try:
            raw = input("Enter sheet number or name: ").strip()
        except (EOFError, KeyboardInterrupt):
            print("", file=sys.stderr)
            return sheet_names[0], False

        if same_columns and raw == str(len(sheet_names) + 1):
            return None, True

        # Numeric choice
        try:
            idx = int(raw) - 1
            if 0 <= idx < len(sheet_names):
                return sheet_names[idx], False
        except ValueError:
            pass

        # Name match
        if raw in sheet_names:
            return raw, False

        print("Invalid choice — try again.", file=sys.stderr)


# ══════════════════════════════════════════════════════════════════════════════
# FORMAT LOADERS
# ══════════════════════════════════════════════════════════════════════════════

def _load_text(path: Path, suffix: str) -> dict:
    """Load .csv, .tsv, or .txt file into a pandas DataFrame."""
    import pandas as pd

    raw      = path.read_bytes()
    enc_info = _detect_encoding_info(raw)
    enc      = enc_info["detected"] if enc_info["detected"] != "unknown" else "utf-8"
    text     = _read_text_safely(raw, enc)

    if suffix == ".tsv":
        delimiter = "\t"
    else:
        delimiter = _detect_delimiter(text)

    if suffix == ".txt":
        _validate_txt_table(text, delimiter)

    sep = r"\|" if delimiter == "|" else delimiter
    try:
        df = pd.read_csv(
            io.StringIO(text),
            dtype=str,
            on_bad_lines="skip",
            sep=sep,
            engine="python",
        )
    except Exception as exc:
        raise ValueError(f"Could not parse {suffix} file: {exc}") from exc

    return {
        "dataframe":        df,
        "detected_format":  suffix.lstrip("."),
        "detected_encoding": enc,
        "encoding_info":    enc_info,
        "delimiter":        delimiter,
        "raw_text":         text,
        "sheet_name":       None,
        "sheet_names":      None,
        "original_rows":    len(df) + 1,
        "original_columns": len(df.columns),
        "warnings":         [],
    }


def _sheets_same_columns(
    path: Path,
    names: list[str],
    engine: Optional[str] = None,
) -> bool:
    """Return True when every listed sheet has identical column headers."""
    import pandas as pd

    header_sets: list[tuple] = []
    for name in names:
        try:
            df = pd.read_excel(path, sheet_name=name, nrows=0, dtype=str, engine=engine)
            header_sets.append(tuple(df.columns.tolist()))
        except Exception:
            pass
    return len(set(header_sets)) == 1


def _load_excel(
    path: Path,
    suffix: str,
    sheet_name: Optional[str] = None,
    consolidate_sheets: Optional[bool] = None,
) -> dict:
    """
    Load .xlsx, .xls, or .xlsm into a pandas DataFrame.

    When multiple sheets exist:
      - If sheet_name is given, load that sheet.
      - If consolidate_sheets is True and all sheets share columns, concatenate.
      - Otherwise prompt the user, or require explicit selection in
        non-interactive mode.
    """
    import pandas as pd

    warnings: list[str] = []

    # .xls requires xlrd; give a clear error if missing.
    if suffix == ".xls":
        try:
            import xlrd  # noqa: F401
        except ImportError:
            raise ImportError(
                ".xls files require xlrd — run: pip install xlrd"
            )

    try:
        with pd.ExcelFile(path) as xf:
            all_sheets = list(xf.sheet_names)
    except Exception as exc:
        raise ValueError(f"Could not open workbook: {exc}") from exc

    if len(all_sheets) == 1:
        chosen_name = all_sheets[0]
        consolidate = False
    elif sheet_name is not None:
        # Caller specified which sheet to use.
        if sheet_name not in all_sheets:
            raise ValueError(
                f"Sheet '{sheet_name}' not found. Available: {all_sheets}"
            )
        chosen_name = sheet_name
        consolidate = False
    else:
        same_cols   = _sheets_same_columns(path, all_sheets)
        consolidate = consolidate_sheets  # might be None

        if not sys.stdin.isatty():
            if consolidate:
                if not same_cols:
                    raise ValueError(
                        "Cannot consolidate workbook sheets with different columns. "
                        f"Available sheets: {all_sheets}"
                    )
                chosen_name = None
            else:
                raise ValueError(_sheet_selection_message(all_sheets, same_cols, suffix))
        else:
            if consolidate is None:
                chosen_name, consolidate = _prompt_sheet_choice(all_sheets, same_cols)
            elif consolidate and not same_cols:
                raise ValueError(
                    "Cannot consolidate workbook sheets with different columns. "
                    f"Available sheets: {all_sheets}"
                )
            elif not consolidate:
                chosen_name, consolidate = _prompt_sheet_choice(all_sheets, same_cols)

    if consolidate:
        frames = []
        for name in all_sheets:
            try:
                frames.append(pd.read_excel(path, sheet_name=name, dtype=str))
            except Exception as exc:
                warnings.append(f"Could not load sheet '{name}': {exc}")
        if not frames:
            raise ValueError("No sheets could be loaded.")
        df = pd.concat(frames, ignore_index=True)
        active_sheet = f"[all {len(frames)} sheets]"
        warnings.append(f"Consolidated {len(frames)} sheets into one table.")
    else:
        try:
            df = pd.read_excel(path, sheet_name=chosen_name, dtype=str)
        except Exception as exc:
            raise ValueError(
                f"Could not load sheet '{chosen_name}': {exc}"
            ) from exc
        active_sheet = chosen_name

        if len(all_sheets) > 1:
            others = [s for s in all_sheets if s != chosen_name]
            warnings.append(
                f"Multiple sheets found ({len(all_sheets)} total); "
                f"used '{active_sheet}'. Ignored: {others}"
            )

    return {
        "dataframe":        df,
        "detected_format":  suffix.lstrip("."),
        "detected_encoding": None,
        "encoding_info":    None,
        "delimiter":        None,
        "raw_text":         None,
        "sheet_name":       active_sheet,
        "sheet_names":      all_sheets,
        "original_rows":    len(df) + 1,
        "original_columns": len(df.columns),
        "warnings":         warnings,
    }


def _load_ods(
    path: Path,
    sheet_name: Optional[str] = None,
    consolidate_sheets: Optional[bool] = None,
) -> dict:
    """
    Load .ods OpenDocument spreadsheet using pandas with the 'odf' engine.

    Requires odfpy: pip install odfpy
    """
    import pandas as pd

    try:
        import odf  # noqa: F401
    except ImportError:
        raise ImportError(
            ".ods files require odfpy — run: pip install odfpy"
        )

    warnings: list[str] = []

    try:
        with pd.ExcelFile(path, engine="odf") as xf:
            all_sheets = list(xf.sheet_names)
    except Exception as exc:
        raise ValueError(f"Could not open .ods file: {exc}") from exc

    if len(all_sheets) == 1:
        chosen_name = all_sheets[0]
        consolidate = False
    elif sheet_name is not None:
        if sheet_name not in all_sheets:
            raise ValueError(
                f"Sheet '{sheet_name}' not found in .ods. Available: {all_sheets}"
            )
        chosen_name = sheet_name
        consolidate = False
    else:
        same_cols   = _sheets_same_columns(path, all_sheets, engine="odf")
        consolidate = consolidate_sheets

        if not sys.stdin.isatty():
            if consolidate:
                if not same_cols:
                    raise ValueError(
                        "Cannot consolidate .ods sheets with different columns. "
                        f"Available sheets: {all_sheets}"
                    )
                chosen_name = None
            else:
                raise ValueError(_sheet_selection_message(all_sheets, same_cols, ".ods"))
        else:
            if consolidate is None:
                chosen_name, consolidate = _prompt_sheet_choice(all_sheets, same_columns=same_cols)
            elif consolidate and not same_cols:
                raise ValueError(
                    "Cannot consolidate .ods sheets with different columns. "
                    f"Available sheets: {all_sheets}"
                )
            elif not consolidate:
                chosen_name, consolidate = _prompt_sheet_choice(all_sheets, same_columns=same_cols)

    if consolidate:
        frames = []
        for name in all_sheets:
            try:
                frames.append(pd.read_excel(path, sheet_name=name, engine="odf", dtype=str))
            except Exception as exc:
                warnings.append(f"Could not load sheet '{name}': {exc}")
        if not frames:
            raise ValueError("No .ods sheets could be loaded.")
        df = pd.concat(frames, ignore_index=True)
        active_sheet = f"[all {len(frames)} sheets]"
        warnings.append(f"Consolidated {len(frames)} sheets into one table.")
    else:
        try:
            df = pd.read_excel(path, sheet_name=chosen_name, engine="odf", dtype=str)
        except Exception as exc:
            raise ValueError(f"Could not load sheet '{chosen_name}': {exc}") from exc
        active_sheet = chosen_name

    return {
        "dataframe":        df,
        "detected_format":  "ods",
        "detected_encoding": None,
        "encoding_info":    None,
        "delimiter":        None,
        "raw_text":         None,
        "sheet_name":       active_sheet,
        "sheet_names":      all_sheets,
        "original_rows":    len(df) + 1,
        "original_columns": len(df.columns),
        "warnings":         warnings,
    }


def _load_json(path: Path) -> dict:
    """
    Load a .json file containing either an array of objects or a nested dict.

    Arrays → directly converted to a DataFrame.
    Dicts  → scans top-level keys for the first list value and uses that;
              falls back to treating the whole dict as a single-row table.
    Nested dicts/lists inside rows are flattened with pd.json_normalize().
    """
    import pandas as pd

    raw      = path.read_bytes()
    enc_info = _detect_encoding_info(raw)
    enc      = enc_info["detected"] if enc_info["detected"] != "unknown" else "utf-8"
    text     = raw.decode(enc, errors="replace")

    try:
        data = _json.loads(text)
    except _json.JSONDecodeError as exc:
        raise ValueError(f"Invalid JSON: {exc}") from exc

    warnings: list[str] = []

    if isinstance(data, list):
        records = data
    elif isinstance(data, dict):
        list_keys = [k for k, v in data.items() if isinstance(v, list)]
        if list_keys:
            key     = list_keys[0]
            records = data[key]
            warnings.append(f"Nested JSON: used array at top-level key '{key}'")
        else:
            records = [data]
            warnings.append("JSON is a single object; treated as a one-row table")
    else:
        raise ValueError(
            f"JSON root must be an array or object, got {type(data).__name__}"
        )

    try:
        df = pd.json_normalize(records)
    except Exception as exc:
        raise ValueError(f"Could not flatten JSON records: {exc}") from exc

    return {
        "dataframe":        df,
        "detected_format":  "json",
        "detected_encoding": enc,
        "encoding_info":    enc_info,
        "delimiter":        None,
        "raw_text":         text,
        "sheet_name":       None,
        "sheet_names":      None,
        "original_rows":    len(df) + 1,
        "original_columns": len(df.columns),
        "warnings":         warnings,
    }


def _load_jsonl(path: Path) -> dict:
    """
    Load a .jsonl (JSON Lines) file — one JSON object per line.

    Blank lines and parse errors are skipped with a warning.
    """
    import pandas as pd

    raw      = path.read_bytes()
    enc_info = _detect_encoding_info(raw)
    enc      = enc_info["detected"] if enc_info["detected"] != "unknown" else "utf-8"
    text     = raw.decode(enc, errors="replace")

    records: list[dict] = []
    parse_errors: list[str] = []

    for line_num, line in enumerate(text.splitlines(), start=1):
        line = line.strip()
        if not line:
            continue
        try:
            obj = _json.loads(line)
            records.append(obj)
        except _json.JSONDecodeError as exc:
            parse_errors.append(f"line {line_num}: {exc}")

    warnings: list[str] = []
    if parse_errors:
        sample = "; ".join(parse_errors[:3])
        extra  = f" (+{len(parse_errors) - 3} more)" if len(parse_errors) > 3 else ""
        warnings.append(f"{len(parse_errors)} lines could not be parsed — {sample}{extra}")

    df = pd.json_normalize(records) if records else pd.DataFrame()

    return {
        "dataframe":        df,
        "detected_format":  "jsonl",
        "detected_encoding": enc,
        "encoding_info":    enc_info,
        "delimiter":        None,
        "raw_text":         text,
        "sheet_name":       None,
        "sheet_names":      None,
        "original_rows":    len(df) + 1,
        "original_columns": len(df.columns),
        "warnings":         warnings,
    }


# ══════════════════════════════════════════════════════════════════════════════
# PUBLIC API
# ══════════════════════════════════════════════════════════════════════════════

def load_file(
    path: "str | Path",
    sheet_name: Optional[str] = None,
    consolidate_sheets: Optional[bool] = None,
) -> dict:
    """
    Load any supported file into a pandas DataFrame.

    Args:
        path:               Path to the file (str or Path).
        sheet_name:         For spreadsheets: which sheet to load.
                            None = auto-detect (prompt if interactive).
        consolidate_sheets: For spreadsheets with identical columns:
                            True  = merge all sheets into one DataFrame.
                            False = pick one sheet.
                            None  = prompt the user (or pick first silently).

    Returns:
        dict with keys: dataframe, detected_format, detected_encoding,
        encoding_info, delimiter, raw_text, sheet_name, sheet_names,
        original_rows, original_columns, warnings.

    Raises:
        FileNotFoundError  if the file does not exist.
        ValueError         if the format is unsupported or unreadable.
        ImportError        if a required optional dependency is missing.
    """
    path   = Path(path)
    suffix = path.suffix.lower()

    if not path.exists():
        raise FileNotFoundError(f"File not found: {path}")

    if suffix not in ALL_FORMATS:
        supported = ", ".join(sorted(ALL_FORMATS))
        raise ValueError(
            f"Unsupported format '{suffix}'. Supported: {supported}"
        )

    if suffix in TEXT_FORMATS:
        return _load_text(path, suffix)

    if suffix in EXCEL_FORMATS:
        return _load_excel(path, suffix, sheet_name, consolidate_sheets)

    if suffix in ODS_FORMATS:
        return _load_ods(path, sheet_name, consolidate_sheets)

    if suffix in JSON_FORMATS:
        return _load_json(path)

    if suffix in JSONL_FORMATS:
        return _load_jsonl(path)

    # Should never reach here given the earlier check, but be explicit.
    raise ValueError(f"Unhandled format: {suffix}")
