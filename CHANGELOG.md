# Changelog

All notable changes to sheet-doctor are documented here.

---

## [Unreleased]

### Added
- **`skills/csv-doctor/README.md`** — standalone developer documentation for the CSV skill:
  - Explains what `csv-doctor` does, how each script works, accepted formats, output structure, manual commands, and the `extreme_mess.csv` example flow
- **`csv-doctor` / `reporter.py`** — human-readable health report generator:
  - Combines `diagnose.py` and `column_detector.py` into a plain-text report and a structured JSON artifact
  - Includes file overview, grouped issues, per-column breakdown, recommended actions, assumptions, and three scoring views: raw, recoverability, and post-heal
  - Emits a dedicated PII warning when likely names, emails, phone numbers, or national-ID-like values are detected
- **`csv-doctor` / `column_detector.py`** — standalone semantic profiler for messy tabular data:
  - Infers likely column meaning even when headers are weak or wrong (`date`, `currency/amount`, `plain number`, `percentage`, `email`, `phone`, `URL`, `country`, `currency code`, `name`, `categorical`, `free text`, `boolean`, `ID/code`, `unknown`)
  - Emits per-column quality stats: null/unique counts and percentages, top values, sample values, numeric/date min/max, mixed-type flag, and suspected issues
  - Detects semantic quality problems such as mixed date formats, inconsistent capitalisation, leading/trailing whitespace, slight-variation duplicates, near-constant values, outliers, and possible PII
- **`web/app.py`** — local Streamlit UI:
  - Upload local files or paste public file URLs, enter a plain-English prompt, preview the parsed table, and run diagnose or make-readable flows
  - Routes text/JSON formats to `csv-doctor`, modern Excel workbooks to `excel-doctor`, and falls back to loader-based readable export for `.xls` / `.ods`
  - Exposes sheet selection / consolidation controls for workbook preview in the UI
  - Supports sequential batch processing with in-app status/progress and per-file download buttons
- **`tests/test_loader.py`** — regression coverage for the universal loader:
  - Local behavior tests for strict `.txt` rejection and explicit multi-sheet workbook selection in non-interactive mode
  - Public corpus integration tests covering `.csv`, `.tsv`, `.xlsx`, `.xls`, `.xlsm`, `.ods`, `.json`, `.jsonl`, plus corrupt `.xls` failure handling
- **`tests/test_column_detector.py`** — regression coverage for semantic inference:
  - Locks down the expected type/issue profile for `sample-data/extreme_mess.csv`
  - Covers generic headers, whitespace and near-duplicate detection, and numeric/percentage ranges
- **`tests/test_heal_edge_cases.py`** — focused healer regression coverage:
  - Formula residue rows are quarantined
  - Leading metadata rows are removed before the actual header and logged
  - Combined amount/currency values are split correctly
  - Notes rows and subtotal rows are quarantined with explicit reasons
  - Merged-cell style blank runs are forward-filled in categorical columns
  - Semantic mode normalizes alternate headers and now covers workbook sheet selection / consolidation entry points

### Changed
- **`csv-doctor` / `loader.py`** — Phase 4 operational hardening:
  - Added large-file guardrails with explicit warning, degraded-mode, and hard-stop thresholds for risky in-memory loads
  - Result payloads now expose `degraded_mode` so callers and the UI can react explicitly to risky inputs
  - Missing optional dependencies now fail with clearer import contracts for `.xls` (`xlrd`) and `.ods` (`odfpy`)
  - Corrupt workbook open failures are now wrapped more quietly so parser noise does not leak into user-facing stderr/stdout
- **`csv-doctor` / `reporter.py`** — Phase 3 scoring and action planning:
  - Added `raw_health_score` for the original file, `recoverability_score` based on the actual clean/quarantine split, and `post_heal_score` for the cleaned output only
  - Kept `health_score` as a backward-compatible alias of the raw score for older consumers
  - Recommended actions now use actual healing projection data (`clean_rows`, `quarantine_rows`, `needs_review_rows`, real fix counts) instead of relying only on diagnostic heuristics
- **`csv-doctor` / `heal.py`** — added reusable in-memory execution via `execute_healing(...)` so reporting can use the real healing outcome without writing a workbook first
- **`csv-doctor` / `heal.py`** — Phase 1 + 2 production hardening:
  - Shared row-accounting and normalized issue reporting now keep `diagnose.py` and `reporter.py` aligned on raw vs parsed row counts
  - Semantic healing mode now handles non-exact headers by inferring likely `name`, `date`, `amount`, `currency`, `status`, `department`, `category`, and `notes` columns
  - Header detection was tightened so real tabular rows are less likely to be misclassified as pre-header metadata
- **`csv-doctor` / `heal.py`** — workbook CLI hardening:
  - Added `--sheet <name>` to heal an explicitly selected workbook sheet
  - Added `--all-sheets` to consolidate compatible sheets before healing
  - Hardened workbook preprocessing so numeric/no-header sheets no longer crash the healer
- **`requirements.txt`** — added `streamlit` and `requests` for the local UI layer and public-file URL imports
- **`csv-doctor` / `diagnose.py`** — now embeds `column_semantics` in the main JSON health report:
  - Includes per-column inferred types, quality stats, and suspected issues alongside the existing structural diagnostics
  - Summary issue counting now includes semantic issue presence and unknown-column detection as light signals without replacing the existing structural verdict model
- **`csv-doctor` / `heal.py`** — hardened for weird real-world export failures:
  - Quarantines formula strings left behind by Excel exports (for example `=SUM(...)`) as `Excel formula found, not data`
  - Detects pre-header metadata/header rows, removes them from the dataset, and logs them as `File Metadata` entries in the Change Log
  - Quarantines sparse TOTAL/SUM/Subtotal rows with numeric amounts as `Calculated subtotal row`
  - Quarantines long single-cell prose rows as `Appears to be a notes row`
  - Repairs merged-cell style blank runs in categorical columns using forward-fill with explicit Change Log entries
  - Splits combined amount/currency values such as `$1,200 USD` back into the correct columns before normalisation
- **`csv-doctor` / `loader.py`** — tightened file-loading contract:
  - `.txt` files now raise a clear error when they are plain text rather than delimited/tabular data
  - Multi-sheet `.xlsx`, `.xls`, and `.ods` files now require explicit `sheet_name` selection in non-interactive mode; `consolidate_sheets=True` is allowed only when columns match
  - Successful spreadsheet loads now include `sheet_names` in the result dict
  - Workbook metadata reads now close file handles cleanly to avoid resource warnings in batch/test runs
- **`csv-doctor` / `column_detector.py`** — extended PII coverage to include Aadhaar-like national-ID patterns in addition to emails, phone numbers, and names
- **`web/app.py`** — improved UI behavior:
  - Disabled upload/URL inputs only during active processing, not immediately after file selection
  - Removed Streamlit's top decoration/status strip and replaced it with an in-app loader message
  - Added public URL rewriting for GitHub, Dropbox, Google Drive, OneDrive, and Box share links
  - Added Google Sheets `/edit` → `.xlsx` export handling and response-based file-type inference when the shared URL hides the extension
- **`web/UI_CHANGELOG.md`** — added a dedicated UI-facing notes log for Streamlit interface changes
- **Docs** — README, SKILL, and CONTRIBUTING updated to match the stricter loader behavior, test command, current UI capabilities, the new `column_semantics` report shape, the new healer edge-case coverage, and the completed `csv-doctor` architecture/docs set

---

## [0.4.1] — 2026-02-28

### Changed
- **`CONTRIBUTING.md`** — updated to reflect `loader.py` as the shared file-reading layer:
  - Replaced "Healer scripts" wanted item (excel-doctor `heal.py` now exists) with "New format support in `loader.py`" and "Better heuristics in existing scripts"
  - Added code snippet showing how new skills should import `loader.py` instead of writing their own file-reading logic
  - Added `loader.py` conventions section: result dict contract, error vs warning handling, encoding fallback chain, optional dependency error messages, stderr-only prompts, `isatty()` check before calling `input()`
  - Removed `excel-doctor heal.py` from skill ideas table (done)

---

## [0.4.0] — 2026-02-28

### Added
- **`csv-doctor` / `loader.py`** — universal file loader shared by `diagnose.py` and `heal.py`:
  - Supports 9 formats: `.csv`, `.tsv`, `.txt`, `.xlsx`, `.xls`, `.xlsm`, `.ods`, `.json`, `.jsonl`
  - Encoding detection via chardet + line-by-line fallback chain (UTF-8 → detected encoding → Latin-1 → CP1252 with replace) — never crashes on encoding
  - Delimiter auto-detection via `csv.Sniffer` with fallback scoring (comma, tab, pipe, semicolon)
  - Multi-sheet Excel/ODS: prompts user interactively when a TTY is attached; silently picks first sheet with a warning when running as a subprocess
  - Consolidation option: when all sheets share the same columns, offers to merge them into one DataFrame
  - JSON support: array of objects → DataFrame directly; nested dict → finds first list key and uses it; single object → one-row table
  - JSON Lines support: parses each line independently, skips blanks and bad lines with warnings
  - Returns a standard dict: `dataframe`, `detected_format`, `detected_encoding`, `encoding_info`, `delimiter`, `raw_text`, `sheet_name`, `original_rows`, `original_columns`, `warnings`

### Changed
- **`csv-doctor` / `diagnose.py`** — refactored to use `loader.py` for all file I/O:
  - Removed: `detect_encoding()`, `read_csv_text_safely()`, `detect_delimiter()`, `load_pandas_df()` (all now in `loader.py`)
  - `check_date_formats()` and `check_columns_quality()` now accept a pandas DataFrame directly instead of rebuilding one internally
  - `main()` calls `load_file()` and unpacks `encoding_info`, `raw_text`, `delimiter`, `dataframe` from the result
- **`csv-doctor` / `heal.py`** — refactored to use `loader.py` for all file I/O:
  - Removed: `detect_delimiter()`, `read_mixed_encoding()` (replaced by `loader.py`)
  - New `read_file()` wrapper calls `load_file()` and returns `(raw_rows, delimiter)` — the rest of the processing pipeline is unchanged
  - `heal.py` can now read any format the loader supports, not just CSV
- **`csv-doctor` / `SKILL.md`** — added full `loader.py` documentation: format table, encoding strategy, multi-sheet behaviour, `load_file()` return dict, JSON handling rules, optional dependencies
- **`README.md`** — added "Supported file formats" section with the 9-format table; `loader.py` added to the skills table and folder structure; optional install commands for `xlrd` and `odfpy`

---

## [0.3.0] — 2026-02-27

### Added
- **`excel-doctor` / `diagnose.py`** — full Excel diagnostic script (13 checks):
  - Sheet inventory (empty + hidden sheets), merged cells, formula errors, formula cache misses, mixed data types, empty rows, empty columns, duplicate headers, whitespace headers, date format consistency, single-value columns, structural rows (TOTAL/subtotal), high-null columns
  - Dual-load pattern: loads workbook twice (with and without `data_only=True`) to detect cached vs. live formula state
  - Outputs JSON to stdout; exit code 0 on success, 1 on failure
- **`excel-doctor` / `heal.py`** — Excel healer that fixes what `diagnose.py` finds:
  - 4 healing passes: unmerge cells + fill from anchor value; normalise and deduplicate headers; clean text (BOM, null bytes, line breaks, smart quotes) + normalise dates to YYYY-MM-DD; remove fully empty rows
  - Writes a healed workbook in-place structure plus an added `Change Log` sheet
  - Prints a plain-text summary report to stdout
- **`sample-data/generate_xlsx.py`** — reproducible generator for `messy_sample.xlsx`

### Changed
- **`csv-doctor` / `heal.py`** — major rewrite: dual-mode processing
  - **Schema-specific mode**: detects the 8-column finance schema and applies deep normalisation (dates, amounts, currencies, names, status)
  - **Generic mode**: works on any CSV shape — auto-detects delimiter, normalises headers (deduplicates with `_2`/`_3` suffixes), repairs alignment, cleans text, removes exact duplicates
  - `SPARSE_THRESHOLD_SCHEMA` (50%) and `SPARSE_THRESHOLD_GENERIC` (25%) extracted as named constants with explanatory comments — single source of truth for both thresholds and their quarantine reason strings
  - `_clean_cell_text()` defined once and shared between schema-specific and generic paths (removed duplicate implementation)
- **`csv-doctor` / `diagnose.py`** — improved encoding and delimiter detection: per-line `_decode_line()` fallback chain (detected encoding → UTF-8 → Latin-1 → replace); `detect_delimiter()` scoring for comma/semicolon/tab/pipe
- **`csv-doctor` / `SKILL.md`** — documented dual-mode behaviour, generic mode fixes, updated examples
- **`excel-doctor` / `SKILL.md`** — documented `heal.py` healing passes, safe-fix list, and assumptions
- **`.gitignore`** — added `.DS_Store` and `sample-data/*_healed.*` (heal.py output files)

---

## [0.2.0] — 2026-02-27

### Added
- **`csv-doctor` / `heal.py`** — full CSV healer that outputs a 3-sheet Excel workbook:
  - Sheet 1 "Clean Data": fixed rows with `was_modified` and `needs_review` columns
  - Sheet 2 "Quarantine": rows that could not be fixed, with `quarantine_reason`
  - Sheet 3 "Change Log": one row per individual change (original value, new value, action, reason)
  - Decision tree: empty rows discarded; sparse/structural rows quarantined; fixable rows cleaned; partially fixable rows cleaned and flagged
  - Fixes encoding (BOM, null bytes, smart quotes, mixed Latin-1/UTF-8), structural problems (misaligned rows, phantom commas, unquoted commas in Notes, short rows), date formats (7 formats → YYYY-MM-DD), amounts (8 formats → float), currencies (7 formats → ISO 3-letter), names (Title Case, Last/First flip), and status values
  - Near-duplicate detection (same Name/Amount/Currency/Category, date within 2 days) — both rows kept, both flagged
  - Exact duplicate removal — first occurrence kept, rest logged as Removed
- **`sample-data/extreme_mess.csv`** — 50-row comprehensive disaster CSV covering every common real-world failure mode: mixed Latin-1/UTF-8 encoding, BOM, null bytes, smart quotes, 2 merged header rows, misaligned columns, phantom commas, 7 date formats, 8 amount formats, 7 currency formats, duplicate rows, near-duplicates, name/status/department inconsistencies, TOTAL subtotal trap, and line breaks inside cells
- **`sample-data/generate_extreme_mess.py`** — reproducible generator for `extreme_mess.csv`, writing raw bytes to produce authentic mixed-encoding chaos
- **`excel-doctor` / `SKILL.md`** — full skill documentation for the Excel diagnostic script
- `openpyxl` added to `requirements.txt`

### Changed
- **`csv-doctor` / `SKILL.md`** — updated to document both `diagnose.py` and `heal.py` as a pair
- **`excel-doctor` / `SKILL.md`** — cleaned up duplicate "Hidden sheets" section, added `data_only=True` caveat, updated JSON example to match `messy_sample.xlsx` exactly, added sample file reference table
- **`README.md`** — `excel-doctor` moved from "coming soon" to the skills table; `heal.py` documented; "Try it immediately" expanded to all three sample files; skills directory structure updated; GitHub URL corrected
- **`CONTRIBUTING.md`** — `excel-doctor` removed from wanted list; `heal.py` for excel-doctor added as top contribution opportunity; clone URL corrected; test commands updated; `heal.py` conventions documented; skill ideas table added

---

## [0.1.0] — 2026-02-27

### Added
- **`csv-doctor` / `diagnose.py`** — CSV diagnostic script covering: encoding detection (chardet), column alignment, date format consistency, empty rows, duplicate headers, whitespace headers, empty columns, single-value columns. Outputs JSON to stdout; Claude formats as a plain-English health report with HEALTHY / NEEDS ATTENTION / CRITICAL verdict
- **`csv-doctor` / `SKILL.md`** — skill definition telling Claude when and how to use the script
- **`sample-data/messy_sample.csv`** — deliberately broken CSV with Latin-1 encoding corruption, misaligned columns, 6 date formats, empty rows, and a duplicate header row
- **`requirements.txt`** — `pandas`, `chardet`, `openpyxl`
- **`.gitignore`** — excludes `.venv`
- **`README.md`**, **`CONTRIBUTING.md`**, **`LICENSE`** (MIT)
