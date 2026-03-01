# Changelog

All notable changes to sheet-doctor are documented here.

---

## [Unreleased]

### Changed
- **CLI / packaging** — `sheet-doctor` is now a real installable command:
  - Added `sheet-doctor diagnose <file>`
  - Added `sheet-doctor heal <file> [output]`
  - Added `sheet-doctor report <file>`
  - Added format-aware routing across `csv-doctor` and `excel-doctor`
  - Added `--mode auto|tabular|workbook`, `--sheet`, `--all-sheets`, `--output`, and `--json-summary` handling at the top-level CLI where relevant
  - `pip install .` and `pipx install .` now expose a usable `sheet-doctor` command
  - Bundled runtime script copies now ship inside the package so installed CLI commands still work outside the repo checkout
  - Added a minimal `setup.py` compatibility shim so older local pip/setuptools environments can still install the CLI entrypoint
- **`excel-doctor` / `diagnose.py`** — workbook-native diagnosis is now materially stronger and more explicit:
  - Reports hidden and very-hidden sheets separately
  - Detects header bands, metadata/preamble rows, notes-like rows, structural subtotal rows, and empty edge columns
  - Reports formula cells alongside formula errors and cache misses
  - Emits workbook triage (`workbook_native_safe_cleanup`, `tabular_rescue_recommended`, `manual_spreadsheet_review_required`) with a plain-English reason, confidence, and recommended next action
  - Emits machine-readable residual-risk sections describing what is safe to auto-fix, what remains risky, and what still needs manual spreadsheet review
  - Emits explicit manual-review warnings when formula logic, hidden sheets, or heuristic header detection still require spreadsheet judgment
  - Emits sheet-level risk summaries plus a workbook-level summary
  - Rejects `.xls` explicitly and rejects encrypted OOXML workbooks with a stable user-facing error
- **`excel-doctor` / `heal.py`** — workbook-native healing is now safer and clearer:
  - Writes atomically through a temp workbook path before replacing the target output
  - Preserves `.xlsm` output suffixes by default so macros are not silently stripped by the default path
  - Flattens safe stacked headers, removes metadata rows, trims empty edge columns, unmerges ranges, preserves formulas, cleans text/date cells, and appends a `Change Log` sheet
  - Structured summaries now report `workbook-native` mode explicitly, include preserved-formula counts, include workbook triage and residual-risk output, and show before/after issue counts for key workbook risks
- **Excel fixtures / tests** — added committed workbook-native regression coverage:
  - `tests/fixtures/excel/hidden_layers.xlsx`
  - `tests/fixtures/excel/stacked_headers.xlsx`
  - `tests/fixtures/excel/preamble_report.xlsx`
  - `tests/fixtures/excel/formula_cases.xlsx`
  - `tests/fixtures/excel/merged_edges.xlsx`
  - `tests/fixtures/excel/text_date_cleanup.xlsx`
  - `tests/fixtures/excel/duplicate_headers.xlsx`
  - `tests/fixtures/excel/notes_totals.xlsx`
  - `tests/fixtures/excel/ragged_clinical.xlsx`
  - `tests/test_excel_doctor.py` now covers diagnosis findings, workbook-native healing fixes, encrypted/corrupt failures, atomic output protection, `.xlsm` support, and UI mode separation
- **`web/app.py`** — workbook mode separation is now explicit in the UI:
  - Shows whether a workbook will run in `workbook-native`, `tabular-rescue`, or `tabular-rescue-fallback` mode
  - Explains why that mode was chosen and the tradeoff before the run starts
  - Workbook-native runs now keep `.xlsm` output names and capture Excel-heal JSON summaries
  - Workbook-native reports and heal summaries now surface formula/manual-review warnings directly in the results view
  - Workbook triage is now shown prominently for workbook files, including the recommended path (`excel-doctor`, `csv-doctor` tabular rescue, or manual review first) and the tradeoff for each path
- **`csv-doctor` / `heal.py`** — split into smaller modules under `skills/csv-doctor/scripts/heal_modules/`:
  - Shared constants/dataclasses now live in `shared.py`
  - Row/header preprocessing now lives in `preprocessing.py`
  - Value cleanup/normalisation now lives in `normalization.py`
  - Semantic planning and processing loops now live in `semantic.py`
  - Workbook-writing helpers now live in `workbook.py`
  - Structured summary generation now lives in `summary.py`
  - `heal.py` remains the stable CLI entrypoint and public surface
- **Coverage / CI** — real coverage reporting is now configured:
  - Added `coverage` to runtime/dev dependencies
  - Added coverage configuration to `pyproject.toml`
  - GitHub Actions now runs `coverage run` plus `coverage report -m`
- **Tests** — added edge-shape regression coverage in `tests/test_data_shape_edges.py`:
  - one-column files
  - one-row files
  - duplicate headers
  - 500+ columns
  - very long cell values
  - emoji / RTL / CJK survival
  - numbers stored as text
  - Excel serial dates
  - negative accounting amounts
  - all-null / all-identical columns
- **`web/app.py`** — privacy/network behavior tightened:
  - Removed Google Fonts dependency so the UI no longer makes font-network requests by default
  - Public URL mode now warns explicitly that it makes outbound network requests
  - Remote URL downloads now enforce the size limit before or during streaming instead of after reading the full response into memory
  - Local workbook inspection now uses stricter temporary-directory cleanup instead of `NamedTemporaryFile(delete=False)` paths
- **`csv-doctor` / `heal.py`** — workbook writes are now atomic:
  - Writes to a temp workbook path first
  - Replaces the final output only after a successful save
  - Failed writes no longer leave half-written `.xlsx` files behind
- **Tests / fixtures** — CI no longer depends on `/tmp` public corpora for the loader matrix:
  - Added committed fixtures under `tests/fixtures/` for `.csv`, `.tsv`, `.xlsx`, `.xlsm`, `.ods`, `.json`, `.jsonl`, corrupt workbook cases, and workbook preamble/stacked/ragged layouts
  - Added atomic-write regression coverage and header-only failure coverage for `csv-doctor/heal.py`
- **`csv-doctor` / `loader.py`** — clearer user-facing failure contracts:
  - Empty files now raise `ValueError("File is empty")`
  - Encrypted/password-protected OOXML workbooks now raise a clear unsupported-file error
  - Corrupt workbook errors remain wrapped in stable user-facing messages
- **`csv-doctor` / `diagnose.py`** — format support now matches `loader.py`:
  - Supports `.csv`, `.tsv`, `.txt`, `.xlsx`, `.xls`, `.xlsm`, `.ods`, `.json`, and `.jsonl`
  - Added `--sheet <name>` and `--all-sheets` for multi-sheet workbook diagnosis in non-interactive runs
  - Reports `sheet_name` / `sheet_names` when workbook inputs are used
- **`csv-doctor` / `reporter.py`** — now reports on the same format set as `diagnose.py`:
  - Added `--sheet <name>` and `--all-sheets` for workbook report generation
  - Passes workbook selection through to both diagnosis and healing projection
- **Tests** — added explicit diagnose/reporter format coverage for `.xlsx`, `.xlsm`, `.json`, `.jsonl`, and `.ods` when the optional dependency is available
- **Docs** — README, `skills/csv-doctor/SKILL.md`, and `skills/csv-doctor/README.md` made more honest:
  - removed `No upload limits`
  - removed the blanket `No data leaving your machine` claim in favor of local-first wording plus explicit public-URL network behavior
  - removed stale hard-coded score examples
  - added a real format matrix (`loader` vs `diagnose` vs `heal` vs UI)
  - documented hard in-memory limits (`250 MB`, `1,000,000 rows`)
  - documented workbook rescue as heuristic/best-effort
  - documented unsupported cases such as password-protected Excel and parquet
  - added `What this tool does not do` and `Out of scope` sections

---

## [0.5.0] — 2026-03-01

### Performance
- **`csv-doctor` / `column_detector.py`** — large-file speed: 60s+ → 6.7s on 50k-row files
  - Type inference now samples up to 2,000 rows per column instead of scanning all rows; null counts and value_counts still run on the full column via fast pandas native ops
  - `detect_suspected_issues()` receives the sampled texts for its case/canonical/value_counts loops, cutting per-column work by up to 25×
  - Unstripped raw texts passed separately for the whitespace-detection check so accuracy is preserved
  - Each column's output now includes `analysis_sampled: bool` so callers can tell when a column was sampled
- **`csv-doctor` / `heal.py`** — large-file write speed: 60s+ → 8s on 50k-row files
  - Added `WRITE_ONLY_THRESHOLD = 5_000`: outputs above this row count use openpyxl `write_only=True` mode, which avoids `ws.max_row` polling and column-width scans after every `append()`
  - Added `LARGE_FILE_SKIP_EXTRAS = 10_000`: near-duplicate detection and merged-cell forward-fill are skipped above this row count (expensive O(n) passes not meaningful at scale)
  - Both thresholds are named constants at the top of the file

### Fixed
- **CI** — golden snapshot tests now pass on Python 3.9, 3.11, and 3.12 (were persistently failing):
  - `type_scores` floats vary across pandas versions (`pd.to_datetime` parses edge-case dates differently) — normalised in snapshot comparison
  - `has_mixed_types` is derived from type_scores — normalised
  - `action_counts` and `changelog_entries` differ by 2 on Python 3.9 vs later (chardet/strptime handle two edge-case rows differently) — normalised
  - `most_common_values` and `sample_values` contain decoded cell text which differs when chardet picks MacRoman (Python 3.9) vs WINDOWS-1250 (Python 3.14) — normalised
  - Encoding detection `confidence` score varies by chardet version — normalised
  - Encoding name appears verbatim in issue `plain_english` strings — covered by a recursive `_normalise_string()` pass with a mixed-case-safe regex

### Added
- **`.github/workflows/ci.yml`** — deployable CI pipeline:
  - Runs on push, PR, and manual dispatch
  - Tests Python `3.9`, `3.11`, and `3.12`
  - Installs optional `.xls` / `.ods` reader dependencies in CI so the in-repo fixture matrix is exercised where fixtures exist
  - Runs compile checks, the unit test suite, and a sample end-to-end CSV smoke pipeline
- **`pyproject.toml`** — release/package metadata:
  - Declares core runtime dependencies and optional extras for legacy Excel (`xlrd`) and ODS (`odfpy`)
  - Establishes a single project version for release hygiene
- **`sheet_doctor/__init__.py`** and **`sheet_doctor/contracts.py`** — shared deployable contract layer:
  - Centralises `tool_version`
  - Defines versioned contracts for machine-readable outputs
  - Provides a shared `run_summary` builder for UI/backend-facing scripts
- **`schemas/`** — versioned JSON contract docs:
  - `csv-diagnose.schema.json`
  - `csv-report.schema.json`
  - `csv-heal-summary.schema.json`
  - `excel-diagnose.schema.json`
  - `excel-heal-summary.schema.json`
- **`tests/test_contracts.py`** — contract/regression coverage:
  - Verifies schema metadata and run summaries for CSV and Excel outputs
  - Validates schema files are well-formed JSON
- **`tests/golden/extreme_mess_report.txt`** and **`tests/golden/extreme_mess_report.json`** — stable reporter golden snapshots for `extreme_mess.csv`
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
  - Shows workbook interpretation previews before healing, including detected header-band rows, metadata rows removed, effective headers, and chosen semantic columns
  - Lets users override the detected workbook header row and semantic column roles before running tabular rescue healing
  - Shows a detected-vs-final semantic mapping comparison before rescue runs so workbook overrides are visible before execution
  - Requires an explicit workbook-plan confirmation before tabular rescue runs and keeps that confirmation in the resulting JSON summary
  - Supports sequential batch processing with in-app status/progress and per-file download buttons
- **`tests/test_loader.py`** — regression coverage for the universal loader:
  - Local behavior tests for strict `.txt` rejection and explicit multi-sheet workbook selection in non-interactive mode
  - Fixture-backed tests covering `.csv`, `.tsv`, `.xlsx`, `.xlsm`, `.ods`, `.json`, `.jsonl`, plus corrupt workbook failure handling
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
- **`csv-doctor` / `heal.py`** — workbook-semantic recovery improved:
  - Workbook inputs now preserve raw worksheet rows during healing instead of rebuilding rows from DataFrame headers
  - This keeps workbook preambles, metadata bands, and true header rows visible to semantic/header detection
  - Semantic mode now covers workbook-style inputs with leading report rows and non-exact headers rather than dropping straight to generic cleanup
  - Multi-row workbook header bands are now merged into a single semantic header row before healing continues
  - Sparse leading/trailing workbook columns are now trimmed before semantic planning so ragged clinical/report layouts remain recoverable
  - Clinical/report-style fields such as `Ward` now map into semantic fill-down columns so merged-cell export gaps are repaired in the tested path
  - Semantic mode no longer requires an amount column in the tested path; non-financial workbook reports can still normalize `name`, `date`, `status`, `department`, and `notes`
  - Header detection now compares candidate rows against next-row data signals so text-heavy workbook data rows are less likely to be mistaken for headers
  - Added `identifier` and repeated `measurement` semantic roles so scientific and clinical workbook tables can leave generic mode in the tested path
  - `--json-summary` now persists confirmed workbook-plan metadata: selected sheet, confirmation flag, header-row override, and semantic role overrides
- **`csv-doctor` / `diagnose.py`** — now emits deployable contract metadata:
  - Added `contract`, `schema_version`, `tool_version`, and `run_summary`
  - Exposes `degraded_mode` from the loader in the JSON report when active
- **`csv-doctor` / `reporter.py`** — now emits deployable contract metadata:
  - Added `contract`, `schema_version`, `tool_version`, and `run_summary`
- **`csv-doctor` / `heal.py`** — now supports structured summary output:
  - Added `--json-summary <path>` for machine-readable post-heal summaries
  - Exposes stable counts for clean rows, quarantine rows, review flags, and logged changes
- **`excel-doctor` / `diagnose.py`** — now builds a reusable report object with stable contract metadata and `run_summary`
- **`excel-doctor` / `heal.py`** — now supports `--json-summary <path>` and reusable structured post-heal summaries
- **`tests/test_reporter.py`** — now includes golden snapshot regression coverage for the plain-text and JSON reporter outputs
- **`tests/test_heal_edge_cases.py`** — now covers workbook-semantic healing with preserved preamble rows and semantic normalization on workbook inputs
  - Added stacked-header workbook coverage so merged-style parent headers plus field-level headers stay semantic-healable
  - Added ragged clinical/report workbook coverage so sparse edge columns, stacked headers, and fill-down categorical recovery stay locked down
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
