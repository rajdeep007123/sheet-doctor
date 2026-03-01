# sheet-doctor

**sheet-doctor** is a local-first spreadsheet triage and cleanup tool for messy CSVs, broken exports, and workbook-shaped reporting files.

It is strongest on:
- messy CSVs and table-like spreadsheet data
- explicit cleanup outputs: `Clean Data`, `Quarantine`, and `Change Log`
- headless, scriptable runs for local workflows and CI
- `.xlsx` / `.xlsm` workbook-native cleanup with explicit residual-risk warnings

CLI now available:

```bash
sheet-doctor diagnose file.xlsx
sheet-doctor heal file.csv
sheet-doctor report file.json
sheet-doctor validate file.csv --schema schema.json
sheet-doctor config init
sheet-doctor explain date_mixed_formats
```

Quick install:

```bash
pipx install .
```

Quick try:

```bash
sheet-doctor diagnose sample-data/extreme_mess.csv
sheet-doctor heal sample-data/extreme_mess.csv
sheet-doctor diagnose sample-data/messy_sample.xlsx
```

What it fixes:
- encoding mess
- shifted or misaligned rows
- duplicate headers and empty rows
- mixed date formats
- table-like workbook exports with preamble rows and stacked headers
- workbook structural issues like merged ranges and empty edge columns

Why this exists:
- spreadsheets are messy, but many cleanup tools are interactive first
- `sheet-doctor` is for repeatable local cleanup with audit artifacts
- it is useful when you want:
  - scriptable CLI runs
  - quarantine instead of silent deletion
  - explicit change logs
  - CI-friendly exit codes and JSON output

Why not just use OpenRefine?
- OpenRefine is interactive first; `sheet-doctor` is scriptable and headless first
- OpenRefine is great for exploratory cleanup; `sheet-doctor` is built for repeatable runs and saved artifacts
- `sheet-doctor` gives you quarantine and change logs as first-class outputs
- `sheet-doctor` is local-first and CI-friendly, but it is not a replacement for every spreadsheet workflow

## What it does

Real-world spreadsheets are usually export debris: wrong encodings, misaligned columns, mixed date formats, duplicate headers, subtotal rows, notes rows, formula residue, or workbooks with report-style preambles.

Core components:

| Component | Scripts | What it does |
|---|---|---|
| `csv-doctor` | `loader.py` | Universal file loader — reads `.csv .tsv .txt .xlsx .xls .xlsm .ods .json .jsonl` into a pandas DataFrame with encoding detection, delimiter sniffing, and explicit multi-sheet handling |
| `csv-doctor` | `diagnose.py` | Structural diagnostics plus column semantics: encoding, delimiter detection, column alignment, date formats, empty rows, duplicate headers, inferred column types, per-column quality stats, suspected issues |
| `csv-doctor` | `column_detector.py` | Standalone per-column semantic inference and quality profiling for messy tabular data, even when headers are weak or wrong |
| `csv-doctor` | `heal.py` | Schema-aware healing (`schema-specific`, `semantic`, `generic`) — outputs a 3-sheet Excel workbook (Clean Data / Quarantine / Change Log) with edge-case handling for formula residue, notes rows, subtotal rows, metadata/header rows, merged-cell export gaps, misplaced currency values, and explicit workbook sheet selection |
| `excel-doctor` | `diagnose.py` | Workbook-native Excel diagnostics for `.xlsx/.xlsm`: hidden/very-hidden sheets, merged cells, header bands, metadata rows, formula cells/errors/cache misses, mixed types, duplicate/whitespace headers, empty edge columns |
| `excel-doctor` | `heal.py` | Workbook-native Excel cleanup for `.xlsx/.xlsm`: unmerge ranges, flatten safe stacked headers, remove metadata rows and empty rows, trim empty edge columns, clean text/date values, preserve formulas, append Change Log |
| `web` | `app.py` | Local Streamlit UI — upload files or paste public file URLs, describe what you want, preview the table, and download a human-readable workbook |

## Who This Is For

sheet-doctor is for people dealing with messy tabular exports and broken spreadsheet structure.

Best fit:
- Analysts cleaning CSV exports from ERP, CRM, HR, finance, and survey tools
- Consultants working through client spreadsheet cleanup
- Ops teams that need a first-pass cleanup before manual review
- People who want local-first processing for sensitive files
- Teams that value:
  - a readable rescue output
  - explicit quarantine
  - a change log of what was touched
  - honest manual-review boundaries

Best current fit by mode:
- `csv-doctor` — ugly CSVs and table-like spreadsheet data
- `excel-doctor` — workbook-native cleanup for `.xlsx` / `.xlsm` when preserving workbook structure matters

## Who This Is Not For

sheet-doctor is not a general “fix any spreadsheet” system.

Not a good fit if you need:
- one-click repair of arbitrary Excel models
- fuzzy entity resolution across vendors, customers, or employees
- multi-file merge and reconciliation
- business-rule validation
- password-protected workbook repair
- streaming-scale processing for very large files
- a fully non-technical no-setup workflow

Be especially careful with:
- formula-heavy workbooks
- hidden-sheet-dependent workbooks
- complex reporting workbooks with spreadsheet logic that must be preserved and recalculated

In those cases, sheet-doctor can help with triage and safe cleanup, but it does not replace manual spreadsheet review.

## csv-doctor Status

- ✅ `loader.py` — universal file format handler
- ✅ `diagnose.py` — file health check
- ✅ `column_detector.py` — smart column analysis
- ✅ `reporter.py` — plain-text and JSON health report generator with recoverability scoring
- ✅ `heal.py` — fixer with Clean Data / Quarantine / Change Log output
- ✅ `skills/csv-doctor/SKILL.md` — Claude invocation guide for the full CSV workflow
- ✅ `skills/csv-doctor/README.md` — standalone developer documentation for the CSV skill
- ✅ `.github/workflows/ci.yml` — reproducible CI checks across supported Python versions
- ✅ `schemas/` — versioned JSON contract docs for deployable machine outputs
- ✅ `pyproject.toml` — package/release metadata with optional extras for `.xls` and `.ods`
- ✅ workbook-semantic healing — workbook preambles and non-exact workbook headers now flow through semantic mode instead of flattening into generic cleanup
- ✅ workbook override controls in the UI — users can override detected header rows and semantic roles before tabular rescue healing

## Architecture

`csv-doctor` is a pipeline, not five unrelated scripts:

1. `loader.py`
   Loads the incoming file into a pandas DataFrame, handling encoding, delimiter detection, and multi-format input.
2. `diagnose.py`
   Runs structural checks on the loaded file and embeds semantic analysis from `column_detector.py`.
3. `column_detector.py`
   Infers what each column likely contains and computes per-column quality signals.
4. `reporter.py`
   Turns the `diagnose.py` + `column_detector.py` JSON output into a non-technical health report and a UI-friendly JSON artifact, including raw, recoverability, and post-heal scoring.
5. `heal.py`
   Uses the same loader foundation to repair what can be repaired, quarantine what should not be trusted, and log every meaningful change.

In practice:
- `loader.py` is the shared ingestion layer
- `diagnose.py` and `column_detector.py` explain what is wrong
- `reporter.py` explains it in human terms
- `heal.py` produces the usable workbook output
- `SKILL.md` tells Claude when to invoke the workflow
- `skills/csv-doctor/README.md` documents the subsystem for developers
- `sheet_doctor/contracts.py` defines stable machine contracts for the UI/backend
- `schemas/` documents those contracts for CI and future API consumers

Report scores:
- `raw_health_score` measures how broken the original file is before repair
- `recoverability_score` measures how much usable data should remain after healing, based on the actual clean/quarantine split
- `post_heal_score` measures the expected quality of the cleaned output only
- `health_score` remains in the JSON as a backward-compatible alias of `raw_health_score`

Healing modes:
- `schema-specific` when the canonical 8-column finance/export shape is detected
- `semantic` when non-exact headers still strongly map to roles like name/date/amount/currency/status
- `generic` when only structural cleanup is safe

Workbook-semantic behavior:
- workbook inputs now preserve raw worksheet rows during healing instead of rebuilding rows from pandas headers
- leading report/preamble rows in workbooks can now be detected as metadata before semantic role inference runs
- semantic mode can now recover workbook tables with pre-header bands plus non-exact headers such as `Emp Name`, `Txn Date`, `Cost`, and `Approval State`
- stacked workbook header bands can now be merged into one semantic header row before normalization
- ragged workbook/report layouts with sparse leading or trailing columns now trim those sparse edges before header and semantic detection
- clinical/report-style workbook columns such as `Ward` now participate in semantic fill-down so merged-cell style blanks remain readable after healing
- non-financial workbook tables can now enter `semantic` mode when they still strongly map to roles such as `name`, `date`, `status`, `department`, and `notes`
- scientific/clinical workbook tables can now infer `identifier` and repeated `measurement` columns, so messy public workbook corpora do not immediately drop to generic mode

Deployable machine outputs:
- JSON-producing scripts now emit `contract`, `schema_version`, `tool_version`, and `run_summary`
- `csv-doctor/heal.py` and `excel-doctor/heal.py` now support `--json-summary <path>` for backend/UI ingestion
- `csv-doctor/heal.py` summaries now persist confirmed workbook rescue choices, including selected sheet, confirmed header row, semantic role overrides, and whether the user explicitly confirmed the plan
- Stable schema docs live in [`schemas/`](./schemas)
- CI validates compile health, unit tests, and the sample CSV pipeline on every push/PR

---

## Supported file formats

`csv-doctor` does not support every format equally. This matrix is for the tabular `csv-doctor` path:

| Format | Loader | Diagnose | Heal | UI |
|--------|--------|----------|------|----|
| `.csv` | ✅ | ✅ | ✅ | ✅ |
| `.tsv` | ✅ | ✅ | ✅ | ✅ |
| `.txt` tabular only | ✅ | ✅ | ✅ | ✅ |
| `.xlsx` | ✅ | ✅ (`--sheet` / `--all-sheets` for multi-sheet workbooks) | ✅ tabular rescue | ✅ |
| `.xls` | ✅ (`xlrd` required) | ✅ (`--sheet` / `--all-sheets`) | ⚠️ tabular rescue only | ⚠️ fallback/tabular rescue |
| `.xlsm` | ✅ | ✅ (`--sheet` / `--all-sheets`) | ✅ tabular rescue | ✅ |
| `.ods` | ✅ (`odfpy` required) | ✅ (`--sheet` / `--all-sheets`) | ⚠️ tabular rescue only | ⚠️ fallback/tabular rescue |
| `.json` | ✅ | ✅ | ✅ | ✅ |
| `.jsonl` | ✅ | ✅ | ✅ | ✅ |
| `.parquet` | ❌ | ❌ | ❌ | ❌ |

For files with **mixed encodings** (Latin-1 and UTF-8 bytes on different rows), the loader decodes line-by-line and never crashes.

For **large inputs**, the loader now applies explicit safety rails:
- warns on large file sizes and row counts before processing becomes risky
- enters a visible `degraded_mode` when the file is likely to be slow or memory-heavy
- rejects files beyond hard in-memory safety limits with a clear error instead of crashing unpredictably

Hard limits:
- `250 MB` file size
- `1,000,000` rows

For **Excel/ODS files with multiple sheets**, the loader prompts you to pick a sheet in interactive sessions. In non-interactive/API use, it requires `sheet_name=...` or `consolidate_sheets=True` and raises a clear error listing the available sheets.

For **optional dependencies and corrupt workbook files**, the loader now fails more cleanly:
- missing `xlrd` for `.xls` raises a clear `ImportError`
- missing `odfpy` for `.ods` raises a clear `ImportError`
- empty files raise `ValueError("File is empty")`
- encrypted/password-protected OOXML workbooks raise a clear unsupported-file error
- corrupt workbook opens are wrapped so parser/library spew does not leak into user-facing output

`csv-doctor/heal.py` now exposes that workbook choice at the CLI layer:
- `--sheet <name>` to heal one workbook sheet
- `--all-sheets` to consolidate compatible sheets before healing

---

## Install

You need Python installed.

Supported Python versions:
- CI-tested: `3.9`, `3.11`, `3.12`
- Project requirement: `>=3.9`

**1. Clone the repo**

```bash
git clone https://github.com/razzo007/sheet-doctor.git
cd sheet-doctor
```

**2. Create a virtual environment and install dependencies**

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

> On Windows: `.venv\Scripts\activate`

Optional extras for additional formats:

```bash
pip install xlrd    # .xls legacy Excel files
pip install odfpy  # .ods OpenDocument files
```

Install options:

```bash
pipx install .
```

or

```bash
pip install .
```

Both now create a working `sheet-doctor` CLI. Repo-first script usage still works for development.

**3. Run the loader tests**

```bash
python -m unittest discover -s tests -v
```

The test suite now uses committed in-repo fixtures under `tests/fixtures/` for:
- `.csv`
- `.tsv`
- `.xlsx`
- `.xlsm`
- `.ods`
- `.json`
- `.jsonl`
- corrupt workbook cases
- workbook preamble / stacked-header / ragged-layout rescue cases

Remaining optional/manual gap:
- happy-path legacy `.xls` loading is still not covered by a committed fixture because the repo does not ship a legacy `.xls` writer toolchain; CI still covers `.xls` missing-dependency and corrupt-file error handling

CI runs the same checks from [`.github/workflows/ci.yml`](/Users/razzo/Documents/For%20Codex/sheet-doctor/.github/workflows/ci.yml), plus compile checks, file-level coverage reporting, workbook/JSON diagnose-report smoke checks, and the sample end-to-end CSV pipeline.

**4. Optional: install the Claude Code skills**

You do not need Claude Code to use the CLI.

If you want to use the repo inside Claude Code as a local skill set:

Copy the skill folders into your Claude Code skills directory:

```bash
cp -r skills/csv-doctor ~/.claude/skills/csv-doctor
cp -r skills/excel-doctor ~/.claude/skills/excel-doctor
```

Or symlink them if you want edits to take effect immediately:

```bash
ln -s "$(pwd)/skills/csv-doctor" ~/.claude/skills/csv-doctor
ln -s "$(pwd)/skills/excel-doctor" ~/.claude/skills/excel-doctor
```

**5. Run it**

CLI:

```bash
sheet-doctor diagnose sample-data/extreme_mess.csv
sheet-doctor heal sample-data/extreme_mess.csv
sheet-doctor report sample-data/extreme_mess.csv
sheet-doctor validate sample-data/extreme_mess.csv --schema schema.json
sheet-doctor diagnose sample-data/messy_sample.xlsx
sheet-doctor heal sample-data/messy_sample.xlsx
```

Top-level commands:
- `sheet-doctor diagnose <input>`
- `sheet-doctor heal <input>`
- `sheet-doctor report <input>`
- `sheet-doctor validate <input> --schema <path.json>`
- `sheet-doctor config init`
- `sheet-doctor explain <rule-id>`
- `sheet-doctor version`

Exit codes:

| Code | Meaning |
|---|---|
| `0` | Success |
| `1` | Command error, bad args, missing file, or unexpected crash |
| `2` | Input parse/read failure |
| `3` | Diagnose/report found issues |
| `4` | Heal succeeded but quarantine rows exist |
| `5` | Validate failed, or heal was run with `--fail-on-quarantine` and quarantine rows exist |

Default output layout:

If you do not pass `--out` or `--output`, the CLI writes into:

```text
./sheet-doctor-output/<stem>-<timestamp>/
```

Typical files:
- `diagnose` -> `report.json`
- `report` -> `report.txt` or `report.json`
- `heal` -> cleaned workbook/file plus `heal-summary.json`

CI-friendly JSON mode:
- human-oriented logs go to `stderr`
- machine-readable JSON goes to `stdout` when `--json` is set

Example:

```bash
sheet-doctor diagnose sample-data/extreme_mess.csv --json > report.json
echo $?  # 3 when issues are found
```

Routing rules:
- `.csv .tsv .txt .json .jsonl` → `csv-doctor`
- `.xlsx .xlsm` → `excel-doctor` by default for `diagnose` / `heal`
- `.xlsx .xlsm` with `--sheet` or `--all-sheets` → `csv-doctor` tabular rescue/report path
- `.xls .ods` → explicit tabular fallback path

Examples:

```bash
# Workbook-native diagnosis
sheet-doctor diagnose sample-data/messy_sample.xlsx

# Tabular rescue view of one workbook sheet
sheet-doctor diagnose sample-data/messy_sample.xlsx --mode tabular --sheet Orders

# Workbook-native healing with machine-readable summary
sheet-doctor heal sample-data/messy_sample.xlsx --json-summary /tmp/messy_sample_heal.json

# Tabular report as JSON
sheet-doctor report sample-data/extreme_mess.csv --format json --output /tmp/extreme_mess_report.json

# Validate against a simple JSON schema
sheet-doctor validate sample-data/extreme_mess.csv --schema /tmp/schema.json

# Generate a starter config file
sheet-doctor config init

# Explain a stable rule id
sheet-doctor explain structural_misaligned_rows
```

In any Claude Code session:

```
/csv-doctor path/to/your/file.csv
/excel-doctor path/to/your/file.xlsx
```

Or just drop a file in and say: *"diagnose this CSV"* / *"fix my spreadsheet"*

**6. Run the local UI**

```bash
streamlit run web/app.py
```

What the UI currently does:
- Upload many files in one batch
- Paste public file URLs one per line
- Add a plain-English prompt like "make this readable for humans"
- Preview the loaded table before acting
- Preview workbook interpretation before healing, including detected header bands, metadata rows removed, effective headers, and chosen semantic columns
- Let the user override the detected workbook header row before healing
- Let the user override semantic column roles in the UI before tabular rescue healing
- Show a detected-vs-final mapping comparison so users can see what changed before they run rescue
- Require explicit confirmation of the workbook rescue plan before tabular rescue runs
- Route to diagnose or heal flows
- Process files sequentially with an in-app progress loader
- Show source-specific notes for special URL handling such as Google Sheets export
- Download a readable `.xlsx` output

Current UI healing matrix:
- `.csv .tsv .txt .json .jsonl` → `csv-doctor/heal.py`
- `.xlsx .xlsm` → `excel-doctor/heal.py` by default, or `csv-doctor/heal.py` in optional tabular rescue mode
- `.xls .ods` → `csv-doctor/heal.py` in tabular rescue mode when the workbook looks tabular enough, otherwise loader-based readable export fallback

Workbook modes in the UI:
- `excel-doctor` = workbook-native cleanup for `.xlsx/.xlsm`
- `csv-doctor` tabular rescue = flatten a workbook sheet into a readable 3-sheet table rescue
- `.xls` / `.ods` do not have workbook-native healing here; they stay on the tabular/fallback path

Public URL support:
- GitHub file URLs (`blob` links are rewritten to raw)
- Dropbox public links
- Google Drive public file links
- Google Sheets public share URLs (`/edit`) exported as `.xlsx`
- OneDrive public links
- Box public links
- Other direct public file URLs that return the file itself

Remote file-type detection:
- If the public URL hides the extension, the UI now infers the file type from the response headers and file signature
- This fixes masked links where the shared URL looks like a web page but actually serves an Excel/CSV/ODS file

## Known limitations

- `csv-doctor` is strongest on flat tabular data. Workbook rescue is a heuristic flatten-and-clean flow, not workbook-native reconstruction.
- `excel-doctor` is workbook-native only for `.xlsx` / `.xlsm`. It does not support workbook-native `.xls` repair.
- `excel-doctor` preserves formula cells as formulas. It does not recalculate formula results or reconstruct missing formula caches.
- `excel-doctor` now surfaces explicit manual-review warnings when formula cells, formula errors, or formula cache misses are present; those warnings are not auto-fix guarantees.
- `.xlsm` macros are only preserved when the output stays `.xlsm`.
- Password-protected / encrypted Excel files are not supported.
- Corrupted workbooks fail with explicit errors; they are not partially reconstructed.
- `.parquet` is not supported.
- Large files are still processed in memory. Guardrails prevent obviously unsafe runs, but this is not a streaming pipeline.
- happy-path `.xls` coverage remains optional/manual in CI because the repo does not currently ship a committed writable legacy `.xls` fixture
- Public URL mode depends on the remote server actually serving a directly downloadable file.
- Health scores are heuristic summaries, not guarantees of business correctness.
- `excel-doctor` is the better fit when preserving workbook structure matters more than flattening the data into a readable table.

## What this tool does not do

- It does not merge multiple files into one reconciled dataset.
- It does not do fuzzy entity resolution for vendors, customers, or employees.
- It does not preserve workbook-native logic when you choose tabular rescue; that path flattens the sheet into rows and columns.
- It does not perform workbook-native healing for `.xls`.
- It does not stream huge files; the current pipeline still loads data in memory.
- It does not promise business-truth validation. It cleans structure and common data mess, not accounting correctness or domain correctness.

## Out of scope

- `.parquet`
- password-protected / encrypted Excel workbooks
- survey multi-select normalization
- Access database import
- multi-file merge workflows
- fuzzy entity resolution beyond the current exact / near-duplicate heuristics

## Security / privacy

- Local files stay local unless you use public URL mode in the UI.
- Public URL mode makes outbound network requests to fetch remote files.
- Public URL mode now rejects remote files above `100 MB` before or during download; it is still not a hardened anti-abuse boundary.
- Uploaded/remote files are processed through temporary files during local UI use; they are not intended to be stored permanently, but this is not a hardened secure-processing boundary.
- No telemetry or analytics calls are built into the Python scripts themselves.

UI notes:
- Ongoing interface-specific notes live in `web/UI_CHANGELOG.md`

---

## Try it immediately

Sample files live in `sample-data/` — all deliberately broken for testing.

Direct script usage still works and is useful for development, but the installable `sheet-doctor` CLI is now the primary interface.

**`messy_sample.csv`** — encoding corruption, misaligned columns, 7 date formats, empty rows, duplicate header.

```bash
source .venv/bin/activate
python skills/csv-doctor/scripts/diagnose.py sample-data/messy_sample.csv
```

**`extreme_mess.csv`** — the full disaster: mixed Latin-1/UTF-8 encoding, BOM, null bytes, smart quotes, phantom commas, 7 date formats, 8 amount formats, near-duplicates, a TOTAL subtotal trap, and more. Small file, high mess.

```bash
# Diagnose it
python skills/csv-doctor/scripts/diagnose.py sample-data/extreme_mess.csv

# Diagnose a workbook sheet through csv-doctor's tabular lens
python skills/csv-doctor/scripts/diagnose.py sample-data/messy_sample.xlsx --sheet "Orders"

# Diagnose JSON
python skills/csv-doctor/scripts/diagnose.py /path/to/export.json

# Diagnose JSON Lines
python skills/csv-doctor/scripts/diagnose.py /path/to/export.jsonl

# Inspect per-column semantics only
python skills/csv-doctor/scripts/column_detector.py sample-data/extreme_mess.csv

# Build a plain-English health report + JSON artifact
python skills/csv-doctor/scripts/reporter.py sample-data/extreme_mess.csv

# Build a report from one workbook sheet
python skills/csv-doctor/scripts/reporter.py sample-data/messy_sample.xlsx /tmp/messy_orders_report.txt /tmp/messy_orders_report.json --sheet "Orders"

# Fix it — outputs extreme_mess_healed.xlsx with 3 sheets
python skills/csv-doctor/scripts/heal.py sample-data/extreme_mess.csv

# Fix it and emit a structured JSON summary for the UI/backend
python skills/csv-doctor/scripts/heal.py sample-data/extreme_mess.csv /tmp/extreme_mess_healed.xlsx --json-summary /tmp/extreme_mess_heal_summary.json

# Heal one workbook sheet explicitly
python skills/csv-doctor/scripts/heal.py /path/to/workbook.xlsx /tmp/healed.xlsx --sheet "Visible"

# Consolidate compatible workbook sheets before healing
python skills/csv-doctor/scripts/heal.py /path/to/workbook.xlsx /tmp/healed.xlsx --all-sheets
```

Machine-readable contracts:
- CSV diagnose: [`schemas/csv-diagnose.schema.json`](/Users/razzo/Documents/For%20Codex/sheet-doctor/schemas/csv-diagnose.schema.json)
- CSV report: [`schemas/csv-report.schema.json`](/Users/razzo/Documents/For%20Codex/sheet-doctor/schemas/csv-report.schema.json)
- CSV heal summary: [`schemas/csv-heal-summary.schema.json`](/Users/razzo/Documents/For%20Codex/sheet-doctor/schemas/csv-heal-summary.schema.json)
- Excel diagnose: [`schemas/excel-diagnose.schema.json`](/Users/razzo/Documents/For%20Codex/sheet-doctor/schemas/excel-diagnose.schema.json)
- Excel heal summary: [`schemas/excel-heal-summary.schema.json`](/Users/razzo/Documents/For%20Codex/sheet-doctor/schemas/excel-heal-summary.schema.json)

Current `extreme_mess.csv` score progression can be regenerated locally with `reporter.py`. Treat the scores as heuristic guidance, not a stability guarantee across every Python / pandas / chardet version combination.

Recent `csv-doctor/heal.py` edge-case coverage:
- Quarantines text formulas like `=SUM(...)` as `Excel formula found, not data`
- Detects and removes leading metadata/header rows before the real header, logging them as `File Metadata` in the Change Log
- Quarantines subtotal/total rows as `Calculated subtotal row`
- Quarantines long single-cell prose rows as `Appears to be a notes row`
- Repairs merged-cell style export gaps in categorical columns with logged forward-fill changes
- Splits combined values like `$1,200 USD` into `Amount` + `Currency`
- Uses semantic role inference to normalize non-exact headers such as `Emp Name`, `Txn Date`, `Cost`, `Curr`, and `Approval State`
- Accepts `--sheet` / `--all-sheets` for workbook inputs so multi-sheet files do not fail at the CLI boundary
- Reporter distinguishes between raw file damage, recoverability, and post-heal quality
- Reporter adds a PII warning when likely names, emails, phones, or national-ID-like values are detected

**`messy_sample.xlsx`** — broken Excel workbook with hidden sheets, merged cells, formula errors, and mixed column types.

```bash
# Diagnose
python skills/excel-doctor/scripts/diagnose.py sample-data/messy_sample.xlsx

# Heal — outputs messy_sample_healed.xlsx with a Change Log tab
python skills/excel-doctor/scripts/heal.py sample-data/messy_sample.xlsx

# Diagnose an .xlsm workbook while keeping workbook-native semantics
python skills/excel-doctor/scripts/diagnose.py /path/to/workbook.xlsm
```

`excel-doctor` scope, honestly:
- Supports workbook-native diagnosis and healing for `.xlsx` and `.xlsm`
- Rejects `.xls` with an explicit “use csv-doctor tabular rescue or convert to .xlsx first” message
- Detects hidden/very-hidden sheets, merged ranges, header bands, metadata/preamble rows, formula cells/errors/cache misses, duplicate/whitespace headers, empty rows/columns, empty edge columns, and mixed-type columns
- Heals safe structural/text/date issues while preserving workbook sheets and formulas
- Does not recalculate formulas, recover missing cache values, or repair encrypted/corrupted workbooks beyond failing cleanly
- Emits workbook triage explicitly:
  - `workbook_native_safe_cleanup`
  - `tabular_rescue_recommended`
  - `manual_spreadsheet_review_required`
- Adds residual-risk reporting so the workbook path says what was fixed, what remains, and what still requires manual spreadsheet review
- Emits workbook-native manual-review warnings when formulas, hidden sheets, or heuristic header-band detection mean spreadsheet context still matters after cleanup

Workbook triage, plainly:
- `workbook_native_safe_cleanup` — workbook-native cleanup is the right first step because the workbook mostly has safe structural/text issues
- `tabular_rescue_recommended` — the workbook looks more like an exported report or messy table, so `csv-doctor` tabular rescue is often clearer than preserving workbook layout
- `manual_spreadsheet_review_required` — spreadsheet logic or hidden workbook context matters enough that cleanup alone is not a safe interpretation

---

## Claude Code integration

`sheet-doctor` is a Python tool first. Claude Code integration is optional.

If you use Claude Code, each skill is a `SKILL.md` plus a `scripts/` folder. Claude reads the skill, runs the script, and uses the output as part of the workflow.

```
skills/
├── csv-doctor/
│   ├── SKILL.md             ← Claude reads this to understand the skill
│   └── scripts/
│       ├── loader.py        ← universal file loader (used by all csv-doctor scripts)
│       ├── diagnose.py      ← structural + semantic analysis, outputs JSON
│       ├── column_detector.py ← per-column type/quality inference, outputs JSON
│       ├── reporter.py      ← plain-text + JSON health report builder
│       └── heal.py          ← fixes all issues, writes .xlsx workbook
└── excel-doctor/
    ├── SKILL.md
    └── scripts/
        ├── diagnose.py      ← analyses the workbook, outputs JSON
        └── heal.py          ← applies safe fixes, writes healed workbook + Change Log
```

---

## Contributing

This is a solo side project from a designer who came back to code. Contributions are genuinely welcome — see [CONTRIBUTING.md](CONTRIBUTING.md).

---

## License

MIT. Free forever.
