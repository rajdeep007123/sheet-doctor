# sheet-doctor

> Designer by heart. Back in code after 8 years. Built this because someone had to — and no one else was doing it for free.

**sheet-doctor** is a free, open-source Claude Code Skills Pack for diagnosing and fixing messy spreadsheet files. Drop a broken file in, get a human-readable health report out. No SaaS subscription. No upload limits. No data leaving your machine.

---

## What it does

Real-world spreadsheets are disasters. Wrong encodings, misaligned columns, five different date formats in the same column, blank rows, duplicate headers, formula errors — sheet-doctor finds all of it and tells you exactly what's wrong and where.

**Current skills:**

| Skill | Scripts | What it does |
|---|---|---|
| `csv-doctor` | `loader.py` | Universal file loader — reads `.csv .tsv .txt .xlsx .xls .xlsm .ods .json .jsonl` into a pandas DataFrame with encoding detection, delimiter sniffing, and explicit multi-sheet handling |
| `csv-doctor` | `diagnose.py` | Structural diagnostics plus column semantics: encoding, delimiter detection, column alignment, date formats, empty rows, duplicate headers, inferred column types, per-column quality stats, suspected issues |
| `csv-doctor` | `column_detector.py` | Standalone per-column semantic inference and quality profiling for messy tabular data, even when headers are weak or wrong |
| `csv-doctor` | `heal.py` | Schema-aware healing (`schema-specific`, `semantic`, `generic`) — outputs a 3-sheet Excel workbook (Clean Data / Quarantine / Change Log) with edge-case handling for formula residue, notes rows, subtotal rows, metadata/header rows, merged-cell export gaps, misplaced currency values, and explicit workbook sheet selection |
| `excel-doctor` | `diagnose.py` | Deep Excel diagnostics: sheet inventory, merged cells, formula errors/cache misses, mixed types, duplicate/whitespace headers, structural rows, sparse columns |
| `excel-doctor` | `heal.py` | Safe workbook fixes: unmerge ranges, standardise/dedupe headers, clean text/date values, remove empty rows, append Change Log |
| `web` | `app.py` | Local Streamlit UI — upload files or paste public file URLs, describe what you want, preview the table, and download a human-readable workbook |

More skills coming: `merge-doctor`, `type-doctor`, `encoding-fixer`.

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

Deployable machine outputs:
- JSON-producing scripts now emit `contract`, `schema_version`, `tool_version`, and `run_summary`
- `csv-doctor/heal.py` and `excel-doctor/heal.py` now support `--json-summary <path>` for backend/UI ingestion
- Stable schema docs live in [`schemas/`](./schemas)
- CI validates compile health, unit tests, and the sample CSV pipeline on every push/PR

---

## Supported file formats

`csv-doctor` reads all of these — no manual conversion needed:

| Format | Notes |
|--------|-------|
| `.csv` | Delimiter auto-detected (comma, tab, pipe, semicolon) |
| `.tsv` | Tab-separated |
| `.txt` | Sniffed like `.csv`; rejects plain-text files that are not tabular |
| `.xlsx` | Excel (modern) |
| `.xls` | Excel (legacy) — requires `pip install xlrd` |
| `.xlsm` | Excel macro-enabled — macros ignored, data loaded |
| `.ods` | OpenDocument spreadsheet — requires `pip install odfpy` |
| `.json` | Array of objects or nested dict (auto-flattened) |
| `.jsonl` | JSON Lines — one object per line |

For files with **mixed encodings** (Latin-1 and UTF-8 bytes on different rows), the loader decodes line-by-line and never crashes.

For **large inputs**, the loader now applies explicit safety rails:
- warns on large file sizes and row counts before processing becomes risky
- enters a visible `degraded_mode` when the file is likely to be slow or memory-heavy
- rejects files beyond hard in-memory safety limits with a clear error instead of crashing unpredictably

For **Excel/ODS files with multiple sheets**, the loader prompts you to pick a sheet in interactive sessions. In non-interactive/API use, it requires `sheet_name=...` or `consolidate_sheets=True` and raises a clear error listing the available sheets.

For **optional dependencies and corrupt workbook files**, the loader now fails more cleanly:
- missing `xlrd` for `.xls` raises a clear `ImportError`
- missing `odfpy` for `.ods` raises a clear `ImportError`
- corrupt workbook opens are wrapped so parser/library spew does not leak into user-facing output

`csv-doctor/heal.py` now exposes that workbook choice at the CLI layer:
- `--sheet <name>` to heal one workbook sheet
- `--all-sheets` to consolidate compatible sheets before healing

---

## Install

You need [Claude Code](https://claude.ai/code) and Python 3.9+ installed.

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

Or install from package metadata:

```bash
pip install .
pip install .[all]
```

Optional extras for additional formats:

```bash
pip install xlrd    # .xls legacy Excel files
pip install odfpy  # .ods OpenDocument files
```

**3. Run the loader tests**

```bash
python -m unittest discover -s tests -v
```

The test suite covers strict `.txt` rejection, multi-sheet workbook selection rules, and integration checks against public sample corpora when those fixtures are available locally.

CI runs the same checks from [`.github/workflows/ci.yml`](/Users/razzo/Documents/For%20Codex/sheet-doctor/.github/workflows/ci.yml), plus compile checks and a sample end-to-end CSV smoke test.

**4. Register the skills with Claude Code**

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
- Route to diagnose or heal flows
- Process files sequentially with an in-app progress loader
- Show source-specific notes for special URL handling such as Google Sheets export
- Download a readable `.xlsx` output

Current UI healing matrix:
- `.csv .tsv .txt .json .jsonl` → `csv-doctor/heal.py`
- `.xlsx .xlsm` → `excel-doctor/heal.py`
- `.xls .ods` → loader-based readable export fallback (preview + clean workbook export, not full workbook-preserving heal)

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

UI notes:
- Ongoing interface-specific notes live in `web/UI_CHANGELOG.md`

---

## Try it immediately

Sample files live in `sample-data/` — all deliberately broken for testing.

**`messy_sample.csv`** — encoding corruption, misaligned columns, 7 date formats, empty rows, duplicate header.

```bash
source .venv/bin/activate
python skills/csv-doctor/scripts/diagnose.py sample-data/messy_sample.csv
```

**`extreme_mess.csv`** — the full disaster: mixed Latin-1/UTF-8 encoding, BOM, null bytes, smart quotes, phantom commas, 7 date formats, 8 amount formats, near-duplicates, a TOTAL subtotal trap, and more. 50 rows of authentic chaos.

```bash
# Diagnose it
python skills/csv-doctor/scripts/diagnose.py sample-data/extreme_mess.csv

# Inspect per-column semantics only
python skills/csv-doctor/scripts/column_detector.py sample-data/extreme_mess.csv

# Build a plain-English health report + JSON artifact
python skills/csv-doctor/scripts/reporter.py sample-data/extreme_mess.csv

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

Current `extreme_mess.csv` score progression:
- Raw Health Score: `32/100`
- Recoverability Score: `84/100`
- Post-Heal Score: `95/100`

This means the source file is badly damaged, but most of it is recoverable automatically and the cleaned output is close to production-usable.

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
```

---

## How it works

Each skill is a `SKILL.md` that tells Claude what the skill does, plus a `scripts/` folder with the Python that does the work. Claude reads the skill, runs the script, interprets the output, and gives you a plain-English report.

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
