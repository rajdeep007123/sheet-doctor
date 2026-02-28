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
| `csv-doctor` | `heal.py` | Schema-aware healing (generic + finance mode) — outputs a 3-sheet Excel workbook (Clean Data / Quarantine / Change Log) |
| `excel-doctor` | `diagnose.py` | Deep Excel diagnostics: sheet inventory, merged cells, formula errors/cache misses, mixed types, duplicate/whitespace headers, structural rows, sparse columns |
| `excel-doctor` | `heal.py` | Safe workbook fixes: unmerge ranges, standardise/dedupe headers, clean text/date values, remove empty rows, append Change Log |
| `web` | `app.py` | Local Streamlit UI — upload files or paste public file URLs, describe what you want, preview the table, and download a human-readable workbook |

More skills coming: `merge-doctor`, `type-doctor`, `encoding-fixer`.

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

For **Excel/ODS files with multiple sheets**, the loader prompts you to pick a sheet in interactive sessions. In non-interactive/API use, it requires `sheet_name=...` or `consolidate_sheets=True` and raises a clear error listing the available sheets.

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

# Fix it — outputs extreme_mess_healed.xlsx with 3 sheets
python skills/csv-doctor/scripts/heal.py sample-data/extreme_mess.csv
```

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
│       ├── loader.py        ← universal file loader (used by both scripts below)
│       ├── diagnose.py      ← structural + semantic analysis, outputs JSON
│       ├── column_detector.py ← per-column type/quality inference, outputs JSON
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
