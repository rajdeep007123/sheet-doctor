# Changelog

All notable changes to sheet-doctor are documented here.

---

## [Unreleased]

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
