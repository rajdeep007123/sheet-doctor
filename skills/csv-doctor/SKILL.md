# csv-doctor

Comprehensive CSV and tabular-file rescue skill for Claude Code.

Use this skill when the user needs any of the following:
- load a messy tabular file in almost any spreadsheet-like format
- diagnose why a CSV or exported spreadsheet is broken
- infer what each column probably means when headers are weak or wrong
- generate a human-readable health report for non-technical users
- repair a messy file into a usable workbook with `Clean Data`, `Quarantine`, and `Change Log` sheets

Typical trigger phrases:
- "diagnose this CSV"
- "what is wrong with this export"
- "make this spreadsheet readable"
- "fix this tabular file"
- "check this file before import"
- "give me a report a human can understand"
- "/csv-doctor"

This skill is appropriate for:
- `.csv`
- `.tsv`
- `.txt` that is actually tabular
- `.xlsx`
- `.xls`
- `.xlsm`
- `.ods`
- `.json`
- `.jsonl`

This skill is not just for literal CSV files. Use it whenever the user has a messy table-like file and the goal is diagnosis, explanation, or repair.

## Scripts

### `scripts/loader.py`
Shared ingestion layer for every other script.

What it does:
- detects delimiter for text formats
- detects encoding and survives mixed-encoding text
- loads Excel, ODS, JSON, and JSONL into pandas
- returns a standard dict with the DataFrame plus loader metadata

Use it whenever you need a DataFrame from a user file.

### `scripts/diagnose.py`
Structural health checker.

What it checks:
- encoding problems
- misaligned rows / broken column counts
- mixed date formats
- empty rows
- duplicate headers / repeated header rows
- whitespace headers
- empty columns
- single-value columns
- `column_semantics` from `column_detector.py`

Primary output:
- JSON to stdout

Use it when the user asks:
- what is broken
- whether the file is safe to import
- what kinds of issues exist

### `scripts/column_detector.py`
Semantic column profiler.

What it infers:
- `date`
- `currency/amount`
- `plain number`
- `percentage`
- `email address`
- `phone number`
- `URL`
- `country name or code`
- `currency code`
- `name`
- `categorical`
- `free text`
- `boolean`
- `ID/code`
- `unknown`

What it reports per column:
- null counts and percentages
- unique counts and percentages
- most common values
- min/max for numeric/date-like columns
- sample values
- mixed-type flag
- suspected issues

Use it when the user cares about:
- column meaning
- PII detection
- semantic profiling
- bad headers / unknown schema

### `scripts/reporter.py`
Human-readable report generator.

What it does:
- combines `diagnose.py` and `column_detector.py`
- produces a plain-text report for humans
- produces a JSON report for UI/API use
- computes a health score
- groups issues by severity
- includes recommended actions and assumptions
- shows a PII warning when likely PII is present

Use it when the user asks for:
- a report
- an explanation a non-technical person can understand
- a summary for product/UI/display

### `scripts/heal.py`
Repair pipeline.

What it outputs:
- `Clean Data`
- `Quarantine`
- `Change Log`

What it fixes:
- encoding junk
- BOM / null bytes / smart quotes / broken line breaks
- delimiter/alignment problems
- short rows and overflow rows
- common date/amount/currency/name/status normalization
- exact duplicates
- some near-duplicate review flags
- metadata/header rows before the real header
- notes rows
- subtotal/total rows
- formula residue rows
- merged-cell style categorical gaps
- combined amount/currency values

What it does not do:
- invent business truth
- silently keep formula text as trusted data
- remove PII automatically

Use it when the user wants:
- a fixed workbook
- a quarantine tab for bad rows
- a change log of what was altered

## Standard workflow

Pick the lightest script that matches the user’s need:

1. `loader.py`
   Use internally when another script needs file ingestion.
2. `diagnose.py`
   Use for technical JSON diagnostics.
3. `column_detector.py`
   Use for semantic column understanding.
4. `reporter.py`
   Use when the user wants a readable report.
5. `heal.py`
   Use when the user wants the file repaired.

Common combinations:
- Diagnose + explain:
  run `diagnose.py`, then optionally `reporter.py`
- Diagnose + repair:
  run `diagnose.py`, then `heal.py`
- Semantic understanding first:
  run `column_detector.py`, then `reporter.py`

## Manual commands

```bash
python skills/csv-doctor/scripts/diagnose.py <path-to-file>
python skills/csv-doctor/scripts/column_detector.py <path-to-file>
python skills/csv-doctor/scripts/reporter.py <path-to-file> [output.txt] [output.json]
python skills/csv-doctor/scripts/heal.py <path-to-file> [output.xlsx]
```

## Claude behavior

When you use this skill:

1. Choose the right script for the user’s actual request.
2. Run the script.
3. Read the output carefully.
4. Translate technical JSON into clear plain English unless the user asked for raw output.
5. If `heal.py` was used, explicitly report:
   - clean row count
   - quarantine row count
   - change log count
   - any `needs_review` implications
6. If PII is detected, call it out clearly and do not imply it was removed automatically.
7. Do not overclaim certainty. If something is inferred, say it is inferred.

## Output expectations

`diagnose.py`:
- JSON health report

`column_detector.py`:
- JSON column semantics report

`reporter.py`:
- plain-text report
- JSON report artifact

`heal.py`:
- healed `.xlsx` workbook with 3 sheets
- printed summary to stdout
