# csv-doctor

`csv-doctor` is the tabular-file rescue layer inside `sheet-doctor`.

It exists to answer five practical questions:
- can we load this file safely?
- what is broken?
- what do these columns probably mean?
- how do we explain the damage to a human?
- what can we fix automatically without pretending uncertain rows are trustworthy?

## What it does

`csv-doctor` handles messy tabular exports across text, spreadsheet, and JSON-like formats.

What it is:
- a tabular rescue tool
- strong on flat files and table-like workbook sheets

What it is not:
- a workbook-native spreadsheet repair engine
- a multi-file merge tool
- a password-protected Excel opener

Current script set:
- `scripts/loader.py`
- `scripts/diagnose.py`
- `scripts/column_detector.py`
- `scripts/reporter.py`
- `scripts/heal.py`

## How the scripts work

### `loader.py`
Universal file ingestion.

Responsibilities:
- detect encoding
- detect delimiter
- load spreadsheet and JSON variants into pandas
- return a standard metadata dict

Accepted formats:
- `.csv`
- `.tsv`
- `.txt`
- `.xlsx`
- `.xls`
- `.xlsm`
- `.ods`
- `.json`
- `.jsonl`

Actual script support matrix:

| Script | `.csv/.tsv/.txt` | `.xlsx/.xls/.xlsm/.ods` | `.json/.jsonl` |
|---|---|---|---|
| `loader.py` | ✅ | ✅ | ✅ |
| `diagnose.py` | ✅ | ✅ (`--sheet` / `--all-sheets` for multi-sheet workbooks) | ✅ |
| `reporter.py` | ✅ | ✅ (`--sheet` / `--all-sheets`) | ✅ |
| `heal.py` | ✅ | ✅ tabular rescue | ✅ |

### `diagnose.py`
Structural health checker.

Reports:
- encoding problems
- row/column misalignment
- date format inconsistency
- empty rows
- duplicate/repeated headers
- whitespace headers
- empty columns
- single-value columns
- `column_semantics`

Workbook note:
- multi-sheet workbook inputs need `--sheet <name>` or `--all-sheets` in non-interactive runs

### `column_detector.py`
Semantic profiler.

Infers likely column type:
- date
- currency/amount
- plain number
- percentage
- email address
- phone number
- URL
- country name or code
- currency code
- name
- categorical
- free text
- boolean
- ID/code
- unknown

Also reports:
- null and unique stats
- sample values
- min/max where meaningful
- suspected issues
- PII signals

### `reporter.py`
Readable report builder.

Consumes the `diagnose.py` and `column_detector.py` views and produces:
- a plain-text report
- a JSON report artifact

Sections:
- file overview
- health score
- issues found
- column breakdown
- recommended actions
- assumptions

Workbook note:
- uses the same sheet-selection rules as `diagnose.py`
- for workbook-native diagnosis/preservation, use `excel-doctor`

### `heal.py`
Repair pipeline.

Produces a workbook with 3 sheets:
- `Clean Data`
- `Quarantine`
- `Change Log`

Repairs and edge cases covered:
- encoding cleanup
- null bytes / BOM / smart quotes / line breaks
- misaligned rows
- short rows and overflow rows
- date / amount / currency / name / status normalization
- exact duplicates
- repeated headers
- metadata rows before the real header
- notes rows
- subtotal/total rows
- formula residue rows
- merged-cell style categorical gaps
- combined amount/currency values

## Output structure

### `Clean Data`
Rows that survived repair and are usable.

Extra flags:
- `was_modified`
- `needs_review`

### `Quarantine`
Rows that should not be trusted automatically.

Extra field:
- `quarantine_reason`

Examples:
- `Excel formula found, not data`
- `Calculated subtotal row`
- `Appears to be a notes row`

### `Change Log`
One row per significant edit or row-level action.

Columns:
- `original_row_number`
- `column_affected`
- `original_value`
- `new_value`
- `action_taken`
- `reason`

## Manual usage

From the repo root:

```bash
source .venv/bin/activate
```

Diagnose:

```bash
python skills/csv-doctor/scripts/diagnose.py sample-data/extreme_mess.csv
python skills/csv-doctor/scripts/diagnose.py sample-data/messy_sample.xlsx --sheet "Orders"
python skills/csv-doctor/scripts/diagnose.py /path/to/export.json
python skills/csv-doctor/scripts/diagnose.py /path/to/export.jsonl
```

Column analysis only:

```bash
python skills/csv-doctor/scripts/column_detector.py sample-data/extreme_mess.csv
```

Human-readable report:

```bash
python skills/csv-doctor/scripts/reporter.py sample-data/extreme_mess.csv
python skills/csv-doctor/scripts/reporter.py sample-data/messy_sample.xlsx /tmp/orders_report.txt /tmp/orders_report.json --sheet "Orders"
```

Repair:

```bash
python skills/csv-doctor/scripts/heal.py sample-data/extreme_mess.csv
```

## Example output from `extreme_mess.csv`

Expected behavior on the sample disaster file:
- `diagnose.py` returns a `CRITICAL`-level JSON report with structural and semantic findings
- `reporter.py` assigns a low health score and warns about PII
- `heal.py` writes `sample-data/extreme_mess_healed.xlsx`

Typical findings from this file:
- broken row alignment
- mixed date formats
- encoding corruption
- a repeated header row
- a subtotal row
- a long metadata/notes row
- PII-like name data

Typical healed workbook result:
- usable cleaned rows in `Clean Data`
- bad structural rows and subtotal/note rows in `Quarantine`
- every meaningful normalization or quarantine event in `Change Log`

## Developer notes

If you build on top of `csv-doctor`:
- treat `loader.py` as the shared entry point
- prefer reusing `diagnose.py`, `column_detector.py`, and `reporter.py` outputs rather than rebuilding similar logic elsewhere
- do not silently turn uncertain rows into trusted facts; quarantine them and log why
- document whether your feature is table rescue or workbook-native preservation
