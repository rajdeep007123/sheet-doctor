# Public Data Evaluation Sheet

Use this file to track spreadsheet-cleaning runs by test class:
- `smoke`
- `stress`
- `parser regression`
- `real mess`

Important constraint:
- The Kaggle datasets listed below require a Kaggle-authenticated download flow.
- They cannot be fetched directly from this environment without your Kaggle session or API credentials.
- The evaluation workflow below is ready now; drop the downloaded files into the repo and reuse the same commands.

## Recommended placement

Put downloaded public datasets here:

```bash
mkdir -p sample-data/public
```

Examples:

```bash
sample-data/public/messy_imdb.csv
sample-data/public/messy_employee.xlsx
sample-data/public/fifa21_raw.csv
```

## Evaluation workflow

Activate the venv first:

```bash
cd "/Users/razzo/Documents/For Codex/sheet-doctor"
source .venv/bin/activate
```

### CSV / tabular rescue flow

```bash
# Step 1 — Diagnose
python skills/csv-doctor/scripts/diagnose.py <file>

# Step 2 — Report (gets health scores)
python skills/csv-doctor/scripts/reporter.py <file> /tmp/report.txt /tmp/report.json

# Step 3 — Heal
python skills/csv-doctor/scripts/heal.py <file> /tmp/healed.xlsx --json-summary /tmp/heal_summary.json

# Step 4 — Inspect outputs
ls -lh /tmp/report.txt /tmp/report.json /tmp/healed.xlsx /tmp/heal_summary.json
```

Metrics to capture:
- raw health score
- clean rows
- quarantine rows
- changes logged
- crash / no crash

### Excel workbook-native flow

```bash
# Step 1 — Diagnose
python skills/excel-doctor/scripts/diagnose.py <file> > /tmp/excel_diag.json

# Step 2 — Heal
python skills/excel-doctor/scripts/heal.py <file> /tmp/healed.xlsx --json-summary /tmp/excel_heal_summary.json

# Step 3 — Re-diagnose healed output
python skills/excel-doctor/scripts/diagnose.py /tmp/healed.xlsx > /tmp/excel_rediag.json

# Step 4 — Inspect outputs
ls -lh /tmp/excel_diag.json /tmp/excel_heal_summary.json /tmp/healed.xlsx /tmp/excel_rediag.json
```

Metrics to capture:
- verdict before heal
- issue count before heal
- changes logged
- verdict after heal
- issue count after heal
- crash / no crash

## Test classes

### Smoke

Goal:
- verify that the standard path works on small/basic files

Good sources:
- `datablist.com` small CSVs
- `learningcontainer.com` sample CSV
- small local sample files

Expected outcome:
- diagnose/report/heal pass cleanly
- no crashes
- low or zero quarantine on clean files

### Stress

Goal:
- verify runtime, degraded mode, and clean rejection behavior

Good sources:
- `customers-500000.csv`
- `customers-2000000.csv`

Expected outcome:
- `500K` rows: degraded mode or slow pass
- `2M` rows: clean rejection if limits are exceeded

### Parser regression

Goal:
- catch delimiter/header/shape regressions across lots of small CSV shapes

Good sources:
- [FSU CSV collection](https://people.sc.fsu.edu/~jburkardt/data/csv/csv.html)

Expected outcome:
- lots of small files should parse consistently
- these files are usually not messy enough to prove healing quality

### Real mess

Goal:
- test whether the tool is actually useful on ugly business-like data

Good sources:
- Kaggle dirty datasets
- ugly workbook fixtures in this repo
- real exports you are allowed to test

## Current baseline results

These are baseline runs from the current repo corpus so there is a stable comparison point before external/public datasets are added.

| Class | File | Tool | Format | Rows | Score | Before | After | Clean | Quarantine | Changes | Crashed? | Notes |
|---|---|---|---:|---:|---|---|---:|---:|---:|---|---|
| `real mess` | `sample-data/extreme_mess.csv` | `csv-doctor` | CSV | 51 | 32 | `Poor — major surgery required` | `post-heal score 79` | 42 | 5 | 105 | No | 11 `needs_review`; tabular rescue flow |
| `real mess` | `sample-data/messy_sample.xlsx` | `excel-doctor` | XLSX | n/a | n/a | `CRITICAL (13 issues)` | `CRITICAL (9 issues)` | n/a | n/a | 10 | No | workbook-native; merged cells/dupe headers/empty rows improved; formula errors remain |
| `real mess` | `tests/fixtures/excel/ragged_clinical.xlsx` | `excel-doctor` | XLSX | n/a | n/a | `CRITICAL (9 issues)` | not re-scored here | n/a | n/a | 7 | No | workbook-native; metadata rows, stacked header, edge columns, dates cleaned |
| `real mess` | `tests/fixtures/excel/formula_cases.xlsx` | `excel-doctor` | XLSX | n/a | n/a | `CRITICAL (9 issues)` | not re-scored here | n/a | n/a | 0 | No | workbook-native preserves formulas; manual review required |

## First public pass

These are measured runs against public files already downloaded under `sample-data/public/`.

| Class | File | Source | Tool | Format | Rows | Score | Before | After | Clean | Quarantine | Changes | Crashed? | Notes |
|---|---|---|---|---:|---:|---|---|---:|---:|---:|---|---|
| `smoke` | `Employee Sample Data.csv` | SpreadsheetGuru | `csv-doctor` | CSV | 1001 | `66 / 87 / 87` | `CRITICAL (4 issues)` | n/a | 1000 | 0 | 2245 | No | semantic mode; heavy normalization, no quarantine |
| `smoke` | `Financials Sample Data.csv` | SpreadsheetGuru | `csv-doctor` | CSV | 352 | `80 / 82 / 82` | `NEEDS ATTENTION (2 issues)` | n/a | 351 | 0 | 351 | No | semantic mode; mostly straightforward cleanup |
| `smoke` | `Employee Sample Data.xlsx` | SpreadsheetGuru | `excel-doctor` | XLSX | n/a | n/a | `HEALTHY (1 issue)` | `HEALTHY (1 issue)` | n/a | n/a | 15 | No | triage `workbook_native_safe_cleanup`; only high-null column remains |
| `smoke` | `Financials Sample Data.xlsx` | SpreadsheetGuru | `excel-doctor` | XLSX | n/a | n/a | `HEALTHY (1 issue)` | `HEALTHY (1 issue)` | n/a | n/a | 0 | No | triage `workbook_native_safe_cleanup`; workbook was already mostly clean |
| `smoke` | `learningcontainer-employee.xlsx` | Learning Container | `excel-doctor` | XLSX | n/a | n/a | `CRITICAL (2 issues)` | `CRITICAL (2 issues)` | n/a | n/a | 0 | No | triage `manual_spreadsheet_review_required`; formula error + mixed types remain untouched |
| `smoke` | `Sample-sales-data-excel.xls` | Learning Container | `csv-doctor` | XLS | 9995 | n/a | `NEEDS ATTENTION (3 issues)` | n/a | 9994 | 0 | 399 | Partial | tabular fallback only; heal worked, but `sheet-doctor report ... --json` timed out after 20s and every clean row was flagged `needs_review` |

## Public targets to run

### Smoke targets

| Source | Expected local filename | Likely path | Format | Expected outcome |
|---|---|---|---|---|
| Datablist | `customers-100.csv` | `sample-data/public/customers-100.csv` | CSV | clean pass |
| Datablist | `customers-10000.csv` | `sample-data/public/customers-10000.csv` | CSV | clean pass |
| Learning Container | `sample-csv-file-for-testing.csv` | `sample-data/public/sample-csv-file-for-testing.csv` | CSV | clean pass |

### Stress targets

| Source | Expected local filename | Likely path | Format | Expected outcome |
|---|---|---|---|---|
| Datablist | `customers-500000.csv` | `sample-data/public/customers-500000.csv` | CSV | degraded mode or slow pass |
| Datablist | `customers-2000000.csv` | `sample-data/public/customers-2000000.csv` | CSV | clean rejection / guardrail |

### Parser regression targets

| Source | Example local filename | Likely path | Format | Expected outcome |
|---|---|---|---|---|
| FSU CSV collection | `addresses.csv` | `sample-data/public/addresses.csv` | CSV | parser sanity |
| FSU CSV collection | `grades.csv` | `sample-data/public/grades.csv` | CSV | parser sanity |
| FSU CSV collection | `cities.csv` | `sample-data/public/cities.csv` | CSV | parser sanity |
| FSU CSV collection | `biostats.csv` | `sample-data/public/biostats.csv` | CSV | parser sanity |

### Real mess targets

#### Kaggle targets to run once downloaded

| Dataset | Expected local filename | Likely path | Format | Status |
|---|---|---|---|---|
| Messy IMDB Dataset | `messy_imdb.csv` | `sample-data/public/messy_imdb.csv` | CSV | Pending download |
| Messy Employee Dataset | `messy_employee.*` | `sample-data/public/` | CSV/XLSX unknown | Pending download |
| Dirty Data Sample | `dirty_data_sample.*` | `sample-data/public/` | CSV/XLSX unknown | Pending download |
| FIFA 21 Raw Messy Dataset | `fifa21_messy_raw.csv` | `sample-data/public/fifa21_messy_raw.csv` | CSV | Pending download |
| Dirty Dataset Practice | `dirty_dataset_practice.*` | `sample-data/public/` | CSV/XLSX unknown | Pending download |

## Run table

Fill this after the files are local:

| Class | File | Tool | Format | Rows | Score | Before | After | Clean | Quarantine | Changes | Crashed? | Notes |
|---|---|---|---:|---:|---|---|---:|---:|---:|---|---|
| `smoke` | `customers-100.csv` | `csv-doctor` | CSV | ? | ? | ? | ? | ? | ? | ? | ? | |
| `smoke` | `customers-10000.csv` | `csv-doctor` | CSV | ? | ? | ? | ? | ? | ? | ? | ? | |
| `stress` | `customers-500000.csv` | `csv-doctor` | CSV | ? | ? | ? | ? | ? | ? | ? | ? | |
| `stress` | `customers-2000000.csv` | `csv-doctor` | CSV | ? | ? | ? | ? | ? | ? | ? | ? | |
| `parser regression` | `addresses.csv` | `csv-doctor` | CSV | ? | ? | ? | ? | ? | ? | ? | ? | |
| `parser regression` | `grades.csv` | `csv-doctor` | CSV | ? | ? | ? | ? | ? | ? | ? | ? | |
| `real mess` | `messy_imdb.csv` | `csv-doctor` | CSV | ? | ? | ? | ? | ? | ? | ? | ? | |
| `real mess` | `messy_employee.*` | `csv-doctor` or `excel-doctor` | ? | ? | ? | ? | ? | ? | ? | ? | ? | |
| `real mess` | `dirty_data_sample.*` | `csv-doctor` or `excel-doctor` | ? | ? | ? | ? | ? | ? | ? | ? | ? | |
| `real mess` | `fifa21_messy_raw.csv` | `csv-doctor` | CSV | ? | ? | ? | ? | ? | ? | ? | ? | |
| `real mess` | `dirty_dataset_practice.*` | `csv-doctor` or `excel-doctor` | ? | ? | ? | ? | ? | ? | ? | ? | ? | |
| `smoke` | `Employee Sample Data.csv` | `csv-doctor` | CSV | 1001 | `66 / 87 / 87` | `CRITICAL (4 issues)` | n/a | 1000 | 0 | 2245 | No | SpreadsheetGuru |
| `smoke` | `Financials Sample Data.csv` | `csv-doctor` | CSV | 352 | `80 / 82 / 82` | `NEEDS ATTENTION (2 issues)` | n/a | 351 | 0 | 351 | No | SpreadsheetGuru |
| `smoke` | `Employee Sample Data.xlsx` | `excel-doctor` | XLSX | n/a | n/a | `HEALTHY (1 issue)` | `HEALTHY (1 issue)` | n/a | n/a | 15 | No | SpreadsheetGuru; safe-cleanup triage |
| `smoke` | `Financials Sample Data.xlsx` | `excel-doctor` | XLSX | n/a | n/a | `HEALTHY (1 issue)` | `HEALTHY (1 issue)` | n/a | n/a | 0 | No | SpreadsheetGuru; already mostly clean |
| `smoke` | `learningcontainer-employee.xlsx` | `excel-doctor` | XLSX | n/a | n/a | `CRITICAL (2 issues)` | `CRITICAL (2 issues)` | n/a | n/a | 0 | No | Learning Container; manual review required |
| `smoke` | `Sample-sales-data-excel.xls` | `csv-doctor` | XLS | 9995 | n/a | `NEEDS ATTENTION (3 issues)` | n/a | 9994 | 0 | 399 | Partial | Learning Container; report JSON timed out after 20s |

## Notes on interpretation

- Use `csv-doctor` when the file should end up as a readable table with `Clean Data / Quarantine / Change Log`.
- Use `excel-doctor` when workbook structure matters and the file is `.xlsx` or `.xlsm`.
- `.xls` is not workbook-native in `excel-doctor`; route it through `csv-doctor` tabular rescue or convert it first.
- For workbook-native runs, `clean` and `quarantine` are not applicable because the output is a cleaned workbook, not a 3-sheet table rescue.
- Formula-heavy Excel files may show low change counts on purpose; preserved formulas are a manual-review case, not an auto-fix failure.
