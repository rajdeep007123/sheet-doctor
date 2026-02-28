# csv-doctor

Four scripts for messy tabular files:

- **`loader.py`** — universal file loader used by the other three scripts
- **`diagnose.py`** — analyses a file and produces a human-readable health report
- **`column_detector.py`** — infers what each column probably contains and reports per-column quality stats
- **`heal.py`** — fixes every issue it can and writes a 3-sheet Excel workbook

---

## Trigger phrases

Use this skill when the user says things like:
- "diagnose this CSV"
- "what's wrong with this file"
- "fix my spreadsheet" / "heal this CSV"
- "check this CSV for errors"
- `/csv-doctor`

---

## loader.py

`loader.py` is the shared file-reading layer. `diagnose.py`, `column_detector.py`, and `heal.py` import it instead of handling file I/O themselves.

### Supported formats

| Extension | Notes |
|-----------|-------|
| `.csv` | Delimiter auto-detected (comma, tab, pipe, semicolon) |
| `.tsv` | Always tab — no sniffing needed |
| `.txt` | Sniffed like `.csv`; rejects plain-text files that are not tabular |
| `.xlsx` | Excel (openpyxl) |
| `.xls` | Excel legacy — requires `pip install xlrd` |
| `.xlsm` | Excel macro-enabled — macros ignored, data loaded |
| `.ods` | OpenDocument spreadsheet — requires `pip install odfpy` |
| `.json` | Array of objects or nested dict (auto-flattened) |
| `.jsonl` | JSON Lines — one object per line |

### Encoding strategy (text files)

Reads raw bytes, decodes **line by line**:
1. Try UTF-8
2. Try the chardet-detected encoding
3. Try Latin-1
4. CP1252 with `replace` (never crashes)

Embedded null bytes are stripped before parsing. This correctly handles files with mixed Latin-1 and UTF-8 on different rows.

### Multi-sheet Excel / ODS

- **One sheet** — loaded silently.
- **Multiple sheets, same columns** — prompts the user to pick one or consolidate all into a single table.
- **Multiple sheets, different columns** — prompts the user to pick one.
- **Non-interactive** (called as a subprocess by Claude Code) — requires explicit `sheet_name` or `consolidate_sheets=True`; otherwise raises a clear error listing the available sheets.

### What load_file() returns

```python
{
  "dataframe":        df,          # pandas DataFrame
  "detected_format":  "csv",       # format string
  "detected_encoding": "latin-1",  # None for binary formats
  "encoding_info": {               # None for binary formats
    "detected":        "latin-1",
    "confidence":      0.99,
    "is_utf8":         False,
    "suspicious_chars": ["row 4: byte b'\\xfc' at position 3"]
  },
  "delimiter":   ",",              # None for non-text formats
  "raw_text":    "...",            # decoded text; None for non-text formats
  "sheet_name":  "Sheet1",        # active sheet; None for non-spreadsheets
  "sheet_names": ["Sheet1"],      # all sheet names; None for non-spreadsheets
  "original_rows":    847,         # row count including header
  "original_columns": 12,
  "warnings":    []                # list of advisory strings
}
```

### JSON handling

- **Array of objects** → converted directly to a DataFrame.
- **Dict with a list value** → uses the first list key as the records array; warns which key was chosen.
- **Single dict** → treated as a one-row table with a warning.
- All nested fields are flattened with `pd.json_normalize()`.

---

## What diagnose.py checks

### 1. Encoding
Detects the file's actual encoding using `chardet` and flags mismatches with UTF-8. Common culprits: Latin-1, Windows-1252, ISO-8859-1. Reports confidence level and which characters look corrupted.

### 2. Column alignment
Counts columns in every row and flags rows where the count differs from the header. Misaligned rows usually mean a rogue comma, an unquoted field containing a comma, or a copy-paste accident.

### 2.5 Delimiter detection
Auto-detects delimiter (`comma`, `semicolon`, `tab`, `pipe`) before analysis. This prevents false "misaligned row" results when a file is not comma-separated.

### 3. Date format consistency
Scans all columns that look like they contain dates and flags any column where more than one date format is in use. Reports every format found and example values for each.

### 4. Empty rows
Finds fully blank rows (all fields empty or whitespace). Reports count and row numbers.

### 5. Duplicate headers
Checks for column names that appear more than once. Duplicate headers silently break most tools that consume the file.

### 6. Bonus checks
- Leading/trailing whitespace in headers
- Completely empty columns
- Columns that are entirely one repeated value (likely a fill-down accident)

### 7. Column semantics
`diagnose.py` now includes a `column_semantics` section powered by `column_detector.py`. For each column it tries to infer the likely semantic type even when the header is weak, generic, or wrong.

Detected types:
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

Per-column quality stats:
- `null_count` / `null_percentage`
- `unique_count` / `unique_percentage`
- `most_common_values`
- `min_value` / `max_value` where applicable
- `sample_values`
- `has_mixed_types`
- `suspected_issues`

Suspected issues currently detected:
- `Mixed date formats detected`
- `Inconsistent capitalisation`
- `Trailing/leading whitespace in X% of values`
- `Possible duplicates with slight differences`
- `Values suspiciously all the same`
- `Outliers detected (values outside 3 standard deviations)`
- `Possible PII detected (emails/phones/names)`

---

## How to invoke

### Diagnose only

```
python skills/csv-doctor/scripts/diagnose.py <path-to-file>
```

Accepts `.csv`, `.tsv`, `.txt`. Claude will run the script, read the JSON output, and turn it into a plain-English health report with:
- A one-line summary verdict (HEALTHY / NEEDS ATTENTION / CRITICAL)
- A numbered list of every issue found, with row numbers and examples
- Suggested fixes for each issue
- A semantic view of each column so the user can see what the file appears to contain, not just how it is broken

### Column semantics only

```
python skills/csv-doctor/scripts/column_detector.py <path-to-file>
```

Accepts any format supported by `loader.py`. Outputs structured JSON focused only on column semantics and quality.

### Diagnose + heal

```
python skills/csv-doctor/scripts/heal.py <path-to-file> [output.xlsx]
```

Accepts any format supported by `loader.py`. Fixes all issues automatically and writes a 3-sheet Excel workbook. Prints a summary report to stdout.

---

## Output format (diagnose.py)

The script outputs a single JSON object to stdout. Claude parses this and formats the final report. Exit code `0` means the script ran successfully (even if issues were found). Exit code `1` means the script itself failed (file not found, completely unparseable, etc.).

```json
{
  "file": "messy_sample.csv",
  "encoding": {
    "detected": "ISO-8859-1",
    "confidence": 0.73,
    "is_utf8": false,
    "suspicious_chars": ["row 4, col 3: café"]
  },
  "column_count": {
    "expected": 5,
    "misaligned_rows": [
      {"row": 7, "count": 4},
      {"row": 12, "count": 6}
    ]
  },
  "date_formats": {
    "order_date": {
      "formats_found": ["DD/MM/YYYY", "MM-DD-YY", "YYYY-MM-DD"],
      "examples": {
        "DD/MM/YYYY": "03/11/2023",
        "MM-DD-YY": "11-03-23",
        "YYYY-MM-DD": "2023-11-03"
      }
    }
  },
  "empty_rows": {
    "count": 3,
    "rows": [6, 14, 21]
  },
  "duplicate_headers": {
    "duplicate_columns": ["customer_id", "notes"],
    "repeated_header_rows": [18]
  },
  "whitespace_headers": ["name ", " email"],
  "column_quality": {
    "empty_columns": ["col_f"],
    "single_value_columns": {
      "status": "active"
    }
  },
  "column_semantics": {
    "columns": {
      "email": {
        "detected_type": "email address",
        "null_count": 1,
        "null_percentage": 8.33,
        "unique_count": 11,
        "unique_percentage": 91.67,
        "most_common_values": [
          {"value": "ops@example.com", "count": 1}
        ],
        "min_value": null,
        "max_value": null,
        "sample_values": ["ops@example.com", "sales@example.com", "hr@example.com"],
        "has_mixed_types": false,
        "suspected_issues": [
          "Possible PII detected (emails/phones/names)"
        ]
      }
    },
    "summary": {
      "total_rows": 12,
      "total_columns": 5,
      "detected_types": {
        "email address": 1,
        "categorical": 2,
        "date": 1,
        "currency/amount": 1
      },
      "issue_counts": {
        "Possible PII detected (emails/phones/names)": 1
      }
    }
  },
  "summary": {
    "verdict": "CRITICAL",
    "issue_count": 7
  }
}
```

---

## Claude's job after running diagnose.py

1. Read the JSON output
2. Write a health report in plain English — no JSON shown to the user
3. Lead with the verdict and total issue count
4. For each issue: describe it, show the affected rows/columns, and give a concrete fix
5. End with a "Next steps" section ordered by severity

---

## heal.py

### What it fixes automatically

`heal.py` runs in two modes:

- **Schema-specific mode** (when headers match the finance sample schema): deep normalisation of dates, amounts, currency, status, near-duplicate detection, etc.
- **Generic mode** (any other file): structural cleaning without assuming what the columns mean.

| Problem | Fix applied |
|---|---|
| Mixed Latin-1 / UTF-8 encoding | Per-line decode: UTF-8 first, Latin-1 fallback |
| BOM characters | Stripped from cell values |
| Null bytes | Removed from cell values |
| Smart / curly quotes | Replaced with straight quotes |
| Line breaks inside cells | Replaced with a space |
| Wrong delimiter assumption | Auto-detected delimiter used for parsing |
| Rows shifted right (ghost leading column) | Leading empty column stripped |
| Phantom comma (ghost field before Notes) | Ghost field removed |
| Unquoted commas in Notes | Overflow columns merged back into Notes |
| Short rows (fewer columns than header) | Padded with empty strings |
| Dates in any of 7 formats | Normalised to YYYY-MM-DD |
| Amounts in any of 8 formats | Normalised to float with 2 decimal places |
| Currency in any of 7 formats | Normalised to 3-letter ISO code |
| Names in any case / "Last, First" order | Title Case, First Last order |
| Inconsistent status values | Normalised to Approved / Rejected / Pending |
| Exact duplicate rows | First occurrence kept; rest removed |
| Repeated header rows / structural totals (generic mode) | Quarantined |

### Output workbook — 3 sheets

**Sheet 1 — Clean Data**
Fixed rows ready to load into a database or BI tool. Extra columns:
- `was_modified` (TRUE/FALSE) — row was changed during healing
- `needs_review` (TRUE/FALSE) — row has a blank amount, unparseable date, was padded, or is a near-duplicate

**Sheet 2 — Quarantine**
Rows that could not be used. Extra column:
- `quarantine_reason` — one of:
  - `Completely empty row`
  - `Row is all whitespace`
  - `Structural row (TOTAL/subtotal/header repeat)`
  - `Less than 50% columns filled`
  - `Impossible date cannot be parsed`
  - `No numeric value found in Amount column`

**Sheet 3 — Change Log**
One row per individual change. Columns: `original_row_number`, `column_affected`, `original_value`, `new_value`, `action_taken` (Fixed / Quarantined / Removed / Flagged), `reason`.

### Claude's job after running heal.py
1. Read the printed summary
2. Report the three tab counts (Clean / Quarantine / Changes logged)
3. List the `needs_review` rows and explain why each was flagged
4. List the Quarantine rows and their reasons
5. Read out the assumptions the script made (printed at the end of the summary)

---

## Dependencies

Core (always required):
- `pandas` — data loading and analysis
- `chardet` — encoding detection
- `openpyxl` — Excel reading and writing

Optional (install only if you need the format):
- `xlrd` — `.xls` legacy Excel files: `pip install xlrd`
- `odfpy` — `.ods` OpenDocument files: `pip install odfpy`

Install all at once: `pip install pandas chardet openpyxl`
