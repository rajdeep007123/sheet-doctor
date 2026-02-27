# csv-doctor

Two scripts for messy CSV files:

- **`diagnose.py`** — analyses a CSV and produces a human-readable health report
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

## What this skill checks

### 1. Encoding
Detects the file's actual encoding using `chardet` and flags mismatches with UTF-8. Common culprits: Latin-1, Windows-1252, ISO-8859-1. Reports confidence level and which characters look corrupted.

### 2. Column alignment
Counts columns in every row and flags rows where the count differs from the header. Misaligned rows usually mean a rogue comma, an unquoted field containing a comma, or a copy-paste accident.

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

---

## How to invoke

### Diagnose only

```
python skills/csv-doctor/scripts/diagnose.py <path-to-csv>
```

Claude will run the script, read the JSON output, and turn it into a plain-English health report with:
- A one-line summary verdict (HEALTHY / NEEDS ATTENTION / CRITICAL)
- A numbered list of every issue found, with row numbers and examples
- Suggested fixes for each issue

### Diagnose + heal

```
python skills/csv-doctor/scripts/heal.py <path-to-csv> [output.xlsx]
```

Fixes all issues automatically and writes a 3-sheet Excel workbook. Prints a summary report to stdout. See **heal.py** section below for details.

---

## Input

A path to any `.csv` file. Both scripts handle mixed encodings (Latin-1, UTF-8, Windows-1252) before reading — the file does not need to be valid UTF-8.

---

## Output format

The script outputs a single JSON object to stdout. Claude parses this and formats the final report. Exit code `0` means the script ran successfully (even if issues were found). Exit code `1` means the script itself failed (file not found, completely unparseable, etc.).

```json
{
  "file": "messy_sample.csv",
  "encoding": {
    "detected": "ISO-8859-1",
    "confidence": 0.73,
    "is_utf8": false,
    "suspicious_chars": ["row 4, col 3: caf\u00e9"]
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
  "duplicate_headers": ["customer_id", "notes"],
  "whitespace_headers": ["name ", " email"],
  "empty_columns": ["col_f"],
  "single_value_columns": {
    "status": "active"
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
| Problem | Fix applied |
|---|---|
| Mixed Latin-1 / UTF-8 encoding | Per-line decode: UTF-8 first, Latin-1 fallback |
| BOM characters | Stripped from cell values |
| Null bytes | Removed from cell values |
| Smart / curly quotes | Replaced with straight quotes |
| Line breaks inside cells | Replaced with a space |
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

- `pandas` — data loading and analysis (diagnose.py)
- `chardet` — encoding detection (diagnose.py)
- `openpyxl` — Excel workbook output (heal.py)

Install: `pip install pandas chardet openpyxl`
