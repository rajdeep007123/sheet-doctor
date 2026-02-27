# csv-doctor

A diagnostic skill for messy CSV files. Produces a human-readable health report covering encoding, structure, dates, and data quality issues.

---

## Trigger phrases

Use this skill when the user says things like:
- "diagnose this CSV"
- "what's wrong with this file"
- "fix my spreadsheet"
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

```
python skills/csv-doctor/scripts/diagnose.py <path-to-csv>
```

Claude will run the script, read the JSON output, and turn it into a plain-English health report with:
- A one-line summary verdict (HEALTHY / NEEDS ATTENTION / CRITICAL)
- A numbered list of every issue found, with row numbers and examples
- Suggested fixes for each issue

---

## Input

A path to any `.csv` file. The file does not need to be valid UTF-8 — the script handles encoding detection before reading.

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

## Claude's job after running the script

1. Read the JSON output
2. Write a health report in plain English — no JSON shown to the user
3. Lead with the verdict and total issue count
4. For each issue: describe it, show the affected rows/columns, and give a concrete fix
5. End with a "Next steps" section ordered by severity

---

## Dependencies

- `pandas` — data loading and analysis
- `chardet` — encoding detection

Install: `pip install pandas chardet`
