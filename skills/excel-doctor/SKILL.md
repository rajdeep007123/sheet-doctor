# excel-doctor

A diagnostic skill for Excel workbooks (.xlsx, .xlsm). Analyses sheet structure, merged cells, formula errors, data type consistency, and other Excel-specific problems. Produces a human-readable health report.

---

## Trigger phrases

Use this skill when the user says things like:
- "diagnose this Excel file"
- "what's wrong with my spreadsheet"
- "check this Excel for errors"
- "analyse my .xlsx"
- `/excel-doctor`

---

## What this skill checks (9 checks)

### 1. Sheet inventory
Lists all sheets in the workbook and flags any that are completely empty or hidden. Hidden sheets are a common source of confusion — formulas may reference data the user can't see.

### 2. Merged cells
Merged cells break pandas, Power Query, and most data pipelines. Reports every merged range per sheet (e.g. `G2:G3`).

### 3. Formula errors
Scans every cell for Excel error values: `#REF!`, `#VALUE!`, `#DIV/0!`, `#NAME?`, `#NULL!`, `#N/A`, `#NUM!`. Catches both cells with an actual error type (`cell.data_type == 'e'`) and string values that look like error text. Reports cell coordinates and error type per sheet.

### 4. Mixed data types in columns
A column containing both numbers and strings (e.g. `125.50` mixed with `"N/A"`) causes silent failures in most tools. Reports affected columns with one example value per type found.

### 5. Empty rows
Finds fully blank rows per sheet. Reports count and row numbers.

### 6. Empty columns
Finds columns where the header exists but every data cell is empty. Reports column names per sheet.

### 7. Duplicate headers
Checks for column names that appear more than once in the header row. Duplicate headers silently break most tools that consume the file.

### 8. Date format consistency
Scans string-typed cells in columns that look date-like and flags any column where more than one format is in use (e.g. `YYYY-MM-DD` mixed with `DD/MM/YYYY`). Note: cells Excel has already typed as `datetime` objects are correctly typed and not flagged.

### 9. Single-value columns
A column where every non-empty row has the same value is likely a fill-down accident. Reports the column name and the repeated value per sheet.

---

## How to invoke

```
python skills/excel-doctor/scripts/diagnose.py <path-to-xlsx>
```

Claude will run the script, read the JSON output, and turn it into a plain-English health report with:
- A one-line summary verdict (HEALTHY / NEEDS ATTENTION / CRITICAL)
- A numbered list of every issue found, with sheet names, cell coordinates, and examples
- Suggested fixes for each issue

---

## Input

A path to any `.xlsx` or `.xlsm` file.

> **Important — `data_only=True`:** The script loads the workbook with `data_only=True`, which reads cached formula results rather than formula strings. This means:
> - Formula errors are reported as their cached values (e.g. `#DIV/0!`) ✓
> - If the workbook has never been opened and calculated in Excel, formula cells may read as `None` instead of their expected values
> - Formula strings themselves (e.g. `=SUM(A1:A10)`) are not visible to the script

---

## Output format

The script outputs a single JSON object to stdout. Claude parses this and formats the final report. Exit code `0` means the script ran successfully (even if issues were found). Exit code `1` means the script itself failed (file not found, unreadable, wrong format, etc.).

```json
{
  "file": "messy_sample.xlsx",
  "sheets": {
    "all": ["Orders", "Summary", "Archive"],
    "count": 3,
    "empty": ["Summary"],
    "hidden": [{"name": "Archive", "state": "hidden"}]
  },
  "merged_cells": {
    "Orders": ["G2:G3"]
  },
  "formula_errors": {
    "Orders": [
      {"cell": "E6", "value": "#DIV/0!"},
      {"cell": "E9", "value": "#REF!"},
      {"cell": "E11", "value": "#VALUE!"}
    ]
  },
  "mixed_types": {
    "Orders": {
      "amount": {
        "types": ["number", "text"],
        "examples": {"number": "250.0", "text": "N/A"}
      }
    }
  },
  "empty_rows": {
    "Orders": {"count": 2, "rows": [5, 10]}
  },
  "empty_columns": {},
  "duplicate_headers": {
    "Orders": ["customer_id"]
  },
  "date_formats": {
    "Orders": {
      "order_date": {
        "formats_found": ["YYYY-MM-DD", "DD/MM/YYYY or MM/DD/YYYY"],
        "examples": {
          "YYYY-MM-DD": "2023-01-15",
          "DD/MM/YYYY or MM/DD/YYYY": "15/02/2023"
        }
      }
    }
  },
  "single_value_columns": {
    "Orders": {"status": "active"}
  },
  "summary": {
    "verdict": "CRITICAL",
    "issue_count": 9
  }
}
```

---

## Claude's job after running the script

1. Read the JSON output
2. Write a health report in plain English — no JSON shown to the user
3. Lead with the verdict and total issue count
4. For each issue: name the sheet, describe the problem, show affected cells/columns, and give a concrete fix
5. End with a "Next steps" section ordered by severity

---

## Sample file

`sample-data/messy_sample.xlsx` is a deliberately broken workbook for testing. It contains:

| Sheet | Problems baked in |
|---|---|
| Orders | Duplicate header (`customer_id` ×2), merged cells (`G2:G3`), 3 formula errors (`#DIV/0!`, `#REF!`, `#VALUE!`), mixed types in `amount` (float + `"N/A"`), 2 empty rows, 6 date formats, `status` column always `"active"` |
| Summary | Completely empty sheet |
| Archive | Hidden sheet |

---

## Dependencies

- `openpyxl` — reads `.xlsx` files and exposes sheet structure, merged cells, cell types, and formula error values

Install: `pip install openpyxl`
