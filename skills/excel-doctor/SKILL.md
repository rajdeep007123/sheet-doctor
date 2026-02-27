# excel-doctor

Excel quality skill for `.xlsx` / `.xlsm` files with two scripts:
- `diagnose.py` for deep diagnostics (JSON health report)
- `heal.py` for safe auto-fixes + a workbook `Change Log` tab

---

## Trigger phrases

Use this skill when the user says things like:
- "diagnose this Excel file"
- "what's wrong with my spreadsheet"
- "check this Excel for errors"
- "analyse my .xlsx"
- "fix this Excel file" / "heal this workbook"
- `/excel-doctor`

---

## What this skill checks (13 checks)

### 1. Sheet inventory
Lists all sheets in the workbook and flags any that are completely empty or hidden. Hidden sheets are a common source of confusion — formulas may reference data the user can't see.

### 2. Merged cells
Merged cells break pandas, Power Query, and most data pipelines. Reports every merged range per sheet (e.g. `G2:G3`).

### 3. Formula errors
Scans every cell for Excel error values: `#REF!`, `#VALUE!`, `#DIV/0!`, `#NAME?`, `#NULL!`, `#N/A`, `#NUM!`. Catches both cells with an actual error type (`cell.data_type == 'e'`) and string values that look like error text. Reports cell coordinates and error type per sheet.

### 4. Formula cache misses
Loads the workbook both with and without `data_only`. Flags formula cells that have no cached value (common when file wasn't recalculated before export). Reports cell and formula.

### 5. Mixed data types in columns
A column containing both numbers and strings (e.g. `125.50` mixed with `"N/A"`) causes silent failures in most tools. Reports affected columns with one example value per type found.

### 6. Empty rows
Finds fully blank rows per sheet. Reports count and row numbers.

### 7. Empty columns
Finds columns where the header exists but every data cell is empty. Reports column names per sheet.

### 8. Duplicate headers
Checks for column names that appear more than once in the header row. Duplicate headers silently break most tools that consume the file.

### 9. Header whitespace
Flags headers with leading/trailing spaces (e.g. `" amount"` or `"status "`), which create hard-to-see schema bugs.

### 10. Date format consistency
Scans string-typed cells in columns that look date-like and flags any column where more than one format is in use (e.g. `YYYY-MM-DD` mixed with `DD/MM/YYYY`). Note: cells Excel has already typed as `datetime` objects are correctly typed and not flagged.

### 11. Single-value columns
A column where every non-empty row has the same value is likely a fill-down accident. Reports the column name and the repeated value per sheet.

### 12. Structural subtotal/total rows
Flags likely subtotal rows (`TOTAL`, `SUBTOTAL`, `GRAND TOTAL`) embedded inside detail tables.

### 13. High-null columns
Flags columns that are mostly empty (>= 80% blanks) but not entirely empty.

---

## How to invoke

Diagnose:
```
python skills/excel-doctor/scripts/diagnose.py <path-to-xlsx>
```

Heal:
```
python skills/excel-doctor/scripts/heal.py <path-to-xlsx> [output.xlsx]
```

For `diagnose.py`, Claude should read the JSON output and turn it into a plain-English health report with:
- A one-line summary verdict (HEALTHY / NEEDS ATTENTION / CRITICAL)
- A numbered list of every issue found, with sheet names, cell coordinates, and examples
- Suggested fixes for each issue

For `heal.py`, Claude should:
- Report output filename and total changes logged
- Call out major automatic fixes applied (merged ranges, headers, dates, empty rows)
- List assumptions and anything that still needs manual review

---

## Input

A path to any `.xlsx` or `.xlsm` file.

`diagnose.py` loads the workbook twice:
- `data_only=True` for cached values and error inspection
- `data_only=False` for formula-string inspection (to detect cache misses)

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
  "formula_cache_misses": {
    "Orders": [
      {
        "cell": "H12",
        "formula": "=SUM(H2:H11)",
        "reason": "No cached value. Open/recalculate in Excel and save."
      }
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
  "whitespace_headers": {},
  "structural_rows": {},
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
  "high_null_columns": {},
  "summary": {
    "verdict": "CRITICAL",
    "issue_count": 15,
    "issue_categories_triggered": 9
  },
  "issue_counts": {
    "formula_errors": 3,
    "duplicate_headers": 1
  }
}
```

---

## heal.py auto-fixes

`heal.py` applies deterministic fixes:
- Unmerge merged ranges and fill child cells from the anchor value
- Standardise headers (trim/collapse spaces), create fallback names for blanks, dedupe duplicates with suffixes (`_2`, `_3`, ...)
- Clean text values (BOM/NULL/line-break/smart-quote cleanup)
- Normalise common date strings to `YYYY-MM-DD`
- Remove fully empty rows
- Append a `Change Log` sheet with one row per edit

---

## Sample file

`sample-data/messy_sample.xlsx` is a deliberately broken workbook for testing. It contains:

| Sheet | Problems baked in |
|---|---|
| Orders | Duplicate header (`customer_id` ×2), merged cells (`G2:G3`), 3 formula errors (`#DIV/0!`, `#REF!`, `#VALUE!`), mixed types in `amount` (float + `"N/A"`), 2 empty rows, 7 date formats, `status` column always `"active"` |
| Summary | Completely empty sheet |
| Archive | Hidden sheet |

---

## Dependencies

- `openpyxl` — reads `.xlsx` files and exposes sheet structure, merged cells, cell types, and formula error values

Install: `pip install openpyxl`
