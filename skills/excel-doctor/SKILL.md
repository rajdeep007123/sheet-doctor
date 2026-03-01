# excel-doctor

Workbook-native Excel quality skill for `.xlsx` and `.xlsm` files.

Use this skill when the user wants to inspect or clean a real Excel workbook while preserving workbook structure as much as safely possible.

Scripts:
- `diagnose.py` — workbook-native diagnostics to JSON
- `heal.py` — workbook-native cleanup plus a `Change Log` sheet

Do not use this skill for:
- `.xls` workbook-native repair
- `.ods`
- generic flat-file rescue where a 3-sheet tabular output is preferred

For those cases, use `csv-doctor` tabular rescue instead.

---

## Trigger phrases

Use this skill when the user says things like:
- "diagnose this Excel workbook"
- "what is wrong with this .xlsx"
- "fix this .xlsm while keeping the workbook"
- "check hidden sheets / merged cells / formulas"
- "analyse this spreadsheet without flattening it"
- `/excel-doctor`

---

## Real scope

Supported workbook-native inputs:
- `.xlsx`
- `.xlsm`

Not supported workbook-native:
- `.xls` — reject with an explicit message telling the user to use `csv-doctor` tabular rescue or convert to `.xlsx`
- password-protected / encrypted OOXML workbooks — reject cleanly
- corrupt workbooks — fail cleanly, do not promise reconstruction

---

## What `diagnose.py` checks

`diagnose.py` is workbook-native. It does not flatten the workbook first.

It reports:
- all sheet names
- hidden sheets
- very-hidden sheets
- empty sheets
- merged ranges
- formula cells
- formula errors
- formula cache misses
- duplicate headers
- whitespace headers
- header bands / stacked headers
- metadata / preamble rows before the real header
- empty rows
- empty columns
- empty edge columns
- mixed-type columns
- mixed date formats in text-like date columns
- likely subtotal / total rows
- likely notes / metadata rows
- per-sheet risk summaries
- workbook-level summary
- workbook triage classification (`workbook_native_safe_cleanup`, `tabular_rescue_recommended`, `manual_spreadsheet_review_required`)
- plain-English triage reason
- triage confidence
- recommended next action
- residual-risk sections describing what is safe to auto-fix, what remains risky, and what still needs manual spreadsheet review
- manual-review warnings when formulas, hidden sheets, or heuristic header detection mean cleanup is not the same as business-safe interpretation

If something cannot be known safely, it should be called out as a limitation rather than guessed.

---

## What `heal.py` does

`heal.py` performs safe workbook-native cleanup:
- unmerges ranges and fills child cells from the anchor value
- flattens stacked header bands when they look like real table headers
- removes metadata / preamble rows before the real table header
- standardises / deduplicates headers
- trims fully empty edge columns
- removes fully empty rows
- cleans obvious text artifacts:
  - BOM
  - null bytes
  - embedded line breaks
  - smart quotes
  - repeated whitespace
- normalises parseable date strings to `YYYY-MM-DD`
- appends a `Change Log` sheet
- writes atomically via temp file + replace

Important limits:
- formulas are preserved, not recalculated
- missing formula cache values are not reconstructed
- `.xlsm` macros are only preserved if the output stays `.xlsm`
- this is not a spreadsheet reconstruction engine

`heal.py` summaries now also include:
- workbook triage after healing
- residual-risk reporting after healing
- before/after issue counts for key workbook problems such as merged ranges, duplicate headers, empty rows, formula errors, and formula cache misses

---

## When to use `excel-doctor` vs `csv-doctor`

Use `excel-doctor` when:
- workbook sheets matter
- hidden sheets matter
- merged ranges / header bands / workbook layout matter
- the user wants to preserve workbook structure
- diagnosis reports `workbook_native_safe_cleanup` as the recommended path

Use `csv-doctor` tabular rescue when:
- the user wants `Clean Data / Quarantine / Change Log`
- the workbook is really just a messy table
- flattening the workbook into rows/columns is acceptable
- the input is `.xls` or `.ods`
- diagnosis reports `tabular_rescue_recommended`

Require manual spreadsheet review first when:
- diagnosis reports `manual_spreadsheet_review_required`
- formula errors or cache misses matter to workbook meaning
- hidden/very-hidden sheets may contain required business context

---

## How to invoke

Diagnose:
```bash
python skills/excel-doctor/scripts/diagnose.py <path-to-xlsx-or-xlsm>
```

Heal:
```bash
python skills/excel-doctor/scripts/heal.py <path-to-xlsx-or-xlsm> [output.xlsx|output.xlsm]
```

Structured healing summary:
```bash
python skills/excel-doctor/scripts/heal.py workbook.xlsx healed.xlsx --json-summary /tmp/excel_heal_summary.json
```

---

## Expected output shape

`diagnose.py` returns JSON with:
- contract metadata
- workbook file info
- sheet inventory
- workbook-native issue sections
- `workbook_triage`
- `residual_risk`
- `sheet_summaries`
- `workbook_summary`
- `summary`
- `issue_counts`
- `run_summary`

`heal.py` returns:
- healed workbook written to disk
- appended `Change Log` sheet
- optional JSON summary with:
  - contract metadata
  - mode = `workbook-native`
  - stats
  - changes logged
  - assumptions
  - `workbook_triage`
  - `residual_risk`
  - `before_after_issue_summary`
  - run summary

---

## Sample commands

```bash
python skills/excel-doctor/scripts/diagnose.py sample-data/messy_sample.xlsx
python skills/excel-doctor/scripts/heal.py sample-data/messy_sample.xlsx
python skills/excel-doctor/scripts/diagnose.py /tmp/messy_sample_healed.xlsx
```

---

## Sample workbook expectations

`sample-data/messy_sample.xlsx` is useful for validating:
- hidden sheet detection
- merged range detection
- formula error detection
- duplicate header detection
- mixed-type column detection
- empty row removal
- workbook-native healing with `Change Log`

---

## Dependencies

- `openpyxl`

Install:
```bash
pip install openpyxl
```
