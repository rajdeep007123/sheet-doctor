# Contributing to sheet-doctor

Thanks for being here. This project is built by one person returning to code after 8 years, so every contribution — big or small — genuinely matters.

---

## What we need most

- **New skills** — `merge-doctor`, `type-doctor`, `encoding-fixer` (see ideas below)
- **New format support in loader.py** — edge cases, better heuristics for delimiter sniffing or encoding detection
- **Better heuristics in existing scripts** — date parsing, amount normalisation, near-duplicate detection
- **More messy sample files** — real-world broken CSVs, .xlsx, .json files (anonymized, please)
- **Documentation fixes** — if something confused you, it'll confuse others

---

## How to contribute

**1. Fork and clone**

```bash
git clone https://github.com/razzo007/sheet-doctor.git
cd sheet-doctor
```

**2. Create a branch**

```bash
git checkout -b my-new-skill
```

**3. Make your changes**

For a new skill, copy the `csv-doctor` structure:

```
skills/
└── your-skill-name/
    ├── SKILL.md
    └── scripts/
        ├── diagnose.py   ← analysis, outputs JSON to stdout
        └── heal.py       ← fixes issues, outputs .xlsx workbook (optional but encouraged)
```

If your skill reads tabular files, **import `loader.py` from `csv-doctor`** instead of writing your own file-reading logic:

```python
import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).parent.parent.parent / "csv-doctor" / "scripts"))
from loader import load_file

result = load_file("path/to/file.csv")
df     = result["dataframe"]
```

This gives you encoding detection, delimiter sniffing, and multi-format support for free.

**4. Test against the sample data**

```bash
# csv-doctor
python skills/csv-doctor/scripts/diagnose.py sample-data/messy_sample.csv
python skills/csv-doctor/scripts/diagnose.py sample-data/extreme_mess.csv
python skills/csv-doctor/scripts/heal.py sample-data/extreme_mess.csv

# excel-doctor
python skills/excel-doctor/scripts/diagnose.py sample-data/messy_sample.xlsx

# loader regression tests
python -m unittest discover -s tests -v
```

If you're adding a new skill, add a matching sample file to `sample-data/` with a `generate_*.py` script so others can reproduce it.

**5. Open a pull request**

Describe what your skill does and what problem it solves. No formal template — just be clear.

---

## Skill structure

Every skill needs:

- **`SKILL.md`** — tells Claude what the skill does, what input it expects, and what output it produces. Follow the format in `skills/csv-doctor/SKILL.md`.
- **`scripts/diagnose.py`** — analysis script. Outputs a single JSON object to stdout. Must run standalone with `python scripts/diagnose.py <file>`.

Healer scripts are optional but follow this pattern:

- **`scripts/heal.py`** — fixes what `diagnose.py` finds. Outputs a 3-sheet Excel workbook: **Clean Data** (fixed rows + `was_modified` + `needs_review`), **Quarantine** (unusable rows + `quarantine_reason`), **Change Log** (one row per individual change). See `skills/csv-doctor/scripts/heal.py` as a reference.

---

## loader.py conventions

`skills/csv-doctor/scripts/loader.py` is the shared file-reading layer. If you're contributing to it:

- `load_file()` must always return the standard result dict — all keys present, missing values as `None` not omitted
- Format loaders (`_load_text`, `_load_excel`, etc.) raise exceptions for unrecoverable errors; warnings go in `result["warnings"]`
- Never crash on encoding — the fallback chain (UTF-8 → detected → latin-1 → cp1252 replace) must always produce a string
- Optional dependencies (`xlrd`, `odfpy`) must fail with a clear `ImportError` message pointing to the install command
- Interactive prompts go to `stderr`; only data goes to `stdout`
- Check `sys.stdin.isatty()` before calling `input()` — scripts are often run as subprocesses by Claude Code
- In non-interactive mode, multi-sheet spreadsheets must require an explicit `sheet_name` or `consolidate_sheets=True`; never silently pick the first sheet
- `.txt` files that are not clearly delimited/tabular must raise a clear error instead of returning a misleading one-column DataFrame

---

## Ground rules

- Keep output **human-readable first**. These reports are read by people, not machines.
- **No telemetry, no network calls.** Everything runs locally.
- `diagnose.py` must exit with code `0` if analysis completed (even if issues were found) and `1` if the script itself failed. Claude needs to know the difference.
- `heal.py` prints a plain-text summary to stdout — row counts, fixes applied, assumptions made.
- Keep dependencies minimal. `pandas`, `chardet`, `openpyxl` are already in scope. Add new ones only if truly necessary and note them in your PR.

---

## Skill ideas

| Skill | What it would do |
|---|---|
| `merge-doctor` | Detect and unmerge merged cells, fill down values, rebuild a flat table |
| `type-doctor` | Detect and fix mixed-type columns — coerce strings to numbers/dates where safe |
| `encoding-fixer` | Re-encode a file to clean UTF-8, repairing Latin-1/Windows-1252 corruption |

---

## Questions

Open an issue or start a discussion. There are no dumb questions here — I'm learning too.
