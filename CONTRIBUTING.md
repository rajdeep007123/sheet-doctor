# Contributing to sheet-doctor

Thanks for being here. This project is built by one person returning to code after 8 years, so every contribution — big or small — genuinely matters.

---

## What we need most

- **New skills** — `merge-doctor`, `type-doctor`, `encoding-fixer` (see ideas below)
- **Healer scripts** — `excel-doctor` has `diagnose.py` but no `heal.py` yet
- **Better heuristics** in existing scripts — edge cases, encoding detection improvements
- **More messy sample files** — real-world broken CSVs and .xlsx files (anonymized, please)
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

**4. Test against the sample data**

```bash
# csv-doctor
python skills/csv-doctor/scripts/diagnose.py sample-data/messy_sample.csv
python skills/csv-doctor/scripts/diagnose.py sample-data/extreme_mess.csv
python skills/csv-doctor/scripts/heal.py sample-data/extreme_mess.csv

# excel-doctor
python skills/excel-doctor/scripts/diagnose.py sample-data/messy_sample.xlsx
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
| `excel-doctor heal.py` | Fix what excel-doctor diagnoses — output a clean .xlsx |
| `merge-doctor` | Detect and unmerge merged cells, fill down values, rebuild a flat table |
| `type-doctor` | Detect and fix mixed-type columns — coerce strings to numbers/dates where safe |
| `encoding-fixer` | Re-encode a file to clean UTF-8, repairing Latin-1/Windows-1252 corruption |

---

## Questions

Open an issue or start a discussion. There are no dumb questions here — I'm learning too.
