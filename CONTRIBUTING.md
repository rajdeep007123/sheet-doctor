# Contributing to sheet-doctor

Thanks for being here. This project is built by one person returning to code after 8 years, so every contribution — big or small — genuinely matters.

---

## What we need most

- **New skills** — `excel-doctor`, `merge-doctor`, `type-doctor`, `encoding-fixer`
- **Better heuristics** in `diagnose.py` — edge cases, encoding detection improvements
- **More messy sample files** — real-world broken CSVs (anonymized, please)
- **Documentation fixes** — if something confused you, it'll confuse others

---

## How to contribute

**1. Fork and clone**

```bash
git clone https://github.com/rajdeep/sheet-doctor.git
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
        └── diagnose.py   (or whatever makes sense)
```

**4. Test against the sample data**

```bash
python skills/csv-doctor/scripts/diagnose.py sample-data/messy_sample.csv
```

If you're adding a new skill, add a matching sample file to `sample-data/`.

**5. Open a pull request**

Describe what your skill does and what problem it solves. No formal template — just be clear.

---

## Skill structure

Every skill needs:

- **`SKILL.md`** — tells Claude what the skill does, what input it expects, and what output it produces. Follow the format in `skills/csv-doctor/SKILL.md`.
- **`scripts/`** — the actual code. Must run standalone with `python scripts/diagnose.py <file>`.

---

## Ground rules

- Keep output **human-readable first**. These reports are read by people, not machines.
- **No telemetry, no network calls.** Everything runs locally.
- Scripts must exit cleanly with a non-zero code if analysis fails — Claude needs to know something went wrong.
- Keep dependencies minimal. `pandas`, `chardet`, `openpyxl` are already in scope. Add new ones only if truly necessary and note them in your PR.

---

## Questions

Open an issue or start a discussion. There are no dumb questions here — I'm learning too.
