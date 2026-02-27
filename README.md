# sheet-doctor

> Designer by heart. Back in code after 8 years. Built this because someone had to — and no one else was doing it for free.

**sheet-doctor** is a free, open-source Claude Code Skills Pack for diagnosing and fixing messy CSV and Excel files. Drop a broken spreadsheet in, get a human-readable health report out. No SaaS subscription. No upload limits. No data leaving your machine.

---

## What it does

Real-world spreadsheets are disasters. Wrong encodings, misaligned columns, five different date formats in the same column, blank rows, duplicate headers — sheet-doctor finds all of it and tells you exactly what's wrong and where.

**Current skills:**

| Skill | What it checks |
|---|---|
| `csv-doctor` | Encoding, column alignment, date format consistency, empty rows, duplicate headers |

More skills coming: `excel-doctor`, `merge-doctor`, `type-doctor`.

---

## Install

sheet-doctor runs as a Claude Code skill. You need [Claude Code](https://claude.ai/code) installed.

**1. Clone the repo**

```bash
git clone https://github.com/rajdeep/sheet-doctor.git
cd sheet-doctor
```

**2. Install Python dependencies**

```bash
pip install pandas chardet openpyxl
```

**3. Register the skill with Claude Code**

Copy the skill folder into your Claude Code skills directory:

```bash
cp -r skills/csv-doctor ~/.claude/skills/csv-doctor
```

Or, if you want to develop and edit the skill in place, symlink it:

```bash
ln -s "$(pwd)/skills/csv-doctor" ~/.claude/skills/csv-doctor
```

**4. Run it**

In any Claude Code session, just ask:

```
/csv-doctor path/to/your/file.csv
```

Or drag a file in and say: *"diagnose this CSV"*

---

## Try it immediately

A deliberately broken sample file lives at `sample-data/messy_sample.csv`. It has every problem baked in on purpose — encoding issues, misaligned columns, mixed date formats, empty rows, duplicate headers.

```bash
python skills/csv-doctor/scripts/diagnose.py sample-data/messy_sample.csv
```

---

## How it works

Each skill in sheet-doctor is a `SKILL.md` file that tells Claude what the skill does, plus a `scripts/` folder with the actual Python that does the work. Claude reads the skill, runs the script, interprets the output, and gives you a clean plain-English health report.

```
skills/
└── csv-doctor/
    ├── SKILL.md          ← Claude reads this to understand the skill
    └── scripts/
        └── diagnose.py   ← Python does the actual analysis
```

---

## Contributing

This is a solo side project from a designer who came back to code. Contributions are genuinely welcome — see [CONTRIBUTING.md](CONTRIBUTING.md).

---

## License

MIT. Free forever.
