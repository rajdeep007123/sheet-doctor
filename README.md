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

You need [Claude Code](https://claude.ai/code) and Python 3.9+ installed.

**1. Clone the repo**

```bash
git clone https://github.com/rajdeep007123/sheet-doctor.git
cd sheet-doctor
```

**2. Create a virtual environment and install dependencies**

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

> On Windows: `.venv\Scripts\activate`

**3. Register the skill with Claude Code**

Copy the skill folder into your Claude Code skills directory:

```bash
cp -r skills/csv-doctor ~/.claude/skills/csv-doctor
```

Or symlink it if you want edits to take effect immediately:

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

A deliberately broken sample file lives at `sample-data/messy_sample.csv` — encoding corruption, misaligned columns, 7 different date formats, empty rows, and a duplicate header row all baked in on purpose.

```bash
source .venv/bin/activate
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
