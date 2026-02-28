# Production Blockers

This file tracks the issues that still block a credible "production-ready"
claim for `sheet-doctor`. The ordering below is strict: fix the earlier items
first because the later items depend on them or matter less to user trust.

## P0

- [x] Workbook sheet selection in `csv-doctor/heal.py`
  - Current behavior: multi-sheet workbooks fail unless the caller already
    selected a sheet through `loader.py`.
  - Evidence: public `any_sheets.xlsx` fails with `ValueError: Multiple sheets found`.
  - Required outcome:
    - CLI supports `--sheet <name>`
    - CLI supports `--all-sheets` or `--consolidate-sheets`
    - failure messages stay clean and actionable
  - Status:
    - `heal.py` now supports `--sheet <name>` and `--all-sheets`
    - public `any_sheets.xlsx` now heals successfully when a sheet is selected explicitly

- [x] `.xlsm` empty-after-header-detection failure
  - Current behavior: some workbook inputs load, then `heal.py` exits with
    `ERROR: File is empty after metadata/header detection.`
  - Evidence: public `issue221.xlsm`
  - Required outcome:
    - workbook data survives preprocessing
    - metadata/header detection does not strip real tabular content
  - Status:
    - `issue221.xlsm` now heals successfully after workbook/header detection hardening

- [x] Clean corrupt-workbook failure contract
  - Current behavior: corrupt `.xls` inputs leak parser noise before the final error.
  - Evidence: public `corrupted_error.xls`
  - Required outcome:
    - loader/healer returns one clean typed error
    - no `xlrd` internal spew reaches the user
  - Status:
    - loader now wraps corrupt workbook opens and suppresses parser spew in the tested path
    - corrupt `.xls` failures now return a single clean workbook-open error

## P1

- [ ] Workbook-semantic healing beyond flat tabular sheets
  - Current behavior: messy workbook layouts mostly fall back to generic cleanup.
  - Evidence:
    - `messy_aki.xlsx`
    - `messy_bp.xlsx`
    - `messy_glucose.xlsx`
  - Required outcome:
    - semantic mode works on workbook-derived tables with preambles, ragged layouts,
      and non-exact headers
    - recover date/amount/status/category semantics when confidence is high enough
  - Status:
    - workbook healing now preserves raw worksheet rows so preambles and true headers survive into semantic detection
    - workbook inputs with leading metadata rows and non-exact headers now heal in `semantic` mode in the tested path

- [x] Post-heal scoring and recoverability reporting
  - Current behavior: report score reflects raw-file damage only.
  - Required outcome:
    - `raw_health_score`
    - `recoverability_score`
    - `post_heal_score`
    - report makes clear what was salvageable vs what still needs review
  - Status:
    - reporter now emits all three scores and uses actual healing projection data

- [x] Reporter golden tests
  - Current behavior: reporter has unit coverage, but not locked text/json snapshots.
  - Required outcome:
    - golden test for `.txt`
    - golden test for `.json`
    - stable output contract for the UI
  - Status:
    - `tests/golden/extreme_mess_report.txt` and `tests/golden/extreme_mess_report.json` now lock the reporter output
    - `tests/test_reporter.py` validates both snapshots after timestamp normalization

## P2

- [x] Large-file guardrails
  - Current behavior: full in-memory processing path with no explicit safety rails.
  - Required outcome:
    - size warnings
    - row-count warnings
    - degraded/sample mode for large inputs
  - Status:
    - loader now warns on large file sizes and row counts
    - loader exposes `degraded_mode` for risky but allowed inputs
    - files above hard safety limits now fail early with a clear error

- [x] Optional dependency error normalization
  - Current behavior: dependency-missing behavior varies by file type and code path.
  - Required outcome:
    - consistent `ImportError`/`ValueError` contracts
    - dependency-aware test skips where appropriate
  - Status:
    - `.xls` now fails with a clear `ImportError` for missing `xlrd`
    - `.ods` now fails with a clear `ImportError` for missing `odfpy`
    - loader tests cover these contracts directly

- [x] CI / schema stability
  - Current behavior: local tests are good, but output stability is not enforced in CI.
  - Required outcome:
    - CI matrix
    - JSON schema/version tests
    - release-quality repeatability
  - Status:
    - GitHub Actions now runs compile checks, unit tests, and a sample CSV smoke pipeline
    - versioned schema files live under `schemas/`
    - contract tests validate stable machine-readable metadata

## Current Real-World Evaluation Snapshot

- Strong:
  - messy CSV with non-exact headers and encoding issues
  - row-accounting consistency
  - semantic healing for flat tabular files
  - large-file warning/degraded-mode behavior
  - cleaner optional-dependency and corrupt-workbook failure paths

- Partial:
  - workbook diagnostics
  - workbook healing for merged cells / header cleanup

- Weak:
  - workbook semantic recovery
