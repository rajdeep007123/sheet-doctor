# Production Blockers

This file tracks the issues that still block a credible "production-ready"
claim for `sheet-doctor`. The ordering below is strict: fix the earlier items
first because the later items depend on them or matter less to user trust.

## P0

- [ ] Workbook sheet selection in `csv-doctor/heal.py`
  - Current behavior: multi-sheet workbooks fail unless the caller already
    selected a sheet through `loader.py`.
  - Evidence: public `any_sheets.xlsx` fails with `ValueError: Multiple sheets found`.
  - Required outcome:
    - CLI supports `--sheet <name>`
    - CLI supports `--all-sheets` or `--consolidate-sheets`
    - failure messages stay clean and actionable

- [ ] `.xlsm` empty-after-header-detection failure
  - Current behavior: some workbook inputs load, then `heal.py` exits with
    `ERROR: File is empty after metadata/header detection.`
  - Evidence: public `issue221.xlsm`
  - Required outcome:
    - workbook data survives preprocessing
    - metadata/header detection does not strip real tabular content

- [ ] Clean corrupt-workbook failure contract
  - Current behavior: corrupt `.xls` inputs leak parser noise before the final error.
  - Evidence: public `corrupted_error.xls`
  - Required outcome:
    - loader/healer returns one clean typed error
    - no `xlrd` internal spew reaches the user

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

- [ ] Post-heal scoring and recoverability reporting
  - Current behavior: report score reflects raw-file damage only.
  - Required outcome:
    - `raw_health_score`
    - `recoverability_score`
    - `post_heal_score`
    - report makes clear what was salvageable vs what still needs review

- [ ] Reporter golden tests
  - Current behavior: reporter has unit coverage, but not locked text/json snapshots.
  - Required outcome:
    - golden test for `.txt`
    - golden test for `.json`
    - stable output contract for the UI

## P2

- [ ] Large-file guardrails
  - Current behavior: full in-memory processing path with no explicit safety rails.
  - Required outcome:
    - size warnings
    - row-count warnings
    - degraded/sample mode for large inputs

- [ ] Optional dependency error normalization
  - Current behavior: dependency-missing behavior varies by file type and code path.
  - Required outcome:
    - consistent `ImportError`/`ValueError` contracts
    - dependency-aware test skips where appropriate

- [ ] CI / schema stability
  - Current behavior: local tests are good, but output stability is not enforced in CI.
  - Required outcome:
    - CI matrix
    - JSON schema/version tests
    - release-quality repeatability

## Current Real-World Evaluation Snapshot

- Strong:
  - messy CSV with non-exact headers and encoding issues
  - row-accounting consistency
  - semantic healing for flat tabular files

- Partial:
  - workbook diagnostics
  - workbook healing for merged cells / header cleanup

- Weak:
  - multi-sheet healing UX
  - corrupt workbook failure path
  - workbook semantic recovery
