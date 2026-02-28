# JSON Contracts

These schemas document the stable machine-readable outputs used by the local UI,
future backend/API layers, and CI contract tests.

Versioning rules:
- `schema_version` changes only for breaking contract changes
- additive fields do not require a major version bump
- scripts also emit a `contract` object with `name` and `version`
- `tool_version` tracks the sheet-doctor release, not the schema version

Current contracts:
- `csv_doctor.diagnose`
- `csv_doctor.report`
- `csv_doctor.heal_summary`
- `excel_doctor.diagnose`
- `excel_doctor.heal_summary`
