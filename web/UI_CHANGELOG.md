# UI Changelog

Notes specific to the local `web/app.py` interface live here so UI work can move without getting lost in the core project changelog.

---

## [Unreleased]

### Added
- Public file URL intake alongside local uploads
- Sequential batch processing with an in-app progress/status loader
- Per-file download buttons that persist after processing
- Workbook interpretation previews in the queue and result views:
  - detected header-band rows
  - metadata rows removed
  - effective headers after workbook preprocessing
  - chosen semantic columns with confidence
  - applied override summary when the user forces a header row or semantic role
- Public URL rewriting for common share links:
  - GitHub `blob` -> raw file
  - Dropbox -> direct download
  - Box -> direct download
  - Google Drive public file links -> direct download
  - Google Sheets share URLs -> `.xlsx` export
  - OneDrive public links -> download mode
- Response-based remote file-type inference when the shared URL hides the extension
- Inline source notes in the queue for special URL handling, including:
  - `Detected Google Sheet. It will be exported to .xlsx before processing.`

### Changed
- Upload and URL inputs stay enabled until processing actually begins
- Streamlit's top decoration/status strip is hidden and replaced with in-app status messaging
- UI styling aligned more closely with the Quietly.tools palette and font direction
- Workbook configuration now exposes:
  - a header-row override control
  - per-column semantic role override controls
  - optional tabular rescue mode for modern workbook files when users want a 3-sheet readable output instead of workbook-preserving healing
