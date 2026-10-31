# GridForge 11.30.0 Feature Matrix

| Area | Status | Current implementation and limitation |
|---|---|---|
| OneDrive | Implemented for desktop | Root-level workbook list, XLSX/JSON open, XLSX create/update, PKCE sign-in, DPAPI-protected refresh token, and eTag guarded updates. |
| Google Drive | Implemented for desktop | `drive.file` scoped file list/open/create/update, PKCE sign-in, DPAPI-protected refresh token, and a version preflight before updates. Existing Drive files may need to be opened/shared with the app; native Google Sheets are not imported. |
| Dropbox | Implemented for desktop | Root-level file list/open, XLSX create/update, PKCE offline sign-in, DPAPI-protected refresh token, and revision guarded updates. Exact redirect registration is required. |
| Cloud limitations | Partial | No folder navigation, background sync/autosave, real-time collaboration, shared-link workflows, cloud version history UI, or browser-only token persistence. File size limit is 80 MB and list cap is 500. |
| Local offline use | Implemented | Local spreadsheet editing and workbook files continue to work without configuring any cloud provider. |
| Excel compatibility | Partial | GridForge remains an independent spreadsheet engine and does not claim full Excel UI, formula, PivotTable, chart, or file-format parity. |

## Code and package verification

- `npm run check` — passed.
- `python -m py_compile desktop/gridforge_desktop.py` — passed.
- `npm run build` — passed as part of the Windows packaging flow.
- Windows installer build — passed with `-SkipTests`.
- Automated tests, Windows application launch, live UI testing, and Excel GUI checks — not run.

