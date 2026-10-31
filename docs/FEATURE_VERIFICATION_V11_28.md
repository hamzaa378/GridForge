# GridForge v11.28 Feature Matrix

| Area | Status | Current implementation and limitation |
|---|---|---|
| Excel-style workbook shell | Improved | Title and quick-access area, command tab row, ribbon, formula bar, worksheet canvas, sheet tabs, and status strip now use a more consistent Excel 2019-like hierarchy. GridForge remains a separate application. |
| Draw tab discoverability | Partial | The Draw tab is positioned with the Excel-style tabs and opens an availability notice. Ink, replay, and ink conversion tools are not implemented. |
| Worksheet focus | Improved | Dashboard metrics are hidden in the default workbook layout, while optional formula insights remain accessible. |
| Responsive and theme behavior | Retained | Light/dark and compact layouts remain supported by the existing responsive styles. No live UI verification was performed. |
| Formula, chart, and workbook compatibility | Unchanged | This release changes presentation only. See the compatibility audit and earlier feature matrices for known behavior and limitations. |

## Code-only verification

- `npm run check` — passed.
- `npm run build` — passed; regenerated `dist/`.
- `desktop/build_windows.ps1 -SkipTests` — installer/package compilation completed with Windows test phases skipped.
- No automated tests, Windows tests, app launch, live UI checks, or Excel GUI validation were run.
