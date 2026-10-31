# GridForge 12.0.4

- Runs XLSX ZIP validation, XML parsing, worksheet construction, style loading, and formula calculation in a module worker so large local and cloud workbook imports do not monopolize the UI thread.
- Relays import progress to the status area. Returned XLSM workbook contents use the same worker path; the original macro-enabled package remains available for the existing save workflow.
- Carries forward v12.0.3 data-bar XLSX round-tripping and left chart legends, plus v12.0.2 ChartEx label settings.
- `npm run check` and Windows packaging passed with test phases skipped. No automated tests, app launch, live UI, or Excel GUI checks were run.
- Existing compressed/expanded file and worksheet size limits remain. The worker reduces UI-thread blocking but does not guarantee imports of arbitrarily large or complex files.
