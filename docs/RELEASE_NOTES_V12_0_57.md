# GridForge v12.0.57

- `FIELDVALUE` now returns `#FIELD!` because GridForge does not implement linked Excel data types; it no longer interprets arbitrary text as a record.
- `INFO("release")` returns the Excel 16.0 compatibility release string, and `INFO("numfile")` returns the current workbook worksheet count. `INFO("recalc")` reports Automatic.
- Static TypeScript check and production build passed. Windows installer and portable packages compiled with `-SkipTests`; no tests, app launch, live UI, or Excel GUI checks were run.
