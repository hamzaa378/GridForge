# GridForge v12.0.135 — OpenDocument compatibility

- Added ODS import for supported numeric, text, boolean, date, time, cached formula values, worksheet names, and repeated rows or columns. Large imports run in a background worker and obey GridForge worksheet and archive limits.
- Added ODS export for calculated cell values across worksheets, including supported dynamic-array spill results. Export is capped to GridForge worksheet and populated-cell limits.
- Added an ODS export confirmation and import notes that explain the feature boundary: formulas are imported as cached values; export does not preserve formulas, formatting, charts, PivotTables, validation rules, macros, or embedded objects.
- Added ODS support to the Windows Save As dialog and documented local ODS behavior. Cloud browsing and save remain XLSX/JSON only.
- Code compilation completed. No automated tests, Windows test phases, or live UI checks were run.
