# GridForge v12.0.56

- XLSX cached error cells import as formula error values so dependent formulas can propagate errors instead of seeing text. Unknown cached error codes remain text with an import warning.
- Formula cells export native OOXML error types and cached scalar results for numbers, text, and booleans, improving previews in spreadsheet readers that do not recalculate formulas immediately.
- Static TypeScript check and production build passed. Windows installer and portable packages compiled with `-SkipTests`; no tests, app launch, live UI, or Excel GUI checks were run.
