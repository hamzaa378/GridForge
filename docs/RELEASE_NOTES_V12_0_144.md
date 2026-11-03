# GridForge v12.0.144 — OpenDocument data validation

- ODS export now writes supported list, whole-number, decimal, date, and text-length validation conditions, list sources, blank-cell behavior, input prompts, and error alerts.
- Validation names are attached to populated cells and compact repeated blank-cell spans, so rules cover their selected ranges without filling the workbook model with empty cells.
- ODS import restores supported validation definitions and their cell ranges, including repeated rows and columns. Unsupported or malformed OpenDocument validation conditions are skipped safely.
- ODS validation export is bounded to 100,000 covered rows per sheet to keep memory and output size predictable. Larger ranges stop with a clear export error.

ODS charts, PivotTables, macros, embedded objects, and many advanced validation extensions remain unsupported. Formula translation is still a subset, and LibreOffice round-trip comparison was not performed.

Build: `npm run build`. The Windows installer build uses `-SkipTests`. No automated tests, app launch, live UI, or LibreOffice checks were run.
