# GridForge v12.0.145 — OpenDocument hyperlinks

- ODS export now writes safe external hyperlinks and in-workbook cell locations as OpenDocument text links.
- Link text stays tied to the cell's stored or calculated value; link tooltips are carried as link titles.
- ODS import restores HTTP, HTTPS, mailto, and supported in-workbook cell links, with the existing hyperlink safety and range checks.
- Relative file links and unsupported internal targets are skipped. Existing validation-rule round-tripping remains available, with a 100,000-row-per-sheet export bound.

ODS charts, PivotTables, macros, embedded objects, complex custom formats, and cloud browsing/saving remain unsupported. No LibreOffice round-trip comparison was performed.

Build: `npm run build`. Windows installer build uses `-SkipTests`. No automated tests, app launch, live UI, or LibreOffice checks were run.
