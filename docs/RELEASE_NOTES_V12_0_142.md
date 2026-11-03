# GridForge v12.0.142 — Chart axis and ODS compatibility

- Added an editable logarithmic vertical axis to supported column, line, area, combo, and scatter charts. Choose a base from 2 through 1000; GridForge previews the scale and exports/imports the logarithmic axis in XLSX. Log scales require positive primary-axis values and positive axis bounds.
- Continued ODS formula round-tripping through OpenFormula for the safely translatable subset, with cached values retained for unsupported formulas. Repeated formula cells adjust relative references as they expand.
- Added ODS interchange for simple scientific custom number formats, including exponent digits and forced positive signs.
- Added thin, medium, and thick border weight controls, preserved in supported XLSX and ODS styles.

ODS charts, PivotTables, data validation, macros, and embedded objects remain unsupported. Formula behavior, locale-specific rules, dates, and numeric precision still need broader LibreOffice comparisons. Thick dashed/dotted line combinations are simplified in XLSX.

Build: `npm run build`. Windows installer compilation uses `desktop/build_windows.ps1 -SkipTests`. No automated tests, app launch, live UI, or LibreOffice checks were run.
