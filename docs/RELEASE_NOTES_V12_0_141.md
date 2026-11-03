# GridForge v12.0.141 — OpenDocument formula round-tripping

- ODS export now writes formulas that GridForge can safely translate into OpenFormula syntax. Formulas outside the supported translation subset continue to export their calculated values.
- ODS import translates supported OpenFormula formulas back into GridForge formula syntax and recalculates them. The importer reports how many formulas were translated and how many were retained as cached values.
- Repeated ODS formula cells now adjust relative references for each expanded row and column, preserving expected copy-fill behavior.
- The ODS export prompt now describes formula handling and the remaining unsupported workbook objects.
- Cell borders now have thin, medium, and thick weights in the Home ribbon, with weight retained in XLSX and ODS for supported border styles.
- ODS number-format interchange now handles simple scientific formats such as `0.00E+00`, including exponent digit count and forced positive signs.

ODS still does not preserve charts, PivotTables, validation rules, macros, embedded objects, or complex multi-section and conditional number formats. OpenFormula function behavior, locale rules, dates, and calculation precision need broader Calc comparisons. XLSX uses the nearest available preset for thick dashed and dotted lines.

Build: `npm run build`. No automated tests, Windows runtime, live UI, or LibreOffice checks were run.
