# GridForge v12.0.159 — ODS worksheet-local names

ODS export now writes worksheet-scoped defined names as native table-local named expressions. GridForge already imported these into the worksheet scope; both directions now preserve supported names and OpenFormula-compatible definitions.

Names whose formulas fall outside the OpenFormula translation subset still use cached values in other applications, with the original GridForge formula retained in GridForge-specific metadata where supported.

Verification: `npm run build` and the Windows desktop release build with `-SkipTests` passed. Automated tests, live UI, and LibreOffice GUI round-trips were not run.
