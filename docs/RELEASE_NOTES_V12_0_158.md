# GridForge v12.0.158 — ODS conditional style import

GridForge now reads supported numeric comparison and interval rules from ODF conditional cell styles and compacts adjacent styled cells into bounded rule ranges. Together with v12.0.157, these rules can round-trip between GridForge and ODF-based applications for populated cells.

Imported conditional cell coverage is capped at 100,000 cells and reconstructed rules at 30 per sheet. Data bars, color scales, blank-cell formatting, formula-driven conditions, and complex rule precedence remain incomplete.

Verification: `npm run build` and the Windows desktop release build with `-SkipTests` passed. Automated tests, live UI, and LibreOffice GUI round-trips were not run.
