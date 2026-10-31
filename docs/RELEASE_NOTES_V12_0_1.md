# GridForge 12.0.1

- Fixed the Draw and Page Layout ribbons expanding vertically to consume unused window space. They now keep a fixed ribbon-height footprint so the worksheet remains the main work area.
- Fixed `TOCOL` and `TOROW` so `ignore=2` or `ignore=3` can discard cell errors in referenced ranges before array construction; blank filtering remains controlled independently by `ignore=1` or `ignore=3`.
- XLSX import now yields between large cell-reading, workbook-update, and formula-calculation batches, with progress shown in the app status area.
- ZIP checksum verification during XLSX import now runs in 1 MB chunks and reports progress instead of checking a large expanded file in one uninterrupted pass.
- CSV import and public CSV query refresh now parse and evaluate formulas in batches, keeping the window responsive on larger files. CSV imports clipped to GridForge's row or column limits now report the source dimensions.
- Cloud workbook opening reports the XLSX import stage and restores the prior status after a failed local import.
- `npm run check` passed. The Windows installer build uses `-SkipTests`; no live UI or Windows tests were run, by request.

