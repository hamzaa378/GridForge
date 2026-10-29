# GridForge v10.9 — Single continuation / handover guide

**Checkpoint:** 23 October 2026. This is the latest source snapshot. Baseline: v10.8 native DrawingML charts + v10.6 offline Macro Studio + v10.4 tabbed ribbon. The current package is Windows-installer **source**, not a compiled installer.

## Important architecture

- TypeScript spreadsheet logic: `src/engine/workbook.ts` (dependency graph and cell transactions), `functions*.ts` (506 local formulas + 10 offline connection placeholders), `parser.ts` (custom formula parser), `document.ts` (multi-sheet workbooks). Do not replace this evaluator with arbitrary JS evaluation.
- Spreadsheet UI: `src/main.ts`, `public/styles.css`; browser app `dist/` included and served by Python `run_gridforge.py` or packaged with pywebview.
- Charts: `src/engine/charts.ts` (13 chart kinds, six palettes, SVG), `src/engine/ooxml_charts.ts` (native Excel chart XML), XLSX import/export in `src/engine/xlsx.ts`.
- Tables: new `src/engine/tables.ts` and `SheetTools.tables`. Ten style names, header/stripe/emphasis flags. `renderGrid` overlays table styles without destroying user cell styles. Native tables included as `xl/tables/tableN.xml`.
- Native macros: `src/engine/macros.ts` with bounded parameters/conditions/repetition. Microsoft VBA requires the separate optional Windows bridge, never auto-launches.
- Windows desktop: `desktop/gridforge_desktop.py`, `desktop/build_windows.ps1`, `desktop/installer/GridForge.iss`, `BUILD_WINDOWS_INSTALLER.bat`. Installer version and Python desktop version now 10.9.0.

## Exact validated checkpoints

- `npm run build` PASS (TypeScript and bundled static assets)
- `npm test` PASS **889/889**
- `python -m unittest discover -s tests -p 'test_desktop*.py'` PASS **7/7**
- `python scripts/smoke_tables_charts_v109.py` PASS: Format as Table, banded style gallery, create combo chart from table, native XLSX/JSON export, XLSX reimport, full gallery importer, font/size/quick color, no JS page errors
- `node scripts/benchmark.mjs` PASS: selective dependency recalculation after isolated edits
- `node scripts/build_excel_migration_demo_v109.mjs` generated valid example with **13 charts / 11 tables / 2 sheets**. ZIP integrity and XML parsing passed.
- No desktop Microsoft Excel COM execution or actual Windows installer build performed here. Do not claim real Excel parity or shipped EXE until tested.

## Commands

```powershell
# No development setup required to launch precompiled UI:
python run_gridforge.py
# Source build and test:
npm install
npm run build
npm test
python scripts\smoke_tables_charts_v109.py
# Build sample:
node scripts\build_excel_migration_demo_v109.mjs
# Excel Windows acceptance:
python -m pip install pywin32
python scripts\verify_excel_migration_windows_v109.py
# Installer: requires Python, Node, Inno Setup 6 and Windows
.\BUILD_WINDOWS_INSTALLER.bat
```

## Risks / outstanding work (prioritized)

1. **Real Excel validation**: open the 13-chart/11-table sample and verify chart-object count and presentation. Native DrawingML XML and GridForge roundtrip do not establish actual Excel rendering. Repair OOXML if Excel complains.
2. **Actual Windows WebView2 UX + installer**: build EXE on Windows, smoke import/export, macro studio, chart resizing, light/dark, write permission and uninstall. Signed installer for public distribution.
3. **Table semantics**: structured references (`TableName[Column]`), interactive header filter controls, table auto-expansion, calculated columns, total rows and full undo of table-layout changes are not implemented; table styling itself is persistent and editable.
4. **Charts**: secondary Y axis, per-series assignments, trendlines, histogram/box/violin/financial charts and advanced Office styles. Combo currently uses last numeric series as line.
5. **Formatting**: rich text within one cell, Excel theme-based colors and borders, conditional-formatting rules, row heights/column width and data validation beyond existing implementation.
6. **Excel function parity**: Excel 506-case comparison previously showed 450/456 matches on frozen comparable cases after v10.3 repairs; limited cases and unresolved edge cases remain. Do not equate all 506 names with complete semantics.

## Reproducibility and safe changes

Before modifying: `npm test`, `node scripts/benchmark.mjs` and the browser smoke. After modifying charts/tables, additionally validate example XLSX OOXML relationships, `importXlsx` round-trip and Windows Excel smoke. Preserve safety boundaries: no implicit VBA, arbitrary user-supplied executable macro code, network calls for external formulas or remote OOXML relationships. Keep `docs/CONTINUATION_GUIDE_V10_9.md` updated with all new modules, test results, known limits and setup steps.
