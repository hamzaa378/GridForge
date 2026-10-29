# GridForge v11.0 — Single-source project handover / continuation guide

## 1. Release identity

Version 11.0.0, based on v10.9 Excel table/chart release. Purpose: usable Excel-like tables, structured-reference formulas, multi-column filters and live conditional colors. Full source plus compiled static app, offline launcher, Windows installer builder, unit/browser tests, and examples packaged. Not a binary installer.

## 2. What's now implemented

- `src/engine/tables.ts`: table validation/style/color/column names; new `resolveStructuredFormula(...)` converts supported structured expressions into A1 ranges for evaluation, without overwriting the original raw text. Double-quoted strings are excluded from conversion. Same-sheet tables only. Table row-local `[@Field]` needs a cell **inside** the named table.
- `src/engine/workbook.ts`: new `setTableContext(tables)` rebuilds formula AST/dependency edges and invalidates caches after table geometry or headers change. New raw formulas are interpreted in that context. The regular parser/506 functions remain intact.
- `src/engine/document.ts`: persisted/duplicated sheets restore table formula context when reopening JSON.
- `src/engine/table_growth.ts`: new `expandTablesAfterEdit(book,tools,edits)` extends only a contiguous next row, checks overlapping tables and worksheet bounds, tracks charts based on the full table range, extends the filter region. Chart size limits of 250 rows and 12 cols remain.
- `src/engine/sheettools.ts`: `FilterSpec.criteria` adds optional `[{column,values}]` AND-filter selections. Legacy `{column,query}` remains supported. Conditional rule `operator='scale'` interpolates colors against live numeric min/max with a per-render cache to avoid repeated range scans.
- `src/main.ts`: chart/table editing, selection/formatting, new table-header filter dropdown, filter unique-value checkboxes, select all/none, per-column clear and multi-column filter composition; live scales dialog; table formula context updates after header edits/table creation/removal/expansion. Existing grid edits and in-app paste trigger growth.
- `src/engine/xlsx.ts`: original structured formulas emitted to native Excel; AutoFilter criteria in table XML (or worksheet AutoFilter if no named table), native 2-color conditional formatting; import of those same constructs. Existing native chart and table OOXML kept.
- `public/styles.css`: dedicated compact filter UI and header buttons; no additional persistent sidebar.

## 3. Test & release commands

```bash
npm run build
npm test
python scripts/smoke_migration_v110.py
python -m unittest discover -s tests -p 'test_desktop*.py'
node scripts/benchmark.mjs
node scripts/build_migration_demo_v110.mjs
python desktop/gridforge_desktop.py --check
```

- Unit tests: `tests/migration_v110.test.mjs`, existing suites. Last full run: **901/901 pass** (to be reconfirmed after packaging).
- Browser smoke: `scripts/smoke_migration_v110.py`; renders compiled app entirely in-memory so localhost policies cannot block it, tests UI actions and round-trips. Last run: **PASS** across table filters, named formulas, growth, charts, native XLSX, two-color scales and no page errors.
- Desktop wrapper tests: `tests/test_desktop*.py`, existing 7 tests (rerun before packaging).
- Excel Windows acceptance: `scripts/verify_excel_migration_windows_v110.py`. Intentionally not run here because no Windows Excel COM.
- Windows binary: `BUILD_WINDOWS_INSTALLER.bat` -> `desktop/build_windows.ps1`, version set to `11.0.0` in `package.json`, `desktop/gridforge_desktop.py`, `desktop/installer/GridForge.iss`. End-user `.exe` needs Windows PyInstaller/Inno Setup.

## 4. Sample data & expectations

`examples/GridForge_v11_0_Table_Migration_Demo.{json,xlsx}` built with `node scripts/build_migration_demo_v110.mjs`. Table `SalesMetrics` extends from A1:D6 to A1:D7. Two native charts track A1:D7, two column value filters and a numeric color scale. `F2` computes 3085 and `D7` computes 600. Browser round-trip verifies these values after importing the XLSX.

## 5. Known limitations & next iterations

- Same-sheet structured refs intentionally cover a subset. Excel special item operators (`#Totals`, `#This Row` in full nested forms), external sheet refs, escape conventions for punctuation in headers, named formulas and function-specific implicit intersection are not complete.
- Edit-triggered table growth is not a full table engine: no worksheet physical row insertion/deletion, calculated-column autofill, pivot totals, or auto-inclusion after every possible macro/remote operation. Chart automatically tracks a changed range only when it **exactly** matched the original table range.
- Filter selectors work in the app and round-trip for simple exact lists. Do not claim all Excel filter types, sorting by color, advanced custom filters, slicers or dynamic filter ranges.
- Conditional formatting uses solid highlights and 2-color scales, not full 3-color/gradient Excel parity.
- Chart structural OOXML tests and GridForge importer pass, but test live Excel renderer with `verify_excel_migration_windows_v110.py` before releasing as Excel-compatible.
- Node runner tests and Chromium do not verify real WebView2 installer behavior. Run the Windows installer, portable EXE, native Save As, opening/reopening, and uninstaller on a clean Windows machine.
- Preserve original 506 parity fixtures. Old 506 frozen cases remain a single-test-per-name audit; avoid claiming 100% Excel formula compatibility.

## 6. Handoff priorities

1. Run real Windows Excel acceptance and inspect warnings/repairs. If table filters appear unexpectedly, verify `.xlsx` `<autoFilter>` is inside the native table part rather than duplicated in worksheet XML.
2. Extend structured references to `#Totals`, calculated columns and consistent table expansion for macro/bulk-paste paths, backed by dependency invalidation tests.
3. Secondary chart axes/trendlines and richer conditional formatting with OOXML interoperability tests.
4. Accessibility/keyboard-focused pass on header filter menus and table formatting.
5. Build/sign installer only after a successful clean-machine Windows acceptance test.
