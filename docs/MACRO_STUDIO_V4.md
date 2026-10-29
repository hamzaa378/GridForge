# GridForge v10.6 — Offline Macro Studio

## Purpose

GridForge macros run **without Microsoft Excel**, Node.js, or an internet connection when using the packaged Windows application. They are *declarative JSON actions*, not JavaScript or Microsoft VBA. This shared TypeScript executor powers the main spreadsheet and the separate Macro Studio shortcut.

## New v4 capabilities

- Reusable named **text and number parameters** with `{{name}}` substitution inside values, conditions and replacement strings. Numeric types reject invalid numeric inputs. Parameters do not control action kinds or executable code.
- `cleanText` operations: `trim`, `upper`, `lower`, `proper`, `collapseSpaces`. By default, formula source is skipped.
- `formatIf` formats cells whose evaluated values meet `equals`, `notEquals`, `contains`, `greaterThan`, `lessThan`, `blank` or `notBlank`.
- `removeDuplicates` compacts the selected rectangle using optionally selected key columns. Formula references shift relative to row moves; cells outside the range are unchanged. This is **not** Excel's full row-deletion feature.
- Existing `cells`, `styles`, `clear`, `fillDown`, `sequence`, `copy`, `sort`, `replace`, `ifCell` and `repeat` actions work in v4.
- All v1, v2 and v3 recorded macros remain loadable.

## Macro Studio Windows shortcut

After building the Windows installer, choose **Start → GridForge → GridForge Macro Studio**. It opens the same offline application directly in the Automate dialog. Standard GridForge is also available as a shortcut. No separate JavaScript engine is packaged: the bundled WebView2 hosts the precompiled TypeScript code.

## No-install Python launch

With only Python and a browser, double-click `START_MACRO_STUDIO_WINDOWS.bat` or run `python run_gridforge.py --macro-workbench`. You can run, preview, import and export macros without Excel.

## Example macros

- `examples/macros/cleanup_contacts.json`
- `examples/macros/parameterized_tax.json`
- `examples/macros/parameter_values.json`

To use a macro: open **Automate**, select **Import macro JSON**, choose the example, optionally fill in parameter values as a JSON object (all values as strings), click **Preview effects**, then run. Save or export your workbook afterward. The preview does not mutate the workbook.

## Developer command line (Node 20+, separate from end-user installer)

```powershell
npm run build
node scripts/run_macro_offline.mjs --workbook examples/sample_gridforge_workbook.json --macro examples/macros/parameterized_tax.json --parameters examples/macros/parameter_values.json --dry-run
node scripts/run_macro_offline.mjs --workbook examples/sample_gridforge_workbook.json --macro examples/macros/parameterized_tax.json --parameters examples/macros/parameter_values.json --output resulting-workbook.json
```

The CLI accepts **GridForge JSON documents** only; it does not directly parse XLSX/XLSM. Import/export workbook formats using the application. It never overwrites the input path and writes the result using a temporary file then atomic rename. An invalid macro leaves the output unwritten.

## Safety and limits

- No arbitrary JavaScript, VBA, shell calls, filesystem reads or outbound requests inside a native macro.
- Entire macro is simulated on a cloned workbook before any writes. Supported mutations on each changed worksheet are committed in one undoable transaction.
- Macros cannot change worksheet structure or access outside the workbook.
- Validation imposes a nesting depth of three, a limit of 20 repetitions, and an estimated limit of 15,000 cell touches. Preview changes are capped to the first 100 entries.
- Imported macros should still be treated as untrusted data. Preview before running and keep workbook backups. Running a macro is an explicit user action.
- Macro storage uses local browser/WebView storage. Export macro JSON for backup or sharing.
- **Existing VBA inside `.xlsm` is not natively executable.** The optional desktop Excel bridge remains separate and requires Microsoft Excel. Legacy VBA-to-native translation is not implemented.

## Scope of additional Excel-like functionality

This release adds familiar **Remove Duplicates, Clean Text and Conditional Formatting** as automation actions. It does not include every Excel feature: PivotTables, full VBA, Power Query, external data sources, structural row insertion and complex dynamic-array spilling remain limitations.

## Verification

Run `npm run check`, `npm test`, `python -m unittest discover -s tests -p 'test_desktop*.py' -v`, and `python run_gridforge.py --check`. For native Windows installer validation, run `BUILD_WINDOWS_INSTALLER.bat` on Windows and smoke-test both shortcuts.

### Browser test on Windows

Install Playwright and Chromium in a development environment, then run `python tests/smoke_macro_studio.py`. This checks the actual local app, the macro-workbench launch route, template selection, parameterized preview, and execution. It was blocked by the hosted Chromium administrator policy during release preparation; it must be run on Windows before claiming desktop GUI acceptance.

The Automate dialog can export a JSON **execution report** after a completed run. This is separate from the reusable macro's JSON definition.
