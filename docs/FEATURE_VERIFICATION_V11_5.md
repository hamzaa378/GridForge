# GridForge v11.5 Feature Matrix and Verification

This release adds a functional Formula ribbon to the standalone Excel-style workspace.

| Area | v11.5 behavior | Verification |
|---|---|---|
| Function library | Searchable function insertion dialog and categorized ribbon menus backed by the existing function catalog | Focused Chromium UI workflow |
| Aggregate insertion | SUM, AVERAGE, COUNT, MAX and MIN formulas for a selected range or the contiguous numeric cells above the active cell | Browser checks verify formula output and stored expression |
| Formula auditing | Highlight direct precedents and dependents; clear highlights | Browser checks against workbook dependency graph |
| Formula display | Toggle formulas in populated worksheet cells | Browser checks both formula and calculated-value views |
| Formula inspection | Shows expression tree, referenced cells and values, and current result | Focused Chromium UI workflow |

## Compatibility limits

GridForge is a standalone spreadsheet application. It does not embed, automate or require Microsoft Excel. The feature set is not full Excel parity. Named ranges, calculation mode controls and stepwise evaluation are not implemented in this release. Native XLSX import/export and chart support are limited to the documented OOXML subset.

## Verification

- `npm run regression` — 9/9 stages passed: 915/915 Node tests; 32 general Chromium checks; existing chart, table, resize and data-bar workflows; the v11.5 Formula ribbon workflow; XLSM preservation; 8 VBA bridge security checks; and both performance benchmarks.
- `python -m unittest discover -s tests -p "test_desktop*.py"` — 7/7 desktop tests passed.
- `desktop/build_windows.ps1 -SkipTests` — portable app and 11.5.0 Windows installer built successfully.
- The installer was installed to an isolated folder on Windows 11; the installed app launched and reported `Responding=True`, then the test installation was uninstalled.
- Browser tests use Chromium; no Microsoft Excel GUI was launched for interoperability verification.

