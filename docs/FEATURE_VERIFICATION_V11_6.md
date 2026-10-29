# GridForge v11.6 Feature Matrix and Verification

This release adds worksheet-scoped named ranges and Name Manager workflows to the standalone Excel-style workspace.

| Area | v11.6 behavior | Verification |
|---|---|---|
| Define Name | Saves a unique worksheet-scoped name for a cell or rectangular range, with an optional comment | Engine and browser workflows |
| Name Manager | Lists names for the active worksheet; edits and deletes names | Browser interaction checks |
| Create from Selection | Creates names from selected header rows and/or columns; resolves duplicate headings with unique suffixes | Browser interaction checks |
| Formula evaluation | Named cells and ranges can be used directly or inside functions; dependency edges point to the named cells | 3 engine tests and browser formula checks |
| History and persistence | Name changes support undo/redo and persist in GridForge JSON | Engine and browser checks |
| XLSX/XLSM | Reads and writes simple worksheet-scoped `definedNames` referring to cells/ranges | XLSX browser round-trip and XLSM package-preservation check |

## Compatibility limits

GridForge is a standalone spreadsheet application. It does not embed, automate or require Microsoft Excel. Names in this release are worksheet-scoped and refer to cells or rectangular ranges. Workbook-scoped names, formula/constant names and stepwise formula evaluation are not supported. Other XLSX/XLSM behavior remains limited to the documented OOXML subset.

## Verification

- `npm run regression` — 10/10 stages passed, including 918/918 Node tests, both performance benchmarks, 32 general Chromium checks, the v11.5 Formula ribbon workflow, v11.6 named-range workflow, XLSM name round-trip, and 8 VBA bridge security checks.
- `python -m unittest discover -s tests -p "test_desktop*.py"` — 7/7 desktop tests passed.
- `desktop/build_windows.ps1 -SkipTests` compiled the 11.6.0 Windows installer successfully.
- The installer was installed in an isolated Windows 11 folder; the app launched and reported `Responding=True`, then the test installation was uninstalled.
- Microsoft Excel itself was not launched for this verification.
