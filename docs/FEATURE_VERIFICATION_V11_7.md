# GridForge v11.7 Feature Matrix and Verification

| Area | v11.7 behavior | Verification |
|---|---|---|
| List validation | Apply an inline list to a selected cell range; reject values outside the list | 4 focused source tests; frontend build |
| Numeric validation | Whole-number or decimal entries with optional minimum and maximum; blank handling | 4 focused source tests |
| Workbook persistence | Rules save and reopen in GridForge JSON | Focused source tests |
| Excel file support | Supported rules write to native XLSX `dataValidations`; supported inline list and numeric records import | XLSX export XML covered by focused source tests |
| Data tab | Open Data Validation for the current selection; edit or clear overlapping rules | TypeScript check and frontend build |

## Compatibility limits

This implementation supports inline dropdown lists (up to 50 values and 255 combined characters), whole numbers and decimals with optional inclusive minimum/maximum bounds, and blank-cell handling. It does not support cell-range-sourced lists, custom formulas, dates, text length or XLSM validation edits. GridForge remains a standalone spreadsheet and does not claim full Microsoft Excel parity.

## Verification

- `npm run check` — passed.
- `npm run build` — passed.
- `node --test tests/data_validation_v117.test.mjs` — 4/4 passed.
- Windows desktop, installer and Windows UI tests were not run, at the user's request. The 11.7 installer was not built.
