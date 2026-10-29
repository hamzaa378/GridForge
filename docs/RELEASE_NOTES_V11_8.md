# GridForge 11.8.0 Release Notes

## Data Validation improvements

- Build dropdowns from a live source range on the current worksheet.
- Configure whole-number and decimal rules with between, not-between, equal, not-equal, greater-than and less-than comparisons (including inclusive variants).
- Add an input prompt title/message that appears when selecting a validated cell.
- Choose Stop, Warning or Information behavior for invalid entries; warning asks whether to keep the value.
- Save supported validation rules in GridForge JSON and native XLSX `dataValidations` records.

Inline lists support up to 50 values and 255 combined characters. Range lists are limited to the active worksheet. Date, text-length and custom-formula rules, cross-sheet ranges, and validation updates in macro-preserving XLSM saves are not included.

## Verification performed

- `npm run check` — TypeScript check passed.
- `npm run build` — frontend build passed.
- `node --test tests/data_validation_v117.test.mjs` — 9 focused source tests passed.
- Windows desktop, installer and Windows UI tests were not run, per the user's instruction; the 11.8 installer was not built.
