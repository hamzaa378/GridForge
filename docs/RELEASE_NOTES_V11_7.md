# GridForge 11.7.0 Release Notes

## Data Validation

- Apply dropdown lists to a selected range from the Excel-style Data tab.
- Reject manually entered values that are not in the dropdown list.
- Restrict entries to whole numbers or decimals with optional minimum and maximum bounds.
- Allow or disallow blank values, edit rules for a selected range, or clear overlapping rules.
- Persist these settings in GridForge JSON and write/read the supported native XLSX Data Validation subset.

Inline dropdown lists support up to 50 values and 255 combined characters. Validation sourced from a cell range, custom formulas, dates and text length are not included. XLSM preservation behavior is unchanged; this release does not add validation editing to macro-preserving XLSM saves.

## Verification performed

- `npm run check` — TypeScript check passed.
- `npm run build` — frontend build passed.
- `node --test tests/data_validation_v117.test.mjs` — 4 focused source tests passed.
- Per the user request, Windows desktop, installer and Windows UI tests were not run. The 11.7 installer has not been built or launch-verified.
