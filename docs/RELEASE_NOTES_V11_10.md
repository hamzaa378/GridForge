# GridForge 11.10.0 Release Notes

## Worksheet visibility

- Hide the selected worksheet and unhide sheets from a worksheet list.
- Require at least one visible worksheet, and prevent selecting or deleting the final visible sheet.
- Persist hidden state in GridForge JSON and native XLSX workbook sheet metadata.

## Earlier additions in this work session

- Text to Columns preview/splitting with delimiters, quoted fields, escaped quotes and overwrite checks.
- Dynamic same-sheet Data Validation list sources, eight numeric criteria, input prompts and Stop / Warning / Information alerts.

GridForge's macro-preserving XLSM cell-edit patcher does not update sheet visibility or validation metadata. Full Excel workbook structure parity is not claimed.

## Verification performed

- `npm run check` — passed.
- `npm run build` — passed.
- Focused source tests — 16/16 passed across validation, Text to Columns and worksheet visibility.
- Windows desktop, installer and Windows UI tests were not run, per the user's instruction; the 11.10 installer was not built.
