# GridForge 11.9.0 Release Notes

## Data tab features

- Text to Columns splits a selected one-column range into adjacent fields using comma, semicolon, tab, space or a custom one-character delimiter.
- Preview the first rows before splitting; support quoted text fields, escaped quotation marks and optional merging of repeated delimiters.
- Choose an output destination and receive an overwrite prompt when existing cells would be replaced.
- Data Validation now supports dynamic list ranges, all eight numeric comparison operators, input prompts and Stop / Warning / Information alert behavior.
- GridForge JSON and supported XLSX Data Validation metadata preserve supported rule settings.

Text to Columns materializes the resulting fields as cell values. It supports one-column source selections; multiline text within a single cell is not treated as separate worksheet rows. Validation lists sourced from another worksheet, date/text-length/custom-formula validations and XLSM validation edits remain unsupported.

## Verification performed

- `npm run check` — TypeScript check passed.
- `npm run build` — frontend build passed.
- Focused source tests — 13/13 passed across data validation and Text to Columns.
- Windows desktop, installer and Windows UI tests were not run, per the user's instruction; the 11.9 installer was not built.
