# GridForge 11.11.0 Release Notes

## Worksheet navigation

- Move the active worksheet one position left or right.
- Assign a preset tab color, clear a color, and retain colors on duplicated sheets.
- Save tab color in GridForge JSON and native XLSX worksheet properties.
- Select whole rows or columns from their headings; hide selected rows/columns and unhide all from View, with state retained in GridForge JSON and native XLSX.

## Find across worksheets

- Find text or formulas across all visible worksheets.
- Search results show the sheet and cell address; selecting a result activates that worksheet and cell.
- Replace All remains scoped to the active worksheet and retains the existing single-sheet undo behavior.

## Multi-level sorting

- Custom Sort supports up to three ordered sort keys, each with its own direction.
- The sort remains a single undoable operation and recorded action macros preserve all sort levels.

## Data Validation

- Restrict entries by date using ISO date input or Excel serial values.
- Restrict text by character count with the supported comparison operators.
- Persist both rule types in GridForge JSON and native XLSX validation records.

## Verification and limits

- Source TypeScript check and frontend build passed.
- Focused source tests passed (39/39).
- Windows desktop, installer and UI tests were not run; installer was not built.
- Native Excel GUI import/export was not tested. Custom-formula validation remains unsupported.
