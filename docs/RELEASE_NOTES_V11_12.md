# GridForge 11.12.0 Release Notes

## Cross-worksheet formulas

- Formula references can read cells and rectangular ranges from another worksheet, for example `=Sheet2!A1`, `=SUM(Data!B2:B20)`, and `='Quarter 1'!C4`.
- Quoted sheet names with spaces and apostrophes are supported. Cross-sheet references resolve case-insensitively and participate in cycle detection.
- Cross-sheet formulas recalculate when read, so source edits are reflected without stale cached values.
- Formula behavior remains a custom implementation. Some reference-oriented functions, defined names scoped to other sheets, renamed-sheet reference repair, external workbook links and differing sheet dimensions are not fully Excel-compatible.

## Pivot Summary

- Create a grouped summary from a headered range, with a row field, optional column field, value field, and Sum, Count, Average, Minimum or Maximum aggregation.
- A grand total is included. Saved definitions can refresh from their original source, and edits made through the normal cell editor refresh summaries automatically.
- The generated report is an ordinary cell range. It can be charted with the existing chart tools.
- GridForge does not create Excel-native PivotTable or PivotChart metadata. Slicers, timelines, Power Query, Data Model, Power Pivot, DAX and OLAP connections remain unsupported.

## Verification and limits

- TypeScript source check and frontend build passed.
- Windows desktop, installer, Excel GUI and Windows UI tests were not run.
- Dynamic-array formulas still do not spill into neighboring cells. The 506 registered local function names are not 506 exact Excel matches; service-dependent functions and formula fidelity gaps remain.
