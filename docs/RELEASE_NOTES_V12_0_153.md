# GridForge v12.0.153 — ODS worksheet-object round-trips

- ODS exports now include a bounded `gridforge:sheet-objects` attribute for GridForge conditional rules, drawings, data-table definitions, Pivot summary definitions, and slicers.
- GridForge imports validate those definitions against the destination workbook limits and restores them when the workbook is reopened. Invalid or oversized metadata is ignored with a warning.
- This is GridForge-specific round-trip support. It does not create native ODF/Calc conditional formats or PivotTables, and other spreadsheet applications continue to see calculated cells and standard objects only.
- Query definitions are excluded because their external source URLs may contain private access details. Macros and unrelated embedded objects remain unsupported.
- Formula source-text preservation and half-point font sizing remain included. Static production build and Windows packaging are run without automated or live UI tests.
