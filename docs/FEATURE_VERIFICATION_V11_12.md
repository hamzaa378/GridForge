# GridForge v11.12 Feature Matrix

| Area | Behavior | Verification |
|---|---|---|
| Cross-worksheet formulas | Direct references and rectangular ranges, including quoted sheet names, with cross-sheet cycle checks | TypeScript check and frontend build |
| Pivot summaries | Group by row and optional column fields; Sum, Count, Average, Minimum and Maximum; grand totals; saved refresh definitions | TypeScript check and frontend build |
| Summary refresh | Manual refresh command plus refresh after normal edits, paste, undo and redo | Source implementation; no UI test run |
| XLSX compatibility | Cross-sheet formula text and ordinary summary cells remain part of standard workbook export | Existing OOXML export path; Excel GUI was not launched |

## Limits

- Pivot summaries are GridForge definitions that render ordinary worksheet cells. They are not native Excel PivotTables; PivotCharts, slicers, timelines, PivotCaches and pivot refresh metadata are not emitted.
- Power Query, Data Model, Power Pivot, DAX and OLAP connections remain unsupported.
- Dynamic arrays still do not spill. Function registration does not prove behavior parity, and the 10 service-dependent function names are not fully usable offline.
- Cross-sheet formulas do not yet cover external workbook links, full Excel reference syntax, all reference-oriented function semantics, or automatic formula rewrite when a worksheet is renamed or deleted.
- Windows desktop, installer, UI and Microsoft Excel interoperability tests were not run.
