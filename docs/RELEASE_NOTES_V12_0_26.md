# GridForge 12.0.26

`GROUPBY` and `PIVOTBY` now evaluate summary formulas in workbook context. Both support named built-in aggregates and one- or two-parameter LAMBDA aggregators, filter arrays, field sorting, grand totals, and `PERCENTOF`/custom two-argument aggregation. `GROUPBY` accepts multiple row-field columns, with a bounded hierarchical grouping depth.

`PIVOTBY` currently supports one row field, one column field, and one values column. Both functions return `#VALUE!` for subtotal depths beyond grand totals; PIVOTBY parent-total modes and GROUPBY's Table field relationship are unsupported. Header inference and more complex aggregation vectors are also limited. This formula-based summary behavior is separate from interactive native PivotTables.

Microsoft syntax references: [GROUPBY](https://support.microsoft.com/en-us/excel/functions/groupby-function) and [PIVOTBY](https://support.microsoft.com/en-us/excel/functions/pivotby-function).

Verification: TypeScript static check and Windows packaging with `-SkipTests`. No tests, app launch, live UI, or Excel GUI checks were run.