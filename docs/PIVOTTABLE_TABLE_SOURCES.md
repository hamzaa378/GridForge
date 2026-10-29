# Table-backed PivotTable sources

When a Pivot Summary uses a formatted GridForge table, selected directly or by entering its exact source range, GridForge saves the table name with the PivotTable definition. This lets refresh follow the table's current extent as it grows.

For XLSX export, the PivotCache uses a native `worksheetSource` table name instead of freezing the cache to a cell range. The exported cache records and worksheet report are built from the current table range. On import, a native table source resolves back to the table name and its current range, so later refreshes continue to follow the table.

The worksheet table itself remains the source of truth. If it is renamed or deleted, the PivotTable cannot resolve the saved table source and reports that the source is missing. A Pivot Summary created from an arbitrary cell range remains range-based.

This implementation is covered by code-level XLSX export/import checks. It has not been validated by opening the result in Microsoft Excel.
