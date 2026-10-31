# GridForge 12.0.28

Formula-based PIVOTBY now accepts multiple measure columns and returns each measure under its column-field category, including grand totals. It remains limited to one row field and one column field.

Workbook restoration can now defer formula evaluation when returning from the XLSX, CSV, or GridForge JSON import workers. Formula dependency graphs and table/name contexts are restored immediately, and values are computed as cells are requested by the interface. This removes a second whole-workbook calculation on the UI thread after worker import. Saved workbook restoration uses the same lazy path. Cloud JSON files now use the JSON import worker rather than parsing and calculating synchronously in the cloud-open handler.

The worker still performs full import and formula evaluation before returning. This change avoids repeating that full pass on the interface thread; visible formulas may calculate when first rendered, and browsers without Worker support retain a synchronous fallback. No live UI, Windows runtime, or Excel GUI checks were run.
