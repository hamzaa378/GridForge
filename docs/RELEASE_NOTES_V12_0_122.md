# GridForge v12.0.122 — Sparse range index

Large local `SUM`, `AVERAGE`, `MIN`, and `MAX` formulas now query a per-column index of populated rows. Calculations no longer allocate a new set containing every stored and spilled cell for each formula; they inspect only populated coordinates in the referenced columns and rows.

The index is maintained when cells are edited and when dynamic-array spill cells are added or cleared. Empty spill cells remain indexed as populated positions but do not contribute numeric values. The previous sparse aggregation behavior for overlapping ranges and referenced formula errors is retained.

Verification is limited to the TypeScript check and release build/package. No automated tests, desktop launch, live UI, or Excel GUI checks are performed.
