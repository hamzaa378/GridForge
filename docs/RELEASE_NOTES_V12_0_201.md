# GridForge v12.0.201 — Typed data-bar bounds

Conditional data bars now support automatic minimum/maximum, number, percent, percentile, and formula bounds. The editor exposes both bounds; validation rejects malformed formulas, out-of-range percentages, incompatible min/max positions, and reversed fixed numeric limits. Rendering evaluates formula and percentile bounds against the selected range and displays the configured middle or automatic zero axis with its selected color.

XLSX and native LibreOffice Calc ODS extension import/export preserve typed data-bar bounds. The rendering model now collects numeric values from the workbook's sparse populated-cell index, so large mostly-empty ranges avoid walking every blank row. The same sparse sampling path now backs color-scale and icon-set calculations.

Verification: 1,090 code-only regression tests, TypeScript check/build, full core regression stages, 10,000-formula dependency benchmark, 10,000-cell Data Studio benchmark, and full-height sparse-range coverage. No live UI or LibreOffice GUI test was run. Calc application rendering and exact data-bar axis semantics still need a real Calc round-trip.
