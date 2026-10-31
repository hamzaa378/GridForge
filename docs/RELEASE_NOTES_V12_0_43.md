# GridForge 12.0.43

Formula-based GROUPBY now supports nested row subtotals for multiple grouping fields. Positive depths place subtotals after their child groups and the grand total at the bottom; negative depths place them before their child groups and the grand total at the top. When subtotals are requested, the groups are ordered by hierarchy so prefix totals remain contiguous. A spill-size preflight rejects oversized summaries before subtotal output rows are created.

Verification is limited to static TypeScript checking, Python syntax compilation, and production packaging with the test phase skipped. No tests, app launch, live UI, Windows runtime, or Excel GUI checks were run. Direct Excel comparisons for subtotal ordering, blank/error handling, and custom LAMBDA aggregators remain outstanding.
