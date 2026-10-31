# GridForge 12.0.42

Formula-based PIVOTBY now supports nested row and column subtotals for multi-level hierarchies. Positive total depths place subtotals after their child groups and the grand total at the bottom; negative depths place them before their child groups and the grand total at the top. Subtotal sorting must preserve hierarchy order, and unsupported sort combinations return a formula error. A spill-size preflight rejects oversized summaries before subtotal output rows are materialized.

Verification is limited to static TypeScript checking, Python syntax compilation, and production packaging with the test phase skipped. No tests, app launch, live UI, Windows runtime, or Excel GUI checks were run. Direct Excel comparisons for subtotal ordering, blank/error handling, and custom LAMBDA aggregators remain outstanding.
