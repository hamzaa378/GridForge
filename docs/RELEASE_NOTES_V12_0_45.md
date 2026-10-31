# GridForge 12.0.45

PIVOTBY builds row and column prefix indexes for grouped records and uses the smaller matching candidate set when calculating detail cells and subtotals. This reduces repeated full-map scans in large cross-tab summaries while retaining the existing output-size guard.

Verification is limited to static TypeScript checking, Python syntax compilation, and production packaging with the test phase skipped. No tests, app launch, live UI, Windows runtime, or Excel GUI checks were run.
