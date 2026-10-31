# GridForge 12.0.38

Formula-based PIVOTBY now supports all five documented `relative_to` bases: column total, row total, grand total, parent column total, and parent row total. Parent totals follow tuple prefixes when multiple row or column fields are provided. Nested subtotal output, total-depth values beyond grand totals, and vector-of-LAMBDA aggregations remain unsupported.

Verification was limited to static TypeScript checking and production packaging with the test phase skipped. No tests, app launch, live UI, Windows runtime, or Excel GUI checks were run.
