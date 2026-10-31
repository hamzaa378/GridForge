# GridForge 12.0.44

Public JSON data queries can now select a nested array before normalization. Enter an optional JSON Pointer path such as `/data/items`; path syntax and depth are bounded, and array traversal accepts numeric indices. The selected content must still be a flat array of rows or a flat array of objects. The refresh keeps the existing HTTPS-only, no-credentials, no-redirect, timeout, and response-size protections.

Verification is limited to static TypeScript checking, Python syntax compilation, and production packaging with the test phase skipped. No tests, app launch, live UI, Windows runtime, or Excel GUI checks were run.
