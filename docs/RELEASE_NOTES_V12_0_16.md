# GridForge 12.0.16

The CSV import worker now applies cell edits in bounded batches instead of keeping a second array containing every imported cell edit. Formula addresses are retained separately and evaluated after the values are installed. This reduces peak memory for larger CSV imports while preserving the limits and clipping behavior introduced in v12.0.14.

The worker still parses CSV rows into memory before building the worksheet, so file size and cell-count limits remain in effect.

Verification: static TypeScript check, Python syntax compilation, production build, and Windows packaging with `-SkipTests`. No application launch, live UI test, or Excel GUI interoperability check was run.
