# GridForge 12.0.46

`CUBESETCOUNT` now appears with GridForge's OLAP-dependent connected functions. It no longer counts comma-delimited text as if that were an Excel cube set; without an OLAP connection, it reports `#N/A` for a set expression and validates the one-argument signature.

The registered formula counts now show 506 locally evaluated or workbook-context names and 10 service-dependent names. This clarifies availability; it does not claim full Excel parity for the remaining local functions.

Verification is limited to static TypeScript checking, Python syntax compilation, and production packaging with the test phase skipped. No tests, app launch, live UI, Windows runtime, or Excel GUI checks were run.
