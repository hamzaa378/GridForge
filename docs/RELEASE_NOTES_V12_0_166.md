# GridForge v12.0.166 — ODS fractional seconds

Date/time number formats with `s.0`, `ss.00`, or `ss.000` now translate to and from OpenDocument second components with the matching decimal precision. ODS supports fractional-second formatting; this implementation is bounded to three decimal places. Unsupported finer precision remains outside the translated subset.

Verification: TypeScript check, production frontend build, Python packaging checks, portable build, and Windows installer compilation. Automated tests, Windows launch, live UI, and LibreOffice GUI round-tripping were not run.
