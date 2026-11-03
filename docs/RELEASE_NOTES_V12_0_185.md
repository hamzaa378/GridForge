# GridForge v12.0.185 — ODS chart analysis options

Native ODS charts now export and import per-series fixed, percentage, standard-deviation, and standard-error bars, with supported plus/minus direction and x/y dimension. They also round-trip linear, exponential, logarithmic, polynomial, and moving-average regression curves, including polynomial order, moving-average period, equation, and R-squared display options.

The mapping uses OpenDocument chart error indicators and regression-curve style properties. Error-cap appearance and forward/backward forecast extensions are not represented by this mapping; power trendlines remain in GridForge metadata and are not emitted as native ODF curves. ODF support in a given Calc version still needs a real application round-trip.

Verification for this source release: TypeScript check, production frontend build, Windows portable and installer packaging with tests skipped, and source archive packaging. No automated tests, live UI, Windows runtime, or LibreOffice GUI round-trip were run.
