# GridForge v12.0.190 — Calc precise rounding formulas

ODS formula import/export now maps the implemented `CEILING.PRECISE`, `FLOOR.PRECISE`, and `ISO.CEILING` functions to Calc’s documented `COM.MICROSOFT` names. Calc has supported these compatibility functions since LibreOffice 4.3. This lets Calc recalculate the formula instead of relying on its cached value. Numerical edge-case equivalence remains unverified.

Verification for this source release: TypeScript check, production frontend build, Windows portable and installer packaging with tests skipped, and source archive packaging. No automated tests, live UI, Windows runtime, or LibreOffice GUI round-trip were run.
