# GridForge v12.0.161 — ODS chart-axis bounds

ODS chart import now retains a specified minimum or maximum axis bound even when the opposite bound is automatic. Previously, it restored bounds only when both were explicitly present, silently losing one-sided axis scaling from Calc charts.

The current bound is retained only when it is finite and does not conflict with the other explicit bound. Invalid or reversed pairs are ignored safely. Existing major-unit handling is unchanged.

Verification: TypeScript check, production frontend build, Python packaging checks, portable build, and Windows installer compilation. Automated tests, Windows launch, live UI, and LibreOffice GUI round-tripping were not run.
