# GridForge v12.0.160 — ODS chart-label fidelity

Native ODS chart styles now preserve supported value-only, percentage-only, and combined value/percentage labels, category and series names, separators, and label placement. Chart import reads those settings back into GridForge for supported native chart types.

This does not add specialized ODF chart classes where the ODF chart model has no matching type, nor does it fully import external chart styling or linked-data refresh settings.

Verification: `npm run build` and the Windows desktop release build with `-SkipTests` passed. Automated tests, live UI, and LibreOffice GUI round-trips were not run.
