# GridForge v12.0.157 — ODS conditional fills

ODS export now carries basic numeric greater-than, greater-or-equal, less-than, less-or-equal, equal, not-equal, between, and not-between conditional fill rules using ODF conditional cell styles. The rules are attached to populated cells in the configured range. GridForge continues to preserve full rule definitions in its own metadata.

This does not yet cover blank cells, data bars, color scales, or exact multi-rule priority behavior. Conditional styles created by other ODF applications are not fully imported into GridForge.

Verification: `npm run build` and the Windows desktop release build with `-SkipTests` passed. Automated tests, live UI, and LibreOffice GUI round-trips were not run.
