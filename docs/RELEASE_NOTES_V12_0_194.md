# GridForge v12.0.194 — ODS horizontal chart axis fidelity

ODS horizontal bar charts now attach category reversal and label rotation to the Y category axis. Value-axis bounds and number formatting remain attached to the X axis. Previously the exporter could place these settings on the wrong axis or omit the category-axis style.

This improves OpenDocument chart interoperability; rendering differences between GridForge and Calc remain possible and no Calc GUI round-trip was performed.

Verification: TypeScript check and production build passed. Windows installer and portable packaging with `-SkipTests` completed successfully. No automated tests, live UI, or LibreOffice GUI checks were run.

