# GridForge v12.0.155 — ODS drawing previews

GridForge worksheet ink and drawing objects now export as a standard embedded SVG picture in the ODS package. LibreOffice Calc and other ODF applications can display the drawing as one image. GridForge also retains validated, bounded stroke definitions in its namespaced worksheet metadata, so reopening the file in GridForge restores editable strokes.

This does not make the imported SVG individually editable as pen strokes in Calc. External ODF drawings are not imported as GridForge stroke objects. SVG data is capped at 2 MB per sheet drawing image.

Verification: `npm run build` and the Windows desktop release build with `-SkipTests` passed. Automated tests, live UI, and LibreOffice GUI round-trips were not run.
