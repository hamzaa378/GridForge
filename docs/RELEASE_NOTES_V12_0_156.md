# GridForge v12.0.156 — ODS drawing and chart fidelity

- ODS exports worksheet drawings as a standard embedded SVG picture. ODF applications can display the image, and GridForge retains bounded editable stroke definitions in its own metadata.
- ODS chart import restores supported value, category, and series labels, along with major and minor value-axis gridline visibility.

SVG drawing previews are capped at 2 MB per worksheet. Other ODF applications see a flattened picture; external ODF vector objects are not imported as editable GridForge strokes. Chart feature coverage still depends on the ODF chart type and modeled settings.

Verification: `npm run build` and the Windows desktop release build with `-SkipTests` passed. Automated tests, live UI, and LibreOffice GUI round-trips were not run.
