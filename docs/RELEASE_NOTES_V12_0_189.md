# GridForge v12.0.189 — Editable chart marker shapes

The chart editor now lets each series choose circle, square, diamond, triangle, plus, X, or star markers. GridForge renders the selected shape and retains it across supported XLSX and ODS chart import/export. Existing marker visibility, series colors, and error-bar/trendline settings remain supported.

The available shapes are the common subset shared by GridForge, DrawingML charts, and OpenDocument named symbols. Image markers, custom glyphs, line/outline styling, and application-specific chart effects remain outside this mapping. The ODS symbol mapping follows the ODF chart symbol-name property.

Verification for this source release: TypeScript check, production frontend build, Windows portable and installer packaging with tests skipped, and source archive packaging. No automated tests, live UI, Windows runtime, or LibreOffice GUI round-trip were run.
