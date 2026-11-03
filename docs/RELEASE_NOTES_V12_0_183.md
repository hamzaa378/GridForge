# GridForge v12.0.183 — ODS chart sources and custom fonts

ODS chart import now reconstructs a bounded contiguous source range from chart-series value and domain references when the plot area does not provide `table:cell-range-address`. It accepts only a same-sheet rectangular layout with a consistent row or column orientation, avoiding a guessed range for irregular, cross-sheet, or oversized sources.

Cell styles now retain safe font family names imported from ODS and XLSX even when they are not in GridForge's preset list. The Home font box is editable and provides preset suggestions; custom input is bounded and rejects markup and control characters.

Verification for this source release: TypeScript check only. No automated tests, Windows tests, live UI, or LibreOffice/Excel GUI round-trip were run.
