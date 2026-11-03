# GridForge v12.0.182 — Cell alignment fidelity

GridForge now supports justified horizontal alignment and bottom vertical alignment through its cell formatting model and ribbon. The settings render in the worksheet and round-trip through supported XLSX and ODS style records. ODF logical `end` alignment imports as right alignment; `justify` is retained directly.

Distributed alignment, text rotation, shrink-to-fit, text direction, and rich mixed formatting within a cell remain partial or unavailable.

Verification for this source release: TypeScript check only. No automated tests, Windows tests, live UI, or Calc/Excel GUI round-trip were run.
