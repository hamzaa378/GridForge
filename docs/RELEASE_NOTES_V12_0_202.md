# GridForge v12.0.202 — Reversed ODS icon sets

Conditional formatting now exposes **Reverse icon order** when creating icon-set rules. The option was already supported by the workbook model, renderer, and XLSX path; it now also writes and reads LibreOffice Calc's `calcext:reverse` attribute so ODS applications can retain the chosen order.

Regression coverage strips GridForge-only metadata to exercise ODS interoperability, then checks native export and import of the reversed setting. No LibreOffice GUI round-trip or live UI test was run.
