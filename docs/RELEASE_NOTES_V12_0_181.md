# GridForge v12.0.181 — ODS array formula round-trips

ODS now exports supported dynamic-array formulas with OpenDocument matrix row/column spans, writes the calculated spill values as cached cells, and retains formatting applied to individual spill cells. On import, cells covered by an importable matrix formula are not stored as independent values, so supported formulas can recalculate and spill again in GridForge. Matrix handling is bounded by GridForge’s existing 20,000-cell spill limit.

This builds on v12.0.180 Calc-compatible formula names for SORT, UNIQUE, and TEXTSPLIT. Calc release support varies by function; formula semantics and array behavior still need cross-application comparison. Complex ODS charts, macros, PivotTables, and unrelated embedded objects remain partial or unsupported.

Verification for this source release: TypeScript check and production build only. No automated tests, Windows tests, live UI, or LibreOffice GUI round-trip were run.
