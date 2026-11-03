# GridForge v12.0.198 — Conditional-rule order fidelity

XLSX conditional-formatting import and export now retain OOXML rule priorities and `stopIfTrue` flags for supported cell comparison, formula, color-scale, and data-bar rules. Imported rules are ordered by priority, and validation rejects invalid priority values or malformed stop flags. XLSX import also retains multiple rules grouped inside one conditional-formatting block.

GridForge evaluates supported formatting rules in stored order and can display one cell fill at a time; full Excel multi-style composition and stop-if-true visual semantics remain incomplete. ODS native style-map conditions do not carry the same priority metadata. No live UI or LibreOffice GUI testing was run.

The ODS formula serialization audit still classifies the 506 frozen fixtures as 404 standard OpenFormula, 84 Microsoft-namespace compatibility, 3 LibreOffice-specific, and 15 GridForge-recovery-only. These categories describe export syntax, not Calc calculation parity. See [ODS formula coverage](ODS_FORMULA_COVERAGE.md).

The saved parity audit now reports the active package version and records complete spill matrices for the four non-scalar fixtures instead of labeling them as gaps. The latest static run matches all 136 deterministic scalar references with zero differences; the four array cases are separately inspectable and are not counted as scalar parity. This does not run or claim desktop Excel comparison.

Verification: `npm run parity:gridforge`, TypeScript check/build, focused XLSX conditional-rule regressions, the full 1,083-test code-only suite, 10,000-formula benchmark, and 10,000-cell Data Studio stress/undo benchmark passed. No live UI or LibreOffice GUI testing was run.
