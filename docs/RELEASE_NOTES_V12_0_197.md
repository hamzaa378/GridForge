# GridForge v12.0.197 — Formula-based conditional formatting

The conditional-formatting editor now includes a formula-specific input for formula rules. Formulas are interpreted relative to the first cell in the selected range; mixed and absolute references shift correctly for each tested cell. Formula results are evaluated without changing worksheet contents, and errors do not apply the format.

XLSX export writes native expression rules with their formatting style, and import restores the formula, range, and color. Import now reads every rule in a shared conditional-formatting block instead of dropping all but the first rule. ODS also writes and imports native ODF `is-true-formula` cell-style maps. In ODS export, this style-map approach currently covers populated cells only; blank-cell coverage remains incomplete. Formula-rule priority and the full conditional-format formula feature set remain incomplete.

Verification: TypeScript build, focused feature regressions, full code-only regression suite (1,083 tests), 10,000-formula benchmark, and 10,000-cell Data Studio stress/undo benchmark passed. No live UI or LibreOffice GUI tests were run.
