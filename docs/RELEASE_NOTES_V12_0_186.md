# GridForge v12.0.186 — ODS lookup and sequence formulas

ODS export now writes Calc’s documented `COM.MICROSOFT` names for the implemented `XLOOKUP`, `XMATCH`, and `SEQUENCE` functions. ODS import removes the compatibility namespace and restores these formulas to GridForge syntax. `SEQUENCE` uses the existing bounded matrix-formula/spill handling. Calc documents `XLOOKUP`, `XMATCH`, and `SEQUENCE` as compatibility functions beginning with LibreOffice 24.8; older Calc versions and formulas whose argument behavior differs still need compatibility fallback or may not recalculate.

This change improves formula interoperability; it does not claim exact calculation, locale, error, or precision parity.

Verification for this source release: TypeScript check, production frontend build, Windows portable and installer packaging with tests skipped, and source archive packaging. No automated tests, live UI, Windows runtime, or LibreOffice GUI round-trip were run.
