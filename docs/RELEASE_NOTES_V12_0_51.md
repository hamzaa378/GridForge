# GridForge v12.0.51

- Formula calls now accept comma or semicolon argument separators, allowing common semicolon-list locales to enter formulas without changing the evaluator.
- Calls reject mixed separators with a syntax hint. XLSX export converts semicolon argument separators to invariant comma separators while preserving commas inside text literals.
- Localized function names and decimal/date parsing remain limited.
- Static TypeScript check, production build, and installer compilation only; no automated tests or live UI checks were run.
