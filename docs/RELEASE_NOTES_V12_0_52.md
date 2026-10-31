# GridForge v12.0.52

- Semicolon-separated formulas now normalize decimal commas outside text literals, so common formulas such as `=SUM(1,5;2,5)` can be parsed as 1.5 plus 2.5.
- XLSX export writes the normalized decimal point and invariant comma list separators while preserving quoted text.
- Localized function names, workbook-specific locale selection, and locale-aware date parsing remain incomplete.
- Static TypeScript check, production build, and installer compilation only; no automated tests or live UI checks were run.
