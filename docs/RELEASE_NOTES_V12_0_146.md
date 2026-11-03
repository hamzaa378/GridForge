# GridForge v12.0.146 — Diagonal cell borders

- Added diagonal up, diagonal down, and crossed borders to the Home border selector.
- Diagonal lines render in the grid and support solid, dashed, dotted, and double styles, thin/medium/thick weights, and color.
- XLSX writes and reads the OOXML diagonal direction flags and line definition. ODS writes and reads the corresponding two diagonal border properties.
- ODS validation rules and hyperlinks from v12.0.144 and v12.0.145 remain included.

OOXML defines a shared line style when both diagonal directions are enabled, so it cannot represent different styles for the slash and backslash independently. Complex table borders remain incomplete. No LibreOffice round-trip comparison or live UI check was performed.

Build: `npm run build`. Windows installer build uses `-SkipTests`; no automated tests were run.
