# GridForge v12.0.138 — OpenDocument number-format fidelity

- ODS import/export now maps common single-section custom number formats, including grouped/decimal placeholders, quoted or escaped literal prefixes and suffixes, and percentage formats.
- ODS continues to preserve built-in number/date styles, common cell formatting, merged regions, sheet visibility, and row/column dimensions. Complex multi-section, conditional, and elapsed-time custom formats remain unsupported.
- `npm run build` and the Windows installer build completed with test phases skipped; source and GitHub repository archives were generated. No automated tests, app launch, live UI checks, or LibreOffice GUI checks were run.
