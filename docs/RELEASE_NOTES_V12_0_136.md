# GridForge v12.0.136 — OpenDocument formatting

- ODS import/export now retains supported basic cell formatting: font family and size, bold, italic, underline, strikethrough, text and background colors, borders, horizontal/vertical alignment, wrapping, and indentation.
- Built-in numeric, integer, currency, and percentage formats now map to native OpenDocument number styles and import back into GridForge. Custom number-format codes remain unsupported.
- ODS date/time cells keep numeric values and receive a compatible display format. The shared Excel number-format reader now recognizes clock-time formats as well as date formats.
- ODS import/export now preserves supported row heights, column widths, and hidden row/column state.
- ODS import/export retains hidden worksheet visibility, while ensuring at least one worksheet remains visible when the source hides every sheet.
- Import remains bounded to supported GridForge worksheet limits. ODS formula cells are still imported from cached values; custom number formats, charts, PivotTables, validation rules, macros, and embedded objects remain unsupported.
- `npm run build` completed. The Windows installer was rebuilt with test phases skipped. No automated tests, app launch, live UI checks, or Excel/LibreOffice GUI checks were run.
