# GridForge v12.0.148 — OpenDocument number-format sections

- ODS import and export now preserve common positive, negative, and zero sections in custom number formats.
- Negative and zero branches can use supported numeric patterns or literal-only output; intentionally blank sections remain blank in GridForge's worksheet display.
- ODS now carries print areas, paper size, orientation, margins, and static left/center/right headers and footers through its page-layout and master-page styles.
- ODS page layouts also preserve fit-to-page width/height, page order, and print gridline and heading options.
- Repeating print titles now round-trip for contiguous leading rows (starting at row 1) and columns (starting at column A).
- Unsupported color and conditional sections remain out of the ODS interchange subset. Existing bounded filters, validations, hyperlinks, diagonal borders, and supported formulas are unchanged.
- Non-leading repeat-title ranges, dynamic header fields, and centered-page print options still need ODS-specific support; ODS export reports non-leading ranges instead of dropping them silently.
- Static build and Windows installer packaging skip tests. No LibreOffice GUI or live UI verification was performed.
