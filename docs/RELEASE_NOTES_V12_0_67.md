# GridForge v12.0.67

## Page Layout and Excel workbook compatibility

- Page Setup now offers Excel's two page-numbering orders: **Down, then over** and **Over, then down**.
- The choice is stored in GridForge documents and read from/written to the XLSX worksheet `pageSetup` element.
- Existing workbooks without a saved order continue to use Excel's default **Down, then over** setting.
- Browser printing still uses browser pagination; the stored Excel order, repeated print columns, headers/footers, and manual page breaks are not rendered there.

## Verification

- TypeScript check and Windows installer/portable compilation are recorded in the release manifest.
- No automated tests, Windows runtime tests, app launch, live UI, or Excel GUI tests were run.
