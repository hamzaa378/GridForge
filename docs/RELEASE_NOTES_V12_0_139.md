# GridForge v12.0.139 — LibreOffice-friendly font formatting

- The Home font picker now includes 28 common fonts, including Liberation, Carlito/Caladea, DejaVu, Noto, Lato, Open Sans, Roboto, and Ubuntu families alongside the existing Office fonts.
- Font sizes from 6 through 72 points can be selected and are preserved by supported GridForge formatting, XLSX import/export, and ODS import/export paths.
- This improves document portability when a workbook moves between GridForge, LibreOffice, and Office. Font appearance still depends on which typefaces are installed on the receiving device; font embedding and advanced typography are not supported.
- TypeScript build and Windows installer build are being refreshed with automated test phases skipped. No application launch, live UI, LibreOffice GUI, or Excel GUI checks are run.
