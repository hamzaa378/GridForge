# GridForge v12.0.134 — Capacity, export, and safer cloud sign-in guidance

- Increased the sparse worksheet canvas to 250,000 rows × 2,048 columns; populated/style storage is capped at 2,000,000 cells per worksheet, and XLSX imports at 1,500,000 cell records per worksheet.
- Added TSV export for the active sheet's calculated values.
- Added printable HTML export for the active sheet, respecting its configured print area and layout. HTML output uses the print renderer's 12,000-cell limit; use system Print for PDF or physical printers.
- Clarified that cloud passwords and MFA codes are entered on the cloud provider's official sign-in page in the system browser. GridForge receives the OAuth callback automatically; users do not paste an authorization code. User-configured provider app IDs are still required.
- Updated compatibility and cloud setup documentation. No automated tests, live UI checks, or Excel GUI checks were run.
