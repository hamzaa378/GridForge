# GridForge 12.0.6

- Draw and Page Layout now hide the unused standard command-ribbon row. Their dedicated toolbars keep a fixed 95 px footprint so the worksheet retains its space.
- `DATEVALUE` recognizes the operating-system locale's numeric date order and digits, localized long and short month names, ISO-style year-first dates, and Excel's default two-digit-year interpretation.
- The release packager refreshes installer and portable archive sizes and hashes in the release manifest after creating the portable archive, before generating source archives.
- `npm run check`, production build, installer compilation, and archive integrity checks passed. Automated tests, application launch, live UI, Windows runtime, and Excel GUI checks were not run.
- Locale parsing uses the current runtime locale; it does not yet support choosing an independent workbook locale. Excel date/time edge cases and full locale parity remain incomplete.
