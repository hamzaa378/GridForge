# GridForge v12.0.164 — ODS date and elapsed-time formats

ODS now translates common Excel date and clock format tokens into native OpenDocument styles instead of replacing them with a generic pattern. Supported patterns include numeric and named months, dates and weekday names, 12/24-hour clocks, AM/PM, and quoted or escaped literals. Elapsed `[h]`, `[m]`, and `[s]` components use OpenDocument's extended time-component range. Negative duration values now survive ODS import/export.

Locale-specific calendars, unsupported token widths, fractional seconds, and other advanced date formats remain partial. The generated ODF structures follow the OASIS 1.4 schema for date/time components and `number:truncate-on-overflow` ([ODF 1.4 schema](https://docs.oasis-open.org/office/OpenDocument/v1.4/OpenDocument-v1.4-part3-schema.pdf)).

Verification: TypeScript check, production frontend build, Python packaging checks, portable build, and Windows installer compilation. Automated tests, Windows launch, live UI, and LibreOffice GUI round-tripping were not run.
