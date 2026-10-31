# GridForge v12.0.60

- Formula parsing accepts the Excel `#DATA!` error literal.
- XLSX formula errors retain valid Excel OOXML error codes; GridForge-only diagnostic codes such as `#PARSE!` and `#CYCLE!` export as cached `#VALUE!` instead of nonstandard error strings.
- Static TypeScript check and production build passed. Windows installer and portable packages compiled with `-SkipTests`; no tests, app launch, live UI, or Excel GUI checks were run.
