# GridForge v12.0.81

## Cloud operation guards

- Cloud downloads reject folder entries before requesting workbook bytes.
- Cloud overwrite operations reject folder entries and files from a different selected provider.
- Destination folders are validated against the selected provider before upload.

## Verification

- TypeScript static check passed.
- Windows packages are built with `-SkipTests`; no runtime or live UI tests were run.
