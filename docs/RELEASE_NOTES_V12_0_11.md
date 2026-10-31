# GridForge 12.0.11

- XLSX export now runs in a module worker for local downloads and explicit cloud saves. The UI displays a preparing/saving status while the workbook is generated.
- The export worker reconstructs a stable workbook snapshot before writing and transfers the completed file buffer back to the interface. This moves OOXML generation and ZIP assembly off the UI thread; snapshot serialization and final download handoff still occur in the main thread.
- `npm run check`, production build, Python source compilation, installer packaging with `-SkipTests`, and archive integrity verification passed. No Windows tests, app launch, live UI, or Excel GUI checks were run.
