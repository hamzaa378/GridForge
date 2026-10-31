# GridForge v12.0.110

## Named in-workbook hyperlinks

- Internal hyperlink navigation now resolves worksheet-local and workbook-scoped defined names whose formulas point to a cell or contiguous range.
- Named targets can resolve to another worksheet; navigation uses the existing visible-sheet, range selection, and viewport handling.
- Formula-generated `HYPERLINK` locations and saved worksheet hyperlinks share the resolver. External HTTP(S)/mailto links retain their existing safe-opening path.
- Multi-area destinations and names that resolve to non-reference formulas remain unsupported.

## Verification scope

- TypeScript static check and production build only.
- Windows installer compilation uses `-SkipTests`.
- No automated tests, app launch, live UI, Windows runtime, or Excel GUI checks were run.
