# GridForge v12.0.85

## Chart accessibility

- Enter an optional accessibility description in the chart editor.
- SVG output includes a title, a description when supplied, and a matching accessible name.
- XLSX import/export preserves chart descriptions in drawing nonvisual properties; imported text is bounded to 500 characters.

## Verification

- TypeScript static validation and Windows release compilation use `-SkipTests`.
- No app launch, live UI, Windows runtime, or Excel GUI tests were run.
