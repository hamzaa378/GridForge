# GridForge v12.0.84

## Chart accessibility descriptions

- Add an optional accessibility description of up to 500 characters in the chart editor.
- SVG charts expose the chart title and description to assistive technology.
- XLSX import/export preserves the description in worksheet drawing nonvisual properties.

## Verification

- TypeScript static validation and Windows release compilation use `-SkipTests`.
- No app launch, live UI, Windows runtime, or Excel GUI tests were run.
