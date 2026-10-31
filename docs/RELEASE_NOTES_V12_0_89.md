# GridForge v12.0.89

## Frozen-pane XLSX metadata

- XLSX export writes the active pane for column-only, row-only, and combined frozen panes.
- The pane metadata continues to include split counts, top-left unfrozen cell, and `state="frozen"`.

## Verification

- TypeScript static validation and Windows release compilation use `-SkipTests`.
- No app launch, live UI, Windows runtime, or Excel GUI tests were run.
