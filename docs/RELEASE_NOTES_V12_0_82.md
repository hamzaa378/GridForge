# GridForge v12.0.82

## Dynamic array operators

- Unary plus, minus, and percent operators now calculate over array results and spill them into cells.
- Arithmetic, comparison, and text-concatenation operators calculate element-wise across matching array dimensions. A scalar operand can be applied to an array.
- Array operator inputs and results are bounded by the existing 20,000-cell spill limit.
- Mismatched array dimensions return `#VALUE!` instead of silently using only the first cell.

## Verification

- TypeScript static validation and release compilation are used with `-SkipTests`.
- No app launch, live UI, Windows runtime, or Excel GUI tests were run.
