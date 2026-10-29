# GridForge v11.9 Feature Matrix and Verification

| Area | v11.9 behavior | Verification |
|---|---|---|
| Text to Columns | Split one-column selections by common/custom delimiters; preview; handle quoted fields and escaped quotes | 4 focused source tests; frontend build |
| Destination safety | Validate source/destination bounds and prompt before replacing existing destination values | TypeScript build; source boundary tests |
| Data Validation lists | Inline lists or live same-sheet range lists | Focused tests for current source values and rejection |
| Numeric rules | Whole/decimal with eight Excel comparison operators | Focused operator and XLSX XML tests |
| Prompts and alerts | Optional input/error titles and messages; Stop, Warning, Information behavior | JSON persistence and metadata tests |
| XLSX / GridForge JSON | Persist supported validation settings | Focused source tests |

## Compatibility limits

Text to Columns operates on one selected worksheet column, writes values, and does not split newline-separated values into new rows. Validation range lists are same-sheet only. Date, text-length, custom-formula and cross-sheet validations are not supported. Macro-preserving XLSM edits do not patch Data Validation records. GridForge remains standalone and does not claim full Excel parity.

## Verification

- `npm run check` — passed.
- `npm run build` — passed.
- Focused source tests — 13/13 passed.
- Windows desktop, installer and Windows UI tests were intentionally not run. Installer not built.
