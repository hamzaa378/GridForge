# GridForge v11.8 Feature Matrix and Verification

| Area | v11.8 behavior | Verification |
|---|---|---|
| Dynamic dropdowns | Source list values from a live rectangular range on the current worksheet | Focused source tests for normalization, lookup and allowed values |
| Numeric criteria | Whole number / decimal; between, not-between, equal, not-equal, greater-than, greater-than-or-equal, less-than, less-than-or-equal | Focused operator and XLSX metadata tests |
| Input prompts | Optional title and message attached to selected validated cells | TypeScript check and JSON persistence test |
| Error alerts | Stop, warning/confirm, and information/allow | Source path plus settings persistence test |
| XLSX and JSON | Supported rules save in GridForge JSON and write native XLSX validation XML | Focused source tests |

## Compatibility limits

Range-sourced lists must be on the current worksheet. Inline lists are limited to 50 values / 255 characters. Date, text-length, custom-formula and cross-sheet validation sources are not included. Macro-preserving XLSM edits do not update validation metadata. GridForge remains standalone and does not claim full Excel parity.

## Verification

- `npm run check` — passed.
- `npm run build` — passed.
- `node --test tests/data_validation_v117.test.mjs` — 9/9 passed.
- Windows desktop, installer and Windows UI tests were intentionally not run. Installer not built.
