# GridForge v11.10 Feature Matrix and Verification

| Area | v11.10 behavior | Verification |
|---|---|---|
| Hide/unhide | Hide active worksheet; restore hidden sheets from an Unhide list; keep one visible | Focused document tests |
| GridForge JSON | Persist each worksheet's visibility and choose a visible active sheet after reopen | Focused persistence tests |
| Native XLSX | Export hidden sheet state and active visible tab | Focused workbook XML test |
| Text to Columns | Delimited split, quote qualifiers, preview and destination overwrite prompt | Focused source tests |
| Data Validation | Range lists, numeric operators, prompts and warning/error styles | Focused source and XLSX tests |

## Compatibility limits

Sheet visibility changes in macro-preserving XLSM patching are not supported. `veryHidden` imports as hidden and exports as a normal hidden sheet. Text to Columns handles one source column. Validation sources must be same-sheet; date, text-length and custom-formula rules are unsupported.

## Verification

- `npm run check` — passed.
- `npm run build` — passed.
- Focused source tests — 16/16 passed.
- Windows desktop, installer and Windows UI tests were intentionally not run. Installer not built.
