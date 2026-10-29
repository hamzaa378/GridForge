# GridForge v11.11 Feature Matrix

| Area | Behavior | Verification |
|---|---|---|
| Sheet order | Move active tab left/right while preserving active sheet identity | Focused document test |
| Tab color | Presets, clear, duplicate, JSON persistence and native XLSX `sheetPr/tabColor` | Focused model and OOXML tests |
| Row/column visibility | Select whole row/column from heading; hide selected rows/columns, unhide all, preserve at least one visible, persist to JSON and XLSX | Focused model and OOXML tests |
| Workbook search | Find All across visible sheets and jump to a result | Source implementation; no UI test run |
| Multi-level sort | Sort complete selected rows on up to three keys with per-key direction; macro replay retains keys | Focused sort and macro tests |
| Date validation | ISO dates and Excel serial values compared against date serial criteria | Focused source tests |
| Text-length validation | Enforce character count criteria | Focused source tests |
| XLSX validation | Export date and text-length validation records | Focused OOXML tests |

## Compatibility notes

- Find All searches visible worksheets. Replace All is still limited to the active worksheet.
- Date rules use Excel's 1900 date serial system. Text length counts the stored input string.
- GridForge JSON version 4 remains backward compatible with document versions 2 and 3.
- XLSX interoperability was checked at the XML/source level; Excel itself was not launched.
- Windows desktop, installer and UI tests were not run, per the user's instruction.
