# GridForge v12.0.77

## Large selection responsiveness

- Clearing cell contents now edits only populated cells inside the selection, rather than materializing an edit for each blank cell.
- Clear-format works over the selected cells' existing styles, even when the selection covers the whole worksheet.
- Applying formatting or indentation to more than 500,000 cells at once is stopped with a clear message.

## Verification

- Static TypeScript validation and installer/portable compilation are used for this release.
- No automated tests, Windows runtime tests, app launch, live UI, or Excel GUI tests were run.
