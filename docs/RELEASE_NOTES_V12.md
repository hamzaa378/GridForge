# GridForge 12.0.0

## A workspace users can adapt

- Added four independently selectable visual presets: Excel 2019, Modern, Compact, and High Contrast.
- Presets persist across launches and are independent from light/dark mode and Excel-style/Guided workspace selection.
- The High Contrast preset strengthens focus indicators, control borders, and selected-cell/ribbon visibility in both light and dark themes.
- The Compact preset reduces top bar, ribbon, formula bar, and sheet-tab height and slightly tightens worksheet cell padding.
- The Modern preset uses a calmer teal and blue accent palette.
- Preset selection affects only the interface. It does not change workbook content or saved formatting.

## Carried forward from 11.30

OneDrive, Google Drive, and Dropbox workbook access remains available in the Windows desktop release. Each provider requires a user-configured OAuth app. See [cloud setup](CLOUD_STORAGE_SETUP.md) for registration requirements and limitations.

## Build and verification

- `npm run check` and `npm run build` passed.
- The Windows installer and portable binaries compiled with `desktop/build_windows.ps1 -SkipTests`.
- No automated tests, Windows application launch, live UI test, or Excel GUI validation was performed.
- This release improves presentation flexibility; GridForge still does not claim full Excel feature, formula, chart, PivotTable, or XLSX fidelity.

