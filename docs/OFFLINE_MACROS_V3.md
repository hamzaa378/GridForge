# GridForge offline macros (v10.5 installer source)

GridForge now has an Excel-independent native macro runtime built into the TypeScript spreadsheet engine. It runs offline in the Windows desktop build and in the Python-served web build. It is **not a VBA interpreter**. Imported VBA projects remain preserved by the existing XLSM editing route but cannot be executed without a compatible external runtime (the current optional bridge uses Microsoft Excel desktop).

## Run a native macro without Excel

1. Start GridForge (`START_HERE_WINDOWS.bat` with just Python, or `GridForge.exe` after Windows compilation).
2. Open **Automate**. Select **Number rows 1–20**, **Simple sales calculation**, or **Flag quantity over 100** in the template menu and click **Add template to saved macros**.
3. Select the new macro and click **Preview effects**. Preview makes *no changes*.
4. Click **Run selected macro** to apply it. Use Undo to reverse changes to the active sheet.
5. Use **Export macro JSON** for a portable copy, or **Import macro JSON** for vetted scripts. Recorded v1/v2 macros continue to work.

**Safety:** No JavaScript or VBA evaluation, subprocesses, filesystem operations, network calls, or DLL execution. Macros operate only on cells and styles in the open GridForge workbook. A full dry-run validates the actions on a copied workbook before mutation. Imported macro files are untrusted and should be previewed before running. Maximum 200 top-level steps, nested depth 3, 20 repetitions, and 15,000 budgeted cell operations. A per-sheet macro run becomes one Undo transaction, not a workbook-wide unified cross-sheet Undo.

## Supported native actions

Version 2 actions: `cells`, `styles`, `clear`, `fillDown`, `copy`, `sort`, `replace`. Version 3 adds `sequence`, `ifCell` and `repeat`. Conditional predicates: `equals`, `notEquals`, `contains`, `greaterThan`, `lessThan`, `blank`, `notBlank`. Conditionals can read calculated values from spreadsheet formulas, including the registered GridForge function engine. Repeats are finite, bounded and do not evaluate dynamic scripts.

### Sequence example

```json
{
  "version": 3,
  "name": "Offline row numbers",
  "sheetName": "Sheet 1",
  "steps": [{"kind":"sequence","range":"A2:A21","start":1,"increment":1}]
}
```

### Conditional example

```json
{
  "version": 3,
  "name": "Mark high quantity",
  "sheetName": "Sheet 1",
  "steps": [{
    "kind":"ifCell", "ref":"B2", "operator":"greaterThan", "value":"100",
    "thenSteps":[{"kind":"styles","edits":[{"ref":"B2","style":{"background":"#fbd0c5","bold":true}}]}],
    "elseSteps":[{"kind":"styles","edits":[{"ref":"B2","style":{"background":"#c5eed6"}}]}]
  }]
}
```

### Loop example

```json
{"version":3,"name":"Repeated text replacement","sheetName":"Sheet 1",
 "steps":[{"kind":"repeat","times":3,"steps":[{"kind":"replace","find":"A","replace":"AA","caseSensitive":true}]}]}
```

Repeat applies the nested actions to the current workbook state at each iteration. Row-relative looping variables, file I/O, scheduled background macro tasks, worksheet creation, Excel COM objects and full VBA compatibility are *not* implemented.

## Test and build

```powershell
npm install
npm run check
npm test
python -m unittest discover -s tests -p "test_desktop*.py" -v
.\BUILD_WINDOWS_INSTALLER.bat
```

For the built application, verify a macro preview, sequence fill, Undo, JSON import/export, and persistence after reopening. The full Windows installer build has not been executed in the Linux validation environment.
