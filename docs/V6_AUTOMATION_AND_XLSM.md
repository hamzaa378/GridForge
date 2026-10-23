# GridForge v6: action macro format and XLSM compatibility

GridForge's macro engine is an **allowlisted, declarative automation runner**. It does not parse or execute any JavaScript snippets, Visual Basic for Applications (VBA), Excel 4.0/XLM macros, Windows shell commands, or Office add-ins.

## In-app workflow

Open **Automate**, enter a name, click **Start recording**, and edit/format cells in one or more worksheet tabs. Custom Sort operations are also recorded. Reopen Automate, click **Save recording**, choose the saved macro, and click **Preview effects**. Preview calculates changes on a copy of the workbook and leaves the real cells unchanged. **Run selected macro** applies the validated plan. Exporting or importing JSON lets you reuse a macro on a compatible GridForge workbook.

The v2 format also supports commands not yet captured automatically by every UI control. Write them in JSON and import:

```json
{
  "version": 2,
  "name": "Recalculate and prepare data",
  "sheetName": "Sheet 1",
  "steps": [
    {"kind": "cells", "edits": [{"ref": "B2", "raw": "5"}]},
    {"kind": "fillDown", "source": "C2", "range": "C3:C10"},
    {"kind": "copy", "sourceRange": "A1:C10", "destination": "F1"},
    {"kind": "sort", "range": "A1:C10", "byColumn": 1, "hasHeader": true},
    {"kind": "replace", "sheetName": "Summary", "find": "Draft", "replace": "Final"},
    {"kind": "styles", "edits": [{"ref": "A1", "style": {"bold": true}}]}
  ]
}
```

`byColumn` is the **zero-based column index**: A=0, B=1, C=2. Sheet names must already exist. Examples are in `examples/v6_safe_multisheet_macro.json` and the tested v2 fixture in `tests/macros_v6.test.mjs`.

Supported step kinds: `cells` (raw contents/formulas), `styles` (complete style snapshots), `clear` (range and optional `clearFormatting`), `fillDown` (source cell and rectangular destination, adjusting relative formulas), `copy` (rectangular source, top-left destination, optional `valuesOnly`), `sort` (rectangle, numeric zero-based key column, header and order flags), `replace` (find and replace contents, optional case sensitivity/formula inclusion). Every step may override `sheetName`.

### Safety and rollback

1. Imported JSON is parsed as data and validated against fixed operation types. Unknown kinds, scripts, invalid references, malformed styles and oversized macros are rejected.
2. Maximum 200 steps, 15,000 estimated cell touches, 5,000 individual edits per edit-list. Bounds of every destination are checked.
3. The runner clones the document, evaluates steps in order, then computes the final value/style differences. Invalid later steps leave the **original** untouched.
4. Preview returns affected sheets and a bounded diff list (up to 100 cell value changes). Running applies one undoable transaction per changed worksheet. **Cross-sheet undo is not globally linked**; undo each affected sheet separately.
5. Original v1 macros remain importable. No custom scripting, filesystem access, networking, external process execution or VBA functions are available in the action macro language.

### What about an old `.xlsm` with VBA?

- Import the file to view/edit **supported worksheet data and formulas only**. Macros are **not executed**, and unsupported workbook objects may be lost when converted.
- Use **Original XLSM** to download the exact original bytes **while the same browser session retains them**. This is a preservation convenience, not a VBA editing feature.
- Export your edited worksheet as `.xlsx`, which **excludes the VBA project**. GridForge does not write those edits back into `.xlsm`, or automatically convert VBA to JSON macros.
- To use a real `Workbook_Open`, `Worksheet_Change`, VBA `Sub`/`Function`, Excel UserForms, COM, Office scripts, add-ins or ActiveX, use Microsoft Excel or a compatible runtime rather than GridForge.
- Treat unknown macro-enabled files as untrusted. Even though GridForge never runs embedded macros, opening them later in a VBA-capable program requires normal macro-security precautions.
