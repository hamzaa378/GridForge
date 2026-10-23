# GridForge v7.0 — native XLSM updates and optional executable Excel VBA

## Capability overview

**Two distinct paths:**

1. **Native Save XLSM (no Excel needed):** Import a macro-enabled Excel file; change supported cell contents and formulas; click **Save XLSM**. GridForge patches the original OOXML package, preserving each unmodified binary part, including `xl/vbaProject.bin`, relationships, graphics, chart parts, and other unsupported OOXML parts instead of rebuilding a macro-free `.xlsx`. It flags recalculation so Excel can refresh formulas. No VBA executes during this process.
2. **Execute an existing VBA Sub (Windows + Microsoft Excel desktop):** Start the optional token-protected Python bridge. In GridForge **Automate**, enter its session token and an existing procedure such as `Module1.GenerateReport`. After confirmation, GridForge sends the edited XLSM over local loopback to the bridge. Microsoft Excel opens it using the user's Trust Center settings, invokes the selected no-argument VBA macro, saves it, and returns a new `.xlsm`. GridForge downloads the result and attempts to re-import its supported worksheet contents.

**Not a VBA interpreter:** The browser does not parse, edit, compile, emulate, or execute VBA. Native execution uses the user's installed Microsoft Excel application. This is the correct path for authentic Excel VBA rather than a misleading JavaScript approximation.

## Python-only path: import, modify, export XLSM

1. Have Python 3.10+ and a modern Chromium/Firefox/Edge browser.
2. Extract the repository and double-click `START_HERE_WINDOWS.bat` or run `python run_gridforge.py`.
3. Open the displayed local URL, usually `http://127.0.0.1:4174`.
4. Click **Import** and choose a `.xlsm` from a source you trust. The import report explains unsupported data and formula evaluation.
5. Edit values or formula text in existing worksheets. The formula engine stays GridForge's own 306-function TypeScript implementation.
6. Click **Save XLSM** to download a newly named `*_edited.xlsm` with supported content changes and the original VBA project preserved.
7. Open a COPY of the new file with Excel and verify values, macros, calculated formulas, and formatting before using it for work.

You can still use **Original XLSM** to recover the byte-identical upload, **XLSX** to export supported content without any VBA, or **JSON** for GridForge-native backups.

### Preservation guarantees and limitations

- XLSM source must be a valid OOXML ZIP archive, at most **25 MB compressed**, with at most 3,000 parts and an uncompressed total of at most 80 MB. The round-trip patch modifies **up to 5,000 changed cell values/formulas**. Maximum imported structure: 12 sheets, 1,000 rows × 104 columns per sheet, and 50,000 imported cells per sheet.
- Existing `xl/vbaProject.bin` and unknown Office/media parts are copied unchanged as binary ZIP entry payloads. The ZIP wrapper itself is reconstructed, so the **overall archive bytes differ**. Unmodified parts retain their original payloads, not their original compressed ZIP bytes.
- XLSM export preserves existing worksheet XML structures for untouched cells, formulas, relationships, charts, shapes, named ranges, images, links and formatting. This **does not guarantee every advanced Excel feature will function correctly** after changing dependent cell contents; verify in desktop Excel.
- The XLSM patch currently merges **raw cell contents and supported formula changes only**. GridForge formatting changes are warned about and **not merged** into XLSM. Existing original Excel formatting is preserved.
- Adding, deleting, reordering or renaming worksheets is blocked for Save XLSM to prevent silently corrupting references or workbook metadata. Use XLSX export for such structure changes.
- Editing Excel shared, array and data-table formula cells is refused. Other advanced formulas and macros may refer to data or external resources GridForge cannot evaluate.
- Excel's cached results might be stale until Excel opens and performs full recalculation. Spreadsheet formula compatibility is not equivalent to Microsoft Excel's calculation engine.
- Document-level digital signatures and trust indicators may be invalidated by changing worksheet XML, even though the embedded VBA binary is preserved. Signed documents should be edited/saved and re-signed using Excel.
- XLSM edits are not automatically merged into the original file on disk: the browser downloads a **new** `.xlsm`. This is intentional to protect source files.

## Optional: execute authentic VBA using installed desktop Excel

### Prerequisites

- **Windows 10/11**, Python 3.10 or later, Microsoft Excel **desktop** installed and licensed. Excel Online, Linux browsers alone, and a Python-only runtime cannot execute actual VBA.
- Install one optional dependency for the bridge in PowerShell:

```powershell
py -3 -m pip install pywin32
```

No npm modules or `pywin32` are needed for the Python-only browser app or the native XLSM file-save path.

### Start GridForge and the Excel bridge

In **PowerShell window 1**, in the extracted project root:

```powershell
py -3 run_gridforge.py
```

Open the local URL printed there, often `http://127.0.0.1:4174`.

In **PowerShell window 2**, same folder:

```powershell
py -3 scripts/excel_vba_bridge.py
```

The bridge binds **only** to `127.0.0.1:4187` and prints a random session token. Copy that token. Do not share it, open your bridge port to the internet, or run it with administrator privileges.

### In GridForge

1. **Import** your trusted macro-enabled `.xlsm` and check the import report.
2. Make supported cell edits if needed.
3. Open **Automate**. Enter the token in **Local bridge session token**.
4. Enter the exact name of an existing **no-argument VBA Sub**, for example `Module1.GenerateReport` or `CalculateReport`. The bridge accepts normal procedure identifiers, not arbitrary script expressions or parameters.
5. Click **Run selected VBA in desktop Excel** and read the confirmation carefully. The workbook may contain dangerous executable code.
6. Excel opens visibly and runs the macro **only if its security settings permit it**. GridForge does not bypass Trust Center. Excel can display security warnings; if policy blocks the macro, the bridge returns an error.
7. After the macro completes, the bridge saves the workbook and GridForge downloads `*_vba_result.xlsm`. If supported, its worksheets are reimported into GridForge. If the macro created unsupported workbook structures, the downloaded result is still available; open it in Excel.
8. Stop the bridge with Ctrl+C when finished.

### Important limitations of execution

- **No unattended execution**: the bridge is never started automatically; its random token must be provided manually, and the user confirms every run.
- The bridge honours Excel's Trust Center (`AutomationSecurity=ByUI`); it **does not disable macro protection**. Auto-event macros are disabled while the bridge opens the workbook (`EnableEvents=False`), so code relying on `Workbook_Open` / `Worksheet_Change` events may behave differently. Explicitly selected VBA procedures can run.
- VBA may access the filesystem, execute shell commands, connect to networks, or make arbitrary changes with the user's privileges. Only run personally trusted macros on backed-up workbooks. The token and confirmation **are not a sandbox**.
- Macros requiring interactive UserForms, add-ins, network shares, external references, specific workbook paths, or custom Excel security policies may not run through this bridge. Open these workbooks directly in desktop Excel.
- Only one bridge request executes at a time. Never expose its loopback port through a proxy or public tunnel.
- Excel might keep the workbook open during a modal trust/security prompt. Resolve such prompts manually in Excel; the client has no built-in execution timeout.
- **This release has not executed a real VBA macro under Microsoft Excel in the Linux test environment.** This step must be checked on your Windows installation before claiming end-to-end VBA validation.

## Windows acceptance test for an actual VBA procedure

Use a new throwaway Excel workbook, save as **Excel Macro-Enabled Workbook (.xlsm)**, insert the VBA procedure below under `Module1`, and confirm it runs manually in Excel before testing the bridge.

```vb
Public Sub GridForgeSmokeTest()
    ThisWorkbook.Worksheets(1).Range("J12").Value = "VBA_EXECUTED"
End Sub
```

Close the workbook in Excel, import it into GridForge, run `Module1.GridForgeSmokeTest` through the bridge, and open the downloaded `*_vba_result.xlsm` in desktop Excel. Verify **J12 contains `VBA_EXECUTED`** and the VBA module still exists. Verify GridForge also displays J12 after re-import. Keep macro security enabled and do not use untrusted workbooks.

## Automated tests provided

- `npm run regression` (developer machine with Node and Playwright): 340 unit tests, 10,000-formula benchmark, existing browser interactions, XLSM binary round-trip, and VBA bridge protocol security.
- `python scripts/test_vba_bridge_protocol.py` (Python only): rejects invalid tokens, non-local origins, unsafe macro names, invalid requests and malformed workbooks. **Does not run Excel**.
- `python scripts/test_xlsm_browser.py` (requires Playwright + Chromium): synthesizes a macro-enabled OOXML package with a placeholder VBA binary, edits cells, exports XLSM, and checks unchanged project/media payloads, recomputation flags and failure on sheet-structure changes. **The placeholder binary is not a valid executable VBA project**, so it does not prove live Excel compatibility.
- Full browser-storage reload and true Microsoft Excel macro execution must be verified on Windows.
