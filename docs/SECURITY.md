> **v7 update (2026-10-21):** GridForge can now save edited `.xlsm` files with supported cell changes and preserve the original VBA project. Real VBA execution is available only through an opt-in local Windows bridge using Microsoft Excel desktop and Excel Trust Center settings, NOT in the browser. The older notes below describe the v6 behavior. Read [V7_NATIVE_XLSM_AND_VBA.md](V7_NATIVE_XLSM_AND_VBA.md) for current instructions and verified limits. Live Excel COM execution has not been tested in the Linux authoring environment.

# GridForge Security Model

GridForge is a local, static frontend; there is no application backend, login, or remote execution endpoint.

- **Imported `.xlsm` VBA is not executed, parsed as executable code, interpreted, or passed to JavaScript evaluation.** Only supported OOXML worksheet parts are read by the importer.
- The unchanged original XLSM is kept **in memory in the open browser page** for byte-identical download. It is not uploaded to a server and is not persisted as a macro-enabled source file. Reloading the page or importing a new workbook clears the option.
- **Edited `.xlsx` files do not include VBA**. This is not equivalent to an edited `.xlsm` with working Excel macros.
- **GridForge action macros contain only validated JSON cell/formatting operations**. No eval, function constructors, shell access or network actions. A macro is constrained by size and address limits and applied through the workbook engine.
- The XLSX ZIP parser validates entry checksums, maximum size, path traversal, supported compression and valid XML. DTD/ENTITY constructs are rejected. Unsupported binary parts are ignored by the importer.
- The browser application sanitizes user-visible HTML. Malicious spreadsheet formulas could still be computationally expensive, so only open trusted spreadsheets and back up existing work.
- Excel macro execution is controlled by the user's desktop Excel installation and its security policies, not by GridForge. Do **not** enable macros from unknown sources merely because a file was opened in GridForge.
- This is a student project, not a security-certified document sandbox. Avoid processing classified or untrusted files and do not assume parity with Excel Protected View.

## v6 declarative macro engine

Action-macro JSON v2 is a fixed allowlisted command schema, not a general-purpose script executor. The validator rejects JavaScript, VBA, arbitrary commands, out-of-bounds ranges, invalid styles, unknown action kinds and excessive batch sizes. `runActionMacro(..., {dryRun:true})` previews changes on a cloned document; execution also preflights the full plan on a clone. This is intentionally separate from preserving unchanged XLSM original bytes. There is no VBA evaluation, ActiveX, shell access, or connection to outside services. Cross-sheet action-macro transactions are undoable **per sheet**, not as a single global undo.
