# GridForge v3.0 — architecture

## High-level design

```text
Browser UI (src/main.ts)
 ├─ Virtualized grid / selection / formula bar / right-click / keyboard
 ├─ Workbook toolbar: import/export / sort / find / filters
 ├─ Format ribbon: styling / AutoSum / freeze / conditional rules
 ├─ Side panels: diagnostics, formula AST, dependency impact, What-if, analytics
 └─ localStorage: JSON v3 snapshots (v2 compatible)
         |
         v
SpreadsheetDocument (document.ts)
 ├─ 1–12 named SheetEntry objects
 ├─ Workbook per sheet (sparse input + styles + local dependency graph)
 └─ SheetTools per sheet (freeze/filter/conditional metadata)
         |
         v
Workbook (workbook.ts)
 ├─ Map<address,raw>, Map<address,style>
 ├─ lexer/Pratt parser AST from parser.ts
 ├─ 56 base functions + 80 more in functions_extra.ts
 ├─ forward/reverse dependency graph, dirty invalidation, cache
 ├─ errors: diagnostics with code, cursor position and suggestion
 └─ grouped undo/redo transaction history
         |
   ┌─────┴──────────────────────────────────────────────────┐
   v                                                        v
xlsx.ts + zip.ts                                     operations.ts / clipboard.ts
 OOXML ZIP import/export                              table operations, relative shifts
 stored ZIP writer & DEFLATE reader                    undoable sorting/cut-paste
```

## Formula engine (required task-specific core)

`parser.ts` tokenizes numeric literals, strings, references, ranges, operators and names. A hand-built Pratt parser applies precedence for arithmetic, concatenation and comparisons and recognizes nested function calls. Each formula cell stores its AST, lexical diagnostics and all local referenced cells. `workbook.ts` maintains reverse edges to calculate exactly the transitive affected set after an edit, invalidates just those cached results and recomputes affected populated cells. Evaluator detects direct/indirect cycles with a recursion `visiting` path and reports a descriptive error rather than hanging.

**Crucially**: XLSX input/output is a file format adapter, not a formula evaluator. No SheetJS, ExcelJS or other formula engine is used. The exported `.xlsx` formula content is produced from raw formula text and Excel is requested to recalculate on load.

## Grid virtualization, freeze and filters

The grid addresses 400×52 by default, up to 1000×104 on import. `main.ts` computes a filtered view-to-data-row mapping for the active SheetTools filter. Only viewport rows/columns with overscan plus frozen top-row/left-column elements are rendered. Clicks, double-clicks and context-menu coordinates map screen positions to the actual row and column; underlying data is not reordered or deleted by a filter. Frozen cells are drawn at viewport-relative offsets with higher stacking order. This is currently restricted to one frozen row/column.

## Sheet metadata and persistence

`SpreadsheetDocument.toJSON()` now emits version 3 and a `tools` object for every sheet. `fromJSON` continues to accept v2 documents and fills missing metadata with safe defaults. All user-origin JSON is validated for sheet names, dimensions, raw strings, color rules, max rule count and filter references. `localStorage` retains the same storage key as v2 to preserve data on the same origin. Browser data is **not** synchronized between different ports or hosts.

## XLSX technical detail

- `zip.ts` constructs valid ZIP local file headers, central-directory entries and end record, uses CRC32 checksums, reads method 0 (stored) and method 8 (raw DEFLATE). Checks size/path limits and CRC on XML entries. Browser DecompressionStream is required for imported compressed XLSX.
- `xlsx.ts` serializes XML 2006 SpreadsheetML worksheet/workbook/style/relationship parts, formula nodes, numeric, boolean, text cells, supported fonts, fills, alignment and numeric formats, frozen panes, basic filter ranges and differential conditional styles.
- On import, it reads workbook relationship targets, shared strings, inline text, numeric/boolean values, raw formulas and shared-formula reference shifts, style indexes and supported metadata, and reconstructs sparse Workbook objects.
- An import report lists unsupported/invalid formulas and compatibility warnings; non-lossless features are never marketed as lossless.

## Test and performance strategy

Node built-in runner (`tests/`) verifies engine functions, parser/cycles, sorting/transactions, archive integrity, OOXML export structure and sheet metadata. The optional Chromium regression exercises real mouse/keyboard editing, dialog selection, frozen cells, filters, conditional rules, AutoSum and compressed XLSX round-trip. `scripts/benchmark.mjs` builds 10,000 formulas across rows and measures an isolated one-input edit's propagated set. See `docs/PERFORMANCE_REPORT.md` and `docs/V3_XLSX_IO.md`.

---

## v5 extensions: categories, bounded formulas, and safe automation

GridForge v5 registers 80 new scalar/range functions in `src/engine/functions_v5.ts`, bringing the shared registry to 216. `src/engine/categories.ts` assigns every registered function to one Excel-style sidebar category and exposes a Recently added collection. Calculation still routes through the custom `Workbook.evaluate` logic, including lazy behavior for `CHOOSE`, `ISERROR`, `ISERR`, and `ISNA`.

The application uses two distinct macro concepts:

1. **GridForge action recipes**, in `src/engine/macros.ts`: constrained JSON `cells` and `styles` steps; validated and replayed through `Workbook.applyTransaction`, with no arbitrary code evaluation or external network access. The UI records only explicitly supported actions.
2. **Imported `.xlsm` files**: the normal OOXML worksheet reader loads supported data while ignoring binary VBA parts. `src/main.ts` retains the original uploaded file bytes in memory for unchanged download. The `.xlsx` exporter builds a fresh document with no VBA parts. There is no VBA execution or merger into edited macro-enabled files.

These deliberate boundaries maintain a browser-only architecture and allow reviewers to inspect macro behavior without permitting arbitrary uploaded code to run.
