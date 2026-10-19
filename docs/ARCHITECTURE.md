# Architecture and Technical Decisions

## The problem we are solving

A spreadsheet is not just a table. It is a dependency graph with a language embedded in its cells. Correctness means that `=A1+B1` must be parsed consistently, dependencies must be updated when formulas change, cycles must not crash the program, and editing one cell must not recalculate unrelated cells. The UI must show a large address space without putting every element in the DOM.

GridForge is organized around **four foundational engine/UI concerns**, supplemented by three v2 modules:

1. `parser.ts` — turn formula source code into a typed abstract syntax tree (AST).
2. `functions.ts` — pure, reusable spreadsheet functions with explicit error behavior.
3. `workbook.ts` — orchestrate evaluation, persistence-ready raw storage, graph tracking, history, and what-if snapshots.
4. `main.ts` — render a virtual window into the workbook and translate UI actions into `applyEdits` calls.

**New v2 modules:** `document.ts` (worksheet container, names, sheet operations, v1 migration), `format.ts` (sparse cell-style model, validation, number display), and `operations.ts` (range sorting, literal find/replace, selection analytics and chart data). Supporting utilities: `csv.ts` for CSV parsing/serialization, `styles.css` for presentation, and `scripts/serve.mjs` for a dependency-free local web server.

### V2 document and style data flow

```text
SpreadsheetDocument { activeId, sheets[] }
                   │
         active sheet's Workbook ─── CellStyle map
                   │                   │
        Parser / evaluator        styledDisplay()
                   │                   │
       dependency / cache         DOM presentation
                   │
      grid virtualization / inspector / impact / what-if
```

Worksheets are **independent** formula domains. No `OtherSheet!A1` syntax is supported. `SpreadsheetDocument` serializes all sheet contents and styles under a v2 document shape. `Workbook.fromJSON` also accepts legacy single-sheet v1 records. Formatting updates use `applyStyles` or `patchStyles` and don't invalidate dependency caches. `applyTransaction` groups values and formats into one undo step for same-app paste.

**Sort behavior:** `sortSelection` moves entire selected rows and translates formulas on moved rows via `shiftFormula`; it is intentionally simpler than Microsoft's formula/row rewriting rules. Find/replace is literal and defaults to excluding formula sources. The Data Quality Lens iterates selected cells and plots up to 16 points from the first two selected columns. These tools operate outside the formula grammar, so the handwritten parser remains the only calculator.

## Formula processing pipeline

Example: `=SUM(A1:A3)+IF(B1>10,B1,0)`

```text
Formula input
    |
    v
Tokenizer (tokens + source offsets)
    |   SUM ( A1 : A3 ) + IF ( B1 > 10 , B1 , 0 )
    v
Pratt expression parser
    |
    v
Binary (+)
    ├── Function SUM
    |     └── Range A1:A3
    └── Function IF
          ├── Operator > (B1, 10)
          ├── Reference B1
          └── Value 0
    |
    v
Reference collector {A1,A2,A3,B1}
    |
    v
Dependency graph + reverse graph
    |
    v
Evaluate with cache and recursion-stack cycle checks
    |
    v
Display a scalar result or a structured diagnostic
```

### 1. Tokenization

`tokenize()` walks the input from left to right. Each token includes `kind`, `text`, `start`, and `end`. Categories include numeric literals (scientific notation supported), double-quoted strings (escaping as `""`), `A1`/`$A$1` references, identifiers, operators, commas, colons, parentheses, and EOF.

Token positions are **zero-based offsets relative to the formula body after `=`**. The sidebar adds one character to align its caret with the visible leading `=`.

### 2. Custom parser

`Parser.expr(minPrec)` implements a Pratt parser. Precedence from low to high:

| Level | Operators | Behavior |
| --- | --- | --- |
| 1 | `=`, `<>`, `!=`, `<`, `>`, `<=`, `>=` | Comparisons |
| 2 | `&` | Text concatenation |
| 3 | `+`, `-` | Add/subtract |
| 4 | `*`, `/` | Multiply/divide |
| 5 | Unary `+`, `-` | Unary sign |
| 6 | `^` | Exponentiation, right associative |

`=2+3*4` yields 14, `=2^3^2` yields 512, and `=-2^2` yields −4. Parentheses override precedence; function calls are nested naturally. Ranges are explicit AST nodes rather than string expressions.

If the parser reaches EOF inside an open function call, it raises `FormulaSyntaxError` with a source offset and correction suggestion.

### 3. Dependency graphs

`Workbook` maintains two directed adjacency maps:

- `deps[formulaCell] → Set<referencedCell>`: which inputs are read by a formula.
- `rev[inputCell] → Set<dependentCell>`: which formulas may need recalculation when an input changes.

When a formula is replaced, the workbook removes old reverse edges, parses new references, installs new edges, and walks `rev` from the changed cells to find **only reachable downstream cells**. These caches are invalidated. References into ranges are expanded as graph edges, with a 20,000-reference cap per range.

Example:

```text
A1 --> B1 --> C1
D1 --> E1
```

Changing A1 invalidates only A1, B1, C1. D1 and E1 keep cached values. The `affected` array and recalculation count are surfaced directly in the UI.

A reverse edge can be removed correctly: if C1 changes from `=A1` to `=B1`, editing A1 later no longer invalidates C1. There is a dedicated regression test for this.

### 4. Evaluation, errors, and cycles

The evaluator walks the AST and returns a scalar or a temporary 2D matrix for range-valued arguments. Literal cell values are inferred as number, boolean, blank, or text (an initial apostrophe forces text).

`Workbook.getValue(ref, visiting)` carries the active evaluation path. A direct or indirect re-entry to the same reference is a cycle. It reports `#CYCLE!` with the actual path (for example, `A1 → B1 → A1`) instead of recursing forever. Recalculation errors propagate back to dependent cells.

`IF` and `IFERROR` are **special lazy functions**. The evaluator does not evaluate a branch that is not needed:

- `=IF(FALSE, 1/0, 42)` → 42
- `=IFERROR(1/0, 99)` → 99

Runtime errors are represented by `SheetError(code,message)`. Display errors include `#DIV/0!`, `#NUM!`, `#REF!`, `#NAME?`, `#VALUE!`, `#ARG!`, `#N/A`, and `#CYCLE!`. Syntax errors use `#PARSE!` and include position and suggestion.

### 5. Virtualized grid

The workbook address space defaults to **400 × 52 = 20,800 cells**; this is not 20,800 DOM elements. The scrolling surface has the correct virtual width and height, but `renderGrid()` only mounts the current visible rows/columns with a small overscan of 4 rows and 2 columns. Headers are dynamically repositioned to stay visible while scrolling.

With 30-pixel row heights, 126-pixel column widths, and ~800-pixel viewport height, a normal viewport renders hundreds rather than tens of thousands of cells. The exact number depends on screen dimensions.

All visible cell content is inserted using `textContent` rather than unsanitized HTML. The sidebar escapes formula strings before embedding them in templates. Invalidation refreshes displayed values without rebuilding the entire address space.

### 6. Editing, selection, and history

Editing is a transaction through `applyEdits(Edit[])`. An edit transaction records previous and updated raw values as one undo unit; v2 also records formatting deltas, and `applyTransaction` can commit both together. Drag fill, range clear, and TSV paste can therefore be undone in one operation. History is capped at 150 transactions.

Selection uses `{anchor, end, active}` coordinates. Keyboard arrows update the active cell; Shift+arrows keep the anchor and extend the selected rectangle. Double-click creates an inline editor; the formula bar edits the original text. Source references shift during fill, preserving `$` locks: `A1` shifts, `$A$1` does not.

One subtle UI choice: the grid is redrawn when selection changes. We handle rapid consecutive `mousedown` events on the same cell to make double-click-to-edit reliable even if the original DOM cell has been replaced. The browser interaction smoke test verified this behavior.

### 7. What-if scenario isolation

`preview(ref,proposed)` performs the following steps:

1. Gather the selected cell and reachable dependents from the real dependency graph.
2. Snapshot their current displayed values.
3. Serialize the real workbook's raw cells to an in-memory JSON object.
4. Construct a separate `Workbook` instance.
5. Apply the hypothetical edit to **the clone only**.
6. Compare before/after values and return a diff and recalculation metrics.
7. Keep the original object untouched until the user explicitly clicks **Apply this change**.

This is a local deterministic sandbox, not an LLM prediction or AI-generated guess. It is safe to test formulas containing errors because clone evaluation reports them as spreadsheet errors.

**Tradeoff:** Cloning costs O(number of populated cells) time/memory per preview. Fine for current challenge sizes; a production implementation could use persistent immutable snapshots or structural sharing.

### 8. Persistence and I/O

- localStorage key `gridforge:document:v2` stores a versioned document with names, all worksheets, raw cells and sparse styles after a 250 ms debounce. If only the original v1 `gridforge:greeting:workbook:v1` is present, it is imported as the first sheet. No accounts or network requests.
- **JSON** preserves every sheet, raw formula string, sparse style, selected sheet and workbook dimensions. Import limits are 1–1000 rows and 1–104 columns per sheet (up to 12 sheets).
- **CSV** import respects quoted delimiters/newlines/escaped quotes. CSV export writes the *evaluated result* of cells, not formula source; it is intentionally values-only. Text that begins with an external spreadsheet's formula prefix (after optional whitespace) is apostrophe-prefixed to reduce CSV formula-injection risk.
- `npm run start` serves static files through a small Node HTTP server. Runtime calculation happens completely in the browser.

### 9. Why vanilla TypeScript, not a large framework?

The brief allows any framework but requires TypeScript and prohibits an existing spreadsheet formula engine. Vanilla DOM + TypeScript keeps the challenge's most important logic visible: the candidate can explain how each interaction mutates workbook state without tracing an external reactive framework. It also eliminates runtime dependency and bundle-loader complexity.

The code is still modular: `Workbook` does not depend on browser APIs except standard `performance`, and its unit tests run directly in Node. Replacing `main.ts` with React is possible without changing the parser, functions, or graph core.

## Data correctness / security considerations

- **No JavaScript `eval`, `new Function`, or dynamic code execution.** Formulas are parsed as a restricted expression language.
- **Untrusted text displayed safely** through `textContent`/HTML escaping in the sidebar.
- **JSON imports validate dimensions, types and cell bounds.** File contents are local and do not grant server access.
- **Formula/CSV injection:** CSV imports beginning with `=` run as formulas in this app; external office software may also interpret values beginning with formula prefixes when CSV is opened. Do not treat CSV as a safe format for exchanging untrusted text without additional sanitization policies.
- **Resource protections:** 20,000-cell range expansion limit, 1000×104 maximum imported dimensions, 150 history operations. Expensive formulas can still affect browser responsiveness because computation is single-threaded.

## Known limitations and possible v3 enhancements

- A Web Worker for background recalculation and cancelable large calculations.
- Cross-sheet formula references and formulas beyond the 56 supported names. Same-app clipboard formulas/styles are supported, but depend on browser custom MIME support; external applications receive plain TSV.
- A typed persistent indexed storage layer instead of browser-local `localStorage`.
- More complete Excel-compatible conversion, error and comparison semantics.
- Multi-cell automatic fill series (e.g. 1,2,3...), instead of formula-relative fill.
- Formula reference highlighting, richer accessible focus behavior, and real-time collaborative editing with a backend.

These are potential future improvements rather than features claimed for v2. Native XLSX exchange, pivot tables, macros, collaborative editing and true Excel-parity sorting are also out of scope.
