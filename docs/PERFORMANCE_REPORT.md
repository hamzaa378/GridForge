# GridForge v3 benchmark verification

**Environment:** Node.js v22.16.0, development container, 2026-10-19. Run `npm run benchmark` on your own hardware before quoting response times.

````json
{
  "populated": 10400,
  "formulas": 10000,
  "initialBuildMs": 95.279,
  "singleSourceEdit": {"affected": 26, "recalculated": 26, "elapsedMs": 0.056},
  "isolatedEdit": {"affected": 1, "recalculated": 1, "elapsedMs": 0.006}
}
````

The primary supported claim is structural: the reverse-dependency graph identifies only **26 affected cells** after the specified row input edit, not all 10,000 formulas. Millisecond observations are illustrative, not an SLA or percentile latency guarantee.

---

## Historical benchmark notes

# Performance and Reliability Report

## Test target

GDG Task 2 requires at least 10,000 cells with grid virtualization and dependency-aware recalculation. We measured both the calculation engine and browser-visible DOM behavior. These are **local development measurements**, not guaranteed production performance across all browsers or devices.

### Environment

- Date: 19 October 2026.
- Node.js **v22.16.0**, TypeScript **5.8.3** used for compilation, 400 rows × 52 columns.
- Engine benchmark script: `scripts/benchmark.mjs`.
- Headless Chromium browser smoke testing on a 1560 × 920 viewport, with compiled application code and local assets in an isolated browser context. Direct headless-browser HTTP navigation was blocked by the environment. The v2 test used an offline assembled DOM with a standards-based import map linking actual compiled JS modules; this checks browser execution but **does not** test network serving or GitHub Pages.

### Reproducible calculation benchmark

Commands:

```bash
npm run build
node scripts/benchmark.mjs
```

Dataset: **10,400 populated cells**, comprising 400 literal row inputs and **10,000 formulas**. Each formula references one input on its own row. We then update one heavily referenced input and one unrelated cell.

| Metric | Result on this run |
| --- | ---: |
| Total addressable cells | 20,800 |
| Populated cells | 10,400 |
| Populated formulas | 10,000 |
| Initial full import + recalc | 95.516 ms |
| Initially recalculated populated cells | 10,400 |
| Source edit downstream reach (A2) | 26 affected cells |
| Recalculated on source edit | 26 cells |
| Core recalculation time for source edit | 0.049 ms |
| Isolated edit downstream reach | 1 affected cell |
| Recalculated on isolated edit | 1 cell |
| Core recalculation time for isolated edit | 0.007 ms |

**Interpretation:** The benchmark establishes selective invalidation, not a universal frame rate. Timings are for the engine's edit calculation and exclude DOM painting, file import parsing, browser reflow, and disk writing. Millisecond/sub-millisecond timings vary with system load and JIT warmup. Re-run the command on the reviewer's machine for fair comparison.

### V2 browser verification (compiled modules)

Chromium rendered a 20,800-cell sheet using 231 visible cell nodes at 1440×900, with no uncaught exceptions during the exercised smoke interactions. D4 initially displayed `4000`; applying Bold changed its rendered font weight to `800`. Adding a second sheet and switching back preserved the sample D4 value. Find All for `P-103` returned 2 matches. An additional extended run verified B4 edit → D4 recalculation → undo, style undo/redo, two-sheet isolation, internal clipboard formula shifts, selection analytics, sort/undo, and downloaded v2 JSON export of both sheets and styles. The offline `about:blank` context could not test localStorage persistence. All tests ran against compiled v2 code from `dist/` with an offline module import map because direct browser navigation was restricted.

### V1 regression browser evidence (earlier version)

- The updated v2 offline browser smoke test at 1440×900 yielded **231 DOM cell elements** for a 20,800-cell active-sheet address space (plus headers).
- Browser returned **D4 = 4000** on the seeded dataset.
- An actual browser double-click exposed the inline cell editor. Changing B4 to 225 updated dependent **D4 from 4000 to 7200**.
- Dragging a source formula `=B4*2` from H4 to H7 generated the values H5 = 176, H6 = 332, H7 = 102, confirming relative-reference shifting.
- A What-if preview produced a difference list without modifying the source workbook.
- No unhandled JavaScript exceptions were observed in the smoke tests exercised.

### V2 engine performance run

Exact output from `node scripts/benchmark.mjs` on Node v22.16.0:

```json
{
  "node":"v22.16.0","rows":400,"columns":52,"capacity":20800,
  "populated":10400,"formulas":10000,"initialBuildMs":95.516,
  "initialRecalculated":10400,
  "singleSourceEdit":{"affected":26,"recalculated":26,"elapsedMs":0.049},
  "isolatedEdit":{"affected":1,"recalculated":1,"elapsedMs":0.007}
}
```

v2 styling/document code does not cause whole-sheet recalculation when formatting changes. Results are indicative, measured locally, and not production service-level guarantees.

### Reliability test results

`npm test` result on the prepared source: **42 tests passed, 0 failed**. Tests include direct/indirect cycles, parser diagnostics, dependency graph replacement, undo/redo, JSON and CSV round trips, drag-fill reference translation, non-destructive What-if and a 12,000-populated-cell isolated edit.

### Limitations and interpretation

- This is not a 10,000-formula stress benchmark with *dense multi-cell range dependencies*. Performance degrades as dependency density and range sizes increase.
- UI viewport-based virtualization avoids creating all DOM cells, but **formula recalculation still runs on the main thread**. For millions of dependency edges, a worker would be desirable.
- `localStorage` writes are debounced and not part of the engine timing.
- No 60 fps claim, production scalability guarantee, or external-browser compatibility matrix is made here.

## Repeatability

The benchmark uses deterministic formula rows, no network access, and no random input. Run `npm run build && node scripts/benchmark.mjs`; share the generated JSON alongside the environment's Node version. The same tests can be run on Windows, macOS or Linux with Node 20+.
