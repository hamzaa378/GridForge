# GridForge v8 — 100 Data Studio Actions

These are **100 individually implemented Excel-style data operations**, not 100 additional Excel formula functions or full feature parity with Microsoft Excel. GridForge retains 306 custom formula functions.

## How to use

Select a range in the worksheet, click **Data Studio** (or use the right-click menu), choose an operation, click **Preview changes**, review the count and any warnings, then **Apply selection**. Every apply is undoable with Ctrl+Z.

## Safety and compatibility

- Text and numeric cleanup skip formulas instead of silently replacing their source. Fill and range transformations intentionally move formulas with relative references adjusted.
- Each operation is limited to 12,000 selected cells. Transpose and rotations require square ranges.
- Quick aggregates write formulas below the selection; if those output cells are populated, GridForge prompts before overwriting.
- For imported XLSM workbooks, the original VBA binary is preserved when exporting supported cell-content changes. Cell style updates, structural changes and shared formula edits have limited round-trip support.
- Embedded VBA is never executed during import. A separate Windows desktop Excel bridge can run trusted macros only after explicit approval.

## Complete 100-action catalog

### Text & case

| # | Action | Behavior |
|---:|---|---|
| 1 | **UPPERCASE** (`upper`) | Make selected text uppercase |
| 2 | **lowercase** (`lower`) | Make selected text lowercase |
| 3 | **Title Case** (`title`) | Capitalize words |
| 4 | **Sentence case** (`sentence`) | Capitalize the first character |
| 5 | **sWAP cASE** (`swapcase`) | Invert letter case |
| 6 | **Reverse text** (`reverse-text`) | Reverse Unicode characters |
| 7 | **Trim ends** (`trim`) | Remove leading/trailing whitespace |
| 8 | **Collapse spaces** (`collapse-spaces`) | Replace repeated whitespace with one space |
| 9 | **Remove all whitespace** (`remove-spaces`) | Remove whitespace throughout |
| 10 | **Clean controls** (`clean-control`) | Remove control characters |

### Text extraction

| # | Action | Behavior |
|---:|---|---|
| 11 | **First word** (`first-word`) | Extract the first word |
| 12 | **Last word** (`last-word`) | Extract the last word |
| 13 | **Before first space** (`before-space`) | Keep characters before the first space |
| 14 | **After first space** (`after-space`) | Keep characters after the first space |
| 15 | **Initials** (`initials`) | Extract first letter of each word |
| 16 | **First 3 characters** (`first-three`) | Extract first three characters |
| 17 | **Last 3 characters** (`last-three`) | Extract last three characters |
| 18 | **Drop first character** (`drop-first`) | Remove first Unicode character |
| 19 | **Drop last character** (`drop-last`) | Remove last Unicode character |
| 20 | **Count words** (`word-count`) | Replace text with its word count |

### Text cleanup

| # | Action | Behavior |
|---:|---|---|
| 21 | **Remove digits** (`remove-digits`) | Remove all numeric digits |
| 22 | **Digits only** (`digits-only`) | Keep only numeric digits |
| 23 | **Letters only** (`letters-only`) | Keep Unicode letters |
| 24 | **Alphanumeric only** (`alnum-only`) | Remove punctuation and spaces |
| 25 | **Remove punctuation** (`remove-punct`) | Remove punctuation characters |
| 26 | **Remove line breaks** (`remove-newlines`) | Replace newlines with spaces |
| 27 | **Tabs to spaces** (`tabs-to-spaces`) | Replace tabs with single spaces |
| 28 | **Normalize Unicode** (`unicode-normalize`) | Apply Unicode NFKC normalization |
| 29 | **Remove accents** (`strip-accents`) | Strip combining diacritical marks |
| 30 | **Strip leading zeros** (`strip-leading-zeros`) | Remove zeros preceding numeric text |

### Numeric calculations

| # | Action | Behavior |
|---:|---|---|
| 31 | **Negate** (`negate`) | Switch number sign |
| 32 | **Absolute value** (`absolute`) | Convert to positive magnitude |
| 33 | **Reciprocal** (`reciprocal`) | Calculate 1/x (skips zero) |
| 34 | **Square** (`square`) | Calculate x² |
| 35 | **Cube** (`cube`) | Calculate x³ |
| 36 | **Square root** (`sqrt`) | Calculate square root (skips negatives) |
| 37 | **Sign** (`sign`) | Return -1, 0 or 1 |
| 38 | **Add one** (`increment`) | Increase numbers by one |
| 39 | **Subtract one** (`decrement`) | Decrease numbers by one |
| 40 | **Double** (`double`) | Multiply numbers by two |

### Rounding & precision

| # | Action | Behavior |
|---:|---|---|
| 41 | **Round integer** (`round0`) | Round to zero decimal places |
| 42 | **Round 1 decimal** (`round1`) | Round to one decimal place |
| 43 | **Round 2 decimals** (`round2`) | Round to two decimal places |
| 44 | **Round 3 decimals** (`round3`) | Round to three decimal places |
| 45 | **Round up** (`ceil`) | Round towards positive infinity |
| 46 | **Round down** (`floor`) | Round towards negative infinity |
| 47 | **Truncate** (`trunc`) | Discard the fractional part |
| 48 | **Nearest 5** (`nearest5`) | Round to multiple of five |
| 49 | **Nearest 10** (`nearest10`) | Round to multiple of ten |
| 50 | **Nearest 100** (`nearest100`) | Round to multiple of one hundred |

### Percent & scaling

| # | Action | Behavior |
|---:|---|---|
| 51 | **Divide by 100** (`divide100`) | Convert percentages to fractions |
| 52 | **Multiply by 100** (`multiply100`) | Convert fractions to percent numbers |
| 53 | **Divide by 1,000** (`divide1000`) | Convert to thousands |
| 54 | **Multiply by 1,000** (`multiply1000`) | Scale values up |
| 55 | **Halve** (`half`) | Divide numeric values by two |
| 56 | **Triple** (`triple`) | Multiply by three |
| 57 | **Increase 10%** (`plus10pct`) | Multiply by 1.10 |
| 58 | **Decrease 10%** (`minus10pct`) | Multiply by 0.90 |
| 59 | **Parse percent text** (`parse-percent`) | Convert text such as 12% to 0.12 |
| 60 | **Parse currency text** (`parse-currency`) | Convert $1,234.50 to a number |

### Fill & series

| # | Action | Behavior |
|---:|---|---|
| 61 | **Fill down** (`fill-down`) | Copy first row downward with formula shifts |
| 62 | **Fill up** (`fill-up`) | Copy last row upward with formula shifts |
| 63 | **Fill right** (`fill-right`) | Copy first column across with formula shifts |
| 64 | **Fill left** (`fill-left`) | Copy last column across with formula shifts |
| 65 | **Series 1, 2, 3…** (`series-one`) | Sequential numbering starting at 1 |
| 66 | **Series 0, 1, 2…** (`series-zero`) | Sequential numbering starting at 0 |
| 67 | **Series 2, 4, 6…** (`series-even`) | Sequential even values |
| 68 | **Series 1, 3, 5…** (`series-odd`) | Sequential odd values |
| 69 | **Series 5, 10, 15…** (`series-fives`) | Sequential multiples of five |
| 70 | **Series 10, 20, 30…** (`series-tens`) | Sequential multiples of ten |

### Range arrangement

| # | Action | Behavior |
|---:|---|---|
| 71 | **Flip vertically** (`flip-vertical`) | Reverse entire rows within selection |
| 72 | **Flip horizontally** (`flip-horizontal`) | Reverse entire columns within selection |
| 73 | **Transpose square** (`transpose-square`) | Exchange rows and columns (square only) |
| 74 | **Rotate clockwise** (`rotate-right`) | Rotate square range clockwise |
| 75 | **Rotate counterclockwise** (`rotate-left`) | Rotate square range counterclockwise |
| 76 | **Shift selection down** (`shift-down`) | Move cells down, wrapping within selection |
| 77 | **Shift selection up** (`shift-up`) | Move cells up, wrapping within selection |
| 78 | **Shift selection right** (`shift-right`) | Move cells right, wrapping within selection |
| 79 | **Shift selection left** (`shift-left`) | Move cells left, wrapping within selection |
| 80 | **Clear main diagonal** (`clear-diagonal`) | Clear diagonal cells in selection |

### Data quality

| # | Action | Behavior |
|---:|---|---|
| 81 | **Blanks to 0** (`blank-zero`) | Fill empty cells with zero |
| 82 | **Blanks to N/A text** (`blank-na`) | Fill empty cells with N/A |
| 83 | **Zeros to blanks** (`zero-blank`) | Remove zero-valued cells |
| 84 | **Negative to zero** (`negative-zero`) | Clamp negative numbers to zero |
| 85 | **Positive to zero** (`positive-zero`) | Clamp positive numbers to zero |
| 86 | **Clear formula errors** (`error-clear`) | Replace error formulas with blanks |
| 87 | **Deduplicate rows** (`dedupe-rows`) | Clear repeated rows within selection |
| 88 | **Deduplicate cells** (`dedupe-cells`) | Clear repeated nonblank values |
| 89 | **Spaces to hyphens** (`space-hyphen`) | Replace spaces with hyphens |
| 90 | **Hyphens to spaces** (`hyphen-space`) | Replace hyphens with spaces |

### Quick aggregates

| # | Action | Behavior |
|---:|---|---|
| 91 | **SUM below selection** (`insert-sum`) | Insert a formula below the selection |
| 92 | **AVERAGE below selection** (`insert-average`) | Insert a formula below the selection |
| 93 | **MIN below selection** (`insert-min`) | Insert a formula below the selection |
| 94 | **MAX below selection** (`insert-max`) | Insert a formula below the selection |
| 95 | **COUNT below selection** (`insert-count`) | Insert a formula below the selection |
| 96 | **COUNTA below selection** (`insert-counta`) | Insert a formula below the selection |
| 97 | **MEDIAN below selection** (`insert-median`) | Insert a formula below the selection |
| 98 | **PRODUCT below selection** (`insert-product`) | Insert a formula below the selection |
| 99 | **STDEV.S below selection** (`insert-stdev`) | Insert a formula below the selection |
| 100 | **COUNTBLANK below selection** (`insert-countblank`) | Insert a formula below the selection |

## Testing

`npm run regression` runs TypeScript checks, unit tests, benchmarks, Chromium UI checks, XLSM round-trip tests, and VBA bridge security checks. `npm run regression:core` runs without browser dependencies. See `docs/V8_CODE_AUDIT.md`.
