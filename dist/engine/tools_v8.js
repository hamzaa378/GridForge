import { colToName, shiftFormula } from './parser.js';
import { parseLocaleCurrencyText, parseLocaleNumberText } from './locale_number.js';
import { parseExcelDateValue, parseExcelTimeValue } from './functions_v5.js';
import { splitTrailingDateTimeText } from './time_text.js';
const groups = [
    ['Text & case', 'upper', 'UPPERCASE', 'Make selected text uppercase'],
    ['Text & case', 'lower', 'lowercase', 'Make selected text lowercase'],
    ['Text & case', 'title', 'Title Case', 'Capitalize words'],
    ['Text & case', 'sentence', 'Sentence case', 'Capitalize the first character'],
    ['Text & case', 'swapcase', 'sWAP cASE', 'Invert letter case'],
    ['Text & case', 'reverse-text', 'Reverse text', 'Reverse Unicode characters'],
    ['Text & case', 'trim', 'Trim ends', 'Remove leading/trailing whitespace'],
    ['Text & case', 'collapse-spaces', 'Collapse spaces', 'Replace repeated whitespace with one space'],
    ['Text & case', 'remove-spaces', 'Remove all whitespace', 'Remove whitespace throughout'],
    ['Text & case', 'clean-control', 'Clean controls', 'Remove control characters'],
    ['Text extraction', 'first-word', 'First word', 'Extract the first word'],
    ['Text extraction', 'last-word', 'Last word', 'Extract the last word'],
    ['Text extraction', 'before-space', 'Before first space', 'Keep characters before the first space'],
    ['Text extraction', 'after-space', 'After first space', 'Keep characters after the first space'],
    ['Text extraction', 'initials', 'Initials', 'Extract first letter of each word'],
    ['Text extraction', 'first-three', 'First 3 characters', 'Extract first three characters'],
    ['Text extraction', 'last-three', 'Last 3 characters', 'Extract last three characters'],
    ['Text extraction', 'drop-first', 'Drop first character', 'Remove first Unicode character'],
    ['Text extraction', 'drop-last', 'Drop last character', 'Remove last Unicode character'],
    ['Text extraction', 'word-count', 'Count words', 'Replace text with its word count'],
    ['Text cleanup', 'remove-digits', 'Remove digits', 'Remove all numeric digits'],
    ['Text cleanup', 'digits-only', 'Digits only', 'Keep only numeric digits'],
    ['Text cleanup', 'letters-only', 'Letters only', 'Keep Unicode letters'],
    ['Text cleanup', 'alnum-only', 'Alphanumeric only', 'Remove punctuation and spaces'],
    ['Text cleanup', 'remove-punct', 'Remove punctuation', 'Remove punctuation characters'],
    ['Text cleanup', 'remove-newlines', 'Remove line breaks', 'Replace newlines with spaces'],
    ['Text cleanup', 'tabs-to-spaces', 'Tabs to spaces', 'Replace tabs with single spaces'],
    ['Text cleanup', 'unicode-normalize', 'Normalize Unicode', 'Apply Unicode NFKC normalization'],
    ['Text cleanup', 'strip-accents', 'Remove accents', 'Strip combining diacritical marks'],
    ['Text cleanup', 'strip-leading-zeros', 'Strip leading zeros', 'Remove zeros preceding numeric text'],
    ['Numeric calculations', 'negate', 'Negate', 'Switch number sign'],
    ['Numeric calculations', 'absolute', 'Absolute value', 'Convert to positive magnitude'],
    ['Numeric calculations', 'reciprocal', 'Reciprocal', 'Calculate 1/x (skips zero)'],
    ['Numeric calculations', 'square', 'Square', 'Calculate x²'],
    ['Numeric calculations', 'cube', 'Cube', 'Calculate x³'],
    ['Numeric calculations', 'sqrt', 'Square root', 'Calculate square root (skips negatives)'],
    ['Numeric calculations', 'sign', 'Sign', 'Return -1, 0 or 1'],
    ['Numeric calculations', 'increment', 'Add one', 'Increase numbers by one'],
    ['Numeric calculations', 'decrement', 'Subtract one', 'Decrease numbers by one'],
    ['Numeric calculations', 'double', 'Double', 'Multiply numbers by two'],
    ['Rounding & precision', 'round0', 'Round integer', 'Round to zero decimal places'],
    ['Rounding & precision', 'round1', 'Round 1 decimal', 'Round to one decimal place'],
    ['Rounding & precision', 'round2', 'Round 2 decimals', 'Round to two decimal places'],
    ['Rounding & precision', 'round3', 'Round 3 decimals', 'Round to three decimal places'],
    ['Rounding & precision', 'ceil', 'Round up', 'Round towards positive infinity'],
    ['Rounding & precision', 'floor', 'Round down', 'Round towards negative infinity'],
    ['Rounding & precision', 'trunc', 'Truncate', 'Discard the fractional part'],
    ['Rounding & precision', 'nearest5', 'Nearest 5', 'Round to multiple of five'],
    ['Rounding & precision', 'nearest10', 'Nearest 10', 'Round to multiple of ten'],
    ['Rounding & precision', 'nearest100', 'Nearest 100', 'Round to multiple of one hundred'],
    ['Percent & scaling', 'divide100', 'Divide by 100', 'Convert percentages to fractions'],
    ['Percent & scaling', 'multiply100', 'Multiply by 100', 'Convert fractions to percent numbers'],
    ['Percent & scaling', 'divide1000', 'Divide by 1,000', 'Convert to thousands'],
    ['Percent & scaling', 'multiply1000', 'Multiply by 1,000', 'Scale values up'],
    ['Percent & scaling', 'half', 'Halve', 'Divide numeric values by two'],
    ['Percent & scaling', 'triple', 'Triple', 'Multiply by three'],
    ['Percent & scaling', 'plus10pct', 'Increase 10%', 'Multiply by 1.10'],
    ['Percent & scaling', 'minus10pct', 'Decrease 10%', 'Multiply by 0.90'],
    ['Percent & scaling', 'parse-percent', 'Parse percent text', 'Convert text such as 12% to 0.12'],
    ['Percent & scaling', 'parse-currency', 'Parse currency text', 'Convert $1,234.50 to a number'],
    ['Date & time', 'parse-date', 'Parse date text', 'Convert a supported locale date to an Excel date serial'],
    ['Date & time', 'parse-time', 'Parse time text', 'Convert a supported locale time to an Excel time fraction'],
    ['Date & time', 'parse-datetime', 'Parse date and time', 'Convert supported locale date-time text to an Excel serial'],
    ['Fill & series', 'fill-down', 'Fill down', 'Copy first row downward with formula shifts'],
    ['Fill & series', 'fill-up', 'Fill up', 'Copy last row upward with formula shifts'],
    ['Fill & series', 'fill-right', 'Fill right', 'Copy first column across with formula shifts'],
    ['Fill & series', 'fill-left', 'Fill left', 'Copy last column across with formula shifts'],
    ['Fill & series', 'series-one', 'Series 1, 2, 3…', 'Sequential numbering starting at 1'],
    ['Fill & series', 'series-zero', 'Series 0, 1, 2…', 'Sequential numbering starting at 0'],
    ['Fill & series', 'series-even', 'Series 2, 4, 6…', 'Sequential even values'],
    ['Fill & series', 'series-odd', 'Series 1, 3, 5…', 'Sequential odd values'],
    ['Fill & series', 'series-fives', 'Series 5, 10, 15…', 'Sequential multiples of five'],
    ['Fill & series', 'series-tens', 'Series 10, 20, 30…', 'Sequential multiples of ten'],
    ['Range arrangement', 'flip-vertical', 'Flip vertically', 'Reverse entire rows within selection'],
    ['Range arrangement', 'flip-horizontal', 'Flip horizontally', 'Reverse entire columns within selection'],
    ['Range arrangement', 'transpose-square', 'Transpose square', 'Exchange rows and columns (square only)'],
    ['Range arrangement', 'rotate-right', 'Rotate clockwise', 'Rotate square range clockwise'],
    ['Range arrangement', 'rotate-left', 'Rotate counterclockwise', 'Rotate square range counterclockwise'],
    ['Range arrangement', 'shift-down', 'Shift selection down', 'Move cells down, wrapping within selection'],
    ['Range arrangement', 'shift-up', 'Shift selection up', 'Move cells up, wrapping within selection'],
    ['Range arrangement', 'shift-right', 'Shift selection right', 'Move cells right, wrapping within selection'],
    ['Range arrangement', 'shift-left', 'Shift selection left', 'Move cells left, wrapping within selection'],
    ['Range arrangement', 'clear-diagonal', 'Clear main diagonal', 'Clear diagonal cells in selection'],
    ['Data quality', 'blank-zero', 'Blanks to 0', 'Fill empty cells with zero'],
    ['Data quality', 'blank-na', 'Blanks to N/A text', 'Fill empty cells with N/A'],
    ['Data quality', 'zero-blank', 'Zeros to blanks', 'Remove zero-valued cells'],
    ['Data quality', 'negative-zero', 'Negative to zero', 'Clamp negative numbers to zero'],
    ['Data quality', 'positive-zero', 'Positive to zero', 'Clamp positive numbers to zero'],
    ['Data quality', 'error-clear', 'Clear formula errors', 'Replace error formulas with blanks'],
    ['Data quality', 'dedupe-rows', 'Deduplicate rows', 'Clear repeated rows within selection'],
    ['Data quality', 'dedupe-cells', 'Deduplicate cells', 'Clear repeated nonblank values'],
    ['Data quality', 'space-hyphen', 'Spaces to hyphens', 'Replace spaces with hyphens'],
    ['Data quality', 'hyphen-space', 'Hyphens to spaces', 'Replace hyphens with spaces'],
    ['Quick aggregates', 'insert-sum', 'SUM below selection', 'Insert a formula below the selection'],
    ['Quick aggregates', 'insert-average', 'AVERAGE below selection', 'Insert a formula below the selection'],
    ['Quick aggregates', 'insert-min', 'MIN below selection', 'Insert a formula below the selection'],
    ['Quick aggregates', 'insert-max', 'MAX below selection', 'Insert a formula below the selection'],
    ['Quick aggregates', 'insert-count', 'COUNT below selection', 'Insert a formula below the selection'],
    ['Quick aggregates', 'insert-counta', 'COUNTA below selection', 'Insert a formula below the selection'],
    ['Quick aggregates', 'insert-median', 'MEDIAN below selection', 'Insert a formula below the selection'],
    ['Quick aggregates', 'insert-product', 'PRODUCT below selection', 'Insert a formula below the selection'],
    ['Quick aggregates', 'insert-stdev', 'STDEV.S below selection', 'Insert a formula below the selection'],
    ['Quick aggregates', 'insert-countblank', 'COUNTBLANK below selection', 'Insert a formula below the selection']
];
const kindOf = (category) => category === 'Numeric calculations' || category === 'Rounding & precision' || category === 'Percent & scaling' || category === 'Date & time' ? 'number' : category === 'Fill & series' ? 'fill' : category === 'Range arrangement' ? 'range' : category === 'Data quality' ? 'cleanup' : category === 'Quick aggregates' ? 'formula' : 'text';
export const DATA_TOOLS = Object.freeze(groups.map(([category, id, label, description]) => ({ category, id, label, description, kind: kindOf(category) })));
export const TOOL_CATEGORIES = Object.freeze([...new Set(DATA_TOOLS.map(x => x.category))]);
const addr = (r, c) => `${colToName(c)}${r + 1}`;
const asLiteral = (value) => value.startsWith('=') ? `'${value}` : value;
const nonformula = (raw) => raw && !raw.startsWith('=');
const safeNumber = (n) => Number.isFinite(n) && Math.abs(n) <= 1e150 ? String(Number(n.toPrecision(14))) : null;
const numberTransform = (id, value) => {
    switch (id) {
        case 'negate': return -value;
        case 'absolute': return Math.abs(value);
        case 'reciprocal': return value === 0 ? null : 1 / value;
        case 'square': return value * value;
        case 'cube': return value ** 3;
        case 'sqrt': return value < 0 ? null : Math.sqrt(value);
        case 'sign': return Math.sign(value);
        case 'increment': return value + 1;
        case 'decrement': return value - 1;
        case 'double': return value * 2;
        case 'round0':
        case 'round1':
        case 'round2':
        case 'round3': {
            const digit = Number(id.slice(-1));
            const factor = 10 ** digit;
            return Math.round((value + Number.EPSILON) * factor) / factor;
        }
        case 'ceil': return Math.ceil(value);
        case 'floor': return Math.floor(value);
        case 'trunc': return Math.trunc(value);
        case 'nearest5': return Math.round(value / 5) * 5;
        case 'nearest10': return Math.round(value / 10) * 10;
        case 'nearest100': return Math.round(value / 100) * 100;
        case 'divide100': return value / 100;
        case 'multiply100': return value * 100;
        case 'divide1000': return value / 1000;
        case 'multiply1000': return value * 1000;
        case 'half': return value / 2;
        case 'triple': return value * 3;
        case 'plus10pct': return value * 1.1;
        case 'minus10pct': return value * 0.9;
        default: return null;
    }
};
const title = (s) => s.toLowerCase().replace(/(^|[\s\-])([\p{L}])/gu, (_, sep, chr) => sep + chr.toUpperCase());
const words = (s) => s.trim().split(/\s+/u).filter(Boolean);
const textTransform = (id, s) => {
    const chars = [...s];
    switch (id) {
        case 'upper': return s.toUpperCase();
        case 'lower': return s.toLowerCase();
        case 'title': return title(s);
        case 'sentence': return s.length ? s[0].toUpperCase() + s.slice(1).toLowerCase() : s;
        case 'swapcase': return chars.map(ch => ch === ch.toLowerCase() ? ch.toUpperCase() : ch.toLowerCase()).join('');
        case 'reverse-text': return chars.reverse().join('');
        case 'trim': return s.trim();
        case 'collapse-spaces': return s.trim().replace(/\s+/gu, ' ');
        case 'remove-spaces': return s.replace(/\s+/gu, '');
        case 'clean-control': return s.replace(/[\x00-\x1f\x7f]/g, '');
        case 'first-word': return words(s)[0] ?? '';
        case 'last-word': return words(s).at(-1) ?? '';
        case 'before-space': return s.split(/\s/u, 1)[0] ?? '';
        case 'after-space': return s.replace(/^\S+\s*/u, '');
        case 'initials': return words(s).map(w => [...w][0]?.toUpperCase() ?? '').join('');
        case 'first-three': return chars.slice(0, 3).join('');
        case 'last-three': return chars.slice(-3).join('');
        case 'drop-first': return chars.slice(1).join('');
        case 'drop-last': return chars.slice(0, -1).join('');
        case 'word-count': return words(s).length;
        case 'remove-digits': return s.replace(/[0-9]/g, '');
        case 'digits-only': return s.replace(/[^0-9]/g, '');
        case 'letters-only': return s.replace(/[^\p{L}]/gu, '');
        case 'alnum-only': return s.replace(/[^\p{L}\p{N}]/gu, '');
        case 'remove-punct': return s.replace(/\p{P}/gu, '');
        case 'remove-newlines': return s.replace(/[\r\n]+/g, ' ');
        case 'tabs-to-spaces': return s.replace(/\t/g, ' ');
        case 'unicode-normalize': return s.normalize('NFKC');
        case 'strip-accents': return s.normalize('NFD').replace(/\p{M}/gu, '').normalize('NFC');
        case 'strip-leading-zeros': return s.replace(/^([+-]?)0+(?=\d)/, '$1');
        case 'space-hyphen': return s.replace(/ /g, '-');
        case 'hyphen-space': return s.replace(/-/g, ' ');
        default: return s;
    }
};
function within(b, rows, cols) {
    if (![b.r1, b.r2, b.c1, b.c2].every(Number.isInteger) || b.r1 < 0 || b.c1 < 0 || b.r1 > b.r2 || b.c1 > b.c2 || b.r2 >= rows || b.c2 >= cols)
        throw Error('Selection is outside this worksheet.');
    const count = (b.r2 - b.r1 + 1) * (b.c2 - b.c1 + 1);
    if (count > 12000)
        throw Error('Data Studio is limited to 12,000 selected cells per operation.');
    return count;
}
/** Prepares a deterministic edit list; does not alter workbook until the user applies it. */
export function previewDataTool(book, b, id) {
    const spec = DATA_TOOLS.find(x => x.id === id);
    if (!spec)
        throw Error(`Unknown Data Studio tool: ${id}.`);
    const count = within(b, book.rows, book.cols), edits = [], old = Array.from({ length: b.r2 - b.r1 + 1 }, (_, ri) => Array.from({ length: b.c2 - b.c1 + 1 }, (_, ci) => book.getRaw(addr(b.r1 + ri, b.c1 + ci))));
    let skipped = 0;
    const send = (r, c, raw) => { const ref = addr(r, c); if (raw !== book.getRaw(ref))
        edits.push({ ref, raw }); };
    if (spec.kind === 'text' || spec.kind === 'number' || spec.kind === 'cleanup') {
        const seen = new Set(), rowsSeen = new Set();
        for (let ri = 0; ri < old.length; ri++) {
            const row = old[ri];
            const rowKey = id === 'dedupe-rows' ? JSON.stringify(row.map((x, ci) => x.startsWith('=') ? book.getValue(addr(b.r1 + ri, b.c1 + ci)).value : x)) : '';
            const duplicateRow = id === 'dedupe-rows' && rowsSeen.has(rowKey);
            if (id === 'dedupe-rows')
                rowsSeen.add(rowKey);
            for (let ci = 0; ci < row.length; ci++) {
                const ref = addr(b.r1 + ri, b.c1 + ci), raw = row[ci], formula = raw.startsWith('='), value = book.getValue(ref);
                let next = null;
                if (spec.kind === 'text') {
                    if (!nonformula(raw)) {
                        skipped++;
                        continue;
                    }
                    const v = textTransform(id, raw.startsWith("'") ? raw.slice(1) : raw);
                    next = typeof v === 'number' ? String(v) : asLiteral(v);
                }
                else if (spec.kind === 'number') {
                    if (formula || !raw.trim()) {
                        skipped++;
                        continue;
                    }
                    if (id === 'parse-percent') {
                        const trimmed = raw.trim(), value = trimmed.endsWith('%') && !trimmed.slice(0, -1).includes('%') ? parseLocaleNumberText(trimmed) : null;
                        next = value === null ? null : safeNumber(value);
                    }
                    else if (id === 'parse-currency') {
                        const value = parseLocaleCurrencyText(raw) ?? parseLocaleNumberText(raw);
                        next = value === null ? null : safeNumber(value);
                    }
                    else if (id === 'parse-date' || id === 'parse-time') {
                        try {
                            next = safeNumber(id === 'parse-date' ? parseExcelDateValue(raw) : parseExcelTimeValue(raw));
                        }
                        catch {
                            next = null;
                        }
                    }
                    else if (id === 'parse-datetime') {
                        const parts = splitTrailingDateTimeText(raw);
                        if (parts)
                            try {
                                next = safeNumber(parseExcelDateValue(parts[0]) + parseExcelTimeValue(parts[1]));
                            }
                            catch {
                                next = null;
                            }
                    }
                    else {
                        const x = Number(raw);
                        next = Number.isFinite(x) ? (safeNumber(numberTransform(id, x) ?? NaN)) : null;
                    }
                    if (next === null) {
                        skipped++;
                        continue;
                    }
                }
                else {
                    switch (id) {
                        case 'blank-zero':
                            next = raw === '' ? '0' : null;
                            break;
                        case 'blank-na':
                            next = raw === '' ? 'N/A' : null;
                            break;
                        case 'zero-blank':
                            next = !formula && raw.trim() !== '' && Number(raw) === 0 ? '' : null;
                            break;
                        case 'negative-zero':
                            next = !formula && raw.trim() !== '' && Number(raw) < 0 ? '0' : null;
                            break;
                        case 'positive-zero':
                            next = !formula && raw.trim() !== '' && Number(raw) > 0 ? '0' : null;
                            break;
                        case 'error-clear':
                            next = formula && value.error ? '' : null;
                            break;
                        case 'dedupe-rows':
                            next = duplicateRow ? '' : null;
                            break;
                        case 'dedupe-cells': {
                            const key = JSON.stringify([typeof value.value, value.value]);
                            if (!raw) {
                                next = null;
                                break;
                            }
                            if (seen.has(key))
                                next = '';
                            else
                                seen.add(key);
                            break;
                        }
                        case 'space-hyphen':
                        case 'hyphen-space':
                            next = !formula && raw ? asLiteral(String(textTransform(id, raw))) : null;
                            break;
                    }
                    if (next === null) {
                        skipped++;
                        continue;
                    }
                }
                if (next !== null)
                    send(b.r1 + ri, b.c1 + ci, next);
            }
        }
    }
    else if (spec.kind === 'fill') {
        const series = { 'series-one': [1, 1], 'series-zero': [0, 1], 'series-even': [2, 2], 'series-odd': [1, 2], 'series-fives': [5, 5], 'series-tens': [10, 10] };
        const setup = series[id];
        if (setup) {
            let index = 0;
            for (let r = b.r1; r <= b.r2; r++)
                for (let c = b.c1; c <= b.c2; c++)
                    send(r, c, String(setup[0] + setup[1] * index++));
        }
        else
            for (let ri = 0; ri < old.length; ri++)
                for (let ci = 0; ci < old[ri].length; ci++) {
                    let srcR = ri, srcC = ci;
                    if (id === 'fill-down')
                        srcR = 0;
                    if (id === 'fill-up')
                        srcR = old.length - 1;
                    if (id === 'fill-right')
                        srcC = 0;
                    if (id === 'fill-left')
                        srcC = old[ri].length - 1;
                    const raw = old[srcR][srcC];
                    const shifted = raw.startsWith('=') ? shiftFormula(raw, ri - srcR, ci - srcC) : raw;
                    send(b.r1 + ri, b.c1 + ci, shifted);
                }
    }
    else if (spec.kind === 'range') {
        const height = old.length, width = old[0].length;
        if (['transpose-square', 'rotate-right', 'rotate-left'].includes(id) && height !== width)
            throw Error('This operation requires a square selection (same number of rows and columns).');
        for (let ri = 0; ri < height; ri++)
            for (let ci = 0; ci < width; ci++) {
                if (id === 'clear-diagonal') {
                    if (ri === ci)
                        send(b.r1 + ri, b.c1 + ci, '');
                    continue;
                }
                let srcR = ri, srcC = ci;
                switch (id) {
                    case 'flip-vertical':
                        srcR = height - 1 - ri;
                        break;
                    case 'flip-horizontal':
                        srcC = width - 1 - ci;
                        break;
                    case 'transpose-square':
                        srcR = ci;
                        srcC = ri;
                        break;
                    case 'rotate-right':
                        srcR = height - 1 - ci;
                        srcC = ri;
                        break;
                    case 'rotate-left':
                        srcR = ci;
                        srcC = width - 1 - ri;
                        break;
                    case 'shift-down':
                        srcR = (ri - 1 + height) % height;
                        break;
                    case 'shift-up':
                        srcR = (ri + 1) % height;
                        break;
                    case 'shift-right':
                        srcC = (ci - 1 + width) % width;
                        break;
                    case 'shift-left':
                        srcC = (ci + 1) % width;
                        break;
                }
                const raw = old[srcR][srcC];
                send(b.r1 + ri, b.c1 + ci, raw.startsWith('=') ? shiftFormula(raw, ri - srcR, ci - srcC) : raw);
            }
    }
    else if (spec.kind === 'formula') {
        if (b.r2 + 1 >= book.rows)
            throw Error('There is no room to insert a summary below the selection.');
        const formulas = { 'insert-sum': 'SUM', 'insert-average': 'AVERAGE', 'insert-min': 'MIN', 'insert-max': 'MAX', 'insert-count': 'COUNT', 'insert-counta': 'COUNTA', 'insert-median': 'MEDIAN', 'insert-product': 'PRODUCT', 'insert-stdev': 'STDEV.S', 'insert-countblank': 'COUNTBLANK' };
        const name = formulas[id];
        if (!name)
            throw Error('Formula tool is not registered.');
        for (let c = b.c1; c <= b.c2; c++)
            send(b.r2 + 1, c, `=${name}(${addr(b.r1, c)}:${addr(b.r2, c)})`);
    }
    return { edits, skipped, summary: `${spec.label}: ${edits.length} cell${edits.length === 1 ? '' : 's'} will change out of ${count} selected.`, warning: spec.kind === 'range' ? 'Formulas are translated relative to their new cell positions; verify complex external references.' : undefined };
}
