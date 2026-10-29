import { colToName, parseRef, shiftFormula } from './parser.js';
import { display } from './workbook.js';
const address = (row, col) => `${colToName(col)}${row + 1}`;
export function analyzeSelection(book, b) {
    let count = 0, numbers = 0, sum = 0, min = Infinity, max = -Infinity, blanks = 0, errors = 0, duplicates = 0;
    const seen = new Set();
    const chart = [];
    for (let r = b.r1; r <= b.r2; r++)
        for (let c = b.c1; c <= b.c2; c++) {
            const result = book.getValue(address(r, c));
            count++;
            if (result.error) {
                errors++;
                continue;
            }
            const v = result.value;
            if (v === null || v === '') {
                blanks++;
                continue;
            }
            const key = JSON.stringify([typeof v, v]);
            if (seen.has(key))
                duplicates++;
            else
                seen.add(key);
            if (typeof v === 'number' && Number.isFinite(v)) {
                numbers++;
                sum += v;
                min = Math.min(min, v);
                max = Math.max(max, v);
            }
        }
    for (let r = b.r1; r <= Math.min(b.r2, b.r1 + 24); r++) {
        const label = display(book.getValue(address(r, b.c1))) || address(r, b.c1);
        const target = b.c2 > b.c1 ? b.c1 + 1 : b.c1;
        const v = book.getValue(address(r, target));
        if (!v.error && typeof v.value === 'number' && Number.isFinite(v.value))
            chart.push({ label, value: v.value });
    }
    return { count, numbers, sum, average: numbers ? sum / numbers : null, min: numbers ? min : null, max: numbers ? max : null, blanks, errors, duplicates, chart };
}
export function findCells(book, needle, caseSensitive = false, includeFormulas = true) {
    if (!needle)
        return [];
    const query = caseSensitive ? needle : needle.toLowerCase();
    const items = book.allEntries().filter(({ raw }) => {
        if (!includeFormulas && raw.startsWith('='))
            return false;
        return (caseSensitive ? raw : raw.toLowerCase()).includes(query);
    });
    return items.sort((a, b) => { const x = parseRef(a.ref), y = parseRef(b.ref); return x.row - y.row || x.col - y.col; });
}
export function findAcrossSheets(sheets, needle, caseSensitive = false, includeFormulas = true) {
    return sheets.filter(sheet => !sheet.hidden).flatMap(sheet => findCells(sheet.book, needle, caseSensitive, includeFormulas).map(match => ({ ...match, sheetId: sheet.id, sheetName: sheet.name })));
}
export function replaceCells(book, needle, replacement, caseSensitive = false, includeFormulas = false) {
    if (!needle)
        throw Error('Enter text to find.');
    return findCells(book, needle, caseSensitive, includeFormulas).map(({ ref, raw }) => {
        let next = '';
        let start = 0;
        const original = caseSensitive ? raw : raw.toLowerCase(), query = caseSensitive ? needle : needle.toLowerCase();
        for (let i = original.indexOf(query, start); i !== -1; i = original.indexOf(query, start)) {
            next += raw.slice(start, i) + replacement;
            start = i + needle.length;
        }
        return { ref, raw: next + raw.slice(start) };
    });
}
export function sortRangeTransaction(book, b, byCol, descending = false, firstRowHeader = false) {
    const keys = typeof byCol === 'number' ? [{ column: byCol, descending }] : byCol;
    if (!keys.length || keys.length > 16 || keys.some(key => key.column < b.c1 || key.column > b.c2 || !Number.isInteger(key.column)) || new Set(keys.map(key => key.column)).size !== keys.length)
        throw Error('Sort keys must be unique columns inside the selection.');
    const from = b.r1 + (firstRowHeader ? 1 : 0);
    if (b.r2 <= from)
        return { edits: [], styles: [] };
    const entries = [];
    for (let r = from; r <= b.r2; r++)
        entries.push({ original: r, keys: keys.map(key => book.getValue(address(r, key.column))), raw: Array.from({ length: b.c2 - b.c1 + 1 }, (_, j) => book.getRaw(address(r, b.c1 + j))), styles: Array.from({ length: b.c2 - b.c1 + 1 }, (_, j) => book.getStyle(address(r, b.c1 + j))) });
    entries.sort((a, b) => {
        for (let index = 0; index < keys.length; index++) {
            const x = a.keys[index], y = b.keys[index], key = keys[index];
            if (x.error && !y.error)
                return 1;
            if (y.error && !x.error)
                return -1;
            if (x.value === null || x.value === '') {
                if (y.value === null || y.value === '')
                    continue;
                return 1;
            }
            if (y.value === null || y.value === '')
                return -1;
            const comparison = typeof x.value === 'number' && typeof y.value === 'number' ? x.value - y.value : String(x.value).localeCompare(String(y.value), undefined, { numeric: true, sensitivity: 'base' });
            if (comparison)
                return key.descending ? -comparison : comparison;
        }
        return a.original - b.original;
    });
    const edits = [], styles = [];
    entries.forEach((item, index) => {
        const destination = from + index;
        item.raw.forEach((raw, delta) => {
            const adjusted = raw.startsWith('=') ? shiftFormula(raw, destination - item.original, 0) : raw;
            if (adjusted !== book.getRaw(address(destination, b.c1 + delta)))
                edits.push({ ref: address(destination, b.c1 + delta), raw: adjusted });
            const ref = address(destination, b.c1 + delta), style = item.styles[delta];
            if (JSON.stringify(style) !== JSON.stringify(book.getStyle(ref)))
                styles.push({ ref, style });
        });
    });
    return { edits, styles };
}
/** Backwards-compatible values-only helper retained for v2 tests and integrations. */
export function sortSelection(book, b, byCol, descending = false, firstRowHeader = false) {
    return sortRangeTransaction(book, b, byCol, descending, firstRowHeader).edits;
}
