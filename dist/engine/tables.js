/** Real, bounded Excel-style table definitions. Data stays in ordinary cells. */
import { parseRef, colToName } from './parser.js';
export const TABLE_STYLES = ['TableStyleMedium2', 'TableStyleMedium4', 'TableStyleMedium7', 'TableStyleMedium9', 'TableStyleMedium15', 'TableStyleMedium21', 'TableStyleLight1', 'TableStyleLight9', 'TableStyleDark1', 'TableStyleDark4'];
export function tableRange(value, rows, cols) {
    const m = /^\s*(\$?[A-Z]{1,3}\$?\d+)\s*:\s*(\$?[A-Z]{1,3}\$?\d+)\s*$/i.exec(value);
    if (!m)
        throw Error('Select a rectangle, for example A1:D12.');
    const a = parseRef(m[1].replace(/\$/g, '').toUpperCase()), b = parseRef(m[2].replace(/\$/g, '').toUpperCase());
    if (!a || !b || a.row > b.row || a.col > b.col || b.row >= rows || b.col >= cols || b.row === a.row)
        throw Error('A table requires a header/data rectangle of at least two rows.');
    if ((b.row - a.row + 1) * (b.col - a.col + 1) > 12000)
        throw Error('A table may cover up to 12,000 cells.');
    return { r1: a.row, r2: b.row, c1: a.col, c2: b.col };
}
export function validateTable(value, rows, cols) {
    if (!value || typeof value !== 'object' || Array.isArray(value))
        throw Error('Invalid table definition.');
    const o = value;
    if (typeof o.id !== 'string' || !/^table-[a-zA-Z0-9_-]{1,65}$/.test(o.id))
        throw Error('Invalid table id.');
    if (typeof o.name !== 'string' || !/^[A-Za-z_][A-Za-z0-9_.]{0,60}$/.test(o.name) || parseRef(o.name.toUpperCase()) !== null || /^R[1-9][0-9]*C[1-9][0-9]*$/i.test(o.name))
        throw Error('Table names must start with a letter or underscore and cannot be cell references.');
    if (typeof o.range !== 'string')
        throw Error('Invalid table range.');
    tableRange(o.range, rows, cols);
    if (!TABLE_STYLES.includes(o.style))
        throw Error('Unknown table style.');
    for (const key of ['header', 'bandedRows', 'bandedColumns', 'firstColumn', 'lastColumn'])
        if (typeof o[key] !== 'boolean')
            throw Error(`Invalid table option ${key}.`);
    let calculatedColumns;
    if (o.calculatedColumns !== undefined) {
        if (!Array.isArray(o.calculatedColumns) || o.calculatedColumns.length > 52)
            throw Error('Invalid calculated table columns.');
        const seen = new Set();
        calculatedColumns = o.calculatedColumns.map(raw => {
            if (!raw || typeof raw !== 'object')
                throw Error('Invalid calculated table column.');
            const c = raw;
            if (!Number.isInteger(c.column) || c.column < 0 || c.column > tableRange(o.range, rows, cols).c2 - tableRange(o.range, rows, cols).c1 || typeof c.formula !== 'string' || !c.formula.startsWith('=') || c.formula.length > 500)
                throw Error('Invalid calculated table column.');
            if (seen.has(c.column))
                throw Error('Duplicate calculated table column.');
            seen.add(c.column);
            const overrides = c.overrides ?? [];
            if (!Array.isArray(overrides) || overrides.length > 12000 || overrides.some(x => !Number.isInteger(x) || x < 0))
                throw Error('Invalid calculated-column overrides.');
            return { column: c.column, formula: c.formula, overrides: [...overrides] };
        });
    }
    return { id: o.id, name: o.name, range: o.range, style: o.style, header: o.header, bandedRows: o.bandedRows, bandedColumns: o.bandedColumns, firstColumn: o.firstColumn, lastColumn: o.lastColumn, ...(calculatedColumns ? { calculatedColumns } : {}) };
}
export function intersectTables(a, b) { return !(a.r2 < b.r1 || b.r2 < a.r1 || a.c2 < b.c1 || b.c2 < a.c1); }
export function tableAt(tables, r, c, rows, cols) {
    return tables.find(t => { const b = tableRange(t.range, rows, cols); return r >= b.r1 && r <= b.r2 && c >= b.c1 && c <= b.c2; });
}
const palette = {
    TableStyleMedium2: { dark: '#1d4f91', light: '#d8e7f7', soft: '#f0f5fb', text: '#ffffff' },
    TableStyleMedium4: { dark: '#167b69', light: '#cceee5', soft: '#ebfaf5', text: '#ffffff' },
    TableStyleMedium7: { dark: '#8051ac', light: '#e6d8f4', soft: '#f7f1fc', text: '#ffffff' },
    TableStyleMedium9: { dark: '#bb552d', light: '#f8ddce', soft: '#fff5ee', text: '#ffffff' },
    TableStyleMedium15: { dark: '#4d647e', light: '#dbe3ed', soft: '#f2f6f9', text: '#ffffff' },
    TableStyleMedium21: { dark: '#ba8b22', light: '#faedc7', soft: '#fffae9', text: '#172c3b' },
    TableStyleLight1: { dark: '#d8e4f0', light: '#eef3f9', soft: '#ffffff', text: '#1d3349' },
    TableStyleLight9: { dark: '#cfe8d7', light: '#ebf7ee', soft: '#ffffff', text: '#1d4031' },
    TableStyleDark1: { dark: '#203859', light: '#41628b', soft: '#233c5c', text: '#ffffff' },
    TableStyleDark4: { dark: '#243b38', light: '#467568', soft: '#2e5149', text: '#ffffff' }
};
export function tableAppearance(table, b, r, c) {
    if (r < b.r1 || r > b.r2 || c < b.c1 || c > b.c2)
        return null;
    const theme = palette[table.style], header = table.header && r === b.r1;
    if (header)
        return { background: theme.dark, color: theme.text, bold: true };
    const rowBand = table.bandedRows && ((r - b.r1 - (table.header ? 1 : 0)) % 2 === 1);
    const colBand = table.bandedColumns && ((c - b.c1) % 2 === 1);
    return { background: rowBand || colBand ? theme.light : theme.soft, bold: table.firstColumn && c === b.c1 || table.lastColumn && c === b.c2 };
}
export function tableColumns(book, table) {
    const r = tableRange(table.range, book.rows, book.cols);
    const names = new Set();
    return Array.from({ length: r.c2 - r.c1 + 1 }, (_, i) => {
        let name = table.header ? String(book.getValue(`${colToName(r.c1 + i)}${r.r1 + 1}`).value ?? '').trim() : '';
        name = name.replace(/[\x00-\x1f]/g, ' ').slice(0, 120) || `Column${i + 1}`;
        const base = name;
        let j = 2;
        while (names.has(name.toLowerCase()))
            name = `${base.slice(0, 110)}_${j++}`;
        names.add(name.toLowerCase());
        return name;
    });
}
/** Convert supported Excel same-sheet structured references to ordinary cell references for
 *  the native formula parser. The original expression is retained for editing and XLSX export.
 *  Supported: Table[Column], Table[#All/#Data/#Headers],
 *  Table[[#All/#Data/#Headers],[Column]], and [@Column] inside a table.
 *  Quoted string literals are never interpreted as references. */
export function resolveStructuredFormula(raw, book, tables, location) {
    if (!raw.startsWith('=') || !raw.includes('[') || !tables.length)
        return raw;
    const context = parseRef(location), active = context ? tableAt(tables, context.row, context.col, book.rows, book.cols) : undefined;
    function findTable(name) { return tables.find(t => t.name.toLowerCase() === name.toLowerCase()); }
    function resolved(table, colName, qualifier) {
        const b = tableRange(table.range, book.rows, book.cols);
        let c1 = b.c1, c2 = b.c2;
        if (colName) {
            const columns = tableColumns(book, table), ix = columns.findIndex(v => v.toLowerCase() === colName.trim().toLowerCase());
            if (ix < 0)
                throw Error(`Unknown table column '${colName}' in ${table.name}.`);
            c1 = c2 = b.c1 + ix;
        }
        const q = (qualifier ?? '#Data').toLowerCase();
        if (!['#all', '#data', '#headers'].includes(q))
            throw Error(`Unsupported table selector: ${qualifier}`);
        const r1 = q === '#all' || q === '#headers' ? b.r1 : b.r1 + (table.header ? 1 : 0);
        const r2 = q === '#headers' ? b.r1 : b.r2;
        return `${colToName(c1)}${r1 + 1}:${colToName(c2)}${r2 + 1}`;
    }
    // Process outside double-quoted Excel string literals, including doubled internal quotes.
    const sections = raw.match(/"(?:[^"]|"")*"|[^\"]+/g) ?? [raw];
    return sections.map(section => {
        if (section.startsWith('"'))
            return section;
        return section
            .replace(/\b([A-Za-z_][\w.]*)\[\[(#All|#Data|#Headers)\],\[([^\]\[]+)\]\]/gi, (all, name, q, col) => { const t = findTable(name); return t ? resolved(t, col, q) : all; })
            .replace(/\b([A-Za-z_][\w.]*)\[(#All|#Data|#Headers)\]/gi, (all, name, q) => { const t = findTable(name); return t ? resolved(t, undefined, q) : all; })
            .replace(/\b([A-Za-z_][\w.]*)\[([^\]\[]+)\]/gi, (all, name, col) => { const t = findTable(name); return t ? resolved(t, col, '#Data') : all; })
            .replace(/\[@([^\]\[]+)\]/g, (all, col) => {
            if (!active || !context || context.row === tableRange(active.range, book.rows, book.cols).r1)
                return all;
            const b = tableRange(active.range, book.rows, book.cols), names = tableColumns(book, active), index = names.findIndex(n => n.toLowerCase() === col.toLowerCase());
            if (index < 0)
                throw Error(`Unknown current-row column '${col}'.`);
            return `${colToName(b.c1 + index)}${context.row + 1}`;
        });
    }).join('');
}
