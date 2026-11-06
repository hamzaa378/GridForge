/** Bounded local tabular model for worksheet-backed tables and a safe DAX measure subset. */
import { colToName, parseRef } from './parser.js';
import { display } from './workbook.js';
import { excelDateSerial, excelSerialDateIso } from './date_serial.js';
const NAME = /^[A-Za-z_][A-Za-z0-9_.]{0,127}$/;
const MAX_MODEL_SOURCE_CELLS = 1_000_000;
const MAX_DAX_FILTER_CHECKS = 8_000_000;
const MAX_DAX_VIRTUAL_CELLS = 1_000_000;
const RELATED_LOOKUPS = new WeakMap();
const RELATED_PATHS = new WeakMap();
const key = (value) => value.toLocaleLowerCase();
export function emptyDataModel() { return { tables: [], relationships: [], measures: [] }; }
function refBounds(value, rows, cols) { const parts = value.replace(/\$/g, '').toUpperCase().split(':'), a = parseRef(parts[0] ?? ''), b = parseRef(parts[1] ?? parts[0] ?? ''); if (parts.length > 2 || !a || !b || a.row >= b.row || a.col > b.col || b.row >= rows || b.col >= cols || (b.row - a.row + 1) * (b.col - a.col + 1) > 500_000)
    throw Error('Model tables need a header row and an in-bounds range of at most 500,000 cells.'); return { a, b, range: `${colToName(a.col)}${a.row + 1}:${colToName(b.col)}${b.row + 1}` }; }
/** Validate persisted model metadata without trusting imported workbook JSON. */
export function validateDataModel(value) {
    if (!value || typeof value !== 'object' || Array.isArray(value))
        throw Error('Invalid Data Model definition.');
    const raw = value, tables = raw.tables, relationships = raw.relationships, measures = raw.measures;
    if (!Array.isArray(tables) || tables.length > 64 || !Array.isArray(relationships) || relationships.length > 256 || !Array.isArray(measures) || measures.length > 256)
        throw Error('A local Data Model supports up to 64 tables, 256 relationships, and 256 measures.');
    const cleanTables = [], names = new Set();
    for (const item of tables) {
        if (!item || typeof item !== 'object')
            throw Error('Invalid Data Model table.');
        const table = item;
        if (typeof table.name !== 'string' || !NAME.test(table.name) || names.has(key(table.name)) || typeof table.sheet !== 'string' || !table.sheet.trim() || table.sheet.length > 31 || /[\\/?*\[\]:]/.test(table.sheet) || typeof table.range !== 'string')
            throw Error('Model table names, worksheet names, or source ranges are invalid.');
        names.add(key(table.name));
        cleanTables.push({ name: table.name, sheet: table.sheet, range: table.range.toUpperCase().replace(/\$/g, '') });
    }
    const cleanRelationships = [], edges = new Set();
    for (const item of relationships) {
        if (!item || typeof item !== 'object')
            throw Error('Invalid model relationship.');
        const rel = item;
        if (typeof rel.manyTable !== 'string' || typeof rel.oneTable !== 'string' || typeof rel.manyColumn !== 'string' || typeof rel.oneColumn !== 'string')
            throw Error('A relationship needs text table and column names.');
        const from = key(rel.manyTable), to = key(rel.oneTable), edge = `${from}.${key(rel.manyColumn)}>${to}.${key(rel.oneColumn)}`;
        if (!names.has(from) || !names.has(to) || from === to || !rel.manyColumn || rel.manyColumn.length > 255 || !rel.oneColumn || rel.oneColumn.length > 255 || rel.active !== undefined && typeof rel.active !== 'boolean' || edges.has(edge))
            throw Error('A relationship needs distinct existing tables and valid, unique columns.');
        edges.add(edge);
        cleanRelationships.push({ manyTable: rel.manyTable, manyColumn: rel.manyColumn, oneTable: rel.oneTable, oneColumn: rel.oneColumn, ...(rel.active === false ? { active: false } : {}) });
    }
    const cleanMeasures = [], measureNames = new Set();
    for (const item of measures) {
        if (!item || typeof item !== 'object')
            throw Error('Invalid DAX measure.');
        const measure = item;
        if (typeof measure.name !== 'string' || !NAME.test(measure.name) || measureNames.has(key(measure.name)) || typeof measure.expression !== 'string' || !measure.expression.trim() || measure.expression.length > 4_000 || measure.format !== undefined && !['general', 'number', 'currency', 'percent', 'integer'].includes(measure.format))
            throw Error('DAX measures need a unique name and a bounded expression.');
        validateDaxExpression(measure.expression);
        measureNames.add(key(measure.name));
        cleanMeasures.push({ name: measure.name, expression: measure.expression.trim(), ...(measure.format ? { format: measure.format } : {}) });
    }
    return { tables: cleanTables, relationships: cleanRelationships, measures: cleanMeasures };
}
/** Read model tables as typed scalar rows from the named worksheet ranges. */
export function buildDataModel(sheets, definition) {
    const spec = validateDataModel(definition), tables = new Map();
    let totalSourceCells = 0;
    for (const source of spec.tables) {
        const sheet = sheets.find(item => key(item.name) === key(source.sheet));
        if (!sheet)
            throw Error(`Model table “${source.name}” refers to missing worksheet “${source.sheet}”.`);
        const { a, b, range } = refBounds(source.range, sheet.book.rows, sheet.book.cols), sourceCells = (b.row - a.row + 1) * (b.col - a.col + 1);
        totalSourceCells += sourceCells;
        if (totalSourceCells > MAX_MODEL_SOURCE_CELLS)
            throw Error('All local Data Model source ranges together are limited to 1,000,000 cells to keep refresh and measure calculation responsive.');
        const columns = [], seen = new Set();
        for (let col = a.col; col <= b.col; col++) {
            const column = display(sheet.book.getValue(`${colToName(col)}${a.row + 1}`)).trim();
            if (!column || seen.has(key(column)))
                throw Error(`Model table “${source.name}” needs nonblank, unique column headers.`);
            seen.add(key(column));
            columns.push(column);
        }
        const rows = [];
        for (let row = a.row + 1; row <= b.row; row++) {
            const values = columns.map((_, index) => { const result = sheet.book.getValue(`${colToName(a.col + index)}${row + 1}`); if (result.error)
                throw Error(`Model table “${source.name}” contains a formula error at row ${row + 1}.`); return result.value; });
            if (values.some(value => value !== null && value !== ''))
                rows.push(values);
        }
        if (rows.length > 500_000)
            throw Error(`Model table “${source.name}” exceeds 500,000 data rows.`);
        tables.set(key(source.name), { name: source.name, columns, rows });
    }
    const runtime = { tables, relationships: spec.relationships.filter(item => item.active !== false), measures: spec.measures };
    for (const rel of runtime.relationships) {
        const many = tables.get(key(rel.manyTable)), one = tables.get(key(rel.oneTable));
        if (!many || !one)
            throw Error('A relationship refers to a table absent from the current model.');
        const manyIndex = many.columns.findIndex(column => key(column) === key(rel.manyColumn)), oneIndex = one.columns.findIndex(column => key(column) === key(rel.oneColumn));
        if (manyIndex < 0 || oneIndex < 0)
            throw Error(`Relationship columns are missing: ${rel.manyTable}[${rel.manyColumn}] → ${rel.oneTable}[${rel.oneColumn}].`);
        const unique = new Set();
        for (const row of one.rows) {
            const value = row[oneIndex];
            if (value === null || value === '')
                continue;
            const token = scalarKey(value);
            if (unique.has(token))
                throw Error(`The one-side relationship key ${rel.oneTable}[${rel.oneColumn}] must contain unique values.`);
            unique.add(token);
        }
    }
    const outgoing = new Map();
    for (const rel of runtime.relationships) {
        const from = key(rel.oneTable), to = key(rel.manyTable);
        outgoing.set(from, [...(outgoing.get(from) ?? []), to]);
    }
    const visiting = new Set(), visited = new Set();
    const visit = (node) => { if (visiting.has(node))
        throw Error('Active model relationships cannot contain a directed cycle in this DAX subset.'); if (visited.has(node))
        return; visiting.add(node); for (const next of outgoing.get(node) ?? [])
        visit(next); visiting.delete(node); visited.add(node); };
    for (const name of tables.keys())
        visit(name);
    return runtime;
}
const scalarKey = (value) => `${typeof value}:${String(value)}`;
const daxString = (value) => value === null ? '' : typeof value === 'boolean' ? value ? 'TRUE' : 'FALSE' : String(value);
const DAX_NUMERIC_FUNCTIONS = new Set(['INT', 'TRUNC', 'ROUNDUP', 'ROUNDDOWN', 'MOD', 'POWER', 'SQRT']);
const DAX_TEXT_FUNCTIONS = new Set(['LEN', 'LEFT', 'RIGHT', 'TRIM', 'UPPER', 'LOWER', 'CONCATENATE']);
const DAX_DATE_FUNCTIONS = new Set(['DATE', 'YEAR', 'MONTH', 'DAY', 'QUARTER', 'WEEKDAY', 'WEEKNUM']);
function evaluateDaxDateFunction(name, values) {
    if (name === 'DATE') {
        if (values.length !== 3 || values.some(value => typeof value !== 'number' || !Number.isFinite(value)))
            throw Error('DATE expects finite numeric year, month, and day arguments.');
        let year = Math.trunc(values[0]);
        const month = Math.trunc(values[1]), day = Math.trunc(values[2]);
        if (year < 0 || year > 9999 || Math.abs(month) > 120_000 || Math.abs(day) > 1_000_000)
            throw Error('DATE arguments are outside the supported calendar range.');
        if (year < 1900)
            year += 1900;
        if (year === 1900 && month === 2 && day === 29)
            return 60;
        const date = new Date(Date.UTC(year, month - 1, day));
        if (!Number.isFinite(date.getTime()) || date.getUTCFullYear() < 1900 || date.getUTCFullYear() > 9999)
            throw Error('DATE result is outside the supported 1900–9999 date range.');
        const iso = `${String(date.getUTCFullYear()).padStart(4, '0')}-${String(date.getUTCMonth() + 1).padStart(2, '0')}-${String(date.getUTCDate()).padStart(2, '0')}`, serial = excelDateSerial(iso);
        if (serial === null)
            throw Error('DATE could not convert its result to an Excel date serial.');
        return serial;
    }
    if (values.length < 1 || values.length > (name === 'WEEKDAY' || name === 'WEEKNUM' ? 2 : 1))
        throw Error(`${name} expects ${name === 'WEEKDAY' || name === 'WEEKNUM' ? 'a date and optional return type' : 'one date argument'}.`);
    const input = values[0];
    if (input === null)
        return null;
    let serial = null;
    if (typeof input === 'number')
        serial = input;
    else if (typeof input === 'string')
        serial = excelDateSerial(input);
    if (serial === null || !Number.isFinite(serial) || serial < 0 || serial > 2_958_465)
        throw Error(`${name} requires an Excel date serial or an unambiguous ISO date (YYYY-MM-DD).`);
    const iso = excelSerialDateIso(serial);
    if (!iso)
        throw Error(`${name} received a date outside the supported 1900 date system.`);
    const year = Number(iso.slice(0, 4)), month = Number(iso.slice(5, 7)), day = Number(iso.slice(8, 10));
    switch (name) {
        case 'YEAR': return year;
        case 'MONTH': return month;
        case 'DAY': return day;
        case 'QUARTER': return Math.floor((month - 1) / 3) + 1;
        case 'WEEKDAY': {
            const returnType = values.length === 2 ? values[1] : 1;
            if (typeof returnType !== 'number' || !Number.isInteger(returnType) || ![1, 2, 3].includes(returnType))
                throw Error('WEEKDAY return type must be 1, 2, or 3.');
            const sundayZero = iso === '1900-02-29' ? 4 : new Date(`${iso}T00:00:00Z`).getUTCDay();
            return returnType === 1 ? sundayZero + 1 : returnType === 2 ? (sundayZero + 6) % 7 + 1 : (sundayZero + 6) % 7;
        }
        default: {
            const returnType = values.length === 2 ? values[1] : 1, weekStarts = { 1: 0, 2: 1, 11: 1, 12: 2, 13: 3, 14: 4, 15: 5, 16: 6, 17: 0 };
            if (typeof returnType !== 'number' || !Number.isInteger(returnType) || returnType !== 21 && !Object.hasOwn(weekStarts, returnType))
                throw Error('WEEKNUM return type must be 1, 2, or 11 through 17, or 21 (ISO 8601).');
            if (returnType === 21) {
                const date = new Date(Date.UTC(year, month - 1, day)), weekday = date.getUTCDay() || 7;
                date.setUTCDate(date.getUTCDate() + 4 - weekday);
                const yearStart = Date.UTC(date.getUTCFullYear(), 0, 1);
                return Math.ceil(((date.getTime() - yearStart) / 86_400_000 + 1) / 7);
            }
            const whole = Math.floor(serial), jan1 = excelDateSerial(`${String(year).padStart(4, '0')}-01-01`);
            if (jan1 === null)
                throw Error('WEEKNUM could not resolve the first day of the date year.');
            const jan1Weekday = new Date(Date.UTC(year, 0, 1)).getUTCDay(), offset = (jan1Weekday - weekStarts[returnType] + 7) % 7, dayOfYear = whole - Math.floor(jan1) + 1;
            return Math.floor((dayOfYear - 1 + offset) / 7) + 1;
        }
    }
}
function evaluateDaxTextFunction(name, values) {
    if (name === 'CONCATENATE') {
        if (values.length !== 2)
            throw Error('CONCATENATE expects exactly two arguments.');
        return daxString(values[0] ?? null) + daxString(values[1] ?? null);
    }
    const expected = name === 'LEFT' || name === 'RIGHT' ? [1, 2] : [1, 1];
    if (values.length < expected[0] || values.length > expected[1])
        throw Error(`${name} expects ${expected[0] === expected[1] ? `exactly ${expected[0]}` : 'one or two'} arguments.`);
    const input = values[0];
    if (input !== null && typeof input !== 'string')
        throw Error(`${name} expects text or BLANK().`);
    const text = input ?? '';
    if (name === 'LEN')
        return text.length;
    if (name === 'UPPER')
        return input === null ? null : text.toUpperCase();
    if (name === 'LOWER')
        return input === null ? null : text.toLowerCase();
    if (name === 'TRIM')
        return input === null ? null : text.trim().replace(/ +/g, ' ');
    const count = values.length === 1 ? 1 : values[1];
    if (typeof count !== 'number' || !Number.isInteger(count) || count < 0 || count > 32767)
        throw Error(`${name} character count must be a whole number from 0 to 32,767.`);
    const chars = Array.from(text);
    return (name === 'LEFT' ? chars.slice(0, count) : count ? chars.slice(-count) : []).join('');
}
function daxPercentileInc(values, percentile) {
    if (typeof percentile !== 'number' || !Number.isFinite(percentile) || percentile < 0 || percentile > 1)
        throw Error('PERCENTILE.INC percentile must be a number from 0 through 1.');
    if (!values.length)
        return null;
    const sorted = [...values].sort((a, b) => a - b), position = percentile * (sorted.length - 1), lower = Math.floor(position), fraction = position - lower, low = sorted[lower], high = sorted[Math.min(lower + 1, sorted.length - 1)];
    const result = low * (1 - fraction) + high * fraction;
    if (!Number.isFinite(result))
        throw Error('PERCENTILE.INC produced a non-finite value.');
    return result;
}
function daxPercentileExc(values, percentile) {
    if (typeof percentile !== 'number' || !Number.isFinite(percentile) || percentile <= 0 || percentile >= 1)
        throw Error('PERCENTILE.EXC percentile must be greater than 0 and less than 1.');
    if (values.length < 2)
        return null;
    const sorted = [...values].sort((a, b) => a - b), position = percentile * (sorted.length + 1), rank = Math.floor(position), fraction = position - rank;
    if (rank < 1 || rank > sorted.length)
        throw Error('PERCENTILE.EXC percentile is outside the range supported by the number of numeric values.');
    const low = sorted[rank - 1], high = sorted[Math.min(rank, sorted.length - 1)], result = fraction === 0 ? low : low * (1 - fraction) + high * fraction;
    if (!Number.isFinite(result))
        throw Error('PERCENTILE.EXC produced a non-finite value.');
    return result;
}
function daxVariance(values, sample, name, standardDeviation) {
    if (sample && values.length < 2)
        throw Error(`${name} requires at least two nonblank numeric values.`);
    if (!values.length)
        return null;
    let count = 0, mean = 0, sumSquares = 0;
    for (const value of values) {
        count++;
        const delta = value - mean;
        mean += delta / count;
        const adjusted = value - mean;
        sumSquares += delta * adjusted;
        if (!Number.isFinite(mean) || !Number.isFinite(sumSquares))
            throw Error(`${name} produced a non-finite result.`);
    }
    const variance = sumSquares / (sample ? count - 1 : count);
    if (!Number.isFinite(variance) || variance < 0)
        throw Error(`${name} produced an invalid variance.`);
    const result = standardDeviation ? Math.sqrt(variance) : variance;
    if (!Number.isFinite(result))
        throw Error(`${name} produced a non-finite result.`);
    return result;
}
function evaluateDaxNumericFunction(name, values) {
    const expected = name === 'TRUNC' || name === 'ROUNDUP' || name === 'ROUNDDOWN' ? [1, 2] : name === 'MOD' || name === 'POWER' ? [2, 2] : [1, 1];
    if (values.length < expected[0] || values.length > expected[1])
        throw Error(`${name} expects ${expected[0] === expected[1] ? `exactly ${expected[0]}` : 'one or two'} arguments.`);
    const numberAt = (index, defaultValue) => { const value = values[index] ?? defaultValue; if (typeof value !== 'number' || !Number.isFinite(value))
        throw Error(`${name} expects finite numeric arguments.`); return value; };
    if (name === 'INT')
        return Math.floor(numberAt(0));
    if (name === 'SQRT') {
        const value = numberAt(0);
        if (value < 0)
            throw Error('DAX SQRT requires a non-negative value.');
        return Math.sqrt(value);
    }
    if (name === 'POWER') {
        const value = Math.pow(numberAt(0), numberAt(1));
        if (!Number.isFinite(value))
            throw Error('DAX POWER produced a non-finite result.');
        return value;
    }
    if (name === 'MOD') {
        const dividend = numberAt(0), divisor = numberAt(1);
        if (divisor === 0)
            throw Error('DAX MOD divisor cannot be zero.');
        const value = dividend - divisor * Math.floor(dividend / divisor);
        if (!Number.isFinite(value))
            throw Error('DAX MOD produced a non-finite result.');
        return Object.is(value, -0) ? 0 : value;
    }
    const value = numberAt(0), digits = numberAt(1, 0);
    if (!Number.isInteger(digits) || Math.abs(digits) > 12)
        throw Error(`${name} digit count must be an integer from -12 to 12.`);
    if (name === 'TRUNC') {
        const factor = 10 ** digits, result = digits >= 0 ? Math.trunc(value * factor) / factor : Math.trunc(value / factor) * factor;
        if (!Number.isFinite(result))
            throw Error('DAX TRUNC produced a non-finite result.');
        return result;
    }
    const factor = 10 ** digits, scaled = digits >= 0 ? Math.abs(value) * factor : Math.abs(value) / factor;
    if (!Number.isFinite(scaled))
        throw Error(`${name} produced a non-finite intermediate result.`);
    const rounded = (name === 'ROUNDUP' ? Math.ceil : Math.floor)(scaled), result = Math.sign(value) * (digits >= 0 ? rounded / factor : rounded * factor);
    if (!Number.isFinite(result))
        throw Error(`${name} produced a non-finite result.`);
    return result;
}
function daxTokens(code) {
    if (code.length > 4000)
        throw Error('DAX expression is limited to 4,000 characters.');
    const out = [];
    let i = 0;
    while (i < code.length) {
        const ch = code[i];
        if (/\s/.test(ch)) {
            i++;
            continue;
        }
        if (ch === '"') {
            let value = '';
            i++;
            let closed = false;
            while (i < code.length) {
                if (code[i] === '"') {
                    if (code[i + 1] === '"') {
                        value += '"';
                        i += 2;
                        continue;
                    }
                    i++;
                    closed = true;
                    break;
                }
                value += code[i++];
            }
            if (!closed)
                throw Error('Unclosed DAX text literal.');
            out.push({ kind: 'string', value });
            continue;
        }
        if (ch === '[') {
            const close = code.indexOf(']', i + 1);
            if (close < 0)
                throw Error('Unclosed DAX column reference.');
            const column = code.slice(i + 1, close).trim();
            if (!column || column.length > 255)
                throw Error('DAX column names must contain 1–255 characters.');
            out.push({ kind: 'symbol', value: '[' }, { kind: 'id', value: column }, { kind: 'symbol', value: ']' });
            i = close + 1;
            continue;
        }
        if (/[A-Za-z_]/.test(ch)) {
            const start = i++;
            while (i < code.length && /[A-Za-z0-9_.]/.test(code[i]))
                i++;
            out.push({ kind: 'id', value: code.slice(start, i) });
            continue;
        }
        if (/[0-9]/.test(ch)) {
            const start = i++;
            while (i < code.length && /[0-9.]/.test(code[i]))
                i++;
            const value = code.slice(start, i);
            if (!Number.isFinite(Number(value)))
                throw Error('Invalid DAX number.');
            out.push({ kind: 'number', value });
            continue;
        }
        if (['<=', '>=', '<>', '&&', '||'].includes(code.slice(i, i + 2))) {
            out.push({ kind: 'symbol', value: code.slice(i, i + 2) });
            i += 2;
            continue;
        }
        if ('[](),+-*/=<>&|'.includes(ch)) {
            out.push({ kind: 'symbol', value: ch });
            i++;
            continue;
        }
        throw Error(`Unsupported DAX character “${ch}”.`);
    }
    if (out.length > 2000)
        throw Error('DAX expression is limited to 2,000 tokens.');
    let nesting = 0, unaryRun = 0;
    for (const token of out) {
        if (token.kind === 'symbol' && token.value === '(') {
            if (++nesting > 64)
                throw Error('DAX expression nesting is limited to 64 levels.');
            unaryRun = 0;
        }
        else if (token.kind === 'symbol' && token.value === ')')
            nesting = Math.max(0, nesting - 1);
        if (token.kind === 'symbol' && (token.value === '+' || token.value === '-') || token.kind === 'id' && token.value.toUpperCase() === 'NOT') {
            if (++unaryRun > 64)
                throw Error('DAX unary operator nesting is limited to 64 levels.');
        }
        else
            unaryRun = 0;
    }
    return out;
}
class DaxParser {
    tokens;
    i = 0;
    constructor(tokens) {
        this.tokens = tokens;
    }
    peek() { return this.tokens[this.i]; }
    take() { return this.tokens[this.i++]; }
    expect(value) { if (this.take().value !== value)
        throw Error(`Expected “${value}” in DAX expression.`); }
    parse(min = 0) { let left = this.primary(); const levels = { '||': 1, '&&': 2, '=': 3, '<>': 3, '<': 3, '>': 3, '<=': 3, '>=': 3, '&': 3.5, '+': 4, '-': 4, '*': 5, '/': 5 }; while ((levels[this.peek()?.value ?? ''] ?? 0) >= min && (levels[this.peek()?.value ?? ''] ?? 0) > 0) {
        const op = this.take().value, p = levels[op];
        left = { kind: 'binary', op, left, right: this.parse(p + 1) };
    } return left; }
    finish() { const result = this.parse(); if (this.peek())
        throw Error('Unexpected token in DAX expression.'); return result; }
    primary() { const token = this.take(); if (token.value === '-' || token.value === '+')
        return { kind: 'unary', op: token.value, value: this.primary() }; if (token.kind === 'id' && token.value.toUpperCase() === 'NOT' && this.peek()?.value !== '(')
        return { kind: 'unary', op: 'NOT', value: this.primary() }; if (token.kind === 'string')
        return { kind: 'literal', value: token.value }; if (token.kind === 'number')
        return { kind: 'literal', value: Number(token.value) }; if (token.value === '(') {
        const node = this.parse();
        this.expect(')');
        return node;
    } if (token.kind !== 'id')
        throw Error(`Expected a DAX name near “${token.value}”.`); if (token.value.toLowerCase() === 'true' || token.value.toLowerCase() === 'false')
        return { kind: 'literal', value: token.value.toLowerCase() === 'true' }; if (this.peek()?.value === '[') {
        this.take();
        const column = this.take();
        if (column.kind !== 'id')
            throw Error('Expected a DAX column name.');
        this.expect(']');
        return { kind: 'column', table: token.value, column: column.value };
    } if (this.peek()?.value === '(') {
        this.take();
        const args = [];
        if (this.peek()?.value !== ')') {
            do {
                args.push(this.parse());
                if (args.length > 16)
                    throw Error('DAX function calls are limited to 16 arguments.');
                if (this.peek()?.value !== ',')
                    break;
                this.take();
            } while (true);
        }
        this.expect(')');
        return { kind: 'call', name: token.value.toUpperCase(), args };
    } return { kind: 'table', value: token.value }; }
}
function validateDaxExpression(expression) { new DaxParser(daxTokens(expression)).finish(); }
function relatedPath(runtime, fromTable, targetTable) {
    let cache = RELATED_PATHS.get(runtime);
    if (!cache) {
        cache = new Map();
        RELATED_PATHS.set(runtime, cache);
    }
    const cacheKey = `${fromTable}\u0000${targetTable}`;
    if (cache.has(cacheKey))
        return cache.get(cacheKey);
    const find = (tableName, visited) => { if (visited.has(tableName))
        return []; visited.add(tableName); const paths = []; for (const rel of runtime.relationships.filter(item => key(item.manyTable) === tableName)) {
        const next = key(rel.oneTable);
        if (next === targetTable)
            paths.push([rel]);
        else
            for (const tail of find(next, new Set(visited)))
                paths.push([rel, ...tail]);
        if (paths.length > 1)
            break;
    } return paths; };
    const paths = find(fromTable, new Set());
    if (paths.length > 1)
        throw Error(`RELATED from “${fromTable}” to “${targetTable}” is ambiguous because multiple active paths exist.`);
    const result = paths[0] ?? null;
    cache.set(cacheKey, result);
    return result;
}
function relatedValue(runtime, from, row, targetTableName, targetColumnName) {
    if (key(from.name) === key(targetTableName))
        throw Error('RELATED must reference a different table from the iterator table.');
    const targetKey = key(targetTableName);
    if (!runtime.tables.has(targetKey))
        throw Error(`Unknown model table “${targetTableName}”.`);
    const path = relatedPath(runtime, key(from.name), targetKey);
    if (!path)
        return null;
    let current = from, currentRow = row;
    for (const rel of path) {
        const one = runtime.tables.get(key(rel.oneTable));
        if (!one)
            return null;
        const manyIndex = current.columns.findIndex(item => key(item) === key(rel.manyColumn)), oneIndex = one.columns.findIndex(item => key(item) === key(rel.oneColumn));
        if (manyIndex < 0 || oneIndex < 0)
            throw Error('RELATED relationship columns are missing.');
        const foreignKey = currentRow[manyIndex];
        if (foreignKey === null || foreignKey === undefined || foreignKey === '')
            return null;
        let runtimeCache = RELATED_LOOKUPS.get(runtime);
        if (!runtimeCache) {
            runtimeCache = new Map();
            RELATED_LOOKUPS.set(runtime, runtimeCache);
        }
        const cacheKey = `${key(rel.oneTable)}\u0000${key(rel.oneColumn)}`;
        let index = runtimeCache.get(cacheKey);
        if (!index) {
            index = new Map();
            for (const candidate of one.rows) {
                const value = candidate[oneIndex];
                if (value === null || value === undefined || value === '')
                    continue;
                const token = scalarKey(value);
                if (!index.has(token))
                    index.set(token, candidate);
            }
            runtimeCache.set(cacheKey, index);
        }
        const match = index.get(scalarKey(foreignKey));
        if (!match)
            return null;
        current = one;
        currentRow = match;
    }
    const targetIndex = current.columns.findIndex(item => key(item) === key(targetColumnName));
    if (targetIndex < 0)
        throw Error(`Unknown model column “${targetTableName}[${targetColumnName}]”.`);
    const value = currentRow[targetIndex];
    return value === null || typeof value === 'string' || typeof value === 'number' || typeof value === 'boolean' ? value : null;
}
function rowsFor(runtime, table, context) {
    // Propagate filters through one-to-many relationships to a fixed point. This
    // supports dimension chains (Date -> Calendar -> Sales) while preserving the
    // model's single-direction relationship semantics.
    const restrictions = new Map();
    for (const [field, values] of context.filters)
        restrictions.set(field, new Set(values));
    const compiled = new Map();
    const tableRestrictions = (modelTable) => { const tableName = key(modelTable.name), cached = compiled.get(tableName); if (cached)
        return cached; const result = []; for (const [field, values] of restrictions) {
        const [name, column] = field.split('\u0000');
        if (name !== tableName)
            continue;
        const index = modelTable.columns.findIndex(item => key(item) === key(column ?? ''));
        if (index < 0)
            throw Error(`A DAX filter refers to a missing column in “${modelTable.name}”.`);
        result.push({ index, values });
    } compiled.set(tableName, result); return result; };
    let filterChecks = 0;
    const rowMatches = (modelTable, row) => { for (const { index, values } of tableRestrictions(modelTable)) {
        if (++filterChecks > MAX_DAX_FILTER_CHECKS)
            throw Error('DAX relationship filtering exceeded its 8-million cell-check budget. Reduce model rows, active relationships, or simultaneous filters.');
        if (!values.has(scalarKey(row[index])))
            return false;
    } return true; };
    const rounds = Math.min(runtime.relationships.length + 1, 64);
    for (let round = 0; round < rounds; round++) {
        let changed = false;
        for (const rel of runtime.relationships) {
            const one = runtime.tables.get(key(rel.oneTable)), many = runtime.tables.get(key(rel.manyTable));
            if (!one || !many)
                continue;
            const oneColumn = one.columns.findIndex(item => key(item) === key(rel.oneColumn)), manyColumn = many.columns.findIndex(item => key(item) === key(rel.manyColumn));
            if (oneColumn < 0 || manyColumn < 0)
                continue;
            const oneHasFilter = tableRestrictions(one).length > 0;
            if (!oneHasFilter)
                continue;
            const selected = new Set(one.rows.filter(row => rowMatches(one, row)).map(row => scalarKey(row[oneColumn]))), field = `${key(many.name)}\u0000${key(many.columns[manyColumn])}`, prior = restrictions.get(field);
            if (prior) {
                const narrowed = new Set([...prior].filter(value => selected.has(value)));
                if (narrowed.size !== prior.size) {
                    restrictions.set(field, narrowed);
                    compiled.delete(key(many.name));
                    changed = true;
                }
            }
            else {
                restrictions.set(field, selected);
                compiled.delete(key(many.name));
                changed = true;
            }
        }
        if (!changed)
            break;
    }
    return table.rows.filter(row => rowMatches(table, row));
}
/** Evaluate scalar DAX inside a single-table iterator row; context transition is intentionally absent. */
function evaluateDaxIteratorExpression(node, runtime, context, table, row) {
    if (node.kind === 'literal')
        return node.value;
    if (node.kind === 'column') {
        if (key(node.table) !== key(table.name))
            throw Error('DAX iterator expressions in this subset can reference columns from their current table only.');
        const index = table.columns.findIndex(item => key(item) === key(node.column));
        if (index < 0)
            throw Error(`Unknown model column “${node.table}[${node.column}]”.`);
        const value = row[index];
        return value === null || typeof value === 'string' || typeof value === 'number' || typeof value === 'boolean' ? value : null;
    }
    if (node.kind === 'table')
        throw Error('DAX iterator expressions cannot use table values as scalars.');
    if (node.kind === 'unary') {
        const value = evaluateDaxIteratorExpression(node.value, runtime, context, table, row);
        if (node.op === 'NOT') {
            if (typeof value !== 'boolean')
                throw Error('DAX iterator NOT requires a Boolean.');
            return !value;
        }
        if (typeof value !== 'number')
            throw Error('DAX iterator unary signs require numbers.');
        return node.op === '-' ? -value : value;
    }
    if (node.kind === 'binary') {
        const a = evaluateDaxIteratorExpression(node.left, runtime, context, table, row);
        if (node.op === '&&' || node.op === '||') {
            if (typeof a !== 'boolean')
                throw Error(`DAX iterator ${node.op} requires Boolean operands.`);
            if (node.op === '&&' && !a || node.op === '||' && a)
                return a;
            const b = evaluateDaxIteratorExpression(node.right, runtime, context, table, row);
            if (typeof b !== 'boolean')
                throw Error(`DAX iterator ${node.op} requires Boolean operands.`);
            return b;
        }
        const b = evaluateDaxIteratorExpression(node.right, runtime, context, table, row);
        if (node.op === '&')
            return daxString(a) + daxString(b);
        if (['=', '<>', '<', '>', '<=', '>='].includes(node.op)) {
            const order = typeof a === 'number' && typeof b === 'number' ? a - b : String(a ?? '').localeCompare(String(b ?? ''), undefined, { numeric: true, sensitivity: 'base' });
            switch (node.op) {
                case '=': return a === b;
                case '<>': return a !== b;
                case '<': return order < 0;
                case '>': return order > 0;
                case '<=': return order <= 0;
                default: return order >= 0;
            }
        }
        if (typeof a !== 'number' || typeof b !== 'number')
            throw Error(`DAX iterator ${node.op} requires numeric values.`);
        let result;
        if (node.op === '+')
            result = a + b;
        else if (node.op === '-')
            result = a - b;
        else if (node.op === '*')
            result = a * b;
        else if (b === 0)
            return null;
        else
            result = a / b;
        if (!Number.isFinite(result))
            throw Error('DAX iterator arithmetic produced a non-finite value.');
        return result;
    }
    const values = node.args, evalArg = (index) => values[index] ? evaluateDaxIteratorExpression(values[index], runtime, context, table, row) : null;
    switch (node.name) {
        case 'RELATED': {
            if (values.length !== 1 || values[0].kind !== 'column')
                throw Error('RELATED expects one related-table column reference.');
            const column = values[0];
            return relatedValue(runtime, table, row, column.table, column.column);
        }
        case 'NOT': {
            if (values.length !== 1)
                throw Error('NOT expects one argument.');
            const value = evalArg(0);
            if (typeof value !== 'boolean')
                throw Error('NOT requires a Boolean.');
            return !value;
        }
        case 'AND':
        case 'OR': {
            if (values.length !== 2)
                throw Error(`${node.name} expects two arguments.`);
            const left = evalArg(0);
            if (typeof left !== 'boolean')
                throw Error(`${node.name} requires Boolean arguments.`);
            if (node.name === 'AND' && !left || node.name === 'OR' && left)
                return left;
            const right = evalArg(1);
            if (typeof right !== 'boolean')
                throw Error(`${node.name} requires Boolean arguments.`);
            return right;
        }
        case 'IF': {
            if (values.length < 2 || values.length > 3)
                throw Error('IF expects two or three arguments.');
            const condition = evalArg(0), truth = typeof condition === 'boolean' ? condition : typeof condition === 'number' ? condition !== 0 : condition !== null && condition !== '';
            return truth ? evalArg(1) : values[2] ? evalArg(2) : false;
        }
        case 'DIVIDE': {
            if (values.length < 2 || values.length > 3)
                throw Error('DIVIDE expects two or three arguments.');
            const numerator = evalArg(0), denominator = evalArg(1);
            if (typeof numerator !== 'number' || typeof denominator !== 'number' || denominator === 0)
                return values[2] ? evalArg(2) : null;
            const result = numerator / denominator;
            if (!Number.isFinite(result))
                throw Error('DAX DIVIDE produced a non-finite iterator value.');
            return result;
        }
        case 'ABS': {
            if (values.length !== 1)
                throw Error('ABS expects one argument.');
            const value = evalArg(0);
            if (typeof value !== 'number')
                throw Error('ABS expects a number.');
            return Math.abs(value);
        }
        case 'ROUND': {
            if (values.length !== 2)
                throw Error('ROUND expects two arguments.');
            const value = evalArg(0), digits = evalArg(1);
            if (typeof value !== 'number' || typeof digits !== 'number' || !Number.isInteger(digits) || Math.abs(digits) > 12)
                throw Error('ROUND needs a number and an integer digit count from -12 to 12.');
            const factor = 10 ** digits, result = Math.sign(value) * Math.floor(Math.abs(value) * factor + 0.5) / factor;
            if (!Number.isFinite(result))
                throw Error('DAX ROUND produced a non-finite iterator value.');
            return result;
        }
        case 'INT':
        case 'TRUNC':
        case 'ROUNDUP':
        case 'ROUNDDOWN':
        case 'MOD':
        case 'POWER':
        case 'SQRT': return evaluateDaxNumericFunction(node.name, values.map((_, index) => evalArg(index)));
        case 'DATE':
        case 'YEAR':
        case 'MONTH':
        case 'DAY':
        case 'QUARTER':
        case 'WEEKDAY':
        case 'WEEKNUM': return evaluateDaxDateFunction(node.name, values.map((_, index) => evalArg(index)));
        case 'LEN':
        case 'LEFT':
        case 'RIGHT':
        case 'TRIM':
        case 'UPPER':
        case 'LOWER': return evaluateDaxTextFunction(node.name, values.map((_, index) => evalArg(index)));
        case 'COALESCE':
            for (let index = 0; index < values.length && index < 16; index++) {
                const value = evalArg(index);
                if (value !== null && value !== '')
                    return value;
            }
            return null;
        default: throw Error(`DAX function “${node.name}” is not supported inside this iterator expression.`);
    }
}
function rowsForDaxIterator(source, runtime, context) {
    if (source.kind === 'table') {
        const table = runtime.tables.get(key(source.value));
        if (!table)
            throw Error(`Unknown model table “${source.value}”.`);
        return { table, rows: rowsFor(runtime, table, context) };
    }
    if (source.kind === 'call' && source.name === 'SUMMARIZE') {
        if (source.args.length < 2 || source.args.length > 16 || source.args.slice(1).some(argument => argument.kind !== 'column'))
            throw Error('SUMMARIZE in this subset expects a model table followed by one to fifteen grouping columns.');
        const base = rowsForDaxIterator(source.args[0], runtime, context), columns = source.args.slice(1), headers = [], indexes = [];
        for (const column of columns) {
            if (key(column.table) !== key(base.table.name))
                throw Error('SUMMARIZE grouping columns must belong to its input table in this subset.');
            const index = base.table.columns.findIndex(item => key(item) === key(column.column));
            if (index < 0)
                throw Error(`Unknown model column “${column.table}[${column.column}]”.`);
            if (indexes.includes(index))
                throw Error('SUMMARIZE grouping columns must be unique.');
            indexes.push(index);
            headers.push(base.table.columns[index]);
        }
        const groups = new Map();
        for (const row of base.rows) {
            const values = indexes.map(index => row[index] ?? null), token = JSON.stringify(values.map(value => value === null || value === '' ? ['blank'] : [typeof value, scalarKey(value)]));
            if (!groups.has(token)) {
                if ((groups.size + 1) * headers.length > MAX_DAX_VIRTUAL_CELLS)
                    throw Error('DAX SUMMARIZE result exceeds the 1,000,000-cell virtual table budget.');
                groups.set(token, values);
            }
        }
        const rows = [...groups.values()];
        return { table: { ...base.table, columns: headers, rows }, rows };
    }
    if (source.kind === 'call' && source.name === 'SELECTCOLUMNS') {
        if (source.args.length < 3 || source.args.length % 2 === 0)
            throw Error('SELECTCOLUMNS expects a model table followed by one or more name/expression pairs.');
        const base = rowsForDaxIterator(source.args[0], runtime, context), headers = [];
        for (let index = 1; index < source.args.length; index += 2) {
            const nameNode = source.args[index];
            if (nameNode.kind !== 'literal' || typeof nameNode.value !== 'string' || !nameNode.value.trim() || nameNode.value.length > 255 || headers.some(header => key(header) === key(nameNode.value)))
                throw Error('SELECTCOLUMNS requires unique text column names of at most 255 characters.');
            headers.push(nameNode.value);
        }
        if (base.rows.length * headers.length > MAX_DAX_VIRTUAL_CELLS)
            throw Error('DAX SELECTCOLUMNS result exceeds the 1,000,000-cell virtual table budget.');
        const rows = base.rows.map(row => headers.map((_, index) => { const value = evaluateDaxIteratorExpression(source.args[index * 2 + 2], runtime, context, base.table, row); if (value !== null && typeof value !== 'string' && typeof value !== 'number' && typeof value !== 'boolean')
            throw Error('SELECTCOLUMNS expressions in this subset must return scalar values.'); return value; }));
        return { table: { ...base.table, columns: headers, rows }, rows };
    }
    if (source.kind === 'call' && source.name === 'ADDCOLUMNS') {
        if (source.args.length < 3 || source.args.length % 2 === 0)
            throw Error('ADDCOLUMNS expects a model table followed by one or more name/expression pairs.');
        let base = rowsForDaxIterator(source.args[0], runtime, context);
        if (base.rows.length * (base.table.columns.length + (source.args.length - 1) / 2) > MAX_DAX_VIRTUAL_CELLS)
            throw Error('DAX ADDCOLUMNS result exceeds the 1,000,000-cell virtual table budget.');
        let table = base.table, rows = base.rows.map(row => [...row]);
        for (let index = 1; index < source.args.length; index += 2) {
            const nameNode = source.args[index], expression = source.args[index + 1];
            if (nameNode.kind !== 'literal' || typeof nameNode.value !== 'string' || !nameNode.value.trim() || nameNode.value.length > 255 || table.columns.some(column => key(column) === key(nameNode.value)))
                throw Error('ADDCOLUMNS requires unique text column names of at most 255 characters.');
            const name = nameNode.value;
            const values = rows.map(row => evaluateDaxIteratorExpression(expression, runtime, context, table, row));
            if (values.some(value => value !== null && typeof value !== 'string' && typeof value !== 'number' && typeof value !== 'boolean'))
                throw Error('ADDCOLUMNS expressions in this subset must return scalar values.');
            rows = rows.map((row, rowIndex) => [...row, values[rowIndex] ?? null]);
            table = { ...table, columns: [...table.columns, name], rows };
        }
        return { table, rows };
    }
    if (source.kind === 'call' && source.name === 'ALL') {
        if (source.args.length !== 1)
            throw Error('ALL as a DAX table input expects one model table or column reference.');
        const reference = source.args[0];
        if (reference.kind === 'table') {
            const table = runtime.tables.get(key(reference.value));
            if (!table)
                throw Error(`Unknown model table “${reference.value}”.`);
            const filters = new Map(context.filters), prefix = `${key(table.name)}\u0000`;
            for (const field of filters.keys())
                if (field.startsWith(prefix))
                    filters.delete(field);
            const next = { filters };
            return { table, rows: rowsFor(runtime, table, next) };
        }
        if (reference.kind === 'column') {
            const table = runtime.tables.get(key(reference.table));
            if (!table)
                throw Error(`Unknown model table “${reference.table}”.`);
            const index = table.columns.findIndex(item => key(item) === key(reference.column));
            if (index < 0)
                throw Error(`Unknown model column “${reference.table}[${reference.column}]”.`);
            const filters = new Map(context.filters);
            filters.delete(`${key(table.name)}\u0000${key(table.columns[index])}`);
            const next = { filters }, values = new Map();
            for (const row of rowsFor(runtime, table, next)) {
                const value = row[index] ?? null, token = value === null || value === '' ? '__DAX_BLANK__' : scalarKey(value);
                if (!values.has(token))
                    values.set(token, [value]);
            }
            const distinctRows = [...values.values()];
            return { table: { ...table, columns: [table.columns[index]], rows: distinctRows }, rows: distinctRows };
        }
        throw Error('ALL as a DAX table input expects one model table or column reference.');
    }
    if (source.kind === 'call' && (source.name === 'VALUES' || source.name === 'DISTINCT')) {
        if (source.args.length !== 1 || source.args[0].kind !== 'column')
            throw Error(`${source.name} as a DAX table input expects one model column reference.`);
        const column = source.args[0], table = runtime.tables.get(key(column.table));
        if (!table)
            throw Error(`Unknown model table “${column.table}”.`);
        const index = table.columns.findIndex(item => key(item) === key(column.column));
        if (index < 0)
            throw Error(`Unknown model column “${column.table}[${column.column}]”.`);
        const values = new Map();
        for (const row of rowsFor(runtime, table, context)) {
            const value = row[index] ?? null, token = value === null || value === '' ? '__DAX_BLANK__' : scalarKey(value);
            if (!values.has(token))
                values.set(token, [value]);
        }
        return { table: { ...table, columns: [table.columns[index]], rows: [...values.values()] }, rows: [...values.values()] };
    }
    if (source.kind !== 'call' || source.name !== 'FILTER' || source.args.length !== 2)
        throw Error('DAX iterators and COUNTROWS accept a model table, FILTER(table, Boolean row expression), or VALUES/DISTINCT(model column).');
    const base = rowsForDaxIterator(source.args[0], runtime, context), rows = base.rows.filter(row => { const predicate = evaluateDaxIteratorExpression(source.args[1], runtime, context, base.table, row); if (typeof predicate !== 'boolean')
        throw Error('DAX FILTER row expressions must return TRUE or FALSE.'); return predicate; });
    return { table: base.table, rows };
}
function daxDateSerial(value) { if (typeof value === 'number')
    return Number.isFinite(value) && value >= 0 && value <= 2_958_465 ? value : null; if (typeof value === 'string')
    return excelDateSerial(value); return null; }
function daxShiftDate(serial, interval, count) {
    if (interval === 'DAY') {
        const shifted = serial + count;
        return Number.isFinite(shifted) && shifted >= 0 && shifted <= 2_958_465 ? shifted : null;
    }
    const whole = Math.floor(serial), fraction = serial - whole, iso = excelSerialDateIso(whole);
    if (!iso)
        return null;
    const year = Number(iso.slice(0, 4)), month = Number(iso.slice(5, 7)), day = Number(iso.slice(8, 10)), months = count * (interval === 'QUARTER' ? 3 : interval === 'YEAR' ? 12 : 1), targetIndex = year * 12 + month - 1 + months, targetYear = Math.floor(targetIndex / 12), targetMonth = ((targetIndex % 12) + 12) % 12 + 1;
    if (targetYear < 1900 || targetYear > 9999)
        return null;
    const lastDay = new Date(Date.UTC(targetYear, targetMonth, 0)).getUTCDate(), targetDay = Math.min(day, lastDay);
    if (targetYear === 1900 && targetMonth === 2 && targetDay === 29)
        return 60 + fraction;
    const target = excelDateSerial(`${String(targetYear).padStart(4, '0')}-${String(targetMonth).padStart(2, '0')}-${String(targetDay).padStart(2, '0')}`);
    return target === null ? null : target + fraction;
}
function daxStartOfPeriod(serial, interval) { const iso = excelSerialDateIso(Math.floor(serial)); if (!iso)
    return null; const year = Number(iso.slice(0, 4)), month = Number(iso.slice(5, 7)), startMonth = interval === 'YEAR' ? 1 : interval === 'QUARTER' ? Math.floor((month - 1) / 3) * 3 + 1 : month; return excelDateSerial(`${String(year).padStart(4, '0')}-${String(startMonth).padStart(2, '0')}-01`); }
function daxTotalPeriodContext(runtime, context, column, period) {
    const table = runtime.tables.get(key(column.table));
    if (!table)
        throw Error(`Unknown DAX date table “${column.table}”.`);
    const index = table.columns.findIndex(item => key(item) === key(column.column));
    if (index < 0)
        throw Error(`Unknown DAX date column “${column.table}[${column.column}]”.`);
    const visible = rowsFor(runtime, table, context).map(row => ({ value: row[index], serial: daxDateSerial(row[index]) }));
    for (const item of visible)
        if (item.value !== null && item.value !== '' && item.serial === null)
            throw Error(`${period.toUpperCase()}-to-date requires Excel serial or ISO date values in ${column.table}[${column.column}].`);
    let maximum = -Infinity;
    for (const item of visible)
        if (item.serial !== null)
            maximum = Math.max(maximum, item.serial);
    if (!Number.isFinite(maximum))
        return null;
    const iso = excelSerialDateIso(maximum);
    if (!iso)
        throw Error('DAX time intelligence received a date outside the supported 1900 date system.');
    const year = Number(iso.slice(0, 4)), month = Number(iso.slice(5, 7)), startMonth = period === 'year' ? 1 : period === 'quarter' ? Math.floor((month - 1) / 3) * 3 + 1 : month, start = excelDateSerial(`${String(year).padStart(4, '0')}-${String(startMonth).padStart(2, '0')}-01`);
    if (start === null)
        throw Error('DAX time intelligence could not resolve the start of its date period.');
    const selected = new Set();
    for (const row of table.rows) {
        const value = row[index], serial = daxDateSerial(value);
        if (serial !== null && serial >= start && serial <= maximum)
            selected.add(scalarKey(value));
    }
    const filters = new Map(context.filters), prefix = `${key(table.name)}\u0000`;
    for (const field of filters.keys())
        if (field.startsWith(prefix))
            filters.delete(field);
    filters.set(`${key(table.name)}\u0000${key(column.column)}`, selected);
    return { filters };
}
/** Evaluate a small documented DAX measure subset; unsupported functions are rejected, never guessed. */
export function evaluateDaxMeasure(runtime, expression, filters = []) {
    const ast = new DaxParser(daxTokens(expression)).finish(), context = { filters: new Map() };
    for (const filter of filters) {
        const table = runtime.tables.get(key(filter.table));
        if (!table || !Array.isArray(filter.values))
            throw Error(`Unknown DAX filter table “${filter.table}”.`);
        if (!table.columns.some(column => key(column) === key(filter.column)))
            throw Error(`Unknown DAX filter column “${filter.table}[${filter.column}]”.`);
        context.filters.set(`${key(filter.table)}\u0000${key(filter.column)}`, new Set(filter.values.map(scalarKey)));
    }
    const evalNode = (node, ctx) => {
        if (node.kind === 'literal')
            return node.value;
        if (node.kind === 'table' || node.kind === 'column')
            throw Error('A DAX table or column reference must be used inside a supported aggregation.');
        if (node.kind === 'unary') {
            const value = evalNode(node.value, ctx);
            if (node.op === 'NOT') {
                if (typeof value !== 'boolean')
                    throw Error('DAX NOT requires a Boolean.');
                return !value;
            }
            if (typeof value !== 'number')
                throw Error('DAX unary signs require a number.');
            return node.op === '-' ? -value : value;
        }
        if (node.kind === 'binary') {
            const a = evalNode(node.left, ctx);
            if (node.op === '&&' || node.op === '||') {
                if (typeof a !== 'boolean')
                    throw Error(`DAX ${node.op} requires Boolean operands.`);
                if (node.op === '&&' && !a || node.op === '||' && a)
                    return a;
                const b = evalNode(node.right, ctx);
                if (typeof b !== 'boolean')
                    throw Error(`DAX ${node.op} requires Boolean operands.`);
                return b;
            }
            const b = evalNode(node.right, ctx);
            if (node.op === '&')
                return daxString(a) + daxString(b);
            if (['=', '<>', '<', '>', '<=', '>='].includes(node.op)) {
                const order = typeof a === 'number' && typeof b === 'number' ? a - b : String(a ?? '').localeCompare(String(b ?? ''), undefined, { numeric: true, sensitivity: 'base' });
                switch (node.op) {
                    case '=': return a === b;
                    case '<>': return a !== b;
                    case '<': return order < 0;
                    case '>': return order > 0;
                    case '<=': return order <= 0;
                    default: return order >= 0;
                }
            }
            if (typeof a !== 'number' || typeof b !== 'number')
                throw Error(`DAX ${node.op} requires numeric values.`);
            let result;
            if (node.op === '+')
                result = a + b;
            else if (node.op === '-')
                result = a - b;
            else if (node.op === '*')
                result = a * b;
            else if (b === 0)
                return null;
            else
                result = a / b;
            if (!Number.isFinite(result))
                throw Error('DAX arithmetic produced a non-finite value.');
            return result;
        }
        const args = node.args, name = node.name;
        if (name === 'NOT') {
            if (args.length !== 1)
                throw Error('NOT expects one argument.');
            const value = evalNode(args[0], ctx);
            if (typeof value !== 'boolean')
                throw Error('NOT requires a Boolean.');
            return !value;
        }
        if (name === 'AND' || name === 'OR') {
            if (args.length !== 2)
                throw Error(`${name} expects two arguments.`);
            const left = evalNode(args[0], ctx);
            if (typeof left !== 'boolean')
                throw Error(`${name} requires Boolean arguments.`);
            if (name === 'AND' && !left || name === 'OR' && left)
                return left;
            const right = evalNode(args[1], ctx);
            if (typeof right !== 'boolean')
                throw Error(`${name} requires Boolean arguments.`);
            return right;
        }
        if (name === 'BLANK') {
            if (args.length)
                throw Error('BLANK expects no arguments.');
            return null;
        }
        if (name === 'ISBLANK') {
            if (args.length !== 1)
                throw Error('ISBLANK expects one argument.');
            return evalNode(args[0], ctx) === null;
        }
        if (['HASONEVALUE', 'ISFILTERED', 'SELECTEDVALUE', 'COUNTBLANK'].includes(name)) {
            const selected = name === 'SELECTEDVALUE';
            if ((args.length !== 1 && !(selected && args.length === 2)) || !args[0] || args[0].kind !== 'column')
                throw Error(`${name} expects a model column${selected ? ' and optional alternate result' : ''}.`);
            const column = args[0], table = runtime.tables.get(key(column.table));
            if (!table)
                throw Error(`Unknown model table “${column.table}”.`);
            const index = table.columns.findIndex(item => key(item) === key(column.column));
            if (index < 0)
                throw Error(`Unknown model column “${column.table}[${column.column}]”.`);
            const field = `${key(column.table)}\u0000${key(column.column)}`;
            if (name === 'ISFILTERED')
                return ctx.filters.has(field);
            const rows = rowsFor(runtime, table, ctx), values = rows.map(row => row[index]), distinct = new Map();
            for (const value of values)
                distinct.set(value === null || value === '' ? '__DAX_BLANK__' : scalarKey(value), value === null || typeof value === 'string' || typeof value === 'number' || typeof value === 'boolean' ? value : null);
            if (name === 'COUNTBLANK')
                return values.filter(value => value === null || value === '').length;
            if (name === 'HASONEVALUE')
                return distinct.size === 1;
            if (selected)
                return distinct.size === 1 ? [...distinct.values()][0] ?? null : args[1] ? evalNode(args[1], ctx) : null;
        }
        if (name === 'DIVIDE') {
            if (args.length < 2 || args.length > 3)
                throw Error('DIVIDE expects two or three arguments.');
            const numerator = evalNode(args[0], ctx), denominator = evalNode(args[1], ctx);
            if (typeof numerator !== 'number' || typeof denominator !== 'number' || denominator === 0)
                return args[2] ? evalNode(args[2], ctx) : null;
            const result = numerator / denominator;
            if (!Number.isFinite(result))
                throw Error('DAX DIVIDE produced a non-finite value.');
            return result;
        }
        if (name === 'IF') {
            if (args.length < 2 || args.length > 3)
                throw Error('IF expects two or three arguments.');
            const condition = evalNode(args[0], ctx), truth = typeof condition === 'boolean' ? condition : typeof condition === 'number' ? condition !== 0 : condition !== null && condition !== '';
            return truth ? evalNode(args[1], ctx) : args[2] ? evalNode(args[2], ctx) : false;
        }
        if (name === 'COALESCE') {
            if (args.length < 1 || args.length > 16)
                throw Error('COALESCE expects one to sixteen arguments.');
            for (const item of args) {
                const value = evalNode(item, ctx);
                if (value !== null && value !== '')
                    return value;
            }
            return null;
        }
        if (name === 'TOTALYTD' || name === 'TOTALQTD' || name === 'TOTALMTD') {
            if (args.length !== 2 || args[1].kind !== 'column')
                throw Error(`${name} in this subset expects a scalar expression and one date-column reference.`);
            const period = name === 'TOTALYTD' ? 'year' : name === 'TOTALQTD' ? 'quarter' : 'month', next = daxTotalPeriodContext(runtime, ctx, args[1], period);
            return next ? evalNode(args[0], next) : null;
        }
        if (name === 'PERCENTILEX.INC') {
            if (args.length !== 3)
                throw Error('PERCENTILEX.INC expects a model table, scalar row expression, and percentile.');
            const source = rowsForDaxIterator(args[0], runtime, ctx), values = source.rows.map(row => evaluateDaxIteratorExpression(args[1], runtime, ctx, source.table, row)).filter((value) => typeof value === 'number' && Number.isFinite(value));
            return daxPercentileInc(values, evalNode(args[2], ctx));
        }
        if (name === 'PERCENTILE.INC') {
            if (args.length !== 2 || args[0].kind !== 'column')
                throw Error('PERCENTILE.INC expects one model column and a percentile.');
            const column = args[0], table = runtime.tables.get(key(column.table));
            if (!table)
                throw Error(`Unknown model table “${column.table}”.`);
            const index = table.columns.findIndex(item => key(item) === key(column.column));
            if (index < 0)
                throw Error(`Unknown model column “${column.table}[${column.column}]”.`);
            const values = rowsFor(runtime, table, ctx).map(row => row[index]).filter((value) => typeof value === 'number' && Number.isFinite(value));
            return daxPercentileInc(values, evalNode(args[1], ctx));
        }
        if (name === 'PERCENTILEX.EXC') {
            if (args.length !== 3)
                throw Error('PERCENTILEX.EXC expects a model table, scalar row expression, and percentile.');
            const source = rowsForDaxIterator(args[0], runtime, ctx), values = source.rows.map(row => evaluateDaxIteratorExpression(args[1], runtime, ctx, source.table, row)).filter((value) => typeof value === 'number' && Number.isFinite(value));
            return daxPercentileExc(values, evalNode(args[2], ctx));
        }
        if (name === 'PERCENTILE.EXC') {
            if (args.length !== 2 || args[0].kind !== 'column')
                throw Error('PERCENTILE.EXC expects one model column and a percentile.');
            const column = args[0], table = runtime.tables.get(key(column.table));
            if (!table)
                throw Error(`Unknown model table “${column.table}”.`);
            const index = table.columns.findIndex(item => key(item) === key(column.column));
            if (index < 0)
                throw Error(`Unknown model column “${column.table}[${column.column}]”.`);
            const values = rowsFor(runtime, table, ctx).map(row => row[index]).filter((value) => typeof value === 'number' && Number.isFinite(value));
            return daxPercentileExc(values, evalNode(args[1], ctx));
        }
        if (['STDEVX.P', 'STDEVX.S', 'VARX.P', 'VARX.S'].includes(name)) {
            if (args.length !== 2)
                throw Error(`${name} expects a model table or FILTER(...) and a scalar row expression.`);
            const source = rowsForDaxIterator(args[0], runtime, ctx), evaluated = source.rows.map(row => evaluateDaxIteratorExpression(args[1], runtime, ctx, source.table, row)), numeric = [];
            for (const value of evaluated) {
                if (value === null || value === '')
                    continue;
                if (typeof value !== 'number' || !Number.isFinite(value))
                    throw Error(`${name} row expressions must return finite numbers or BLANK().`);
                numeric.push(value);
            }
            return daxVariance(numeric, name.endsWith('.S'), name, name.startsWith('STDEV'));
        }
        if (name === 'RANKX') {
            if (args.length < 3 || args.length > 5)
                throw Error('RANKX in this subset expects a table, numeric row expression, comparison value, and optional order and tie mode.');
            const source = rowsForDaxIterator(args[0], runtime, ctx);
            if (!source.rows.length)
                return null;
            const comparisonValue = evalNode(args[2], ctx);
            if (comparisonValue !== null && typeof comparisonValue !== 'number')
                throw Error('RANKX comparison value must be numeric or BLANK().');
            const target = comparisonValue ?? 0, values = source.rows.map(row => { const value = evaluateDaxIteratorExpression(args[1], runtime, ctx, source.table, row); if (value !== null && typeof value !== 'number')
                throw Error('RANKX row expressions in this subset must return numbers or BLANK().'); return value ?? 0; });
            let ascending = false;
            if (args[3]) {
                const order = args[3];
                if (order.kind === 'table' && key(order.value) === 'asc' || order.kind === 'literal' && order.value === 1)
                    ascending = true;
                else if (order.kind === 'table' && key(order.value) === 'desc' || order.kind === 'literal' && order.value === 0)
                    ascending = false;
                else
                    throw Error('RANKX order must be ASC/1 or DESC/0.');
            }
            let dense = false;
            if (args[4]) {
                const ties = args[4];
                if (ties.kind === 'table' && key(ties.value) === 'dense')
                    dense = true;
                else if (ties.kind === 'table' && key(ties.value) === 'skip')
                    dense = false;
                else
                    throw Error('RANKX tie mode must be DENSE or SKIP.');
            }
            const precede = (value) => ascending ? value < target : value > target;
            return dense ? new Set(values.filter(precede)).size + 1 : values.filter(precede).length + 1;
        }
        if (name === 'ABS') {
            if (args.length !== 1)
                throw Error('ABS expects one argument.');
            const value = evalNode(args[0], ctx);
            if (typeof value !== 'number')
                throw Error('ABS expects a number.');
            return Math.abs(value);
        }
        if (name === 'ROUND') {
            if (args.length !== 2)
                throw Error('ROUND expects a number and a digit count.');
            const value = evalNode(args[0], ctx), digits = evalNode(args[1], ctx);
            if (typeof value !== 'number' || typeof digits !== 'number' || !Number.isInteger(digits) || Math.abs(digits) > 12)
                throw Error('ROUND needs a number and an integer digit count from -12 to 12.');
            const factor = 10 ** digits, result = Math.sign(value) * Math.floor(Math.abs(value) * factor + 0.5) / factor;
            if (!Number.isFinite(result))
                throw Error('DAX ROUND produced a non-finite value.');
            return result;
        }
        if (DAX_NUMERIC_FUNCTIONS.has(name))
            return evaluateDaxNumericFunction(name, args.map(item => evalNode(item, ctx)));
        if (DAX_DATE_FUNCTIONS.has(name))
            return evaluateDaxDateFunction(name, args.map(item => evalNode(item, ctx)));
        if (DAX_TEXT_FUNCTIONS.has(name))
            return evaluateDaxTextFunction(name, args.map(item => evalNode(item, ctx)));
        if (name === 'CALCULATE') {
            if (args.length < 1 || args.length > 16)
                throw Error('CALCULATE expects an expression and up to fifteen filter arguments.');
            const next = { filters: new Map(ctx.filters) };
            const validateColumn = (column) => { const table = runtime.tables.get(key(column.table)); if (!table || !table.columns.some(item => key(item) === key(column.column)))
                throw Error(`Unknown DAX filter column “${column.table}[${column.column}]”.`); return `${key(column.table)}\u0000${key(column.column)}`; };
            const dateFilterNames = ['DATESYTD', 'DATESQTD', 'DATESMTD', 'DATESBETWEEN', 'DATESINPERIOD', 'DATEADD', 'SAMEPERIODLASTYEAR', 'PARALLELPERIOD'], timeFilters = args.slice(1).filter(filter => filter.kind === 'call' && dateFilterNames.includes(filter.name));
            const dateTableValues = new Map();
            for (const filter of timeFilters) {
                if (filter.kind !== 'call')
                    continue;
                const between = filter.name === 'DATESBETWEEN', inPeriod = filter.name === 'DATESINPERIOD', dateAdd = filter.name === 'DATEADD' || filter.name === 'SAMEPERIODLASTYEAR', samePeriod = filter.name === 'SAMEPERIODLASTYEAR', parallel = filter.name === 'PARALLELPERIOD';
                if (filter.args.length !== (between ? 3 : inPeriod ? 4 : dateAdd ? (samePeriod ? 1 : 3) : parallel ? 3 : 1) || filter.args[0].kind !== 'column')
                    throw Error(`${filter.name} in this subset expects ${between ? 'a date-column reference and start/end date values' : inPeriod ? 'a date-column reference, start date, interval count, and DAY/MONTH/QUARTER/YEAR interval' : dateAdd ? (samePeriod ? 'one date-column reference' : 'a date-column reference, interval count, and DAY/MONTH/QUARTER/YEAR interval') : parallel ? 'a date-column reference, interval count, and MONTH/QUARTER/YEAR interval' : 'one date-column reference'}.`);
                const dateColumn = filter.args[0], field = validateColumn(dateColumn), tableName = key(dateColumn.table), table = runtime.tables.get(tableName), dateIndex = table.columns.findIndex(column => key(column) === key(dateColumn.column));
                let dateValues;
                if (between) {
                    const startInput = evalNode(filter.args[1], ctx), endInput = evalNode(filter.args[2], ctx), startBound = startInput === null ? null : daxDateSerial(startInput), endBound = endInput === null ? null : daxDateSerial(endInput);
                    if (startInput !== null && startBound === null || endInput !== null && endBound === null)
                        throw Error('DATESBETWEEN bounds must be BLANK(), Excel date serials, or ISO dates.');
                    let minimum = Infinity, maximum = -Infinity;
                    const values = table.rows.map(row => ({ value: row[dateIndex], serial: daxDateSerial(row[dateIndex]) }));
                    for (const item of values) {
                        if (item.value !== null && item.value !== '' && item.serial === null)
                            throw Error(`DATESBETWEEN requires Excel serial or ISO dates in ${dateColumn.table}[${dateColumn.column}].`);
                        if (item.serial !== null) {
                            minimum = Math.min(minimum, item.serial);
                            maximum = Math.max(maximum, item.serial);
                        }
                    }
                    const start = startBound ?? minimum, end = endBound ?? maximum;
                    dateValues = new Set();
                    if (start <= end)
                        for (const item of values)
                            if (item.serial !== null && item.serial >= start && item.serial <= end)
                                dateValues.add(scalarKey(item.value));
                }
                else if (inPeriod) {
                    const startInput = evalNode(filter.args[1], ctx), countValue = evalNode(filter.args[2], ctx), intervalNode = filter.args[3], interval = intervalNode.kind === 'table' ? intervalNode.value.toUpperCase() : intervalNode.kind === 'literal' && typeof intervalNode.value === 'string' ? intervalNode.value.toUpperCase() : '';
                    if (intervalNode.kind === 'column' || !['DAY', 'MONTH', 'QUARTER', 'YEAR'].includes(interval))
                        throw Error('DATESINPERIOD interval must be DAY, MONTH, QUARTER, or YEAR.');
                    const start = daxDateSerial(startInput), count = typeof countValue === 'number' && Number.isSafeInteger(countValue) ? countValue : null;
                    if (start === null)
                        throw Error('DATESINPERIOD start date must be an Excel serial or ISO date.');
                    if (count === null || Math.abs(count) > 1_000_000)
                        throw Error('DATESINPERIOD interval count must be an integer from -1,000,000 to 1,000,000.');
                    const shifted = daxShiftDate(start, interval, count);
                    if (shifted === null)
                        throw Error('DATESINPERIOD range is outside the supported Excel 1900 date system.');
                    const lower = Math.min(start, shifted), upper = Math.max(start, shifted), values = table.rows.map(row => ({ value: row[dateIndex], serial: daxDateSerial(row[dateIndex]) }));
                    dateValues = new Set();
                    for (const item of values) {
                        if (item.value !== null && item.value !== '' && item.serial === null)
                            throw Error(`DATESINPERIOD requires Excel serial or ISO dates in ${dateColumn.table}[${dateColumn.column}].`);
                        if (item.serial !== null && item.serial >= lower && item.serial <= upper)
                            dateValues.add(scalarKey(item.value));
                    }
                }
                else if (dateAdd) {
                    const countValue = samePeriod ? -1 : evalNode(filter.args[1], ctx), intervalNode = samePeriod ? null : filter.args[2], interval = samePeriod ? 'YEAR' : intervalNode?.kind === 'table' ? intervalNode.value.toUpperCase() : intervalNode?.kind === 'literal' && typeof intervalNode.value === 'string' ? intervalNode.value.toUpperCase() : '';
                    if (intervalNode?.kind === 'column' || !['DAY', 'MONTH', 'QUARTER', 'YEAR'].includes(interval))
                        throw Error(`${filter.name} interval must be DAY, MONTH, QUARTER, or YEAR.`);
                    const count = typeof countValue === 'number' && Number.isSafeInteger(countValue) ? countValue : null;
                    if (count === null || Math.abs(count) > 1_000_000)
                        throw Error(`${filter.name} interval count must be an integer from -1,000,000 to 1,000,000.`);
                    const visible = rowsFor(runtime, table, ctx).map(row => ({ value: row[dateIndex], serial: daxDateSerial(row[dateIndex]) }));
                    for (const item of visible)
                        if (item.value !== null && item.value !== '' && item.serial === null)
                            throw Error(`${filter.name} requires Excel serial or ISO dates in ${dateColumn.table}[${dateColumn.column}].`);
                    const shifted = new Set();
                    for (const item of visible)
                        if (item.serial !== null) {
                            const target = daxShiftDate(item.serial, interval, count);
                            if (target === null)
                                throw Error(`${filter.name} result is outside the supported Excel 1900 date system.`);
                            shifted.add(target);
                        }
                    dateValues = new Set();
                    for (const row of table.rows) {
                        const value = row[dateIndex], serial = daxDateSerial(value);
                        if (value !== null && value !== '' && serial === null)
                            throw Error(`${filter.name} requires Excel serial or ISO dates in ${dateColumn.table}[${dateColumn.column}].`);
                        if (serial !== null && shifted.has(serial))
                            dateValues.add(scalarKey(value));
                    }
                }
                else if (parallel) {
                    const countValue = evalNode(filter.args[1], ctx), intervalNode = filter.args[2], interval = intervalNode.kind === 'table' ? intervalNode.value.toUpperCase() : intervalNode.kind === 'literal' && typeof intervalNode.value === 'string' ? intervalNode.value.toUpperCase() : '';
                    if (intervalNode.kind === 'column' || !['MONTH', 'QUARTER', 'YEAR'].includes(interval))
                        throw Error('PARALLELPERIOD interval must be MONTH, QUARTER, or YEAR.');
                    const count = typeof countValue === 'number' && Number.isSafeInteger(countValue) ? countValue : null;
                    if (count === null || Math.abs(count) > 1_000_000)
                        throw Error('PARALLELPERIOD interval count must be an integer from -1,000,000 to 1,000,000.');
                    const visible = rowsFor(runtime, table, ctx).map(row => ({ value: row[dateIndex], serial: daxDateSerial(row[dateIndex]) }));
                    for (const item of visible)
                        if (item.value !== null && item.value !== '' && item.serial === null)
                            throw Error(`PARALLELPERIOD requires Excel serial or ISO dates in ${dateColumn.table}[${dateColumn.column}].`);
                    let low = Infinity, high = -Infinity;
                    for (const item of visible)
                        if (item.serial !== null) {
                            low = Math.min(low, item.serial);
                            high = Math.max(high, item.serial);
                        }
                    dateValues = new Set();
                    if (Number.isFinite(low)) {
                        const periodStart = daxStartOfPeriod(low, interval), lastPeriodStart = daxStartOfPeriod(high, interval), shiftedStart = periodStart === null ? null : daxShiftDate(periodStart, interval, count), shiftedLast = lastPeriodStart === null ? null : daxShiftDate(lastPeriodStart, interval, count), nextPeriod = shiftedLast === null ? null : daxShiftDate(shiftedLast, interval, 1);
                        if (shiftedStart === null || nextPeriod === null)
                            throw Error('PARALLELPERIOD range is outside the supported Excel 1900 date system.');
                        for (const row of table.rows) {
                            const value = row[dateIndex], serial = daxDateSerial(value);
                            if (value !== null && value !== '' && serial === null)
                                throw Error(`PARALLELPERIOD requires Excel serial or ISO dates in ${dateColumn.table}[${dateColumn.column}].`);
                            if (serial !== null && serial >= shiftedStart && serial < nextPeriod)
                                dateValues.add(scalarKey(value));
                        }
                    }
                }
                else {
                    const period = filter.name === 'DATESYTD' ? 'year' : filter.name === 'DATESQTD' ? 'quarter' : 'month', periodContext = daxTotalPeriodContext(runtime, ctx, dateColumn, period);
                    dateValues = periodContext?.filters.get(field) ?? new Set();
                }
                const tableFilters = dateTableValues.get(tableName) ?? new Map(), previous = tableFilters.get(field);
                tableFilters.set(field, previous ? new Set([...previous].filter(value => dateValues.has(value))) : dateValues);
                dateTableValues.set(tableName, tableFilters);
            }
            for (const [tableName, tableFilters] of dateTableValues) {
                const prefix = `${tableName}\u0000`;
                for (const field of next.filters.keys())
                    if (field.startsWith(prefix))
                        next.filters.delete(field);
                for (const [field, values] of tableFilters)
                    next.filters.set(field, values);
            }
            for (const filter of args.slice(1)) {
                if (filter.kind === 'call' && dateFilterNames.includes(filter.name))
                    continue;
                if (filter.kind === 'call' && filter.name === 'ALLEXCEPT') {
                    if (filter.args.length < 2 || filter.args.length > 16 || filter.args[0].kind !== 'table')
                        throw Error('ALLEXCEPT expects a model table and one to fifteen columns from that table.');
                    const tableName = key(filter.args[0].value), table = runtime.tables.get(tableName);
                    if (!table)
                        throw Error(`Unknown model table “${filter.args[0].value}”.`);
                    const preserved = new Set();
                    for (const column of filter.args.slice(1)) {
                        if (column.kind !== 'column' || key(column.table) !== tableName)
                            throw Error('ALLEXCEPT columns must belong to its first table argument.');
                        preserved.add(validateColumn(column));
                    }
                    const prefix = `${tableName}\u0000`;
                    for (const field of next.filters.keys())
                        if (field.startsWith(prefix) && !preserved.has(field))
                            next.filters.delete(field);
                    continue;
                }
                if (filter.kind === 'call' && (filter.name === 'REMOVEFILTERS' || filter.name === 'ALL')) {
                    if (filter.args.length > 16)
                        throw Error(`${filter.name} accepts up to sixteen table or column references.`);
                    if (!filter.args.length) {
                        if (filter.name === 'ALL')
                            throw Error('ALL in this DAX subset needs a model table or column; use REMOVEFILTERS() to clear all filters.');
                        next.filters.clear();
                        continue;
                    }
                    for (const reference of filter.args) {
                        if (reference.kind === 'column') {
                            next.filters.delete(validateColumn(reference));
                            continue;
                        }
                        if (reference.kind === 'table') {
                            const table = runtime.tables.get(key(reference.value));
                            if (!table)
                                throw Error(`Unknown model table “${reference.value}”.`);
                            const prefix = `${key(reference.value)}\u0000`;
                            for (const field of next.filters.keys())
                                if (field.startsWith(prefix))
                                    next.filters.delete(field);
                            continue;
                        }
                        throw Error(`${filter.name} in CALCULATE expects model table or column references.`);
                    }
                    continue;
                }
                const keep = filter.kind === 'call' && filter.name === 'KEEPFILTERS';
                if (keep && (filter.args.length !== 1 || filter.args[0].kind !== 'binary'))
                    throw Error('KEEPFILTERS expects one equality filter such as KEEPFILTERS(Sales[Region] = "West").');
                const predicate = keep ? filter.args[0] : filter;
                if (predicate.kind !== 'binary' || predicate.op !== '=' || predicate.left.kind !== 'column')
                    throw Error('This DAX subset accepts CALCULATE filters written as Table[Column] = scalar, KEEPFILTERS, or REMOVEFILTERS/ALL.');
                const field = validateColumn(predicate.left), value = evalNode(predicate.right, ctx), token = scalarKey(value);
                if (keep) {
                    const prior = next.filters.get(field);
                    next.filters.set(field, prior ? new Set([...prior].filter(item => item === token)) : new Set([token]));
                }
                else
                    next.filters.set(field, new Set([token]));
            }
            return evalNode(args[0], next);
        }
        if (name === 'CONCATENATEX') {
            if (args.length < 2 || args.length > 3)
                throw Error('CONCATENATEX in this subset expects a model table, a scalar row expression, and an optional delimiter; order-by arguments are not supported.');
            const source = rowsForDaxIterator(args[0], runtime, ctx);
            if (!source.rows.length)
                return null;
            const delimiter = args[2] ? daxString(evalNode(args[2], ctx)) : '';
            if (delimiter.length > 32767)
                throw Error('CONCATENATEX delimiter exceeds the 32,767-character cell limit.');
            const parts = [];
            let length = 0;
            for (const row of source.rows) {
                const value = evaluateDaxIteratorExpression(args[1], runtime, ctx, source.table, row), part = daxString(value);
                length += part.length + (parts.length ? delimiter.length : 0);
                if (length > 32767)
                    throw Error('CONCATENATEX result exceeds the 32,767-character cell limit.');
                parts.push(part);
            }
            return parts.join(delimiter);
        }
        if (['SUMX', 'AVERAGEX', 'MINX', 'MAXX', 'MEDIANX', 'PRODUCTX', 'COUNTX', 'COUNTAX'].includes(name)) {
            if (args.length !== 2)
                throw Error(`${name} expects a model table or FILTER(...) and a scalar row expression.`);
            const source = rowsForDaxIterator(args[0], runtime, ctx), table = source.table, evaluated = source.rows.map(row => evaluateDaxIteratorExpression(args[1], runtime, ctx, table, row)), numbers = evaluated.filter((value) => typeof value === 'number' && Number.isFinite(value));
            if (name === 'COUNTAX')
                return evaluated.filter(value => value !== null && value !== '').length;
            if (name === 'COUNTX')
                return numbers.length;
            if (!numbers.length)
                return null;
            if (name === 'MINX')
                return Math.min(...numbers);
            if (name === 'MAXX')
                return Math.max(...numbers);
            if (name === 'MEDIANX') {
                numbers.sort((a, b) => a - b);
                const middle = Math.floor(numbers.length / 2), median = numbers.length % 2 ? numbers[middle] : numbers[middle - 1] / 2 + numbers[middle] / 2;
                if (!Number.isFinite(median))
                    throw Error('DAX MEDIANX produced a non-finite value.');
                return median;
            }
            if (name === 'PRODUCTX') {
                let product = 1;
                for (const value of numbers) {
                    product *= value;
                    if (!Number.isFinite(product))
                        throw Error('DAX PRODUCTX produced a non-finite value.');
                }
                return product;
            }
            const total = numbers.reduce((sum, value) => sum + value, 0);
            if (!Number.isFinite(total))
                throw Error(`DAX ${name} produced a non-finite value.`);
            return name === 'SUMX' ? total : total / numbers.length;
        }
        if (name === 'COUNTROWS') {
            if (args.length !== 1)
                throw Error('COUNTROWS expects a model table or FILTER(...) expression.');
            const count = rowsForDaxIterator(args[0], runtime, ctx).rows.length;
            return count ? count : null;
        }
        if (!['SUM', 'AVERAGE', 'MEDIAN', 'PRODUCT', 'MIN', 'MAX', 'COUNT', 'COUNTA', 'DISTINCTCOUNT', 'DISTINCTCOUNTNOBLANK', 'STDEV.P', 'STDEV.S', 'VAR.P', 'VAR.S'].includes(name) || args.length !== 1 || args[0].kind !== 'column')
            throw Error(`DAX function “${name}” is outside GridForge’s supported measure subset.`);
        const column = args[0], table = runtime.tables.get(key(column.table));
        if (!table)
            throw Error(`Unknown model table “${column.table}”.`);
        const index = table.columns.findIndex(item => key(item) === key(column.column));
        if (index < 0)
            throw Error(`Unknown model column “${column.table}[${column.column}]”.`);
        const values = rowsFor(runtime, table, ctx).map(row => row[index]), numeric = values.filter((value) => typeof value === 'number' && Number.isFinite(value));
        if (['STDEV.P', 'STDEV.S', 'VAR.P', 'VAR.S'].includes(name)) {
            if (values.some(value => value !== null && value !== '' && (typeof value !== 'number' || !Number.isFinite(value))))
                throw Error(`${name} column values must be finite numbers or BLANK().`);
            return daxVariance(numeric, name.endsWith('.S'), name, name.startsWith('STDEV'));
        }
        switch (name) {
            case 'SUM': {
                if (!numeric.length)
                    return null;
                const result = numeric.reduce((sum, value) => sum + value, 0);
                if (!Number.isFinite(result))
                    throw Error('DAX SUM produced a non-finite value.');
                return result;
            }
            case 'AVERAGE': {
                if (!numeric.length)
                    return null;
                const result = numeric.reduce((sum, value) => sum + value, 0) / numeric.length;
                if (!Number.isFinite(result))
                    throw Error('DAX AVERAGE produced a non-finite value.');
                return result;
            }
            case 'MEDIAN': {
                if (!numeric.length)
                    return null;
                const sorted = [...numeric].sort((a, b) => a - b), middle = Math.floor(sorted.length / 2), median = sorted.length % 2 ? sorted[middle] : sorted[middle - 1] / 2 + sorted[middle] / 2;
                return Number.isFinite(median) ? median : null;
            }
            case 'PRODUCT': {
                if (!numeric.length)
                    return null;
                let product = 1;
                for (const value of numeric) {
                    product *= value;
                    if (!Number.isFinite(product))
                        throw Error('DAX PRODUCT produced a non-finite value.');
                }
                return product;
            }
            case 'MIN': return numeric.length ? Math.min(...numeric) : null;
            case 'MAX': return numeric.length ? Math.max(...numeric) : null;
            case 'COUNT': return numeric.length;
            case 'COUNTA': return values.filter(value => value !== null && value !== '').length;
            case 'DISTINCTCOUNTNOBLANK':
                if (!values.length)
                    return null;
                return new Set(values.filter(value => value !== null && value !== '').map(scalarKey)).size;
            default:
                if (!values.length)
                    return null;
                return new Set(values.map(value => value === null || value === '' ? '__DAX_BLANK__' : scalarKey(value))).size;
        }
    };
    return evalNode(ast, context);
}
export function evaluateSavedDaxMeasure(runtime, name, filters = []) { const matches = runtime.measures.filter(item => key(item.name) === key(name)); if (matches.length !== 1)
    throw Error(`DAX measure “${name}” was not found.`); return evaluateDaxMeasure(runtime, matches[0].expression, filters); }
