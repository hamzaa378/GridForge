/** GridForge automation is a validated declarative action format, not VBA/JavaScript.
 * v1/v2 action macros remain supported. v3 adds finite conditionals, bounded
 * repeat blocks, and numeric sequences. Playback works entirely offline.
 * No networking, filesystem access, code execution or user-defined scripting.
 */
import { SpreadsheetDocument } from './document.js';
import { parseRef, colToName, shiftFormula } from './parser.js';
import { cleanStyle } from './format.js';
import { snapshotCells, pasteSnapshot } from './clipboard.js';
import { sortRangeTransaction, replaceCells } from './operations.js';
const address = (r, c) => `${colToName(c)}${r + 1}`;
function rectangle(source) {
    if (typeof source !== 'string' || source.length > 30)
        throw Error('Invalid macro range.');
    const xs = source.toUpperCase().split(':');
    if (xs.length > 2)
        throw Error('Invalid rectangular range.');
    const a = parseRef(xs[0]), b = parseRef(xs[1] ?? xs[0]);
    if (!a || !b)
        throw Error(`Invalid macro address ${source}.`);
    if (Math.abs(a.row - b.row) + 1 > 1000 || Math.abs(a.col - b.col) + 1 > 104)
        throw Error('Range too large.');
    return { r1: Math.min(a.row, b.row), r2: Math.max(a.row, b.row), c1: Math.min(a.col, b.col), c2: Math.max(a.col, b.col) };
}
const isObj = (x) => !!x && typeof x === 'object' && !Array.isArray(x);
const cleanName = (s, kind) => { if (typeof s !== 'string' || !s.trim() || s.length > (kind === 'name' ? 64 : 31))
    throw Error(`Invalid ${kind}.`); return s.trim(); };
function parseEdits(v, kind) {
    if (!Array.isArray(v) || v.length > 5000)
        throw Error('Invalid macro edit list.');
    return v.map(item => {
        if (!isObj(item) || typeof item.ref !== 'string' || !parseRef(item.ref))
            throw Error('Invalid edit address.');
        const ref = item.ref.toUpperCase().replace(/\$/g, '');
        if (kind === 'cells') {
            if (typeof item.raw !== 'string' || item.raw.length > 30000)
                throw Error('Invalid macro cell contents.');
            return { ref, raw: item.raw };
        }
        return { ref, style: cleanStyle(item.style) };
    });
}
function parseStep(step, version, depth = 0) {
    if (depth > 3)
        throw Error('Macro nesting depth limit is 3.');
    if (!isObj(step) || typeof step.kind !== 'string')
        throw Error('Invalid macro action.');
    const kind = step.kind;
    if (version === 1 && !['cells', 'styles'].includes(kind))
        throw Error('Legacy macro has unsupported action.');
    const sheetName = version !== 1 && step.sheetName !== undefined ? cleanName(step.sheetName, 'sheet name') : undefined;
    if (kind === 'cells' || kind === 'styles')
        return { kind, sheetName, edits: parseEdits(step.edits, kind) };
    if (version === 1)
        throw Error('Unsupported v1 macro step.');
    if (version >= 3 && kind === 'sequence') {
        const range = cleanRange(step.range), start = step.start, increment = step.increment;
        if (typeof start !== 'number' || !Number.isFinite(start) || Math.abs(start) > 1e12 || typeof increment !== 'number' || !Number.isFinite(increment) || Math.abs(increment) > 1e9)
            throw Error('Invalid sequence numbers.');
        return { kind, sheetName, range, start, increment };
    }
    if (version >= 3 && kind === 'ifCell') {
        const allowed = ['equals', 'notEquals', 'contains', 'greaterThan', 'lessThan', 'blank', 'notBlank'];
        const operator = step.operator;
        if (typeof operator !== 'string' || !allowed.includes(operator))
            throw Error('Invalid condition operator.');
        if (!['blank', 'notBlank'].includes(operator) && (typeof step.value !== 'string' || step.value.length > 200))
            throw Error('Invalid condition value.');
        if (!Array.isArray(step.thenSteps) || step.thenSteps.length > 100 || (step.elseSteps !== undefined && (!Array.isArray(step.elseSteps) || step.elseSteps.length > 100)))
            throw Error('Invalid conditional branches.');
        return { kind, sheetName, ref: cleanRef(step.ref), operator: operator, value: typeof step.value === 'string' ? step.value : undefined,
            thenSteps: step.thenSteps.map(x => parseStep(x, version, depth + 1)),
            elseSteps: Array.isArray(step.elseSteps) ? step.elseSteps.map(x => parseStep(x, version, depth + 1)) : [] };
    }
    if (version >= 3 && kind === 'repeat') {
        if (!Number.isInteger(step.times) || Number(step.times) < 1 || Number(step.times) > 20 || !Array.isArray(step.steps) || step.steps.length > 100)
            throw Error('Repeat must have 1-20 iterations and at most 100 steps.');
        return { kind, times: Number(step.times), steps: step.steps.map(x => parseStep(x, version, depth + 1)) };
    }
    if (version === 4 && kind === 'cleanText') {
        const operation = step.operation;
        if (!['trim', 'upper', 'lower', 'proper', 'collapseSpaces'].includes(String(operation)))
            throw Error('Unsupported text cleanup operation.');
        return { kind, sheetName, range: cleanRange(step.range), operation: operation, skipFormulas: step.skipFormulas !== false };
    }
    if (version === 4 && kind === 'formatIf') {
        const allowed = ['equals', 'notEquals', 'contains', 'greaterThan', 'lessThan', 'blank', 'notBlank'];
        if (!allowed.includes(step.operator))
            throw Error('Invalid format condition.');
        if (!['blank', 'notBlank'].includes(String(step.operator)) && (typeof step.value !== 'string' || step.value.length > 1000))
            throw Error('Invalid conditional format value.');
        return { kind, sheetName, range: cleanRange(step.range), operator: step.operator, value: typeof step.value === 'string' ? step.value : undefined, style: cleanStyle(step.style) };
    }
    if (version === 4 && kind === 'removeDuplicates') {
        const range = cleanRange(step.range), selected = rectangle(range);
        if (step.keyColumns !== undefined && (!Array.isArray(step.keyColumns) || !step.keyColumns.length || step.keyColumns.length > 20 || step.keyColumns.some(x => !Number.isInteger(x) || x < selected.c1 || x > selected.c2) || new Set(step.keyColumns).size !== step.keyColumns.length))
            throw Error('Duplicate keys must be column indices within the selected range.');
        return { kind, sheetName, range, keyColumns: step.keyColumns, hasHeader: step.hasHeader === true };
    }
    if (kind === 'clear')
        return { kind, sheetName, range: cleanRange(step.range), clearFormatting: step.clearFormatting === true };
    if (kind === 'fillDown')
        return { kind, sheetName, source: cleanRef(step.source), range: cleanRange(step.range) };
    if (kind === 'copy')
        return { kind, sheetName, sourceRange: cleanRange(step.sourceRange), destination: cleanRef(step.destination), valuesOnly: step.valuesOnly === true };
    if (kind === 'sort') {
        const range = cleanRange(step.range), byColumn = step.byColumn, selected = rectangle(range);
        if (typeof byColumn !== 'number' || !Number.isInteger(byColumn) || byColumn < 0 || byColumn >= 104 || byColumn < selected.c1 || byColumn > selected.c2)
            throw Error('Invalid sort column outside selection.');
        let keys;
        if (step.keys !== undefined) {
            if (!Array.isArray(step.keys) || step.keys.length < 1 || step.keys.length > 16 || step.keys.some((key) => !key || typeof key !== 'object' || !Number.isInteger(key.column) || (key.column < selected.c1 || key.column > selected.c2) || typeof key.descending !== 'boolean') || new Set(step.keys.map((key) => key.column)).size !== step.keys.length || step.keys[0]?.column !== byColumn)
                throw Error('Invalid multi-level sort keys.');
            keys = step.keys.map((key) => ({ column: key.column, descending: key.descending }));
        }
        return { kind, sheetName, range, byColumn, descending: step.descending === true, ...(keys ? { keys } : {}), hasHeader: step.hasHeader === true };
    }
    if (kind === 'replace') {
        if (typeof step.find !== 'string' || !step.find || step.find.length > 200 || typeof step.replace !== 'string' || step.replace.length > 2000)
            throw Error('Invalid replace expression.');
        return { kind, sheetName, find: step.find, replace: step.replace, caseSensitive: step.caseSensitive === true, includeFormulas: step.includeFormulas === true };
    }
    throw Error(`Unsupported macro action '${kind}'. No custom code is executed.`);
}
function cleanRef(x) { if (typeof x !== 'string' || !parseRef(x))
    throw Error('Invalid macro address.'); return x.toUpperCase().replace(/\$/g, ''); }
function cleanRange(x) { if (typeof x !== 'string')
    throw Error('Invalid macro range.'); rectangle(x); return x.toUpperCase().replace(/\$/g, ''); }
export function validateMacro(data) {
    if (!isObj(data) || ![1, 2, 3, 4].includes(data.version) || !Array.isArray(data.steps) || data.steps.length > 200)
        throw Error('Invalid GridForge macro header.');
    const version = data.version, name = cleanName(data.name, 'name'), sheetName = cleanName(data.sheetName, 'sheet name');
    const steps = data.steps.map(x => parseStep(x, version));
    const parameters = [];
    if (version === 4) {
        if (data.parameters !== undefined && (!Array.isArray(data.parameters) || data.parameters.length > 12))
            throw Error('At most 12 macro parameters are supported.');
        for (const param of data.parameters ?? []) {
            if (!isObj(param) || typeof param.name !== 'string' || !/^[a-zA-Z][a-zA-Z0-9_]{0,31}$/.test(param.name) || !['text', 'number'].includes(String(param.type)))
                throw Error('Invalid macro parameter definition.');
            if (parameters.some(p => p.name === param.name))
                throw Error('Duplicate macro parameter.');
            if (param.defaultValue !== undefined && (typeof param.defaultValue !== 'string' || param.defaultValue.length > 1000))
                throw Error('Invalid parameter default.');
            parameters.push({ name: param.name, type: param.type, required: param.required === true, defaultValue: param.defaultValue });
        }
    }
    const costs = (items) => items.reduce((total, step) => {
        let cost = 0;
        if (step.kind === 'cells' || step.kind === 'styles')
            cost = step.edits.length;
        else if (step.kind === 'clear' || step.kind === 'fillDown' || step.kind === 'sort' || step.kind === 'sequence' || step.kind === 'cleanText' || step.kind === 'formatIf' || step.kind === 'removeDuplicates') {
            const b = rectangle(step.range);
            cost = (b.r2 - b.r1 + 1) * (b.c2 - b.c1 + 1);
        }
        else if (step.kind === 'copy') {
            const b = rectangle(step.sourceRange);
            cost = (b.r2 - b.r1 + 1) * (b.c2 - b.c1 + 1);
        }
        else if (step.kind === 'replace')
            cost = 1000;
        else if (step.kind === 'ifCell')
            cost = 1 + Math.max(costs(step.thenSteps), costs(step.elseSteps ?? []));
        else if (step.kind === 'repeat')
            cost = step.times * costs(step.steps);
        const sum = total + cost;
        if (sum > 15000)
            throw Error('Macro operation limit is 15,000 cell touches.');
        return sum;
    }, 0);
    costs(steps);
    if (version === 1)
        return { version, name, sheetName, steps: steps };
    if (version === 2)
        return { version, name, sheetName, steps: steps };
    if (version === 3)
        return { version, name, sheetName, steps: steps };
    return { version: 4, name, sheetName, parameters, steps: steps };
}
function checkRect(book, b) { if (b.r1 < 0 || b.c1 < 0 || b.r2 >= book.rows || b.c2 >= book.cols)
    throw Error('Macro range exceeds worksheet boundaries.'); }
const refs = (b) => { const out = []; for (let r = b.r1; r <= b.r2; r++)
    for (let c = b.c1; c <= b.c2; c++)
        out.push(address(r, c)); return out; };
/** Resolve only validated data fields, never action kinds or arbitrary code.
 * Parameters are strings or finite numbers. Macro source is revalidated afterward. */
function materialize(macro, supplied) {
    if (supplied !== undefined && (!isObj(supplied) || Object.keys(supplied).some(key => !macro.parameters.some(p => p.name === key))))
        throw Error('Unknown macro parameter.');
    const values = new Map();
    for (const p of macro.parameters) {
        const val = supplied?.[p.name] ?? p.defaultValue;
        if ((val === undefined || val === '') && p.required)
            throw Error(`Missing required parameter ${p.name}.`);
        if (val !== undefined && typeof val !== 'string')
            throw Error(`Invalid parameter ${p.name}.`);
        if ((val ?? '').length > 1000)
            throw Error('Macro parameter too long.');
        if (p.type === 'number' && val !== undefined && val !== '' && (!Number.isFinite(Number(val)) || val.trim() === ''))
            throw Error(`Parameter ${p.name} must be numeric.`);
        values.set(p.name, val ?? '');
    }
    const replace = (input) => input.replace(/\{\{([a-zA-Z][a-zA-Z0-9_]{0,31})\}\}/g, (_token, key) => {
        if (!values.has(key))
            throw Error(`Unknown macro placeholder ${key}.`);
        return values.get(key);
    });
    const transform = (value, key = '') => {
        if (typeof value === 'string' && ['raw', 'find', 'replace', 'value'].includes(key))
            return replace(value);
        if (Array.isArray(value))
            return value.map(x => transform(x));
        if (isObj(value))
            return Object.fromEntries(Object.entries(value).map(([k, v]) => [k, transform(v, k)]));
        return value;
    };
    const substituted = transform(macro.steps);
    return validateMacro({ version: 4, name: macro.name, sheetName: macro.sheetName, parameters: macro.parameters, steps: substituted });
}
function conditionMatches(raw, current, operator, target = '') {
    const a = Number(current), b = Number(target);
    switch (operator) {
        case 'blank': return raw === '';
        case 'notBlank': return raw !== '';
        case 'equals': return current === target;
        case 'notEquals': return current !== target;
        case 'contains': return current.toLowerCase().includes(target.toLowerCase());
        case 'greaterThan': return current !== '' && target !== '' && Number.isFinite(a) && Number.isFinite(b) && a > b;
        case 'lessThan': return current !== '' && target !== '' && Number.isFinite(a) && Number.isFinite(b) && a < b;
    }
}
function simulate(doc, macro) {
    const clone = SpreadsheetDocument.fromJSON(doc.toJSON());
    const runStep = (step) => {
        if (step.kind === 'repeat') {
            for (let i = 0; i < step.times; i++)
                for (const action of step.steps)
                    runStep(action);
            return;
        }
        const name = ('sheetName' in step && step.sheetName) || macro.sheetName;
        const entry = clone.sheets.find(x => x.name === name);
        if (!entry)
            throw Error(`Worksheet '${name}' does not exist.`);
        const book = entry.book;
        if (step.kind === 'ifCell') {
            const p = parseRef(step.ref);
            checkRect(book, { r1: p.row, r2: p.row, c1: p.col, c2: p.col });
            const val = book.getValue(step.ref);
            const current = val.error ? val.error.code : String(val.value ?? '');
            const target = step.value ?? '';
            const outcome = conditionMatches(book.getRaw(step.ref), current, step.operator, target);
            for (const nested of outcome ? step.thenSteps : step.elseSteps ?? [])
                runStep(nested);
            return;
        }
        if (step.kind === 'cleanText') {
            const b = rectangle(step.range);
            checkRect(book, b);
            const normalize = (raw) => {
                if (step.operation === 'trim')
                    return raw.trim();
                if (step.operation === 'upper')
                    return raw.toUpperCase();
                if (step.operation === 'lower')
                    return raw.toLowerCase();
                if (step.operation === 'proper')
                    return raw.toLowerCase().replace(/\b[\p{L}]/gu, s => s.toUpperCase());
                return raw.trim().replace(/\s+/g, ' ');
            };
            book.applyEdits(refs(b).filter(ref => !(step.skipFormulas && book.getRaw(ref).startsWith('='))).map(ref => ({ ref, raw: normalize(book.getRaw(ref)) })), false);
            return;
        }
        if (step.kind === 'formatIf') {
            const b = rectangle(step.range);
            checkRect(book, b);
            const edits = [];
            for (const ref of refs(b)) {
                const current = book.getValue(ref), value = current.error ? current.error.code : String(current.value ?? '');
                if (conditionMatches(book.getRaw(ref), value, step.operator, step.value))
                    edits.push({ ref, style: cleanStyle({ ...book.getStyle(ref), ...step.style }) });
            }
            book.applyStyles(edits, false);
            return;
        }
        if (step.kind === 'removeDuplicates') {
            const b = rectangle(step.range);
            checkRect(book, b);
            const from = b.r1 + (step.hasHeader ? 1 : 0), keys = step.keyColumns ?? Array.from({ length: b.c2 - b.c1 + 1 }, (_, i) => b.c1 + i);
            const seen = new Set();
            const entries = [];
            for (let r = from; r <= b.r2; r++) {
                const key = JSON.stringify(keys.map(c => { const v = book.getValue(address(r, c)); return v.error ? ['error', v.error.code] : [typeof v.value, v.value]; }));
                if (seen.has(key))
                    continue;
                seen.add(key);
                entries.push({ sourceRow: r, data: Array.from({ length: b.c2 - b.c1 + 1 }, (_, i) => book.getRaw(address(r, b.c1 + i))), styles: Array.from({ length: b.c2 - b.c1 + 1 }, (_, i) => book.getStyle(address(r, b.c1 + i))) });
            }
            const edits = [], styles = [];
            for (let r = from; r <= b.r2; r++)
                for (let c = b.c1; c <= b.c2; c++) {
                    const entry = entries[r - from], raw = entry?.data[c - b.c1] ?? '';
                    edits.push({ ref: address(r, c), raw: raw.startsWith('=') && entry ? shiftFormula(raw, r - entry.sourceRow, 0) : raw });
                    styles.push({ ref: address(r, c), style: entry?.styles[c - b.c1] ?? {} });
                }
            book.applyTransaction(edits, styles);
            return;
        }
        if (step.kind === 'sequence') {
            const b = rectangle(step.range);
            checkRect(book, b);
            let index = 0;
            book.applyEdits(refs(b).map(ref => ({ ref, raw: String(step.start + index++ * step.increment) })), false);
            return;
        }
        if (step.kind === 'cells') {
            for (const edit of step.edits) {
                const p = parseRef(edit.ref);
                checkRect(book, { r1: p.row, r2: p.row, c1: p.col, c2: p.col });
            }
            book.applyEdits(step.edits, false);
        }
        if (step.kind === 'styles') {
            for (const edit of step.edits) {
                const p = parseRef(edit.ref);
                checkRect(book, { r1: p.row, r2: p.row, c1: p.col, c2: p.col });
            }
            book.applyStyles(step.edits, false);
        }
        if (step.kind === 'clear') {
            const b = rectangle(step.range);
            checkRect(book, b);
            const cells = refs(b);
            book.applyEdits(cells.map(ref => ({ ref, raw: '' })), false);
            if (step.clearFormatting)
                book.applyStyles(cells.map(ref => ({ ref, style: {} })), false);
        }
        if (step.kind === 'fillDown') {
            const b = rectangle(step.range), source = parseRef(step.source);
            checkRect(book, b);
            checkRect(book, { r1: source.row, r2: source.row, c1: source.col, c2: source.col });
            const raw = book.getRaw(step.source), style = book.getStyle(step.source);
            book.applyTransaction(refs(b).map(ref => { const p = parseRef(ref); return { ref, raw: raw.startsWith('=') ? shiftFormula(raw, p.row - source.row, p.col - source.col) : raw }; }), refs(b).map(ref => ({ ref, style })));
        }
        if (step.kind === 'copy') {
            const b = rectangle(step.sourceRange), destination = parseRef(step.destination);
            checkRect(book, b);
            const snapshot = snapshotCells(book, b, entry.id), tx = pasteSnapshot(book, snapshot, { row: destination.row, col: destination.col }, entry.id, step.valuesOnly);
            book.applyTransaction(tx.edits, tx.styles);
        }
        if (step.kind === 'sort') {
            const b = rectangle(step.range);
            checkRect(book, b);
            const tx = sortRangeTransaction(book, b, step.keys ?? step.byColumn, step.descending, step.hasHeader);
            book.applyTransaction(tx.edits, tx.styles);
        }
        if (step.kind === 'replace')
            book.applyEdits(replaceCells(book, step.find, step.replace, step.caseSensitive, step.includeFormulas), false);
    };
    for (const step of macro.steps)
        runStep(step);
    return clone;
}
/** Complete dry-run before applying any changes; safe even for cross-sheet macros.
 * Changes to each sheet become one undoable transaction. */
export function runActionMacro(doc, data, options) {
    const source = validateMacro(data), macro = source.version === 4 ? materialize(source, options?.parameters) : source, cloned = simulate(doc, macro), changes = [];
    let values = 0, styles = 0;
    const changedSheets = [];
    const transactions = [];
    for (const updated of cloned.sheets) {
        const original = doc.sheets.find(x => x.name === updated.name);
        const oldRaw = new Map(original.book.allEntries().map(x => [x.ref, x.raw]));
        const newRaw = new Map(updated.book.allEntries().map(x => [x.ref, x.raw]));
        const oldStyles = new Map(original.book.allStyles().map(x => [x.ref, x.style]));
        const newStyles = new Map(updated.book.allStyles().map(x => [x.ref, x.style]));
        const edits = [], styleEdits = [];
        for (const ref of new Set([...oldRaw.keys(), ...newRaw.keys()])) {
            const before = oldRaw.get(ref) ?? '', after = newRaw.get(ref) ?? '';
            if (before !== after) {
                edits.push({ ref, raw: after });
                if (changes.length < 100)
                    changes.push({ sheet: updated.name, ref, before, after });
            }
        }
        for (const ref of new Set([...oldStyles.keys(), ...newStyles.keys()])) {
            const before = oldStyles.get(ref) ?? {}, after = newStyles.get(ref) ?? {};
            if (JSON.stringify(before) !== JSON.stringify(after))
                styleEdits.push({ ref, style: after });
        }
        if (edits.length || styleEdits.length) {
            changedSheets.push(updated.name);
            transactions.push({ name: updated.name, edits, styleEdits });
            values += edits.length;
            styles += styleEdits.length;
        }
    }
    if (values + styles > 15000)
        throw Error('Macro would modify too many cells.');
    if (!options?.dryRun)
        for (const tx of transactions) {
            const sheet = doc.sheets.find(s => s.name === tx.name);
            sheet.book.applyTransaction(tx.edits, tx.styleEdits);
        }
    if (macro.version === 1 && !options?.dryRun)
        return { values, styles };
    return { values, styles, sheets: changedSheets, changes, dryRun: !!options?.dryRun };
}
