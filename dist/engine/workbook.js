import { colToName, normalizeRef, parseFormula, parseRef, references, FormulaSyntaxError } from './parser.js';
import { bool, error, functions, number, SheetError, text } from './functions.js';
import { cleanStyle } from './format.js';
const displayNumber = (n) => Number.isFinite(n) ? String(Math.round((n + Number.EPSILON) * 1e10) / 1e10) : '#NUM!';
export function display(result) {
    if (result.error)
        return result.error.code;
    const x = result.value;
    return x === null ? '' : typeof x === 'boolean' ? (x ? 'TRUE' : 'FALSE') : typeof x === 'number' ? displayNumber(x) : x;
}
function tree(n, indent = '') {
    if (n.type === 'literal')
        return `${indent}Value ${JSON.stringify(n.value)}`;
    if (n.type === 'ref')
        return `${indent}Reference ${n.ref}`;
    if (n.type === 'range')
        return `${indent}Range ${n.from}:${n.to}`;
    if (n.type === 'unary')
        return `${indent}Unary ${n.op}\n${tree(n.expr, indent + '  ')}`;
    if (n.type === 'binary')
        return `${indent}Operator ${n.op}\n${tree(n.left, indent + '  ')}\n${tree(n.right, indent + '  ')}`;
    return `${indent}Function ${n.name}\n${n.args.map(arg => tree(arg, indent + '  ')).join('\n')}`;
}
const known = (ref, rows, cols) => {
    const r = parseRef(ref);
    if (!r || r.row >= rows || r.col >= cols)
        return error('#REF!', `Cell '${ref}' is outside this workbook (${rows} rows × ${cols} columns).`);
    return normalizeRef(ref);
};
const simple = (raw) => {
    if (raw.startsWith("'"))
        return raw.slice(1);
    if (/^(true|false)$/i.test(raw))
        return raw.toLowerCase() === 'true';
    const trimmed = raw.trim();
    if (trimmed !== '' && /^[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?$/.test(trimmed))
        return Number(trimmed);
    return raw === '' ? null : raw;
};
const diag = (code, message, position = 0, suggestion = 'Check the formula and its referenced cells.') => ({ code, message, position, length: 1, suggestion });
export class Workbook {
    rows;
    cols;
    raw = new Map();
    styles = new Map();
    ast = new Map();
    parseErrors = new Map();
    deps = new Map();
    rev = new Map();
    cache = new Map();
    dirty = new Set();
    undoStack = [];
    redoStack = [];
    lastStats = { updated: 0, recalculated: 0, affected: [], durationMs: 0, totalPopulated: 0 };
    constructor(rows = 400, cols = 52) { this.rows = rows; this.cols = cols; }
    getRaw(ref) { return this.raw.get(known(ref, this.rows, this.cols)) ?? ''; }
    allEntries() { return [...this.raw].map(([ref, raw]) => ({ ref, raw })); }
    allStyles() { return [...this.styles].map(([ref, style]) => ({ ref, style: { ...style } })); }
    getStyle(ref) { return { ...(this.styles.get(known(ref, this.rows, this.cols)) ?? {}) }; }
    /** Apply complete style snapshots. No calculation is invalidated by formatting changes. */
    applyStyles(edits, recordHistory = true) {
        const changes = new Map();
        for (const item of edits)
            changes.set(known(item.ref, this.rows, this.cols), cleanStyle(item.style));
        const changed = [...changes].filter(([ref, style]) => JSON.stringify(this.getStyle(ref)) !== JSON.stringify(style));
        if (!changed.length)
            return 0;
        const before = changed.map(([ref]) => ({ ref, style: this.getStyle(ref) }));
        for (const [ref, style] of changed) {
            if (Object.keys(style).length)
                this.styles.set(ref, style);
            else
                this.styles.delete(ref);
        }
        if (recordHistory) {
            this.undoStack.push({ before: [], after: [], styleBefore: before, styleAfter: changed.map(([ref, style]) => ({ ref, style })) });
            if (this.undoStack.length > 150)
                this.undoStack.shift();
            this.redoStack = [];
        }
        return changed.length;
    }
    patchStyles(refs, patch) {
        // Validate all incoming fields, but retain explicit false/default values to clear an attribute.
        cleanStyle(patch);
        return this.applyStyles(refs.map(ref => {
            const result = { ...this.getStyle(ref) };
            for (const [key, value] of Object.entries(patch)) {
                if (value === false || (key === 'numberFormat' && value === 'general') || (key === 'align' && value === 'left'))
                    delete result[key];
                else
                    result[key] = value;
            }
            return { ref, style: result };
        }));
    }
    clearStyles(refs) { return this.applyStyles(refs.map(ref => ({ ref, style: {} }))); }
    update(ref, raw) {
        for (const oldRef of this.deps.get(ref) ?? []) {
            this.rev.get(oldRef)?.delete(ref);
            if (this.rev.get(oldRef)?.size === 0)
                this.rev.delete(oldRef);
        }
        this.deps.delete(ref);
        this.ast.delete(ref);
        this.parseErrors.delete(ref);
        if (raw)
            this.raw.set(ref, raw);
        else
            this.raw.delete(ref);
        if (raw.startsWith('=')) {
            try {
                const parsed = parseFormula(raw);
                const refs = references(parsed, this.rows, this.cols);
                this.ast.set(ref, parsed);
                this.deps.set(ref, refs);
                for (const dep of refs) {
                    if (!this.rev.has(dep))
                        this.rev.set(dep, new Set());
                    this.rev.get(dep).add(ref);
                }
            }
            catch (e) {
                const d = e instanceof FormulaSyntaxError ? e.diagnostic : diag('#PARSE!', e instanceof Error ? e.message : String(e));
                this.parseErrors.set(ref, d);
            }
        }
    }
    invalidate(refs) {
        const queue = [...refs], affected = new Set(queue);
        for (let i = 0; i < queue.length; i++)
            for (const dependent of this.rev.get(queue[i]) ?? [])
                if (!affected.has(dependent)) {
                    affected.add(dependent);
                    queue.push(dependent);
                }
        for (const ref of affected) {
            this.cache.delete(ref);
            this.dirty.add(ref);
        }
        return [...affected];
    }
    applyEdits(edits, recordHistory = true) {
        const sanitized = new Map();
        for (const e of edits)
            sanitized.set(known(e.ref, this.rows, this.cols), e.raw);
        const changed = [...sanitized].filter(([ref, raw]) => this.getRaw(ref) !== raw);
        if (!changed.length)
            return { ...this.lastStats, updated: 0, recalculated: 0, affected: [], durationMs: 0 };
        const before = changed.map(([ref]) => ({ ref, raw: this.getRaw(ref) }));
        for (const [ref, raw] of changed)
            this.update(ref, raw);
        const affected = this.invalidate(changed.map(([ref]) => ref));
        const t0 = performance.now();
        let recalculated = 0;
        for (const ref of affected)
            if (this.raw.has(ref)) {
                this.getValue(ref);
                recalculated++;
            }
        this.lastStats = { updated: changed.length, recalculated, affected, durationMs: Math.round((performance.now() - t0) * 1000) / 1000, totalPopulated: this.raw.size };
        if (recordHistory) {
            this.undoStack.push({ before, after: changed.map(([ref, raw]) => ({ ref, raw })), styleBefore: [], styleAfter: [] });
            if (this.undoStack.length > 150)
                this.undoStack.shift();
            this.redoStack = [];
        }
        return this.lastStats;
    }
    /** One undoable transaction for operations that change both values and formatting. */
    applyTransaction(edits, styleEdits) {
        const oldValues = new Map();
        const oldStyles = new Map();
        for (const item of edits)
            oldValues.set(known(item.ref, this.rows, this.cols), this.getRaw(item.ref));
        for (const item of styleEdits)
            oldStyles.set(known(item.ref, this.rows, this.cols), this.getStyle(item.ref));
        // Validate styles before making any changes.
        for (const item of styleEdits)
            cleanStyle(item.style);
        const result = this.applyEdits(edits, false);
        this.applyStyles(styleEdits, false);
        const before = [...oldValues].map(([ref, raw]) => ({ ref, raw }));
        const after = [...oldValues.keys()].map(ref => ({ ref, raw: this.getRaw(ref) }));
        const styleBefore = [...oldStyles].map(([ref, style]) => ({ ref, style }));
        const styleAfter = [...oldStyles.keys()].map(ref => ({ ref, style: this.getStyle(ref) }));
        if (before.some((b, i) => b.raw !== after[i]?.raw) || styleBefore.some((b, i) => JSON.stringify(b.style) !== JSON.stringify(styleAfter[i]?.style))) {
            this.undoStack.push({ before, after, styleBefore, styleAfter });
            if (this.undoStack.length > 150)
                this.undoStack.shift();
            this.redoStack = [];
        }
        return result;
    }
    setCell(ref, raw) { return this.applyEdits([{ ref, raw }]); }
    undo() { const op = this.undoStack.pop(); if (!op)
        return false; this.applyEdits(op.before, false); this.applyStyles(op.styleBefore, false); this.redoStack.push(op); return true; }
    redo() { const op = this.redoStack.pop(); if (!op)
        return false; this.applyEdits(op.after, false); this.applyStyles(op.styleAfter, false); this.undoStack.push(op); return true; }
    get canUndo() { return this.undoStack.length > 0; }
    get canRedo() { return this.redoStack.length > 0; }
    getValue(ref, visiting = []) {
        try {
            ref = known(ref, this.rows, this.cols);
        }
        catch (e) {
            return { value: null, error: diag('#REF!', e.message) };
        }
        const old = this.cache.get(ref);
        if (old && !this.dirty.has(ref))
            return old;
        if (visiting.includes(ref)) {
            return { value: null, error: diag('#CYCLE!', `Circular reference: ${[...visiting.slice(visiting.indexOf(ref)), ref].join(' → ')}`, 0, 'Remove a self-reference or break the loop between these cells.') };
        }
        const parseError = this.parseErrors.get(ref);
        if (parseError) {
            const output = { value: null, error: parseError };
            this.cache.set(ref, output);
            this.dirty.delete(ref);
            return output;
        }
        const ast = this.ast.get(ref);
        let output;
        if (!ast)
            output = { value: simple(this.raw.get(ref) ?? '') };
        else {
            try {
                const value = this.evaluate(ast, [...visiting, ref]);
                if (Array.isArray(value))
                    error('#VALUE!', 'A range cannot be the final value of a single cell. Use SUM, INDEX or another function.');
                output = { value: value };
            }
            catch (e) {
                const x = e instanceof SheetError ? e : new SheetError('#ERROR!', e instanceof Error ? e.message : String(e));
                output = { value: null, error: diag(x.code, x.message, 0, x.code === '#NAME?' ? 'Check the spelling of the function name.' : 'Check the formula and its referenced cells.') };
            }
        }
        this.cache.set(ref, output);
        this.dirty.delete(ref);
        return output;
    }
    evaluate(n, visiting) {
        if (n.type === 'literal')
            return n.value;
        if (n.type === 'ref') {
            const key = known(n.ref, this.rows, this.cols);
            const v = this.getValue(key, visiting);
            if (v.error)
                error(v.error.code, `${key}: ${v.error.message}`);
            return v.value;
        }
        if (n.type === 'range') {
            const a = parseRef(n.from), b = parseRef(n.to);
            if (Math.max(a.row, b.row) >= this.rows || Math.max(a.col, b.col) >= this.cols)
                error('#REF!', 'Range extends beyond sheet boundaries.');
            const matrix = [];
            for (let row = Math.min(a.row, b.row); row <= Math.max(a.row, b.row); row++) {
                const line = [];
                for (let col = Math.min(a.col, b.col); col <= Math.max(a.col, b.col); col++) {
                    const key = `${colToName(col)}${row + 1}`;
                    const val = this.getValue(key, visiting);
                    if (val.error)
                        error(val.error.code, `${key}: ${val.error.message}`);
                    line.push(val.value);
                }
                matrix.push(line);
            }
            return matrix;
        }
        if (n.type === 'unary') {
            const v = this.evaluateScalar(n.expr, visiting);
            return n.op === '-' ? -number(v) : number(v);
        }
        if (n.type === 'binary') {
            const l = this.evaluateScalar(n.left, visiting), r = this.evaluateScalar(n.right, visiting);
            if (n.op === '&')
                return text(l) + text(r);
            if (['=', '<>', '!=', '<', '>', '<=', '>='].includes(n.op)) {
                const comparable = typeof l === 'number' && typeof r === 'number'
                    ? (l === r ? 0 : l < r ? -1 : 1) : text(l).toLowerCase().localeCompare(text(r).toLowerCase());
                switch (n.op) {
                    case '=': return comparable === 0;
                    case '<>':
                    case '!=': return comparable !== 0;
                    case '<': return comparable < 0;
                    case '>': return comparable > 0;
                    case '<=': return comparable <= 0;
                    default: return comparable >= 0;
                }
            }
            const a = number(l), b = number(r);
            let v;
            switch (n.op) {
                case '+':
                    v = a + b;
                    break;
                case '-':
                    v = a - b;
                    break;
                case '*':
                    v = a * b;
                    break;
                case '/':
                    if (b === 0)
                        error('#DIV/0!', 'Division by zero.');
                    v = a / b;
                    break;
                case '^':
                    v = a ** b;
                    break;
                default: return error('#PARSE!', `Unsupported operator ${n.op}.`);
            }
            if (!Number.isFinite(v))
                error('#NUM!', 'Arithmetic result is not finite.');
            return v;
        }
        if (n.type === 'call') {
            const f = functions[n.name];
            if (!f) {
                const candidates = Object.keys(functions).filter(x => x.startsWith(n.name.slice(0, 2))).slice(0, 4);
                return error('#NAME?', `Unknown function ${n.name}.${candidates.length ? ` Did you mean ${candidates.join(' or ')}?` : ''}`);
            }
            if (n.name === 'IF') {
                if (n.args.length < 2 || n.args.length > 3)
                    error('#ARG!', 'IF(condition, true_value, false_value) takes 2–3 arguments.');
                const condition = bool(this.evaluateScalar(n.args[0], visiting));
                return condition ? this.evaluate(n.args[1], visiting) : n.args[2] ? this.evaluate(n.args[2], visiting) : false;
            }
            if (n.name === 'IFERROR') {
                if (n.args.length !== 2)
                    error('#ARG!', 'IFERROR(value, fallback) takes exactly 2 arguments.');
                try {
                    return this.evaluate(n.args[0], visiting);
                }
                catch {
                    return this.evaluate(n.args[1], visiting);
                }
            }
            return f.run(n.args.map(arg => this.evaluate(arg, visiting)));
        }
        return error('#ERROR!', 'Unexpected AST node.');
    }
    evaluateScalar(n, visiting) {
        const v = this.evaluate(n, visiting);
        return Array.isArray(v) ? error('#VALUE!', 'Cannot perform scalar arithmetic on a whole range.') : v;
    }
    inspect(ref) {
        ref = known(ref, this.rows, this.cols);
        const ds = [...(this.deps.get(ref) ?? [])];
        const rvs = [...(this.rev.get(ref) ?? [])];
        return { ref, raw: this.getRaw(ref), result: this.getValue(ref), dependencies: ds, dependents: rvs,
            expressionTree: this.ast.has(ref) ? tree(this.ast.get(ref)) : this.parseErrors.has(ref) ? 'Unable to parse formula' : 'Literal / empty cell',
            referencedValues: ds.slice(0, 100).map(r => ({ ref: r, raw: this.getRaw(r), display: display(this.getValue(r)) })) };
    }
    preview(ref, proposed) {
        const target = known(ref, this.rows, this.cols);
        const impacted = this.transitiveDependents(target);
        const before = new Map(impacted.map(id => [id, display(this.getValue(id))]));
        const cloned = Workbook.fromJSON(this.toJSON());
        const stats = cloned.applyEdits([{ ref: target, raw: proposed }], false);
        return { changes: impacted.map(id => ({ ref: id, before: before.get(id) ?? '', after: display(cloned.getValue(id)) })).filter(c => c.before !== c.after), stats };
    }
    transitiveDependents(ref) {
        const queue = [known(ref, this.rows, this.cols)], seen = new Set(queue);
        for (let i = 0; i < queue.length; i++)
            for (const k of this.rev.get(queue[i]) ?? [])
                if (!seen.has(k)) {
                    seen.add(k);
                    queue.push(k);
                }
        return queue;
    }
    toJSON() { return { version: 2, rows: this.rows, cols: this.cols, cells: Object.fromEntries(this.raw), styles: Object.fromEntries(this.styles) }; }
    static fromJSON(data) {
        if (!data || typeof data !== 'object')
            throw Error('Invalid workbook data.');
        const d = data;
        if (!Number.isInteger(d.rows) || !Number.isInteger(d.cols) || (d.rows ?? 0) < 1 || (d.rows ?? 0) > 1000 || (d.cols ?? 0) < 1 || (d.cols ?? 0) > 104)
            throw Error('Workbook dimensions must be 1–1000 rows and 1–104 columns.');
        if (!d.cells || typeof d.cells !== 'object' || Array.isArray(d.cells))
            throw Error('Workbook cells must be an object.');
        const book = new Workbook(d.rows, d.cols);
        const edits = Object.entries(d.cells).map(([ref, raw]) => { if (typeof raw !== 'string')
            throw Error(`Invalid content in ${ref}.`); return { ref, raw }; });
        if (edits.length > 104000)
            throw Error('Workbook contains too many populated cells.');
        book.applyEdits(edits, false);
        if (d.styles !== undefined) {
            if (!d.styles || typeof d.styles !== 'object' || Array.isArray(d.styles))
                throw Error('Invalid workbook formatting.');
            if (Object.keys(d.styles).length > 104000)
                throw Error('Too many styled cells.');
            book.applyStyles(Object.entries(d.styles).map(([ref, style]) => ({ ref, style: cleanStyle(style) })), false);
        }
        return book;
    }
}
