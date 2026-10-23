export class FormulaSyntaxError extends Error {
    diagnostic;
    constructor(message, position, suggestion, length = 1) {
        super(message);
        this.name = 'FormulaSyntaxError';
        this.diagnostic = { code: '#PARSE!', message, position, length, suggestion };
    }
}
export const nameToCol = (name) => {
    let n = 0;
    for (const c of name.toUpperCase()) {
        if (c < 'A' || c > 'Z')
            return -1;
        n = n * 26 + (c.charCodeAt(0) - 64);
    }
    return n - 1;
};
export const colToName = (column) => {
    let n = column + 1, out = '';
    if (!Number.isInteger(n) || n <= 0)
        return '#REF!';
    while (n > 0) {
        n--;
        out = String.fromCharCode(65 + (n % 26)) + out;
        n = Math.floor(n / 26);
    }
    return out;
};
export function parseRef(ref) {
    const m = /^(\$?)([A-Za-z]{1,3})(\$?)([1-9][0-9]*)$/.exec(ref);
    if (!m)
        return null;
    const col = nameToCol(m[2]);
    const row = Number(m[4]) - 1;
    if (col < 0 || !Number.isSafeInteger(row))
        return null;
    return { col, row, colLock: m[1] === '$', rowLock: m[3] === '$' };
}
export const normalizeRef = (ref) => {
    const p = parseRef(ref);
    return p ? `${colToName(p.col)}${p.row + 1}` : '#REF!';
};
export function tokenize(input) {
    const result = [];
    let i = 0;
    while (i < input.length) {
        if (/\s/.test(input[i])) {
            i++;
            continue;
        }
        const start = i;
        const tail = input.slice(i);
        if (input[i] === '"') {
            i++;
            let complete = false;
            while (i < input.length) {
                if (input[i] === '"') {
                    if (input[i + 1] === '"') {
                        i += 2;
                        continue;
                    }
                    i++;
                    complete = true;
                    break;
                }
                i++;
            }
            if (!complete)
                throw new FormulaSyntaxError('Unterminated text string.', start, 'Add a closing double quotation mark.');
            result.push({ kind: 'string', text: input.slice(start, i), start, end: i });
            continue;
        }
        const number = /^(?:\d+(?:\.\d*)?|\.\d+)(?:[Ee][+-]?\d+)?/.exec(tail);
        if (number) {
            i += number[0].length;
            result.push({ kind: 'number', text: number[0], start, end: i });
            continue;
        }
        const ref = /^\$?[A-Za-z]{1,3}\$?[1-9][0-9]*/.exec(tail);
        // LOG10(100) is a function call even though LOG10 also resembles a cell address.
        if (ref && tail[ref[0].length] !== '(' && !/[A-Za-z0-9_]/.test(tail[ref[0].length] ?? ' ')) {
            i += ref[0].length;
            result.push({ kind: 'ref', text: ref[0].toUpperCase(), start, end: i });
            continue;
        }
        const identifier = /^[A-Za-z_][A-Za-z0-9_.]*/.exec(tail);
        if (identifier) {
            i += identifier[0].length;
            result.push({ kind: 'name', text: identifier[0].toUpperCase(), start, end: i });
            continue;
        }
        const op = /^(<=|>=|<>|!=|[+\-*/^&=<>(),:])/.exec(tail);
        if (op) {
            i += op[0].length;
            result.push({ kind: 'operator', text: op[0], start, end: i });
            continue;
        }
        throw new FormulaSyntaxError(`Unexpected character '${input[i]}'.`, i, 'Remove or replace this character.');
    }
    result.push({ kind: 'eof', text: '', start: i, end: i });
    return result;
}
const precedence = { '=': 1, '<>': 1, '!=': 1, '<': 1, '>': 1, '<=': 1, '>=': 1, '&': 2, '+': 3, '-': 3, '*': 4, '/': 4, '^': 6 };
class Parser {
    tokens;
    source;
    index = 0;
    constructor(tokens, source) {
        this.tokens = tokens;
        this.source = source;
    }
    peek() { return this.tokens[this.index]; }
    take() { return this.tokens[this.index++]; }
    consume(value) { if (this.peek().text === value) {
        this.take();
        return true;
    } return false; }
    parse() {
        if (this.peek().kind === 'eof')
            throw new FormulaSyntaxError('Formula is empty.', 0, 'Enter an expression after =.');
        const node = this.expr(0);
        const extra = this.peek();
        if (extra.kind !== 'eof')
            throw new FormulaSyntaxError(`Unexpected '${extra.text}'.`, extra.start, 'Remove the extra expression or insert an operator.', Math.max(1, extra.text.length));
        return node;
    }
    expr(minPrec) {
        let left = this.primary();
        while (true) {
            const operator = this.peek();
            const prec = precedence[operator.text];
            if (operator.kind !== 'operator' || prec === undefined || prec < minPrec)
                break;
            this.take();
            const right = this.expr(operator.text === '^' ? prec : prec + 1);
            left = { type: 'binary', op: operator.text, left, right, start: left.start, end: right.end };
        }
        return left;
    }
    primary() {
        const t = this.take();
        if (t.kind === 'number')
            return { type: 'literal', value: Number(t.text), start: t.start, end: t.end };
        if (t.kind === 'string')
            return { type: 'literal', value: t.text.slice(1, -1).replace(/""/g, '"'), start: t.start, end: t.end };
        if (t.kind === 'ref') {
            if (this.consume(':')) {
                const other = this.take();
                if (other.kind !== 'ref')
                    throw new FormulaSyntaxError('A range needs a cell reference after the colon.', other.start, 'Use a range such as A1:B10.');
                return { type: 'range', from: t.text, to: other.text, start: t.start, end: other.end };
            }
            return { type: 'ref', ref: t.text, start: t.start, end: t.end };
        }
        if (t.text === '-' || t.text === '+') {
            const e = this.expr(5);
            return { type: 'unary', op: t.text, expr: e, start: t.start, end: e.end };
        }
        if (t.text === '(') {
            const e = this.expr(0);
            if (!this.consume(')'))
                throw new FormulaSyntaxError('Missing closing parenthesis.', this.peek().start, "Add ')' at the end of this expression.");
            return e;
        }
        if (t.kind === 'name') {
            if ((t.text === 'TRUE' || t.text === 'FALSE') && this.peek().text !== '(')
                return { type: 'literal', value: t.text === 'TRUE', start: t.start, end: t.end };
            if (!this.consume('('))
                throw new FormulaSyntaxError(`Unknown name '${t.text}'.`, t.start, 'Use a cell reference or a supported function such as SUM(...).', t.text.length);
            const args = [];
            if (!this.consume(')')) {
                while (true) {
                    if (this.peek().kind === 'eof')
                        throw new FormulaSyntaxError('Missing closing parenthesis.', this.peek().start, "Add ')' at the end of the formula.");
                    args.push(this.expr(0));
                    if (this.consume(')'))
                        break;
                    if (!this.consume(',')) {
                        if (this.peek().kind === 'eof')
                            throw new FormulaSyntaxError('Missing closing parenthesis.', this.peek().start, "Add ')' at the end of the formula.");
                        throw new FormulaSyntaxError("Expected ',' or ')' between function arguments.", this.peek().start, "Separate arguments with ',' and close with ')'.");
                    }
                }
            }
            return { type: 'call', name: t.text, args, start: t.start, end: this.tokens[this.index - 1].end };
        }
        if (t.kind === 'eof')
            throw new FormulaSyntaxError('Incomplete expression.', t.start, 'Add a value, reference, or function.');
        throw new FormulaSyntaxError(`Unexpected '${t.text}'.`, t.start, 'Enter a number, cell reference, or function here.');
    }
}
export const parseFormula = (source) => new Parser(tokenize(source.startsWith('=') ? source.slice(1) : source), source).parse();
export function references(ast, rows, cols) {
    const refs = new Set();
    const visit = (n) => {
        if (n.type === 'ref') {
            const p = parseRef(n.ref);
            if (p && p.row < rows && p.col < cols)
                refs.add(normalizeRef(n.ref));
            return;
        }
        if (n.type === 'range') {
            const a = parseRef(n.from), b = parseRef(n.to);
            if (!a || !b)
                return;
            const lowR = Math.min(a.row, b.row), highR = Math.max(a.row, b.row);
            const lowC = Math.min(a.col, b.col), highC = Math.max(a.col, b.col);
            if ((highR - lowR + 1) * (highC - lowC + 1) > 20000)
                throw new FormulaSyntaxError('Range exceeds 20,000 cells.', n.start, 'Use a smaller range.');
            for (let r = lowR; r <= Math.min(highR, rows - 1); r++)
                for (let c = lowC; c <= Math.min(highC, cols - 1); c++)
                    refs.add(`${colToName(c)}${r + 1}`);
            return;
        }
        if (n.type === 'unary')
            visit(n.expr);
        else if (n.type === 'binary') {
            visit(n.left);
            visit(n.right);
        }
        else if (n.type === 'call')
            n.args.forEach(visit);
    };
    visit(ast);
    return refs;
}
export function shiftFormula(formula, deltaRows, deltaCols) {
    if (!formula.startsWith('='))
        return formula;
    const body = formula.slice(1);
    const parts = tokenize(body);
    let cursor = 0, result = '=';
    for (const t of parts) {
        if (t.kind === 'eof')
            break;
        result += body.slice(cursor, t.start);
        cursor = t.end;
        if (t.kind !== 'ref') {
            result += t.text;
            continue;
        }
        const p = parseRef(t.text);
        const r = p.row + (p.rowLock ? 0 : deltaRows), c = p.col + (p.colLock ? 0 : deltaCols);
        result += r < 0 || c < 0 ? '#REF!' : `${p.colLock ? '$' : ''}${colToName(c)}${p.rowLock ? '$' : ''}${r + 1}`;
    }
    return result + body.slice(cursor);
}
