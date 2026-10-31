/** A hand-written tokenizer + Pratt parser. No spreadsheet/formula libraries are used. */
export type Value = string | number | boolean | null;
export type Node =
  | { type: 'literal'; value: Value; start: number; end: number }
  | { type: 'error'; code: string; start:number; end:number }
  | { type: 'ref'; ref: string; sheet?: string; start: number; end: number }
  | { type: 'spillRef'; ref:string; sheet?:string; start:number; end:number }
  | { type: 'range'; from: string; to: string; sheet?: string; start: number; end: number }
  | { type: 'definedName'; name: string; start: number; end: number }
  | { type: 'unary'; op: string; expr: Node; start: number; end: number }
  | { type: 'binary'; op: string; left: Node; right: Node; start: number; end: number }
  | { type: 'call'; name: string; args: Node[]; start: number; end: number };
export interface Token { kind: 'number' | 'string' | 'error' | 'ref' | 'sheetref' | 'spillref' | 'name' | 'operator' | 'eof'; text: string; start: number; end: number }
export interface Diagnostic { code: string; message: string; position: number; length: number; suggestion: string }
export class FormulaSyntaxError extends Error {
  readonly diagnostic: Diagnostic;
  constructor(message: string, position: number, suggestion: string, length = 1) {
    super(message); this.name = 'FormulaSyntaxError';
    this.diagnostic = { code: '#PARSE!', message, position, length, suggestion };
  }
}
export const nameToCol = (name: string): number => {
  let n = 0;
  for (const c of name.toUpperCase()) { if (c < 'A' || c > 'Z') return -1; n = n * 26 + (c.charCodeAt(0) - 64); }
  return n - 1;
};
export const colToName = (column: number): string => {
  let n = column + 1, out = '';
  if (!Number.isInteger(n) || n <= 0) return '#REF!';
  while (n > 0) { n--; out = String.fromCharCode(65 + (n % 26)) + out; n = Math.floor(n / 26); }
  return out;
};
export interface RefParts { col: number; row: number; colLock: boolean; rowLock: boolean }
export function parseRef(ref: string): RefParts | null {
  const m = /^(\$?)([A-Za-z]{1,3})(\$?)([1-9][0-9]*)$/.exec(ref);
  if (!m) return null;
  const col = nameToCol(m[2]!); const row = Number(m[4]!) - 1;
  if (col < 0 || !Number.isSafeInteger(row)) return null;
  return { col, row, colLock: m[1] === '$', rowLock: m[3] === '$' };
}
export const normalizeRef = (ref: string): string => {
  const p = parseRef(ref); return p ? `${colToName(p.col)}${p.row + 1}` : '#REF!';
};
export function tokenize(input: string): Token[] {
  const result: Token[] = []; let i = 0;
  while (i < input.length) {
    if (/\s/.test(input[i]!)) { i++; continue; }
    const start = i; const tail = input.slice(i);
    if (input[i] === '"') {
      i++; let complete = false;
      while (i < input.length) {
        if (input[i] === '"') { if (input[i + 1] === '"') { i += 2; continue; } i++; complete = true; break; }
        i++;
      }
      if (!complete) throw new FormulaSyntaxError('Unterminated text string.', start, 'Add a closing double quotation mark.');
      result.push({ kind: 'string', text: input.slice(start, i), start, end: i }); continue;
    }
    const errorToken=/^#(?:NULL!|DIV\/0!|VALUE!|REF!|NAME\?|NUM!|N\/A|GETTING_DATA|SPILL!|CALC!|DATA!|CYCLE!|FIELD!|BLOCKED!|CONNECT!|UNKNOWN!|BUSY!|PYTHON!)/i.exec(tail);
    if(errorToken){i+=errorToken[0].length;result.push({kind:'error',text:errorToken[0].toUpperCase(),start,end:i});continue;}
    const sheet = /^(?:'((?:[^']|'')+)'|([A-Za-z_][A-Za-z0-9_.]*))!/.exec(tail);
    if (sheet) {
      const cell = /^\$?[A-Za-z]{1,3}\$?[1-9][0-9]*/.exec(tail.slice(sheet[0].length));
      if (cell) {
        i += sheet[0].length + cell[0].length;
        let range = '';
        const rest = input.slice(i), colon = /^:\$?[A-Za-z]{1,3}\$?[1-9][0-9]*/.exec(rest);
        if (colon) { range = colon[0]; i += range.length; }
        const isSpill=!range&&input[i]==='#';if(isSpill)i++;
        result.push({ kind: isSpill?'spillref':'sheetref', text: `${(sheet[1] ?? sheet[2]!).replace(/''/g, "'")}!${cell[0]}${range}`, start, end: i }); continue;
      }
    }
    const number = /^(?:\d+(?:\.\d*)?|\.\d+)(?:[Ee][+-]?\d+)?/.exec(tail);
    if (number) { i += number[0].length; result.push({ kind: 'number', text: number[0], start, end: i }); continue; }
    const ref = /^\$?[A-Za-z]{1,3}\$?[1-9][0-9]*/.exec(tail);
    if(ref&&tail[ref[0].length]==='#'&&tail[ref[0].length+1]!=='#') {i+=ref[0].length+1;result.push({kind:'spillref',text:ref[0].toUpperCase(),start,end:i});continue;}
    // LOG10(100) is a function call even though LOG10 also resembles a cell address.
    if (ref && tail[ref[0].length] !== '(' && !/[A-Za-z0-9_]/.test(tail[ref[0].length] ?? ' ')) {
      i += ref[0].length; result.push({ kind: 'ref', text: ref[0].toUpperCase(), start, end: i }); continue;
    }
    const identifier = /^[A-Za-z_][A-Za-z0-9_.]*/.exec(tail);
    if (identifier) { i += identifier[0].length; result.push({ kind: 'name', text: identifier[0].toUpperCase(), start, end: i }); continue; }
    const op = /^(<=|>=|<>|!=|[+\-*/^&=<>(),:;%@])/.exec(tail);
    if (op) { i += op[0].length; result.push({ kind: 'operator', text: op[0], start, end: i }); continue; }
    throw new FormulaSyntaxError(`Unexpected character '${input[i]}'.`, i, 'Remove or replace this character.');
  }
  result.push({ kind: 'eof', text: '', start: i, end: i }); return result;
}
/** Normalize decimal commas when the formula uses semicolons as its list separator. */
export function normalizeFormulaLocale(source:string):string {
  let quoted=false,hasSemicolon=false;
  for(let i=0;i<source.length;i++){
    if(source[i]==='"'){if(quoted&&source[i+1]==='"'){i++;continue;}quoted=!quoted;continue;}
    if(!quoted&&source[i]===';'){hasSemicolon=true;break;}
  }
  if(!hasSemicolon)return source;
  quoted=false;let out='';
  for(let i=0;i<source.length;i++){
    const character=source[i]!;
    if(character==='"'){out+=character;if(quoted&&source[i+1]==='"'){out+='"';i++;continue;}quoted=!quoted;continue;}
    out+=!quoted&&character===','&&/\d/.test(source[i-1]??'')&&/\d/.test(source[i+1]??'')?'.':character;
  }
  return out;
}
const precedence: Record<string, number> = { '=': 1, '<>': 1, '!=': 1, '<': 1, '>': 1, '<=': 1, '>=': 1, '&': 2, '+': 3, '-': 3, '*': 4, '/': 4, '^': 6 };
class Parser {
  private index = 0;
  constructor(private readonly tokens: Token[], private readonly source: string) {}
  private peek(): Token { return this.tokens[this.index]!; }
  private take(): Token { return this.tokens[this.index++]!; }
  private consume(value: string): boolean { if (this.peek().text === value) { this.take(); return true; } return false; }
  private callArguments():Node[]{
    const args:Node[]=[];let separator:','|';'|undefined;
    if(this.consume(')'))return args;
    while(true){
      if(this.peek().text===','||this.peek().text===')')args.push({type:'literal',value:null,start:this.peek().start,end:this.peek().start});
      else args.push(this.expr(0));
      if(this.consume(')'))break;
      const next=this.peek();if(next.text!==','&&next.text!==';')throw new FormulaSyntaxError("Expected ',' or ';' between function arguments.",next.start,"Separate arguments with one consistent list separator and close with ')'.");
      if(separator&&separator!==next.text)throw new FormulaSyntaxError('Do not mix comma and semicolon argument separators in one function call.',next.start,'Use either commas or semicolons throughout this function call.');
      separator=next.text as ','|';';this.take();
    }
    return args;
  }
  parse(): Node {
    if (this.peek().kind === 'eof') throw new FormulaSyntaxError('Formula is empty.', 0, 'Enter an expression after =.');
    const node = this.expr(0); const extra = this.peek();
    if (extra.kind !== 'eof') throw new FormulaSyntaxError(`Unexpected '${extra.text}'.`, extra.start, 'Remove the extra expression or insert an operator.', Math.max(1, extra.text.length));
    return node;
  }
  private expr(minPrec: number): Node {
    let left = this.primary();
    while (true) {
      const operator = this.peek(); const prec = precedence[operator.text];
      if(operator.kind==='operator'&&operator.text==='%') {this.take();left={type:'unary',op:'%',expr:left,start:left.start,end:operator.end};continue;}
      if (operator.kind !== 'operator' || prec === undefined || prec < minPrec) break;
      this.take(); const right = this.expr(operator.text === '^' ? prec : prec + 1);
      left = { type: 'binary', op: operator.text, left, right, start: left.start, end: right.end };
    }
    return left;
  }
  private primary(): Node {
    const t = this.take();
    if (t.kind === 'number') return { type: 'literal', value: Number(t.text), start: t.start, end: t.end };
    if (t.kind === 'string') return { type: 'literal', value: t.text.slice(1, -1).replace(/""/g, '"'), start: t.start, end: t.end };
    if (t.kind === 'error') return {type:'error',code:t.text,start:t.start,end:t.end};
    if (t.kind === 'ref') {
      if (this.consume(':')) {
        const other = this.take();
        if (other.kind !== 'ref') throw new FormulaSyntaxError('A range needs a cell reference after the colon.', other.start, 'Use a range such as A1:B10.');
        return { type: 'range', from: t.text, to: other.text, start: t.start, end: other.end };
      }
      return { type: 'ref', ref: t.text, start: t.start, end: t.end };
    }
    if (t.kind === 'sheetref') {
      const bang = t.text.indexOf('!'), sheet = t.text.slice(0, bang), ref = t.text.slice(bang + 1), colon = ref.indexOf(':');
      if (colon >= 0) return { type: 'range', from: ref.slice(0, colon), to: ref.slice(colon + 1), sheet, start: t.start, end: t.end };
      return { type: 'ref', ref, sheet, start: t.start, end: t.end };
    }
    if(t.kind==='spillref'){const bang=t.text.indexOf('!');return bang<0?{type:'spillRef',ref:t.text,start:t.start,end:t.end}:{type:'spillRef',sheet:t.text.slice(0,bang),ref:t.text.slice(bang+1),start:t.start,end:t.end};}
    if (t.text === '-' || t.text === '+' || t.text === '@') {
      const e = this.expr(5); return { type: 'unary', op: t.text, expr: e, start: t.start, end: e.end };
    }
    if (t.text === '(') {
      const e = this.expr(0);
      if (!this.consume(')')) throw new FormulaSyntaxError('Missing closing parenthesis.', this.peek().start, "Add ')' at the end of this expression.");
      return e;
    }
    if (t.kind === 'name') {
      if ((t.text === 'TRUE' || t.text === 'FALSE') && this.peek().text !== '(') return { type: 'literal', value: t.text === 'TRUE', start: t.start, end: t.end };
      if (!this.consume('(')) return { type: 'definedName', name:t.text, start:t.start, end:t.end };
      const args=this.callArguments(),callee:Node={ type: 'call', name: t.text, args, start: t.start, end: this.tokens[this.index - 1]!.end };
      if(t.text==='LAMBDA'&&this.consume('(')){
        const invocationArgs=this.callArguments();
        return {type:'call',name:'__INVOKE_LAMBDA__',args:[callee,...invocationArgs],start:t.start,end:this.tokens[this.index-1]!.end};
      }
      return callee;
    }
    if (t.kind === 'eof') throw new FormulaSyntaxError('Incomplete expression.', t.start, 'Add a value, reference, or function.');
    throw new FormulaSyntaxError(`Unexpected '${t.text}'.`, t.start, 'Enter a number, cell reference, or function here.');
  }
}
export const parseFormula = (source: string): Node => {const normalized=normalizeFormulaLocale(source);return new Parser(tokenize(normalized.startsWith('=') ? normalized.slice(1) : normalized), source).parse();};
export function references(ast: Node, rows: number, cols: number): Set<string> {
  const refs = new Set<string>();
  const visit = (n: Node): void => {
    if (n.type === 'ref'||n.type==='spillRef') { const p = parseRef(n.ref); if (!n.sheet && p && p.row < rows && p.col < cols) refs.add(normalizeRef(n.ref)); return; }
    if (n.type === 'range') {
      if (n.sheet) return;
      const a = parseRef(n.from), b = parseRef(n.to); if (!a || !b) return;
      const lowR = Math.min(a.row, b.row), highR = Math.max(a.row, b.row);
      const lowC = Math.min(a.col, b.col), highC = Math.max(a.col, b.col);
      if ((highR - lowR + 1) * (highC - lowC + 1) > 20000) throw new FormulaSyntaxError('Range exceeds 20,000 cells.', n.start, 'Use a smaller range.');
      for (let r = lowR; r <= Math.min(highR, rows - 1); r++) for (let c = lowC; c <= Math.min(highC, cols - 1); c++) refs.add(`${colToName(c)}${r+1}`);
      return;
    }
    if (n.type === 'unary') visit(n.expr);
    else if (n.type === 'binary') { visit(n.left); visit(n.right); }
    else if (n.type === 'call') n.args.forEach(visit);
  };
  visit(ast); return refs;
}
export function externalReferences(ast:Node,rows:number,cols:number):{sheet:string;ref:string}[]{
  const refs:{sheet:string;ref:string}[]=[];
  const visit=(node:Node)=>{
    if(node.type==='ref'||node.type==='spillRef'){const p=parseRef(node.ref);if(node.sheet&&p&&p.row<rows&&p.col<cols)refs.push({sheet:node.sheet,ref:normalizeRef(node.ref)});return;}
    if(node.type==='range'){if(!node.sheet)return;const a=parseRef(node.from),b=parseRef(node.to);if(!a||!b)return;const lowR=Math.min(a.row,b.row),highR=Math.min(Math.max(a.row,b.row),rows-1),lowC=Math.min(a.col,b.col),highC=Math.min(Math.max(a.col,b.col),cols-1);if((highR-lowR+1)*(highC-lowC+1)>20000)throw new FormulaSyntaxError('Range exceeds 20,000 cells.',node.start,'Use a smaller range.');for(let r=lowR;r<=highR;r++)for(let c=lowC;c<=highC;c++)refs.push({sheet:node.sheet,ref:`${colToName(c)}${r+1}`});return;}
    if(node.type==='unary')visit(node.expr);else if(node.type==='binary'){visit(node.left);visit(node.right);}else if(node.type==='call')node.args.forEach(visit);
  };visit(ast);return refs;
}
export function shiftFormula(formula: string, deltaRows: number, deltaCols: number): string {
  if (!formula.startsWith('=')) return formula;
  const body = formula.slice(1); const parts = tokenize(body);
  let cursor = 0, result = '=';
  for (const t of parts) {
    if (t.kind === 'eof') break;
    result += body.slice(cursor, t.start); cursor = t.end;
    if (t.kind !== 'ref' && t.kind !== 'sheetref' && t.kind !== 'spillref') { result += t.text; continue; }
    if(t.kind==='sheetref'||t.kind==='spillref'){
      const bang=t.text.indexOf('!'),sheet=t.text.slice(0,bang),range=t.text.slice(bang+1),parts=range.split(':');
      const shifted=parts.map(part=>{const p=parseRef(part)!;const r=p.row+(p.rowLock?0:deltaRows),c=p.col+(p.colLock?0:deltaCols);return r<0||c<0?'#REF!':`${p.colLock?'$':''}${colToName(c)}${p.rowLock?'$':''}${r+1}`;});
      const spill=t.kind==='spillref',cleanSheet=bang<0?'':sheet,prefix=/\s/.test(cleanSheet)||cleanSheet.includes("'")?`'${cleanSheet.replace(/'/g,"''")}'`:cleanSheet;
      result+=`${bang<0?'':`${prefix}!`}${shifted.join(':')}${spill?'#':''}`;continue;
    }
    const p = parseRef(t.text)!; const r = p.row + (p.rowLock ? 0 : deltaRows), c = p.col + (p.colLock ? 0 : deltaCols);
    result += r < 0 || c < 0 ? '#REF!' : `${p.colLock ? '$' : ''}${colToName(c)}${p.rowLock ? '$' : ''}${r+1}`;
  }
  return result + body.slice(cursor);
}

/** Rewrite direct references to a worksheet during rename/delete, leaving string literals untouched. */
export function rewriteSheetReferences(formula:string,oldName:string,newName:string|null):string {
  if(!formula.startsWith('='))return formula;
  const body=formula.slice(1),tokens=tokenize(body),replacements:{start:number;end:number;text:string}[]=[];
  for(const token of tokens){if(token.kind!=='sheetref'&&token.kind!=='spillref')continue;const bang=token.text.indexOf('!'),sheet=token.text.slice(0,bang);if(sheet.toLowerCase()!==oldName.toLowerCase())continue;
    const ref=token.text.slice(bang+1),prefix=newName===null?'':`'${newName.replace(/'/g,"''")}'!`,text=newName===null?'#REF!':prefix+ref+(token.kind==='spillref'?'#':'');
    replacements.push({start:token.start,end:token.end,text});
  }
  let cursor=0,out='=';for(const item of replacements){out+=body.slice(cursor,item.start)+item.text;cursor=item.end;}return out+body.slice(cursor);
}
