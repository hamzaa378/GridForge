import { colToName, normalizeRef, parseFormula, parseRef, references, type Diagnostic, type Node, type Value, FormulaSyntaxError } from './parser.js';
import { bool, error, functions, number, SheetError, text, type Datum } from './functions.js';
import {cleanStyle, type CellStyle, type StyleEdit} from './format.js';
export type CellResult = { value: Value; error?: Diagnostic };
export type Edit = { ref: string; raw: string };
export interface RecalcStats { updated: number; recalculated: number; affected: string[]; durationMs: number; totalPopulated: number }
export interface Inspect { ref: string; raw: string; result: CellResult; dependencies: string[]; dependents: string[]; expressionTree: string; referencedValues: { ref: string; raw: string; display: string }[] }
const displayNumber = (n: number): string => Number.isFinite(n) ? String(Math.round((n + Number.EPSILON) * 1e10) / 1e10) : '#NUM!';
export function display(result: CellResult): string {
  if (result.error) return result.error.code;
  const x = result.value; return x === null ? '' : typeof x === 'boolean' ? (x ? 'TRUE' : 'FALSE') : typeof x === 'number' ? displayNumber(x) : x;
}
function tree(n: Node, indent = ''): string {
  if (n.type === 'literal') return `${indent}Value ${JSON.stringify(n.value)}`;
  if (n.type === 'ref') return `${indent}Reference ${n.ref}`;
  if (n.type === 'range') return `${indent}Range ${n.from}:${n.to}`;
  if (n.type === 'unary') return `${indent}Unary ${n.op}\n${tree(n.expr, indent + '  ')}`;
  if (n.type === 'binary') return `${indent}Operator ${n.op}\n${tree(n.left, indent + '  ')}\n${tree(n.right, indent + '  ')}`;
  return `${indent}Function ${n.name}\n${n.args.map(arg => tree(arg, indent + '  ')).join('\n')}`;
}
const known = (ref: string, rows: number, cols: number): string => {
  const r = parseRef(ref); if (!r || r.row >= rows || r.col >= cols) return error('#REF!', `Cell '${ref}' is outside this workbook (${rows} rows × ${cols} columns).`);
  return normalizeRef(ref);
};
const simple = (raw: string): Value => {
  if (raw.startsWith("'")) return raw.slice(1);
  if (/^(true|false)$/i.test(raw)) return raw.toLowerCase() === 'true';
  const trimmed = raw.trim();
  if (trimmed !== '' && /^[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?$/.test(trimmed)) return Number(trimmed);
  return raw === '' ? null : raw;
};
const diag = (code: string, message: string, position = 0, suggestion = 'Check the formula and its referenced cells.'): Diagnostic => ({code, message, position, length: 1, suggestion});
export class Workbook {
  readonly rows: number; readonly cols: number;
  private raw = new Map<string,string>();
  private styles = new Map<string,CellStyle>();
  private ast = new Map<string,Node>();
  private parseErrors = new Map<string,Diagnostic>();
  private deps = new Map<string,Set<string>>();
  private rev = new Map<string,Set<string>>();
  private dynamicDeps = new Map<string,Set<string>>();
  private cache = new Map<string,CellResult>();
  private dirty = new Set<string>();
  private undoStack: { before:Edit[]; after:Edit[]; styleBefore:StyleEdit[]; styleAfter:StyleEdit[] }[] = [];
  private redoStack: { before:Edit[]; after:Edit[]; styleBefore:StyleEdit[]; styleAfter:StyleEdit[] }[] = [];
  lastStats: RecalcStats = {updated:0,recalculated:0,affected:[],durationMs:0,totalPopulated:0};
  constructor(rows = 400, cols = 52) {this.rows=rows;this.cols=cols;}
  getRaw(ref: string): string {return this.raw.get(known(ref,this.rows,this.cols)) ?? '';}
  allEntries(): Edit[] {return [...this.raw].map(([ref,raw])=>({ref,raw}));}
  allStyles():StyleEdit[] {return [...this.styles].map(([ref,style])=>({ref,style:{...style}}));}
  getStyle(ref:string):CellStyle{return {...(this.styles.get(known(ref,this.rows,this.cols))??{})};}
  /** Apply complete style snapshots. No calculation is invalidated by formatting changes. */
  applyStyles(edits:StyleEdit[],recordHistory=true):number {
    const changes=new Map<string,CellStyle>();
    for(const item of edits)changes.set(known(item.ref,this.rows,this.cols),cleanStyle(item.style));
    const changed=[...changes].filter(([ref,style])=>JSON.stringify(this.getStyle(ref))!==JSON.stringify(style));
    if(!changed.length)return 0;
    const before=changed.map(([ref])=>({ref,style:this.getStyle(ref)}));
    for(const [ref,style]of changed){if(Object.keys(style).length)this.styles.set(ref,style);else this.styles.delete(ref);}
    if(recordHistory){this.undoStack.push({before:[],after:[],styleBefore:before,styleAfter:changed.map(([ref,style])=>({ref,style}))});if(this.undoStack.length>150)this.undoStack.shift();this.redoStack=[];}
    return changed.length;
  }
  patchStyles(refs:string[],patch:Partial<CellStyle>):number {
    // Validate all incoming fields, but retain explicit false/default values to clear an attribute.
    cleanStyle(patch);
    return this.applyStyles(refs.map(ref=>{
      const result:Record<string,unknown>={...this.getStyle(ref)};
      for(const [key,value] of Object.entries(patch)){
        if(value===false||(key==='numberFormat'&&value==='general')||(key==='align'&&value==='left'))delete result[key];
        else result[key]=value;
      }
      return {ref,style:result as CellStyle};
    }));
  }
  clearStyles(refs:string[]):number{return this.applyStyles(refs.map(ref=>({ref,style:{}})));}
  private update(ref: string, raw: string): void {
    for (const oldRef of this.deps.get(ref) ?? []) {
      this.rev.get(oldRef)?.delete(ref);
      if (this.rev.get(oldRef)?.size === 0) this.rev.delete(oldRef);
    }
    this.deps.delete(ref);this.clearDynamic(ref);this.ast.delete(ref);this.parseErrors.delete(ref);
    if (raw) this.raw.set(ref,raw); else this.raw.delete(ref);
    if (raw.startsWith('=')) {
      try {
        const parsed=parseFormula(raw);
        const refs=references(parsed,this.rows,this.cols);
        this.ast.set(ref,parsed);this.deps.set(ref,refs);
        for(const dep of refs) {if(!this.rev.has(dep))this.rev.set(dep,new Set());this.rev.get(dep)!.add(ref);}
      } catch(e) {
        const d=e instanceof FormulaSyntaxError?e.diagnostic:diag('#PARSE!',e instanceof Error?e.message:String(e));
        this.parseErrors.set(ref,d);
      }
    }
  }
  private clearDynamic(owner:string):void {
    for(const dep of this.dynamicDeps.get(owner)??[]) {
      this.rev.get(dep)?.delete(owner);
      if(this.rev.get(dep)?.size===0)this.rev.delete(dep);
    }
    this.dynamicDeps.delete(owner);
  }
  private trackDynamic(owner:string,dependency:string):void {
    const ref=known(dependency,this.rows,this.cols);
    if(!this.dynamicDeps.has(owner))this.dynamicDeps.set(owner,new Set());
    if(this.deps.get(owner)?.has(ref))return;
    this.dynamicDeps.get(owner)!.add(ref);
    if(!this.rev.has(ref))this.rev.set(ref,new Set());
    this.rev.get(ref)!.add(owner);
  }
  private referenceArg(n:Node|undefined,visiting:string[],allowText=false):string {
    if(!n)error('#ARG!','Missing reference argument.');
    if(n!.type==='ref')return known(n!.ref,this.rows,this.cols);
    if(allowText){const val=this.evaluateScalar(n!,visiting);return known(text(val).replace(/\$/g,''),this.rows,this.cols);}
    return error('#VALUE!','A cell reference is required.');
  }
  private invalidate(refs: string[]): string[] {
    const queue=[...refs], affected=new Set(queue);
    for(let i=0;i<queue.length;i++) for(const dependent of this.rev.get(queue[i]!)??[]) if(!affected.has(dependent)) {affected.add(dependent);queue.push(dependent);}
    for(const ref of affected) {this.cache.delete(ref);this.dirty.add(ref);}
    return [...affected];
  }
  applyEdits(edits: Edit[], recordHistory = true): RecalcStats {
    const sanitized=new Map<string,string>();
    for(const e of edits) sanitized.set(known(e.ref,this.rows,this.cols),e.raw);
    const changed=[...sanitized].filter(([ref,raw])=>this.getRaw(ref)!==raw);
    if(!changed.length) return {...this.lastStats,updated:0,recalculated:0,affected:[],durationMs:0};
    const before=changed.map(([ref])=>({ref,raw:this.getRaw(ref)}));
    for(const [ref,raw] of changed)this.update(ref,raw);
    const affected=this.invalidate(changed.map(([ref])=>ref));
    const t0=performance.now();let recalculated=0;
    for(const ref of affected) if(this.raw.has(ref)) {this.getValue(ref);recalculated++;}
    this.lastStats={updated:changed.length,recalculated,affected,durationMs:Math.round((performance.now()-t0)*1000)/1000,totalPopulated:this.raw.size};
    if(recordHistory) {
      this.undoStack.push({before,after:changed.map(([ref,raw])=>({ref,raw})),styleBefore:[],styleAfter:[]});
      if(this.undoStack.length>150)this.undoStack.shift();this.redoStack=[];
    }
    return this.lastStats;
  }
  /** One undoable transaction for operations that change both values and formatting. */
  applyTransaction(edits:Edit[],styleEdits:StyleEdit[]):RecalcStats {
    const oldValues=new Map<string,string>();const oldStyles=new Map<string,CellStyle>();
    for(const item of edits)oldValues.set(known(item.ref,this.rows,this.cols),this.getRaw(item.ref));
    for(const item of styleEdits)oldStyles.set(known(item.ref,this.rows,this.cols),this.getStyle(item.ref));
    // Validate styles before making any changes.
    for(const item of styleEdits)cleanStyle(item.style);
    const result=this.applyEdits(edits,false);this.applyStyles(styleEdits,false);
    const before=[...oldValues].map(([ref,raw])=>({ref,raw}));
    const after=[...oldValues.keys()].map(ref=>({ref,raw:this.getRaw(ref)}));
    const styleBefore=[...oldStyles].map(([ref,style])=>({ref,style}));
    const styleAfter=[...oldStyles.keys()].map(ref=>({ref,style:this.getStyle(ref)}));
    if(before.some((b,i)=>b.raw!==after[i]?.raw)||styleBefore.some((b,i)=>JSON.stringify(b.style)!==JSON.stringify(styleAfter[i]?.style))){
      this.undoStack.push({before,after,styleBefore,styleAfter});if(this.undoStack.length>150)this.undoStack.shift();this.redoStack=[];
    }
    return result;
  }
  setCell(ref:string,raw:string):RecalcStats{return this.applyEdits([{ref,raw}]);}
  undo():boolean {const op=this.undoStack.pop();if(!op)return false;this.applyEdits(op.before,false);this.applyStyles(op.styleBefore,false);this.redoStack.push(op);return true;}
  redo():boolean {const op=this.redoStack.pop();if(!op)return false;this.applyEdits(op.after,false);this.applyStyles(op.styleAfter,false);this.undoStack.push(op);return true;}
  get canUndo(): boolean {return this.undoStack.length>0;}
  get canRedo(): boolean {return this.redoStack.length>0;}
  getValue(ref:string, visiting:string[]=[]):CellResult {
    try { ref=known(ref,this.rows,this.cols); }catch(e){return {value:null,error:diag('#REF!',(e as Error).message)};}
    const old=this.cache.get(ref);if(old && !this.dirty.has(ref))return old;
    if(visiting.includes(ref)) {
      return {value:null,error:diag('#CYCLE!',`Circular reference: ${[...visiting.slice(visiting.indexOf(ref)),ref].join(' → ')}`,0,'Remove a self-reference or break the loop between these cells.')};
    }
    const parseError=this.parseErrors.get(ref);
    if(parseError){const output={value:null,error:parseError};this.cache.set(ref,output);this.dirty.delete(ref);return output;}
    const ast=this.ast.get(ref);let output:CellResult;
    if(!ast)output={value:simple(this.raw.get(ref)??'')};
    else {
      this.clearDynamic(ref);
      try {
        const value=this.evaluate(ast,[...visiting,ref]);
        if(Array.isArray(value))error('#VALUE!','A range cannot be the final value of a single cell. Use SUM, INDEX or another function.');
        output={value:value as Value};
      }catch(e) {
        const x=e instanceof SheetError?e:new SheetError('#ERROR!',e instanceof Error?e.message:String(e));
        output={value:null,error:diag(x.code,x.message,0, x.code==='#NAME?'?'Check the spelling of the function name.':'Check the formula and its referenced cells.')};
      }
    }
    this.cache.set(ref,output);this.dirty.delete(ref);return output;
  }
  private evaluate(n:Node,visiting:string[]):Datum {
    if(n.type==='literal')return n.value;
    if(n.type==='ref') {
      const key=known(n.ref,this.rows,this.cols);const v=this.getValue(key,visiting);
      if(v.error)error(v.error.code,`${key}: ${v.error.message}`);return v.value;
    }
    if(n.type==='range') {
      const a=parseRef(n.from)!,b=parseRef(n.to)!;
      if(Math.max(a.row,b.row)>=this.rows || Math.max(a.col,b.col)>=this.cols)error('#REF!','Range extends beyond sheet boundaries.');
      const matrix:Datum[][]=[];
      for(let row=Math.min(a.row,b.row);row<=Math.max(a.row,b.row);row++) {
        const line:Datum[]=[];
        for(let col=Math.min(a.col,b.col);col<=Math.max(a.col,b.col);col++) {
          const key=`${colToName(col)}${row+1}`;const val=this.getValue(key,visiting);
          if(val.error)error(val.error.code,`${key}: ${val.error.message}`);line.push(val.value);
        }
        matrix.push(line);
      }
      return matrix;
    }
    if(n.type==='unary') {const v=this.evaluateScalar(n.expr,visiting);return n.op==='-'?-number(v):number(v);}
    if(n.type==='binary') {
      const l=this.evaluateScalar(n.left,visiting),r=this.evaluateScalar(n.right,visiting);
      if(n.op==='&')return text(l)+text(r);
      if(['=','<>','!=','<','>','<=','>='].includes(n.op)) {
        const comparable = typeof l==='number' && typeof r==='number'
          ? (l===r?0:l<r?-1:1) : text(l).toLowerCase().localeCompare(text(r).toLowerCase());
        switch(n.op){case '=':return comparable===0;case '<>':case '!=':return comparable!==0;case '<':return comparable<0;case '>':return comparable>0;case '<=':return comparable<=0;default:return comparable>=0;}
      }
      const a=number(l),b=number(r);let v:number;
      switch(n.op){case '+':v=a+b;break;case '-':v=a-b;break;case '*':v=a*b;break;case '/':if(b===0)error('#DIV/0!','Division by zero.');v=a/b;break;case '^':v=a**b;break;default:return error('#PARSE!',`Unsupported operator ${n.op}.`);}
      if(!Number.isFinite(v))error('#NUM!','Arithmetic result is not finite.');return v;
    }
    if(n.type==='call') {
      const f=functions[n.name];
      if(!f) {
        const candidates=Object.keys(functions).filter(x=>x.startsWith(n.name.slice(0,2))).slice(0,4);
        return error('#NAME?',`Unknown function ${n.name}.${candidates.length?` Did you mean ${candidates.join(' or ')}?`:''}`);
      }
      if(n.name==='IFNA') {
        if(n.args.length!==2)error('#ARG!','IFNA takes a value and a fallback.');
        try{return this.evaluate(n.args[0]!,visiting);}catch(e){if(e instanceof SheetError&&e.code==='#N/A')return this.evaluate(n.args[1]!,visiting);throw e;}
      }
      if(n.name==='ERROR.TYPE') {
        if(n.args.length!==1)error('#ARG!','ERROR.TYPE needs one argument.');
        try{this.evaluate(n.args[0]!,visiting);}catch(e){if(e instanceof SheetError){const errors:Record<string,number>={'#NULL!':1,'#DIV/0!':2,'#VALUE!':3,'#REF!':4,'#NAME?':5,'#NUM!':6,'#N/A':7,'#GETTING_DATA':8};return errors[e.code]??error('#N/A','Unrecognized error code.');}throw e;}
        return error('#N/A','The argument is not an error.');
      }
      if(n.name==='ISREF') {if(n.args.length!==1)error('#ARG!','ISREF needs one argument.');return n.args[0]!.type==='ref'||n.args[0]!.type==='range';}
      if(n.name==='ROW'||n.name==='COLUMN'){
        if(n.args.length>1)error('#ARG!',`${n.name} takes zero or one reference.`);
        const ref=n.args.length?this.referenceArg(n.args[0],visiting):visiting[visiting.length-1]!;
        const p=parseRef(ref)!;return n.name==='ROW'?p.row+1:p.col+1;
      }
      if(n.name==='ISFORMULA'||n.name==='FORMULATEXT'){
        if(n.args.length!==1)error('#ARG!',`${n.name} takes one reference.`);
        const ref=this.referenceArg(n.args[0],visiting),raw=this.getRaw(ref);
        if(n.name==='ISFORMULA')return raw.startsWith('=');
        if(!raw.startsWith('='))return error('#N/A','Referenced cell does not contain a formula.');
        return raw;
      }
      if(n.name==='CELL'){
        if(n.args.length!==2)error('#ARG!','CELL(info_type, reference) requires 2 arguments.');
        const kind=text(this.evaluateScalar(n.args[0]!,visiting)).toLowerCase(),ref=this.referenceArg(n.args[1],visiting),p=parseRef(ref)!;
        if(kind==='address')return `$${colToName(p.col)}$${p.row+1}`;
        if(kind==='row')return p.row+1;
        if(kind==='col')return p.col+1;
        if(kind==='contents')return this.getValue(ref,visiting).value;
        if(kind==='type'){const v=this.getValue(ref,visiting);return v.error?'e':v.value===null?'b':typeof v.value==='string'?'l':'v';}
        return error('#VALUE!',`CELL info type '${kind}' not supported.`);
      }
      if(n.name==='INDIRECT'||n.name==='OFFSET') {
        const owner=visiting[visiting.length-1]!;
        let ref:string;
        if(n.name==='INDIRECT') {
          if(n.args.length<1||n.args.length>2)error('#ARG!','INDIRECT(address, [a1]) takes 1–2 arguments.');
          if(n.args.length===2&&!bool(this.evaluateScalar(n.args[1]!,visiting)))error('#VALUE!','R1C1 references are not supported.');
          ref=known(text(this.evaluateScalar(n.args[0]!,visiting)).replace(/\$/g,''),this.rows,this.cols);
        }else{
          if(n.args.length<3||n.args.length>5)error('#ARG!','OFFSET(reference, rows, cols, [height], [width]) requires 3–5 arguments.');
          const base=parseRef(this.referenceArg(n.args[0],visiting))!;
          if((n.args.length>=4&&number(this.evaluateScalar(n.args[3]!,visiting))!==1)||(n.args.length===5&&number(this.evaluateScalar(n.args[4]!,visiting))!==1))error('#VALUE!','GridForge supports scalar 1×1 OFFSET results.');
          const r=base.row+Math.trunc(number(this.evaluateScalar(n.args[1]!,visiting))),c=base.col+Math.trunc(number(this.evaluateScalar(n.args[2]!,visiting)));
          ref=known(`${colToName(c)}${r+1}`,this.rows,this.cols);
        }
        if(ref===owner||visiting.includes(ref))error('#CYCLE!',`Circular dynamic reference through ${ref}.`);
        this.trackDynamic(owner,ref);
        const res=this.getValue(ref,visiting);if(res.error)error(res.error.code,res.error.message);
        return res.value;
      }
      if(n.name==='IF') {
        if(n.args.length<2||n.args.length>3)error('#ARG!','IF(condition, true_value, false_value) takes 2–3 arguments.');
        const condition=bool(this.evaluateScalar(n.args[0]!,visiting));
        return condition?this.evaluate(n.args[1]!,visiting):n.args[2]?this.evaluate(n.args[2],visiting):false;
      }
      if(n.name==='IFERROR') {
        if(n.args.length!==2)error('#ARG!','IFERROR(value, fallback) takes exactly 2 arguments.');
        try{return this.evaluate(n.args[0]!,visiting);}catch{return this.evaluate(n.args[1]!,visiting);}
      }
      if(n.name==='XLOOKUP' && n.args.length===4) {
        const firstThree=n.args.slice(0,3).map(arg=>this.evaluate(arg,visiting));
        try{return f.run(firstThree);}catch(e){
          if(e instanceof SheetError && e.code==='#N/A')return this.evaluate(n.args[3]!,visiting);
          throw e;
        }
      }
      if(['ISERR','ISERROR','ISNA'].includes(n.name)) {
        if(n.args.length!==1)error('#ARG!',`${n.name} takes one argument.`);
        try {this.evaluate(n.args[0]!,visiting);return false;}
        catch(e){const code=e instanceof SheetError?e.code:e instanceof FormulaSyntaxError?'#PARSE!':'#ERROR!';return n.name==='ISERROR'||(n.name==='ISERR'&&code!=='#N/A')||(n.name==='ISNA'&&code==='#N/A');}
      }
      if(n.name==='CHOOSE') {
        if(n.args.length<2)error('#ARG!','CHOOSE requires index and at least one value.');
        const idx=Math.trunc(number(this.evaluateScalar(n.args[0]!,visiting)));
        if(idx<1||idx>=n.args.length)error('#VALUE!','CHOOSE index is out of range.');
        return this.evaluate(n.args[idx]!,visiting);
      }
      if(n.name==='IFS') {
        if(n.args.length<2 || n.args.length%2!==0)error('#ARG!','IFS needs condition, value pairs.');
        for(let i=0;i<n.args.length;i+=2)if(bool(this.evaluateScalar(n.args[i]!,visiting)))return this.evaluate(n.args[i+1]!,visiting);
        return error('#N/A','IFS did not match any condition.');
      }
      if(n.name==='SWITCH') {
        if(n.args.length<3)error('#ARG!','SWITCH(expression, case, result, [default]) needs at least 3 arguments.');
        const target=this.evaluateScalar(n.args[0]!,visiting),hasDefault=n.args.length%2===0;
        const finish=hasDefault?n.args.length-1:n.args.length;
        for(let i=1;i<finish;i+=2) {
          const candidate=this.evaluateScalar(n.args[i]!,visiting);
          if(typeof target==='number' && typeof candidate==='number'?target===candidate:text(target).toLowerCase()===text(candidate).toLowerCase())return this.evaluate(n.args[i+1]!,visiting);
        }
        return hasDefault?this.evaluate(n.args[n.args.length-1]!,visiting):error('#N/A','SWITCH has no matching case.');
      }
      // Excel treats numeric text and booleans typed directly into aggregation
      // calls differently from the same values held in referenced cells.
      // Preserve range/reference semantics while coercing direct literals.
      if(['SUM','COUNT','AVERAGE','PRODUCT','MIN','MAX'].includes(n.name)) {
        const evaluated=n.args.map(arg=>{
          const value=this.evaluate(arg,visiting);
          if(arg.type!=='literal')return value;
          if(typeof value==='boolean')return value?1:0;
          if(typeof value==='string' && value.trim()!=='') {
            const parsed=Number(value);
            if(Number.isFinite(parsed))return parsed;
            if(n.name!=='COUNT')error('#VALUE!',`Cannot convert direct text argument '${value}' to a number.`);
          }
          return value;
        });
        return f.run(evaluated);
      }
      return f.run(n.args.map(arg=>this.evaluate(arg,visiting)));
    }
    return error('#ERROR!','Unexpected AST node.');
  }
  private evaluateScalar(n:Node,visiting:string[]):Value {
    const v=this.evaluate(n,visiting);return Array.isArray(v)?error('#VALUE!','Cannot perform scalar arithmetic on a whole range.'):v;
  }
  inspect(ref:string):Inspect {
    ref=known(ref,this.rows,this.cols);const ds=[...(this.deps.get(ref)??[])];const rvs=[...(this.rev.get(ref)??[])];
    return {ref,raw:this.getRaw(ref),result:this.getValue(ref),dependencies:ds,dependents:rvs,
      expressionTree:this.ast.has(ref)?tree(this.ast.get(ref)!):this.parseErrors.has(ref)?'Unable to parse formula':'Literal / empty cell',
      referencedValues:ds.slice(0,100).map(r=>({ref:r,raw:this.getRaw(r),display:display(this.getValue(r))}))};
  }
  preview(ref:string,proposed:string):{changes:{ref:string;before:string;after:string}[];stats:RecalcStats} {
    const target=known(ref,this.rows,this.cols);const impacted=this.transitiveDependents(target);
    const before=new Map(impacted.map(id=>[id,display(this.getValue(id))]));
    const cloned=Workbook.fromJSON(this.toJSON());const stats=cloned.applyEdits([{ref:target,raw:proposed}],false);
    return {changes:impacted.map(id=>({ref:id,before:before.get(id)??'',after:display(cloned.getValue(id))})).filter(c=>c.before!==c.after),stats};
  }
  transitiveDependents(ref:string):string[] {
    const queue=[known(ref,this.rows,this.cols)], seen=new Set(queue);
    for(let i=0;i<queue.length;i++)for(const k of this.rev.get(queue[i]!)??[])if(!seen.has(k)){seen.add(k);queue.push(k);}
    return queue;
  }
  toJSON():{version:2;rows:number;cols:number;cells:Record<string,string>;styles:Record<string,CellStyle>} {return {version:2,rows:this.rows,cols:this.cols,cells:Object.fromEntries(this.raw),styles:Object.fromEntries(this.styles)};}
  static fromJSON(data:unknown):Workbook {
    if(!data||typeof data!=='object')throw Error('Invalid workbook data.');
    const d=data as {rows?:number;cols?:number;cells?:Record<string,unknown>;styles?:Record<string,unknown>};
    if(!Number.isInteger(d.rows)||!Number.isInteger(d.cols)||(d.rows??0)<1||(d.rows??0)>1000||(d.cols??0)<1||(d.cols??0)>104)throw Error('Workbook dimensions must be 1–1000 rows and 1–104 columns.');
    if(!d.cells||typeof d.cells!=='object'||Array.isArray(d.cells))throw Error('Workbook cells must be an object.');
    const book=new Workbook(d.rows!,d.cols!);
    const edits=Object.entries(d.cells).map(([ref,raw])=>{if(typeof raw!=='string')throw Error(`Invalid content in ${ref}.`);return {ref,raw};});
    if(edits.length>104000)throw Error('Workbook contains too many populated cells.');
    book.applyEdits(edits,false);
    if(d.styles!==undefined){if(!d.styles||typeof d.styles!=='object'||Array.isArray(d.styles))throw Error('Invalid workbook formatting.');
      if(Object.keys(d.styles).length>104000)throw Error('Too many styled cells.');
      book.applyStyles(Object.entries(d.styles).map(([ref,style])=>({ref,style:cleanStyle(style)})),false);
    }
    return book;
  }
}
