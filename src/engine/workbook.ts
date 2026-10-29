import { colToName, normalizeRef, parseFormula, parseRef, references, externalReferences, type Diagnostic, type Node, type Value, FormulaSyntaxError } from './parser.js';
import { bool, error, functions, number, SheetError, text, type Datum } from './functions.js';
import {cleanStyle, isExcelDateNumberFormat, type CellStyle, type StyleEdit} from './format.js';
import {resolveStructuredFormula, type TableSpec} from './tables.js';
export type CellResult = { value: Value; error?: Diagnostic };
export type Edit = { ref: string; raw: string };
export interface RecalcStats { updated: number; recalculated: number; affected: string[]; durationMs: number; totalPopulated: number }
export interface Inspect { ref: string; raw: string; result: CellResult; dependencies: string[]; dependents: string[]; expressionTree: string; referencedValues: { ref: string; raw: string; display: string }[] }
export interface DefinedName {name:string;ref:string;comment?:string}
export interface PivotDataCriterion {field:Value;item:Value}
export type PivotRangeTracker=(sheet:string,range:string,cycleRange?:string)=>void;
export type PivotDataResolver=(pivotSheet:string,pivotRef:string,dataField:Value,criteria:PivotDataCriterion[],trackRange:PivotRangeTracker)=>Value;
type DynamicRangeDependency={sheet:string;from:string;to:string};
type HistoryEntry={before:Edit[];after:Edit[];styleBefore:StyleEdit[];styleAfter:StyleEdit[];namesBefore?:DefinedName[];namesAfter?:DefinedName[]};
const displayNumber = (n: number): string => Number.isFinite(n) ? String(Math.round((n + Number.EPSILON) * 1e10) / 1e10) : '#NUM!';
export function display(result: CellResult): string {
  if (result.error) return result.error.code;
  const x = result.value; return x === null ? '' : typeof x === 'boolean' ? (x ? 'TRUE' : 'FALSE') : typeof x === 'number' ? displayNumber(x) : x;
}
function tree(n: Node, indent = ''): string {
  if (n.type === 'literal') return `${indent}Value ${JSON.stringify(n.value)}`;
  if (n.type === 'error') return `${indent}Error ${n.code}`;
  if (n.type === 'ref') return `${indent}Reference ${n.sheet?`${n.sheet}!`:''}${n.ref}`;
  if (n.type === 'spillRef') return `${indent}Spill reference ${n.sheet?`${n.sheet}!`:''}${n.ref}#`;
  if (n.type === 'range') return `${indent}Range ${n.sheet?`${n.sheet}!`:''}${n.from}:${n.to}`;
  if (n.type === 'definedName') return `${indent}Defined name ${n.name}`;
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
  private valueRevisionCounter=0;
  get valueRevision():number{return this.valueRevisionCounter;}
  private sheetName='';
  private documentSheetNames:string[]=[];
  private hiddenRowsForSheet:(sheet:string,includeManuallyHidden:boolean)=>readonly number[]=()=>[];
  private nestedAggregateForSheet:(sheet:string,ref:string)=>boolean=()=>false;
  private externalSheetResolver:((sheet:string,ref:string,visiting:string[])=>CellResult)|null=null;
  private externalRawResolver:((sheet:string,ref:string)=>string|undefined)|null=null;
  private externalSpillResolver:((sheet:string,ref:string,visiting:string[])=>Datum[][])|null=null;
  private pivotDataResolver:PivotDataResolver|null=null;
  private tableContext: readonly TableSpec[] = [];
  private definedNames=new Map<string,DefinedName>();
  private localFormulaNames:{owner:string;values:Map<string,Datum>}[]=[];
  private raw = new Map<string,string>();
  private styles = new Map<string,CellStyle>();
  private ast = new Map<string,Node>();
  private parseErrors = new Map<string,Diagnostic>();
  private deps = new Map<string,Set<string>>();
  private rev = new Map<string,Set<string>>();
  private dynamicDeps = new Map<string,Set<string>>();
  private dynamicRanges=new Map<string,DynamicRangeDependency[]>();
  private externalDeps=new Map<string,{sheet:string;ref:string}[]>();
  private onCellsChanged:((refs:string[])=>void)|null=null;
  private cache = new Map<string,CellResult>();
  private spillAreas=new Map<string,string[]>();
  private spillChildren=new Map<string,{owner:string;value:Value}>();
  private dirty = new Set<string>();
  private undoStack: HistoryEntry[] = [];
  private redoStack: HistoryEntry[] = [];
  lastStats: RecalcStats = {updated:0,recalculated:0,affected:[],durationMs:0,totalPopulated:0};
  constructor(rows = 400, cols = 52) {this.rows=rows;this.cols=cols;}
  /** Link this worksheet to its containing document so formulas can resolve sheet-qualified references. */
  setSheetContext(name:string,resolver:(sheet:string,ref:string,visiting:string[])=>CellResult,onCellsChanged?:(refs:string[])=>void,spillResolver?:(sheet:string,ref:string,visiting:string[])=>Datum[][],documentSheetNames:string[]=[],hiddenRowsForSheet?:(sheet:string,includeManuallyHidden:boolean)=>readonly number[],nestedAggregateForSheet?:(sheet:string,ref:string)=>boolean,rawResolver?:(sheet:string,ref:string)=>string|undefined):void {
    const namesChanged=this.sheetName!==name||this.documentSheetNames.join('\u0000')!==documentSheetNames.join('\u0000');this.sheetName=name;this.externalSheetResolver=resolver;this.externalRawResolver=rawResolver??null;this.onCellsChanged=onCellsChanged??null;this.externalSpillResolver=spillResolver??null;this.documentSheetNames=[...documentSheetNames];this.hiddenRowsForSheet=hiddenRowsForSheet??(()=>[]);this.nestedAggregateForSheet=nestedAggregateForSheet??((sheet,ref)=>sheet.toLocaleLowerCase()===this.sheetName.toLocaleLowerCase()&&this.hasNestedAggregateFormula(ref));if(!namesChanged)return;
    const usesSheetInfo=(node:Node):boolean=>node.type==='call'?(node.name==='SHEET'||node.name==='SHEETS'||node.args.some(usesSheetInfo)):node.type==='unary'?usesSheetInfo(node.expr):node.type==='binary'?usesSheetInfo(node.left)||usesSheetInfo(node.right):false;
    const formulas=[...this.ast].filter(([,node])=>usesSheetInfo(node)).map(([ref])=>ref);if(!formulas.length)return;const affected=this.invalidate(formulas);this.valueRevisionCounter++;for(const ref of affected)if(this.raw.has(ref))this.getValue(ref);this.onCellsChanged?.(affected);
  }
  hasNestedAggregateFormula(ref:string):boolean {
    const root=this.ast.get(normalizeRef(ref));if(!root)return false;const visit=(node:Node):boolean=>node.type==='call'?(node.name==='SUBTOTAL'||node.name==='AGGREGATE'||node.args.some(visit)):node.type==='unary'?visit(node.expr):node.type==='binary'?visit(node.left)||visit(node.right):false;return visit(root);
  }
  /** Refresh SUBTOTAL/AGGREGATE formulas after row visibility changes, which are not cell edits. */
  recalculateAggregatesForRowVisibilityChange():void {
    const usesAggregate=(node:Node):boolean=>node.type==='call'?(node.name==='AGGREGATE'||node.args.some(usesAggregate)):node.type==='unary'?usesAggregate(node.expr):node.type==='binary'?usesAggregate(node.left)||usesAggregate(node.right):false;
    const formulas=[...this.ast].filter(([,node])=>usesAggregate(node)).map(([ref])=>ref);if(!formulas.length)return;const affected=this.invalidate(formulas);this.valueRevisionCounter++;for(const ref of affected)if(this.raw.has(ref))this.getValue(ref);this.onCellsChanged?.(affected);
  }
  /** Attach the containing document's PivotTable lookup after worksheet contexts are bound. */
  setPivotDataContext(resolver:PivotDataResolver|null):void {
    this.pivotDataResolver=resolver;
    const usesPivotData=(node:Node):boolean=>node.type==='call'?(node.name==='GETPIVOTDATA'||node.args.some(usesPivotData)):node.type==='unary'?usesPivotData(node.expr):node.type==='binary'?usesPivotData(node.left)||usesPivotData(node.right):false;
    const formulas=[...this.ast].filter(([,node])=>usesPivotData(node)).map(([ref])=>ref);
    if(!formulas.length)return;
    const affected=this.invalidate(formulas);this.valueRevisionCounter++;for(const ref of affected)if(this.raw.has(ref))this.getValue(ref);
    this.onCellsChanged?.(affected);
  }
  getDefinedNames():DefinedName[]{return [...this.definedNames.values()].map(item=>({...item}));}
  private cleanDefinedNames(items:unknown):Map<string,DefinedName>{
    if(!Array.isArray(items))throw Error('Defined names must be a list.');
    if(items.length>500)throw Error('A worksheet can contain at most 500 defined names.');
    const result=new Map<string,DefinedName>();
    for(const item of items){
      if(!item||typeof item!=='object')throw Error('Invalid defined name.');
      const raw=item as Record<string,unknown>,name=typeof raw.name==='string'?raw.name.trim():'';
      const ref=typeof raw.ref==='string'?raw.ref.replace(/\$/g,'').toUpperCase():'';
      if(name.length>255||!/^[_A-Za-z\\][A-Za-z0-9_.]*$/.test(name)||parseRef(name)||/^[RC]\d+$/i.test(name)||/^(TRUE|FALSE)$/i.test(name)||functions[name.toUpperCase()])throw Error(`'${name||'Name'}' is not a valid defined name.`);
      const key=name.toUpperCase();if(result.has(key))throw Error(`The name '${name}' is already defined on this worksheet.`);
      const pieces=ref.split(':');if(pieces.length>2||pieces.some(part=>!parseRef(part)))throw Error(`'${ref}' must refer to a cell or rectangular range such as A1:B5.`);
      for(const part of pieces){const p=parseRef(part)!;if(p.row>=this.rows||p.col>=this.cols)throw Error(`'${ref}' is outside this worksheet.`);}
      const a=parseRef(pieces[0]!)!,b=parseRef(pieces[1]??pieces[0]!)!;
      const first=`${colToName(Math.min(a.col,b.col))}${Math.min(a.row,b.row)+1}`,last=`${colToName(Math.max(a.col,b.col))}${Math.max(a.row,b.row)+1}`;
      const clean:DefinedName={name,ref:first===last?first:`${first}:${last}`};
      if(typeof raw.comment==='string'&&raw.comment.trim())clean.comment=raw.comment.trim().slice(0,255);
      result.set(key,clean);
    }
    return result;
  }
  setDefinedNames(items:unknown,recordHistory=true):void{
    const next=this.cleanDefinedNames(items),before=this.getDefinedNames(),after=[...next.values()];
    if(JSON.stringify(before)===JSON.stringify(after))return;
    this.definedNames=next;this.rebuildNamedFormulas();this.valueRevisionCounter++;
    if(recordHistory){this.undoStack.push({before:[],after:[],styleBefore:[],styleAfter:[],namesBefore:before,namesAfter:after});if(this.undoStack.length>150)this.undoStack.shift();this.redoStack=[];}
  }
  private rebuildNamedFormulas():void{
    this.ast.clear();this.parseErrors.clear();this.deps.clear();this.rev.clear();this.dynamicDeps.clear();this.externalDeps.clear();
    for(const {ref,raw} of this.allEntries())if(raw.startsWith('='))this.update(ref,raw);
    this.cache.clear();this.dirty=new Set(this.raw.keys());
    const t0=performance.now(),affected=[...this.raw.keys()];let recalculated=0;
    for(const ref of affected){if(this.raw.has(ref)){this.getValue(ref);recalculated++;}}
    this.lastStats={updated:0,recalculated,affected,durationMs:Math.round((performance.now()-t0)*1000)/1000,totalPopulated:this.raw.size};
  }
  /** Update same-sheet Excel table references while retaining the original formula text. */
  setTableContext(tables:readonly TableSpec[]):void {
    this.tableContext=tables.map(t=>({...t}));
    // Rebuild dependency edges because the table may have been resized or renamed.
    const formulas=this.allEntries().filter(e=>e.raw.startsWith('='));
    for(const {ref,raw} of formulas)this.update(ref,raw);
    this.cache.clear();this.dirty=new Set(this.allEntries().map(e=>e.ref));this.valueRevisionCounter++;
  }
  getRaw(ref: string): string {return this.raw.get(known(ref,this.rows,this.cols)) ?? '';}
  /** Return safe navigation metadata for a direct HYPERLINK formula, when its target is understood. */
  getHyperlinkInfo(ref:string):{target?:string;location?:string;display?:string}|null {
    const key=known(ref,this.rows,this.cols),node=this.ast.get(key);if(node?.type!=='call'||node.name!=='HYPERLINK'||node.args.length<1||node.args.length>2)return null;
    try{const destination=text(this.evaluateScalar(node.args[0]!,[key])).trim();if(!destination||destination.length>2048||/[\u0000-\u001f\\]/.test(destination))return null;let info:{target?:string;location?:string;display?:string}|null=null;
      if(destination.startsWith('#'))info={location:destination.slice(1)};else if(/^[a-z][a-z0-9+.-]*:/i.test(destination)){if(!/\s/.test(destination)){const url=new URL(destination);if((url.protocol==='https:'||url.protocol==='http:')&&!!url.hostname)info={target:destination};else if(url.protocol==='mailto:'&&/^mailto:[^\s@]+@[^\s@]+$/i.test(destination))info={target:destination};}}else if(parseRef(destination.replace(/\$/g,''))||destination.includes('!'))info={location:destination};
      if(!info)return null;const label=node.args.length===2?text(this.evaluateScalar(node.args[1]!,[key])):text(this.getValue(key).value);return {...info,...(label?{display:label}:{})};
    }catch{return null;}
  }
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
        if(value===undefined||value===false||(key==='numberFormat'&&value==='general')||(key==='align'&&value==='left'))delete result[key];
        else result[key]=value;
      }
      return {ref,style:result as CellStyle};
    }));
  }
  clearStyles(refs:string[]):number{return this.applyStyles(refs.map(ref=>({ref,style:{}})));}
  private update(ref: string, raw: string): void {
    this.clearSpill(ref);
    for (const oldRef of this.deps.get(ref) ?? []) {
      this.rev.get(oldRef)?.delete(ref);
      if (this.rev.get(oldRef)?.size === 0) this.rev.delete(oldRef);
    }
    this.deps.delete(ref);this.externalDeps.delete(ref);this.clearDynamic(ref);this.ast.delete(ref);this.parseErrors.delete(ref);
    if (raw) this.raw.set(ref,raw); else this.raw.delete(ref);
    if (raw.startsWith('=')) {
      try {
        const parsed=this.resolveDefinedNames(parseFormula(this.tableContext.length?resolveStructuredFormula(raw,this,this.tableContext,ref):raw));
        const refs=references(parsed,this.rows,this.cols);
        const crossRefs=externalReferences(parsed,this.rows,this.cols);if(crossRefs.length)this.externalDeps.set(ref,crossRefs);
        this.ast.set(ref,parsed);this.deps.set(ref,refs);
        for(const dep of refs) {if(!this.rev.has(dep))this.rev.set(dep,new Set());this.rev.get(dep)!.add(ref);}
      } catch(e) {
        const d=e instanceof FormulaSyntaxError?e.diagnostic:diag('#PARSE!',e instanceof Error?e.message:String(e));
        this.parseErrors.set(ref,d);
      }
    }
  }
  private resolveDefinedNames(node:Node,locals:ReadonlySet<string>=new Set()):Node{
    if(node.type==='definedName'){
      if(locals.has(node.name.toUpperCase()))return node;
      const def=this.definedNames.get(node.name.toUpperCase());
      if(!def)return node;
      const parts=def.ref.split(':');
      if(parts.length===1)return {type:'ref',ref:parts[0]!,start:node.start,end:node.end};
      return {type:'range',from:parts[0]!,to:parts[1]!,start:node.start,end:node.end};
    }
    if(node.type==='unary')return {...node,expr:this.resolveDefinedNames(node.expr,locals)};
    if(node.type==='binary')return {...node,left:this.resolveDefinedNames(node.left,locals),right:this.resolveDefinedNames(node.right,locals)};
    if(node.type==='call'&&node.name==='LET'){
      if(node.args.length<3||node.args.length>253||node.args.length%2!==1)throw new FormulaSyntaxError('LET needs 1 to 126 name/value pairs and a final expression.',node.start,'Use LET(name, value, calculation).',node.end-node.start);
      const scope=new Set(locals),args:Node[]=[],bound=new Set<string>();
      for(let index=0;index<node.args.length-1;index+=2){
        const name=node.args[index]!,value=node.args[index+1]!,key=name.type==='definedName'?name.name.toUpperCase():'';
        if(!key||!/^[A-Z_][A-Z0-9_.]*$/.test(key)||parseRef(key))throw new FormulaSyntaxError('LET variable names must be valid identifiers, not cell references.',name.start,'Use a name such as total or _rate.',name.end-name.start);
        if(bound.has(key))throw new FormulaSyntaxError(`LET variable '${key}' is defined more than once.`,name.start,'Give each LET binding a different name.',name.end-name.start);
        bound.add(key);args.push(name,this.resolveDefinedNames(value,scope));scope.add(key);
      }
      args.push(this.resolveDefinedNames(node.args.at(-1)!,scope));return {...node,args};
    }
    if(node.type==='call')return {...node,args:node.args.map(arg=>this.resolveDefinedNames(arg,locals))};
    return node;
  }
  private clearDynamic(owner:string):void {
    for(const dep of this.dynamicDeps.get(owner)??[]) {
      this.rev.get(dep)?.delete(owner);
      if(this.rev.get(dep)?.size===0)this.rev.delete(dep);
    }
    this.dynamicDeps.delete(owner);
    this.dynamicRanges.delete(owner);
  }
  private trackDynamic(owner:string,dependency:string):void {
    const ref=known(dependency,this.rows,this.cols);
    if(!this.dynamicDeps.has(owner))this.dynamicDeps.set(owner,new Set());
    if(this.deps.get(owner)?.has(ref))return;
    this.dynamicDeps.get(owner)!.add(ref);
    if(!this.rev.has(ref))this.rev.set(ref,new Set());
    this.rev.get(ref)!.add(owner);
  }
  private trackExternalRange(ownerIdentity:string,sheet:string,range:string,cycleRange=range):void {
    const [fromRaw,toRaw]=range.replace(/\$/g,'').toUpperCase().split(':'),from=fromRaw?parseRef(fromRaw):null,to=parseRef(toRaw??fromRaw??''),cycleParts=cycleRange.replace(/\$/g,'').toUpperCase().split(':'),cycleFrom=cycleParts[0]?parseRef(cycleParts[0]):null,cycleTo=parseRef(cycleParts[1]??cycleParts[0]??'');
    if(!from||!to||!cycleFrom||!cycleTo)return error('#REF!','PivotTable output range is invalid.');
    const owner=ownerIdentity.split('!').at(-1)!,own=parseRef(owner);
    if(!own)return error('#REF!','Formula owner reference is invalid.');
    if(sheet.toLocaleLowerCase()===this.sheetName.toLocaleLowerCase()&&own!.row>=Math.min(cycleFrom.row,cycleTo!.row)&&own!.row<=Math.max(cycleFrom.row,cycleTo!.row)&&own!.col>=Math.min(cycleFrom.col,cycleTo!.col)&&own!.col<=Math.max(cycleFrom.col,cycleTo!.col))error('#CYCLE!','GETPIVOTDATA cannot be placed inside the PivotTable output it reads.');
    const existing=this.dynamicRanges.get(owner)??[];
    if(!existing.some(item=>item.sheet.toLocaleLowerCase()===sheet.toLocaleLowerCase()&&item.from===fromRaw&&item.to===(toRaw??fromRaw)))existing.push({sheet,from:fromRaw!,to:toRaw??fromRaw!});
    this.dynamicRanges.set(owner,existing);
  }
  private trackDynamicRange(owner:string,sheet:string,range:string):void {
    const [fromRaw,toRaw]=range.replace(/\$/g,'').toUpperCase().split(':'),from=fromRaw?parseRef(fromRaw):null,to=parseRef(toRaw??fromRaw??'');
    if(!from||!to||from.row>=this.rows||to.row>=this.rows||from.col>=this.cols||to.col>=this.cols)error('#REF!','INDIRECT reference is outside the worksheet.');
    const existing=this.dynamicRanges.get(owner)??[];
    if(!existing.some(item=>item.sheet.toLocaleLowerCase()===sheet.toLocaleLowerCase()&&item.from===fromRaw&&item.to===(toRaw??fromRaw)))existing.push({sheet,from:fromRaw!,to:toRaw??fromRaw!});
    this.dynamicRanges.set(owner,existing);
  }
  private referenceArg(n:Node|undefined,visiting:string[],allowText=false):string {
    if(!n)error('#ARG!','Missing reference argument.');
    if(n!.type==='ref'||n!.type==='spillRef')return known(n!.ref,this.rows,this.cols);
    if(allowText){const val=this.evaluateScalar(n!,visiting);return known(text(val).replace(/\$/g,''),this.rows,this.cols);}
    return error('#VALUE!','A cell reference is required.');
  }
  private invalidate(refs: string[]): string[] {
    const queue=[...refs], affected=new Set(queue);
    for(let i=0;i<queue.length;i++) {
      const changed=parseRef(queue[i]!);
      for(const dependent of this.rev.get(queue[i]!)??[])if(!affected.has(dependent)){affected.add(dependent);queue.push(dependent);}
      if(changed)for(const [owner,ranges] of this.dynamicRanges)if(ranges.some(range=>range.sheet.toLocaleLowerCase()===this.sheetName.toLocaleLowerCase()&&(()=>{const from=parseRef(range.from),to=parseRef(range.to);return !!from&&!!to&&changed.row>=Math.min(from.row,to.row)&&changed.row<=Math.max(from.row,to.row)&&changed.col>=Math.min(from.col,to.col)&&changed.col<=Math.max(from.col,to.col);})())&&!affected.has(owner)){affected.add(owner);queue.push(owner);}
    }
    for(const ref of affected) {this.cache.delete(ref);this.dirty.add(ref);}
    return [...affected];
  }
  applyEdits(edits: Edit[], recordHistory = true): RecalcStats {
    const sanitized=new Map<string,string>();
    for(const e of edits) sanitized.set(known(e.ref,this.rows,this.cols),e.raw);
    const changed=[...sanitized].filter(([ref,raw])=>this.getRaw(ref)!==raw);
    if(!changed.length) return {...this.lastStats,updated:0,recalculated:0,affected:[],durationMs:0};
    this.valueRevisionCounter++;
    const blocked=changed.find(([ref])=>this.spillChildren.has(ref));if(blocked)throw Error(`Cannot edit ${blocked[0]} because it belongs to a spilled array. Clear or edit the formula in ${this.spillChildren.get(blocked[0])!.owner}.`);
    const before=changed.map(([ref])=>({ref,raw:this.getRaw(ref)})),spillAffected=changed.flatMap(([ref])=>this.clearSpill(ref));
    for(const [ref,raw] of changed)this.update(ref,raw);
    const affected=this.invalidate([...changed.map(([ref])=>ref),...spillAffected]);
    const t0=performance.now();let recalculated=0;
    for(const ref of affected) if(this.raw.has(ref)) {this.getValue(ref);recalculated++;}
    this.lastStats={updated:changed.length,recalculated,affected,durationMs:Math.round((performance.now()-t0)*1000)/1000,totalPopulated:this.raw.size};
    if(recordHistory) {
      this.undoStack.push({before,after:changed.map(([ref,raw])=>({ref,raw})),styleBefore:[],styleAfter:[]});
      if(this.undoStack.length>150)this.undoStack.shift();this.redoStack=[];
    }
    if(affected.length)this.onCellsChanged?.([...new Set([...affected,...spillAffected])]);
    return this.lastStats;
  }
  /** Invalidate formulas that reference cells changed on another worksheet, and their local dependents. */
  invalidateExternal(sheet:string,refs:string[]):string[]{
    const targets=new Set(refs.map(ref=>ref.toUpperCase())),owners:string[]=[];
    for(const [owner,deps]of this.externalDeps)if(deps.some(dep=>dep.sheet.toLowerCase()===sheet.toLowerCase()&&targets.has(dep.ref)))owners.push(owner);
    for(const [owner,ranges]of this.dynamicRanges)if(ranges.some(range=>range.sheet.toLowerCase()===sheet.toLowerCase()&&[...targets].some(ref=>{const cell=parseRef(ref),from=parseRef(range.from),to=parseRef(range.to);return !!cell&&!!from&&!!to&&cell.row>=Math.min(from.row,to.row)&&cell.row<=Math.max(from.row,to.row)&&cell.col>=Math.min(from.col,to.col)&&cell.col<=Math.max(from.col,to.col);}))&&!owners.includes(owner))owners.push(owner);
    if(!owners.length)return [];
    const affected=this.invalidate(owners);if(affected.length)this.valueRevisionCounter++;for(const ref of affected)if(this.raw.has(ref))this.getValue(ref);
    return affected;
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
  undo():boolean {const op=this.undoStack.pop();if(!op)return false;if(op.namesBefore)this.setDefinedNames(op.namesBefore,false);this.applyEdits(op.before,false);this.applyStyles(op.styleBefore,false);this.redoStack.push(op);return true;}
  redo():boolean {const op=this.redoStack.pop();if(!op)return false;if(op.namesAfter)this.setDefinedNames(op.namesAfter,false);this.applyEdits(op.after,false);this.applyStyles(op.styleAfter,false);this.undoStack.push(op);return true;}
  get canUndo(): boolean {return this.undoStack.length>0;}
  get canRedo(): boolean {return this.redoStack.length>0;}
  getValue(ref:string, visiting:string[]=[]):CellResult {
    try { ref=known(ref,this.rows,this.cols); }catch(e){return {value:null,error:diag('#REF!',(e as Error).message)};}
    const child=this.spillChildren.get(ref);
    if(child){const owner=this.getValue(child.owner,visiting);if(owner.error)return owner;const current=this.spillChildren.get(ref);if(current)return {value:current.value};return {value:null,error:diag('#REF!',`Spill cell ${ref} is no longer available.`)};}
    const identity=this.sheetName?`${this.sheetName}!${ref}`:ref;
    const old=this.cache.get(ref);if(old && !this.dirty.has(ref) && !this.hasExternalReference(this.ast.get(ref)))return old;
    if(visiting.includes(identity)||(!this.sheetName&&visiting.includes(ref))) {
      const at=visiting.indexOf(identity);
      return {value:null,error:diag('#CYCLE!',`Circular reference: ${[...visiting.slice(at<0?0:at),identity].join(' → ')}`,0,'Remove a self-reference or break the loop between these cells.')};
    }
    const parseError=this.parseErrors.get(ref);
    if(parseError){const output={value:null,error:parseError};this.cache.set(ref,output);this.dirty.delete(ref);return output;}
    const ast=this.ast.get(ref);let output:CellResult;
    if(!ast)output={value:simple(this.raw.get(ref)??'')};
    else {
      const priorSpill=this.clearSpill(ref);
      this.clearDynamic(ref);
      try {
        const value=this.evaluate(ast,[...visiting,identity]);
        if(Array.isArray(value))output=this.placeSpill(ref,value,ast);
        else output={value:value as Value};
      }catch(e) {
        const x=e instanceof SheetError?e:new SheetError('#ERROR!',e instanceof Error?e.message:String(e));
        output={value:null,error:diag(x.code,x.message,0, x.code==='#NAME?'?'Check the spelling of the function name.':'Check the formula and its referenced cells.')};
      }
      if(priorSpill.length)this.invalidate(priorSpill);
    }
    if(!this.hasExternalReference(ast))this.cache.set(ref,output);else this.cache.delete(ref);this.dirty.delete(ref);return output;
  }
  private clearSpill(owner:string):string[]{
    const refs=this.spillAreas.get(owner)??[];this.spillAreas.delete(owner);
    for(const ref of refs){if(this.spillChildren.get(ref)?.owner===owner)this.spillChildren.delete(ref);this.cache.delete(ref);this.dirty.add(ref);}
    return refs;
  }
  private placeSpill(owner:string,value:Datum[][],ast:Node):CellResult {
    const rows=value.length,cols=Math.max(0,...value.map(row=>row.length));
    if(!rows||!cols)return {value:null,error:diag('#CALC!','Array formula returned an empty result.')};
    const origin=parseRef(owner)!;
    if(origin.row+rows>this.rows||origin.col+cols>this.cols)return {value:null,error:diag('#SPILL!','The array result extends beyond the worksheet boundary.')};
    const refs:string[]=[];
    for(let r=0;r<rows;r++)for(let c=0;c<cols;c++){
      const ref=`${colToName(origin.col+c)}${origin.row+r+1}`;
      if(ref!==owner){
        if(this.raw.has(ref)||this.spillChildren.has(ref)){this.trackDynamic(owner,ref);return {value:null,error:diag('#SPILL!',`The spill range is blocked at ${ref}. Clear that cell and recalculate.`)};}
        if(this.deps.get(owner)?.has(ref)||this.dynamicDeps.get(owner)?.has(ref))return {value:null,error:diag('#CYCLE!',`Array formula in ${owner} depends on its own spill cell ${ref}.`)};
      }
      refs.push(ref);
    }
    const children:string[]=[];
    for(let r=0;r<rows;r++)for(let c=0;c<cols;c++){
      const ref=`${colToName(origin.col+c)}${origin.row+r+1}`;if(ref===owner)continue;
      const cell=value[r]![c]??null;this.spillChildren.set(ref,{owner,value:cell as Value});children.push(ref);
    }
    this.spillAreas.set(owner,children);
    if(children.length)this.invalidate(children);
    return {value:(value[0]![0]??null) as Value};
  }
  private hasExternalReference(node:Node|undefined):boolean {
    if(!node)return false;
    if((node.type==='ref'||node.type==='range'||node.type==='spillRef')&&!!node.sheet)return true;
    if(node.type==='unary')return this.hasExternalReference(node.expr);
    if(node.type==='binary')return this.hasExternalReference(node.left)||this.hasExternalReference(node.right);
    if(node.type==='call')return node.args.some(arg=>this.hasExternalReference(arg));
    return false;
  }
  getSpillValues(ref:string,visiting:string[]=[]):Datum[][]{
    const key=known(ref,this.rows,this.cols),result=this.getValue(key,visiting);if(result.error)error(result.error.code,result.error.message);
    const children=this.spillAreas.get(key);if(!children)error('#REF!',`${key} is not the anchor of a spilled array.`);
    const origin=parseRef(key)!,positions=children!.map(cell=>parseRef(cell)!),height=Math.max(1,...positions.map(p=>p.row-origin.row+1)),width=Math.max(1,...positions.map(p=>p.col-origin.col+1));
    const matrix:Datum[][]=Array.from({length:height},()=>Array.from({length:width},()=>null));matrix[0]![0]=result.value;
    for(const child of children!){const p=parseRef(child)!,value=this.spillChildren.get(child)?.value??null;matrix[p.row-origin.row]![p.col-origin.col]=value;}
    return matrix;
  }
  private evaluate(n:Node,visiting:string[]):Datum {
    if(n.type==='literal')return n.value;
    if(n.type==='error')return error(n.code,`Formula contains the error value ${n.code}.`);
    if(n.type==='definedName'){
      const key=n.name.toUpperCase(),owner=visiting.at(-1)??'';for(let index=this.localFormulaNames.length-1;index>=0;index--){const scope=this.localFormulaNames[index]!;if(scope.owner===owner&&scope.values.has(key))return scope.values.get(key)!;}
      return error('#NAME?',`Unknown name '${n.name}'.`);
    }
    if(n.type==='ref') {
      if(n.sheet){const result=this.externalSheetResolver?.(n.sheet,n.ref,visiting);if(!result)return error('#REF!',`Worksheet '${n.sheet}' does not exist.`);if(result.error)error(result.error.code,`${n.sheet}!${n.ref}: ${result.error.message}`);return result.value;}
      const key=known(n.ref,this.rows,this.cols);const v=this.getValue(key,visiting);
      if(v.error)error(v.error.code,`${key}: ${v.error.message}`);return v.value;
    }
    if(n.type==='spillRef'){
      if(n.sheet){if(!this.externalSpillResolver)return error('#REF!',`Worksheet '${n.sheet}' does not exist.`);return this.externalSpillResolver(n.sheet,n.ref,visiting);}
      return this.getSpillValues(n.ref,visiting);
    }
    if(n.type==='range') {
      const a=parseRef(n.from)!,b=parseRef(n.to)!;
      if(n.sheet){
        if(Math.max(a.row,b.row)>=this.rows||Math.max(a.col,b.col)>=this.cols)error('#REF!','Range extends beyond worksheet boundaries.');
        const matrix:Datum[][]=[];
        for(let row=Math.min(a.row,b.row);row<=Math.max(a.row,b.row);row++){const line:Datum[]=[];for(let col=Math.min(a.col,b.col);col<=Math.max(a.col,b.col);col++){
          const ref=`${colToName(col)}${row+1}`,result=this.externalSheetResolver?.(n.sheet,ref,visiting);if(!result)return error('#REF!',`Worksheet '${n.sheet}' does not exist.`);if(result.error)error(result.error.code,`${n.sheet}!${ref}: ${result.error.message}`);line.push(result.value);
        }matrix.push(line);}return matrix;
      }
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
    if(n.type==='unary') {
      if(n.op==='@'){
        if(n.expr.type==='range'){
          const from=parseRef(n.expr.from),to=parseRef(n.expr.to),owner=parseRef((visiting.at(-1)??'').split('!').at(-1)??'');if(!from||!to||!owner)return error('#VALUE!','Implicit intersection needs a cell formula context.');
          const r1=Math.min(from.row,to.row),r2=Math.max(from.row,to.row),c1=Math.min(from.col,to.col),c2=Math.max(from.col,to.col);let row:number,col:number;
          if(r1===r2&&c1===c2){row=r1;col=c1;}
          else if(c1===c2){if(owner.row<r1||owner.row>r2)return error('#VALUE!','Implicit intersection found no row match.');row=owner.row;col=c1;}
          else if(r1===r2){if(owner.col<c1||owner.col>c2)return error('#VALUE!','Implicit intersection found no column match.');row=r1;col=owner.col;}
          else{if(owner.row<r1||owner.row>r2||owner.col<c1||owner.col>c2)return error('#VALUE!','Implicit intersection found no matching cell.');row=owner.row;col=owner.col;}
          return this.evaluate({type:'ref',ref:`${colToName(col)}${row+1}`,...(n.expr.sheet?{sheet:n.expr.sheet}:{}),start:n.start,end:n.end},visiting);
        }
        const value=this.evaluate(n.expr,visiting);if(!Array.isArray(value))return value;if(value.length===1&&value[0]!.length===1)return value[0]![0]??null;return error('#VALUE!','Implicit intersection of computed arrays is not supported.');
      }
      const v=number(this.evaluateScalar(n.expr,visiting));return n.op==='-'?-v:n.op==='%'?v/100:v;
    }
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
      if(n.name==='LET'){
        const scope:{owner:string;values:Map<string,Datum>}={owner:visiting.at(-1)??'',values:new Map()};this.localFormulaNames.push(scope);
        try{for(let index=0;index<n.args.length-1;index+=2){const binding=n.args[index];if(!binding||binding.type!=='definedName')return error('#NAME?','LET variable name is invalid.');scope.values.set(binding.name.toUpperCase(),this.evaluate(n.args[index+1]!,visiting));}return this.evaluate(n.args.at(-1)!,visiting);}
        finally{this.localFormulaNames.pop();}
      }
      if(n.name==='SHEET'||n.name==='SHEETS'){
        if(n.args.length>1)return error('#ARG!',`${n.name} takes zero or one reference.`);
        if(n.name==='SHEETS'){
          if(!n.args.length)return this.documentSheetNames.length||1;
          const argument=n.args[0]!;if(argument.type!=='ref'&&argument.type!=='range'&&argument.type!=='spillRef')return error('#VALUE!','SHEETS requires a worksheet reference.');
          const sheetName=argument.sheet??this.sheetName,from=parseRef(argument.type==='range'?argument.from:argument.ref),to=parseRef(argument.type==='range'?argument.to:argument.ref);if(!from||!to||from.row>=this.rows||to.row>=this.rows||from.col>=this.cols||to.col>=this.cols)return error('#REF!','SHEETS reference is outside the worksheet.');if(!this.documentSheetNames.some(name=>name.toLocaleLowerCase()===sheetName.toLocaleLowerCase()))return error('#REF!',`Worksheet '${sheetName}' does not exist.`);return 1;
        }
        if(!n.args.length){const index=this.documentSheetNames.findIndex(name=>name.toLocaleLowerCase()===this.sheetName.toLocaleLowerCase());return index<0?1:index+1;}
        const argument=n.args[0]!;if(argument.type!=='ref'&&argument.type!=='range'&&argument.type!=='spillRef')return error('#VALUE!','SHEET requires a worksheet reference.');
        const ref=argument.type==='range'?argument.from:argument.ref,point=parseRef(ref);if(!point||point.row>=this.rows||point.col>=this.cols)return error('#REF!','SHEET reference is outside the worksheet.');
        const sheetName=argument.sheet??this.sheetName,index=this.documentSheetNames.findIndex(name=>name.toLocaleLowerCase()===sheetName.toLocaleLowerCase());return index<0?error('#REF!',`Worksheet '${sheetName}' does not exist.`):index+1;
      }
      const f=functions[n.name];
      if(!f) {
        const candidates=Object.keys(functions).filter(x=>x.startsWith(n.name.slice(0,2))).slice(0,4);
        return error('#NAME?',`Unknown function ${n.name}.${candidates.length?` Did you mean ${candidates.join(' or ')}?`:''}`);
      }
      if(n.name==='SUBTOTAL'){
        if(n.args.length<2||n.args.length>255)return error('#ARG!','SUBTOTAL(function_num, ref1, [ref2], …) needs 2 to 255 arguments.');
        const functionNumber=number(this.evaluateScalar(n.args[0]!,visiting));if(!Number.isInteger(functionNumber)||!((functionNumber>=1&&functionNumber<=11)||(functionNumber>=101&&functionNumber<=111)))return error('#VALUE!','SUBTOTAL function_num must be an integer from 1 to 11 or 101 to 111.');
        const ignoreHiddenRows=functionNumber>=101,functionCode=ignoreHiddenRows?functionNumber-100:functionNumber,evaluated:Datum[]=[functionCode];
        for(const argument of n.args.slice(1)){
          if(argument.type!=='ref'&&argument.type!=='range'){evaluated.push(this.evaluate(argument,visiting));continue;}
          const first=parseRef(argument.type==='range'?argument.from:argument.ref),last=parseRef(argument.type==='range'?argument.to:argument.ref);if(!first||!last||Math.max(first.row,last.row)>=this.rows||Math.max(first.col,last.col)>=this.cols)return error('#REF!','SUBTOTAL reference is outside worksheet boundaries.');
          const sheet=argument.sheet??this.sheetName,hiddenRows=new Set(this.hiddenRowsForSheet(sheet,ignoreHiddenRows));
          const matrix:Datum[][]=[];
          for(let row=Math.min(first.row,last.row);row<=Math.max(first.row,last.row);row++){
            const line:Datum[]=[];for(let col=Math.min(first.col,last.col);col<=Math.max(first.col,last.col);col++){
              const ref=`${colToName(col)}${row+1}`;if(hiddenRows.has(row)||this.nestedAggregateForSheet(sheet,ref)){line.push(null);continue;}
              const result=argument.sheet?this.externalSheetResolver?.(argument.sheet,ref,visiting):this.getValue(ref,visiting);if(!result)return error('#REF!',`Worksheet '${argument.sheet}' does not exist.`);if(result.error)return error(result.error.code,`${sheet}!${ref}: ${result.error.message}`);line.push(result.value);
            }matrix.push(line);
          }
          evaluated.push(matrix);
        }
        return f.run(evaluated);
      }
      if(n.name==='AGGREGATE'&&n.args.length>=3){
        const source=n.args[2]!;
        if(source.type==='ref'||source.type==='range'){
          const functionNumber=this.evaluateScalar(n.args[0]!,visiting),optionValue=this.evaluateScalar(n.args[1]!,visiting),functionCode=number(functionNumber),option=number(optionValue);if(!Number.isInteger(functionCode)||functionCode<1||functionCode>19)return error('#VALUE!','AGGREGATE function_num must be a supported integer from 1 to 19.');if(!Number.isInteger(option)||option<0||option>7)return error('#VALUE!','AGGREGATE option must be an integer from 0 to 7.');const ignoreHiddenRows=[1,3,5,7].includes(option),ignoreErrors=[2,3,6,7].includes(option),ignoreNested=option>=0&&option<=3,sheetName=source.sheet??this.sheetName,from=parseRef(source.type==='range'?source.from:source.ref),to=parseRef(source.type==='range'?source.to:source.ref);
          if(!from||!to||from.row>=this.rows||to.row>=this.rows||from.col>=this.cols||to.col>=this.cols)return error('#REF!','AGGREGATE reference is outside the worksheet.');
          const hiddenRows=new Set(this.hiddenRowsForSheet(sheetName,ignoreHiddenRows)),data:Datum[][]=[];
          for(let row=Math.min(from.row,to.row);row<=Math.max(from.row,to.row);row++){
            const line:Datum[]=[];for(let col=Math.min(from.col,to.col);col<=Math.max(from.col,to.col);col++){
              const ref=`${colToName(col)}${row+1}`;
              if(hiddenRows.has(row)||ignoreNested&&this.nestedAggregateForSheet(sheetName,ref)){line.push(null);continue;}
              const cell:Node={type:'ref',ref,...(source.sheet?{sheet:source.sheet}:{}),start:source.start,end:source.end};
              try{line.push(this.evaluate(cell,visiting));}catch(e){if(ignoreErrors&&e instanceof SheetError){line.push(null);continue;}throw e;}
            }data.push(line);
          }
          const args:Datum[]=[functionNumber,optionValue,data];if(n.args[3])args.push(this.evaluateScalar(n.args[3],visiting));return f.run(args);
        }
      }
      if(n.name==='GETPIVOTDATA'){
        if(n.args.length<2||n.args.length>255||(n.args.length-2)%2!==0)return error('#ARG!','GETPIVOTDATA(data_field, pivot_table, [field1, item1], …) needs a data field, a PivotTable reference, and complete field/item pairs.');
        const pivotNode=n.args[1]!;
        if(pivotNode.type!=='ref')return error('#VALUE!','The pivot_table argument must be a reference to a cell inside a PivotTable.');
        const pivotRef=normalizeRef(pivotNode.ref);if(pivotRef==='#REF!')return error('#REF!','The pivot_table argument is not a valid cell reference.');
        const dataField=this.evaluateScalar(n.args[0]!,visiting),criteria:PivotDataCriterion[]=[];
        for(let index=2;index<n.args.length;index+=2)criteria.push({field:this.evaluateScalar(n.args[index]!,visiting),item:this.evaluateScalar(n.args[index+1]!,visiting)});
        const resolver=this.pivotDataResolver;if(!resolver)return error('#REF!','GETPIVOTDATA needs a containing workbook PivotTable context.');
        const owner=visiting.at(-1)??'';
        return resolver(pivotNode.sheet??this.sheetName,pivotRef,dataField,criteria,(sheet,range,cycleRange)=>this.trackExternalRange(owner,sheet,range,cycleRange));
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
      if(n.name==='ISBLANK') {
        if(n.args.length!==1)error('#ARG!','ISBLANK needs one argument.');
        const argument=n.args[0]!;
        if(argument.type!=='ref')return false;
        if(!argument.sheet||argument.sheet.toLocaleLowerCase()===this.sheetName.toLocaleLowerCase())return this.getRaw(known(argument.ref,this.rows,this.cols))==='';
        try{return this.evaluate(argument,visiting)===null;}catch{return false;}
      }
      if(n.name==='ISREF') {if(n.args.length!==1)error('#ARG!','ISREF needs one argument.');return n.args[0]!.type==='ref'||n.args[0]!.type==='range'||n.args[0]!.type==='spillRef';}
      if(n.name==='ROW'||n.name==='COLUMN'){
        if(n.args.length>1)error('#ARG!',`${n.name} takes zero or one reference.`);
        const ref=n.args.length?this.referenceArg(n.args[0],visiting):(visiting[visiting.length-1]??'').split('!').pop()!;
        const p=parseRef(ref)!;return n.name==='ROW'?p.row+1:p.col+1;
      }
      if(n.name==='ISFORMULA'||n.name==='FORMULATEXT'){
        if(n.args.length!==1)error('#ARG!',`${n.name} takes one reference.`);
        const argument=n.args[0]!,ref=this.referenceArg(argument,visiting),sheet='sheet'in argument?argument.sheet??this.sheetName:this.sheetName,raw=sheet.toLocaleLowerCase()===this.sheetName.toLocaleLowerCase()?this.getRaw(ref):this.externalRawResolver?.(sheet,ref);
        if(raw===undefined)return error('#REF!',`Worksheet '${sheet}' does not exist.`);
        if(n.name==='ISFORMULA')return raw.startsWith('=');
        if(!raw.startsWith('='))return error('#N/A','Referenced cell does not contain a formula.');
        return raw;
      }
      if(n.name==='CELL'){
        if(n.args.length<1||n.args.length>2)error('#ARG!','CELL(info_type, [reference]) takes one or two arguments.');
        const kind=text(this.evaluateScalar(n.args[0]!,visiting)).toLowerCase(),argument=n.args[1],ref=argument?this.referenceArg(argument,visiting):(visiting.at(-1)??'').split('!').at(-1)!,p=parseRef(ref)!;
        if(argument&&argument.type!=='ref'&&argument.type!=='spillRef')error('#VALUE!','CELL requires a single cell reference.');
        const sheet=argument&&'sheet'in argument?argument.sheet??this.sheetName:this.sheetName,local=sheet.toLocaleLowerCase()===this.sheetName.toLocaleLowerCase(),style=local?this.getStyle(ref):{},resolved=local?this.getValue(ref,visiting):this.externalSheetResolver?.(sheet,ref,visiting);
        if(kind==='address')return `$${colToName(p.col)}$${p.row+1}`;
        if(kind==='row')return p.row+1;
        if(kind==='col')return p.col+1;
        if(kind==='contents'){if(!resolved)return error('#REF!',`Worksheet '${sheet}' does not exist.`);if(resolved.error)error(resolved.error.code,resolved.error.message);return resolved.value;}
        if(kind==='type'){if(!resolved)return error('#REF!',`Worksheet '${sheet}' does not exist.`);return resolved.error?'e':resolved.value===null?'b':typeof resolved.value==='string'?'l':'v';}
        if(kind==='format')return style.dateFormatCode||style.customNumberFormatCode&&isExcelDateNumberFormat(style.customNumberFormatCode)?'D4':style.customNumberFormatCode?'G':style.numberFormat==='integer'?'F0':style.numberFormat==='number'?'F2':style.numberFormat==='currency'?'C2':style.numberFormat==='percent'?'P2':'G';
        if(kind==='prefix'){if(!resolved)return error('#REF!',`Worksheet '${sheet}' does not exist.`);if(resolved.error||typeof resolved.value!=='string')return '';return style.align==='center'?'^':style.align==='right'?'"':'\'';}
        if(kind==='protect')return 1;
        return error('#VALUE!',`CELL info type '${kind}' not supported.`);
      }
      if(n.name==='INDIRECT'||n.name==='OFFSET') {
        const ownerIdentity=visiting[visiting.length-1]!,owner=ownerIdentity.split('!').at(-1)!;
        if(n.name==='INDIRECT') {
          if(n.args.length<1||n.args.length>2)error('#ARG!','INDIRECT(address, [a1]) takes 1–2 arguments.');
          const address=text(this.evaluateScalar(n.args[0]!,visiting)).trim(),a1=n.args.length<2||n.args[1]!.type==='literal'&&n.args[1]!.value===null||bool(this.evaluateScalar(n.args[1]!,visiting));
          if(!address)error('#REF!','INDIRECT address is empty.');
          let targetAddress=address;
          if(!a1){
            const prefix=/^(?:'((?:[^']|'')+)'|([A-Za-z_][A-Za-z0-9_.]*))!/.exec(address),targetSheet=prefix?(prefix[1]??prefix[2]!).replace(/''/g,"'"):undefined,reference=prefix?address.slice(prefix[0].length):address,parts=reference.split(':');
            if(parts.length>2)error('#REF!','INDIRECT requires a single R1C1 cell or range reference.');
            const origin=parseRef(owner);if(!origin)error('#REF!','INDIRECT formula location is invalid.');
            const convert=(part:string):string=>{
              const coordinate=/^R(\[(?:[+-]?\d+)\]|\d*)C(\[(?:[+-]?\d+)\]|\d*)$/i.exec(part);if(!coordinate)error('#REF!','INDIRECT R1C1 address is invalid.');
              const resolve=(token:string,base:number):number=>token.startsWith('[')?base+Number(token.slice(1,-1)):token===''?base:Number(token)-1;
              const row=resolve(coordinate![1]!,origin!.row),col=resolve(coordinate![2]!,origin!.col);if(row<0||row>=this.rows||col<0||col>=this.cols)error('#REF!','INDIRECT R1C1 address is outside the worksheet.');return `${colToName(col)}${row+1}`;
            };
            const resolved=parts.map(convert).join(':');targetAddress=targetSheet?`'${targetSheet.replace(/'/g,"''")}'!${resolved}`:resolved;
          }
          const parsedTarget=this.resolveDefinedNames(parseFormula(`=${targetAddress}`));
          if(parsedTarget.type!=='ref'&&parsedTarget.type!=='range'&&parsedTarget.type!=='spillRef')error('#REF!','INDIRECT requires a valid A1-style cell or range reference.');
          const target=parsedTarget as Extract<Node,{type:'ref'|'range'|'spillRef'}>,targetSheet=target.sheet??this.sheetName,targetRange=target.type==='range'?`${target.from}:${target.to}`:target.ref;
          this.trackDynamicRange(owner,targetSheet,targetRange);
          return this.evaluate(target,visiting);
        }else{
          if(n.args.length<3||n.args.length>5)error('#ARG!','OFFSET(reference, rows, cols, [height], [width]) requires 3–5 arguments.');
          const baseNode=n.args[0]!;
          if(baseNode.type!=='ref'&&baseNode.type!=='range')error('#VALUE!','OFFSET requires a cell or range reference.');
          const base=baseNode as Extract<Node,{type:'ref'|'range'}>,first=base.type==='ref'?parseRef(base.ref):parseRef(base.from),last=base.type==='ref'?first:parseRef(base.to);if(!first||!last)error('#REF!','OFFSET base reference is invalid.');
          const row=Math.min(first!.row,last!.row)+Math.trunc(number(this.evaluateScalar(n.args[1]!,visiting))),col=Math.min(first!.col,last!.col)+Math.trunc(number(this.evaluateScalar(n.args[2]!,visiting)));
          const height=n.args.length>=4&&!(n.args[3]!.type==='literal'&&n.args[3]!.value===null)?Math.trunc(number(this.evaluateScalar(n.args[3]!,visiting))):Math.abs(last!.row-first!.row)+1;
          const width=n.args.length>=5&&!(n.args[4]!.type==='literal'&&n.args[4]!.value===null)?Math.trunc(number(this.evaluateScalar(n.args[4]!,visiting))):Math.abs(last!.col-first!.col)+1;
          if(height<1||width<1)error('#VALUE!','OFFSET height and width must be positive.');
          if(row<0||col<0||row+height>this.rows||col+width>this.cols)error('#REF!','OFFSET result extends outside the worksheet.');
          const to=`${colToName(col+width-1)}${row+height}`;let target:Node;
          if(height===1&&width===1)target={type:'ref',ref:`${colToName(col)}${row+1}`,...(base.sheet?{sheet:base.sheet}:{}),start:base.start,end:base.end};
          else target={type:'range',from:`${colToName(col)}${row+1}`,to,...(base.sheet?{sheet:base.sheet}:{}),start:base.start,end:base.end};
          this.trackDynamicRange(owner,base.sheet??this.sheetName,height===1&&width===1?`${colToName(col)}${row+1}`:`${colToName(col)}${row+1}:${to}`);
          return this.evaluate(target,visiting);
        }
      }
      if(n.name==='IF') {
        if(n.args.length<2||n.args.length>3)error('#ARG!','IF(condition, true_value, false_value) takes 2–3 arguments.');
        const condition=bool(this.evaluateScalar(n.args[0]!,visiting));
        return condition?this.evaluate(n.args[1]!,visiting):n.args[2]?this.evaluate(n.args[2],visiting):false;
      }
      if(n.name==='IFERROR') {
        if(n.args.length!==2)error('#ARG!','IFERROR(value, fallback) takes exactly 2 arguments.');
        try{return this.evaluate(n.args[0]!,visiting);}catch(e){if(e instanceof SheetError||e instanceof FormulaSyntaxError)return this.evaluate(n.args[1]!,visiting);throw e;}
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
      if(n.name==='COUNT') {
        if(n.args.length<1||n.args.length>255)error('#ARG!','COUNT requires 1 to 255 arguments.');
        let count=0;
        const countNumber=(value:Datum):number=>{
          if(!Array.isArray(value))return typeof value==='number'&&Number.isFinite(value)?1:0;
          let total=0;for(const row of value)for(const cell of row)total+=countNumber(cell);return total;
        };
        for(const argument of n.args){
          if(argument.type==='ref'){
            const value=argument.sheet?this.externalSheetResolver?.(argument.sheet,argument.ref,visiting)??{value:null,error:{code:'#REF!',message:'Referenced worksheet is unavailable.',position:0,length:1,suggestion:'Check the worksheet name.'}}:this.getValue(argument.ref,visiting);
            if(!value.error&&typeof value.value==='number'&&Number.isFinite(value.value))count++;
            continue;
          }
          if(argument.type==='range'){
            const first=parseRef(argument.from)!,last=parseRef(argument.to)!,r1=Math.min(first.row,last.row),r2=Math.max(first.row,last.row),c1=Math.min(first.col,last.col),c2=Math.max(first.col,last.col);
            if(r2>=this.rows||c2>=this.cols)error('#REF!','Range extends beyond sheet boundaries.');
            for(let row=r1;row<=r2;row++)for(let col=c1;col<=c2;col++){
              const ref=`${colToName(col)}${row+1}`,value=argument.sheet?this.externalSheetResolver?.(argument.sheet,ref,visiting)??{value:null,error:{code:'#REF!',message:'Referenced worksheet is unavailable.',position:0,length:1,suggestion:'Check the worksheet name.'}}:this.getValue(ref,visiting);
              if(!value.error&&typeof value.value==='number'&&Number.isFinite(value.value))count++;
            }
            continue;
          }
          try{
            const value=this.evaluate(argument,visiting);
            if(argument.type==='literal'){
              if(typeof value==='number'||typeof value==='boolean')count++;
              else if(typeof value==='string'&&value.trim()!==''&&Number.isFinite(Number(value)))count++;
            }else count+=countNumber(value);
          }catch{/* Excel COUNT ignores error values. */}
        }
        return count;
      }
      if(n.name==='COUNTA') {
        if(n.args.length<1||n.args.length>255)error('#ARG!','COUNTA requires 1 to 255 arguments.');
        let count=0;
        const localSheet=(sheet:string|undefined)=>!sheet||sheet.toLocaleLowerCase()===this.sheetName.toLocaleLowerCase();
        const countRef=(ref:string,sheet:string|undefined)=>{
          if(localSheet(sheet)){
            const key=known(ref,this.rows,this.cols),spilled=this.spillChildren.get(key);
            if(this.getRaw(key)!==''||spilled&&spilled.value!==null)count++;
            return;
          }
          const value=this.externalSheetResolver?.(sheet!,ref,visiting);
          if(!value)return error('#REF!',`Worksheet '${sheet}' does not exist.`);
          if(value.error||value.value!==null)count++;
        };
        for(const argument of n.args){
          if(argument.type==='ref'){countRef(argument.ref,argument.sheet);continue;}
          if(argument.type==='range'){
            const first=parseRef(argument.from)!,last=parseRef(argument.to)!,r1=Math.min(first.row,last.row),r2=Math.max(first.row,last.row),c1=Math.min(first.col,last.col),c2=Math.max(first.col,last.col);
            if(r2>=this.rows||c2>=this.cols)error('#REF!','Range extends beyond sheet boundaries.');
            for(let row=r1;row<=r2;row++)for(let col=c1;col<=c2;col++)countRef(`${colToName(col)}${row+1}`,argument.sheet);
            continue;
          }
          try{
            const countValues=(item:Datum):number=>{if(!Array.isArray(item))return item===null?0:1;let subtotal=0;for(const child of item as unknown as Datum[])subtotal+=countValues(child);return subtotal;};
            count+=countValues(this.evaluate(argument,visiting));
          }catch{count++;}
        }
        return count;
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
    const cloned=Workbook.fromJSON(this.toJSON());if(this.tableContext.length)cloned.setTableContext(this.tableContext);const stats=cloned.applyEdits([{ref:target,raw:proposed}],false);
    return {changes:impacted.map(id=>({ref:id,before:before.get(id)??'',after:display(cloned.getValue(id))})).filter(c=>c.before!==c.after),stats};
  }
  transitiveDependents(ref:string):string[] {
    const queue=[known(ref,this.rows,this.cols)], seen=new Set(queue);
    for(let i=0;i<queue.length;i++)for(const k of this.rev.get(queue[i]!)??[])if(!seen.has(k)){seen.add(k);queue.push(k);}
    return queue;
  }
  toJSON():{version:2;rows:number;cols:number;cells:Record<string,string>;styles:Record<string,CellStyle>;definedNames:DefinedName[]} {return {version:2,rows:this.rows,cols:this.cols,cells:Object.fromEntries(this.raw),styles:Object.fromEntries(this.styles),definedNames:this.getDefinedNames()};}
  static fromJSON(data:unknown):Workbook {
    if(!data||typeof data!=='object')throw Error('Invalid workbook data.');
    const d=data as {rows?:number;cols?:number;cells?:Record<string,unknown>;styles?:Record<string,unknown>;definedNames?:unknown};
    if(!Number.isInteger(d.rows)||!Number.isInteger(d.cols)||(d.rows??0)<1||(d.rows??0)>1000||(d.cols??0)<1||(d.cols??0)>104)throw Error('Workbook dimensions must be 1–1000 rows and 1–104 columns.');
    if(!d.cells||typeof d.cells!=='object'||Array.isArray(d.cells))throw Error('Workbook cells must be an object.');
    const book=new Workbook(d.rows!,d.cols!);if(d.definedNames!==undefined)book.setDefinedNames(d.definedNames,false);
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
