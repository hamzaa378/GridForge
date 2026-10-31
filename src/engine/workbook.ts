import { colToName, normalizeRef, parseFormula, parseRef, externalReferences, type Diagnostic, type Node, type Value, FormulaSyntaxError } from './parser.js';
import { bool, error, functions, matchCriterion, number, SheetError, text, type Datum } from './functions.js';
import {cleanStyle, isExcelDateNumberFormat, type CellStyle, type StyleEdit} from './format.js';
import {resolveStructuredFormula, type TableSpec} from './tables.js';
import {MAX_WORKSHEET_COLUMNS,MAX_WORKSHEET_POPULATED_CELLS,MAX_WORKSHEET_ROWS} from './limits.js';
export type CellResult = { value: Value; error?: Diagnostic };
export type Edit = { ref: string; raw: string };
export interface RecalcStats { updated: number; recalculated: number; affected: string[]; durationMs: number; totalPopulated: number }
export interface Inspect { ref: string; raw: string; result: CellResult; dependencies: string[]; dependents: string[]; expressionTree: string; referencedValues: { ref: string; raw: string; display: string }[] }
export interface DefinedName {name:string;ref:string;comment?:string;hidden?:boolean}
export interface PivotDataCriterion {field:Value;item:Value}
export type PivotRangeTracker=(sheet:string,range:string,cycleRange?:string)=>void;
export type PivotDataResolver=(pivotSheet:string,pivotRef:string,dataField:Value,criteria:PivotDataCriterion[],trackRange:PivotRangeTracker)=>Value;
type DynamicRangeDependency={sheet:string;from:string;to:string};
type LocalRangeDependency={from:string;to:string};
function allowsSparseLargeLocalRanges(node:Node,rows:number,cols:number,sheetName:string):boolean {
  if(node.type!=='call')return false;
  const localRange=(argument:Node)=>((argument.type==='ref'||argument.type==='range')&&(!argument.sheet||argument.sheet.toLocaleLowerCase()===sheetName.toLocaleLowerCase()));
  if(['SUM','AVERAGE','MIN','MAX','PRODUCT','STDEV','VAR','STDEV.S','VAR.S','STDEV.P','VAR.P'].includes(node.name))return node.args.length>=1&&node.args.length<=255&&node.args.every(localRange);
  if(node.name==='MEDIAN')return node.args.length>=1&&node.args.length<=255&&node.args.every(localRange);
  if(['LARGE','SMALL'].includes(node.name))return node.args.length===2&&localRange(node.args[0]!)&&node.args[1]!.type!=='range'&&!(node.args[1]!.type==='ref'&&node.args[1]!.sheet&&node.args[1]!.sheet.toLocaleLowerCase()!==sheetName.toLocaleLowerCase());
  if(['COUNT','COUNTA','COUNTBLANK'].includes(node.name))return node.args.length>=1&&node.args.length<=255&&node.args.every(localRange);
  if(node.name==='COUNTIFS')return node.args.length>=2&&node.args.length%2===0&&node.args.filter((_,index)=>index%2===0).every(localRange);
  if(['SUMIFS','AVERAGEIFS','MAXIFS','MINIFS'].includes(node.name))return node.args.length>=3&&node.args.length%2===1&&node.args.every((argument,index)=>index%2===0?localRange(argument):true);
  return false;
}
function collectLocalFormulaDependencies(node:Node,rows:number,cols:number,allowLargeRanges=false):{cells:Set<string>;ranges:LocalRangeDependency[]} {
  const cells=new Set<string>(),ranges:LocalRangeDependency[]=[];
  const visit=(current:Node):void=>{
    if(current.type==='ref'||current.type==='spillRef'){
      const point=parseRef(current.ref);if(!current.sheet&&point&&point.row<rows&&point.col<cols)cells.add(normalizeRef(current.ref));return;
    }
    if(current.type==='range'){
      if(current.sheet)return;
      const first=parseRef(current.from),last=parseRef(current.to);if(!first||!last)return;
      const top=Math.min(first.row,last.row),bottom=Math.max(first.row,last.row),left=Math.min(first.col,last.col),right=Math.max(first.col,last.col);
      if((bottom-top+1)*(right-left+1)>(allowLargeRanges?rows*cols:20_000))throw new FormulaSyntaxError(`Range exceeds ${allowLargeRanges?(rows*cols).toLocaleString():'20,000'} cells.`,current.start,'Use a smaller range.',current.end-current.start);
      const clippedBottom=Math.min(bottom,rows-1),clippedRight=Math.min(right,cols-1);if(top>clippedBottom||left>clippedRight)return;
      ranges.push({from:`${colToName(left)}${top+1}`,to:`${colToName(clippedRight)}${clippedBottom+1}`});return;
    }
    if(current.type==='unary')visit(current.expr);
    else if(current.type==='binary'){visit(current.left);visit(current.right);}
    else if(current.type==='call')current.args.forEach(visit);
  };
  visit(node);return {cells,ranges};
}
type HistoryEntry={before:Edit[];after:Edit[];styleBefore:StyleEdit[];styleAfter:StyleEdit[];namesBefore?:DefinedName[];namesAfter?:DefinedName[]};
const displayNumber = (n: number): string => Number.isFinite(n) ? (Object.is(n,-0)?'0':String(Number(n.toPrecision(15)))) : '#NUM!';
const MAX_SPILL_CELLS = 20_000;
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
  private hiddenColumnsForSheet:(sheet:string)=>readonly number[]=()=>[];
  private nestedAggregateForSheet:(sheet:string,ref:string)=>boolean=()=>false;
  private externalSheetResolver:((sheet:string,ref:string,visiting:string[])=>CellResult)|null=null;
  private externalRawResolver:((sheet:string,ref:string)=>string|undefined)|null=null;
  private externalSpillResolver:((sheet:string,ref:string,visiting:string[])=>Datum[][])|null=null;
  private pivotDataResolver:PivotDataResolver|null=null;
  private tableContext: readonly TableSpec[] = [];
  private definedNames=new Map<string,DefinedName>();
  private workbookDefinedNames=new Map<string,DefinedName>();
  private localFormulaNames:{owner:string;values:Map<string,Datum>;omitted:Set<string>}[]=[];
  private raw = new Map<string,string>();
  /** Sparse occupied-row index keeps aggregate range scans proportional to stored cells in the requested columns. */
  private populatedRowsByColumn=new Map<number,Set<number>>();
  private styles = new Map<string,CellStyle>();
  private ast = new Map<string,Node>();
  private parseErrors = new Map<string,Diagnostic>();
  private deps = new Map<string,Set<string>>();
  private rev = new Map<string,Set<string>>();
  /** Local range references stay compact; column buckets find impacted formulas on edits. */
  private staticRangeDeps=new Map<string,LocalRangeDependency[]>();
  private staticRangeColumns=new Map<string,Set<string>>();
  private dynamicDeps = new Map<string,Set<string>>();
  private dynamicRanges=new Map<string,DynamicRangeDependency[]>();
  /** Column buckets avoid scanning every INDIRECT/OFFSET range after each edit. */
  private dynamicRangeColumns=new Map<string,Set<string>>();
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
  private indexPopulatedCell(ref:string):void{const point=parseRef(ref);if(!point)return;let rows=this.populatedRowsByColumn.get(point.col);if(!rows){rows=new Set();this.populatedRowsByColumn.set(point.col,rows);}rows.add(point.row);}
  private unindexPopulatedCell(ref:string):void{if(this.raw.has(ref)||this.spillChildren.has(ref))return;const point=parseRef(ref);if(!point)return;const rows=this.populatedRowsByColumn.get(point.col);rows?.delete(point.row);if(rows?.size===0)this.populatedRowsByColumn.delete(point.col);}
  /** Link this worksheet to its containing document so formulas can resolve sheet-qualified references. */
  setSheetContext(name:string,resolver:(sheet:string,ref:string,visiting:string[])=>CellResult,onCellsChanged?:(refs:string[])=>void,spillResolver?:(sheet:string,ref:string,visiting:string[])=>Datum[][],documentSheetNames:string[]=[],hiddenRowsForSheet?:(sheet:string,includeManuallyHidden:boolean)=>readonly number[],nestedAggregateForSheet?:(sheet:string,ref:string)=>boolean,rawResolver?:(sheet:string,ref:string)=>string|undefined,deferRecalculation=false,hiddenColumnsForSheet?:(sheet:string)=>readonly number[]):void {
    const namesChanged=this.sheetName!==name||this.documentSheetNames.join('\u0000')!==documentSheetNames.join('\u0000');this.sheetName=name;this.externalSheetResolver=resolver;this.externalRawResolver=rawResolver??null;this.onCellsChanged=onCellsChanged??null;this.externalSpillResolver=spillResolver??null;this.documentSheetNames=[...documentSheetNames];this.hiddenRowsForSheet=hiddenRowsForSheet??(()=>[]);this.hiddenColumnsForSheet=hiddenColumnsForSheet??(()=>[]);this.nestedAggregateForSheet=nestedAggregateForSheet??((sheet,ref)=>sheet.toLocaleLowerCase()===this.sheetName.toLocaleLowerCase()&&this.hasNestedAggregateFormula(ref));if(!namesChanged)return;
    this.staticRangeColumns.clear();for(const [owner,ranges] of this.staticRangeDeps)for(const range of ranges){const from=parseRef(range.from),to=parseRef(range.to);if(!from||!to)continue;for(let col=Math.min(from.col,to.col);col<=Math.max(from.col,to.col);col++){const key=`${this.sheetName.toLocaleLowerCase()}!${col}`;if(!this.staticRangeColumns.has(key))this.staticRangeColumns.set(key,new Set());this.staticRangeColumns.get(key)!.add(owner);}}
    const usesSheetInfo=(node:Node):boolean=>node.type==='call'?(node.name==='SHEET'||node.name==='SHEETS'||node.args.some(usesSheetInfo)):node.type==='unary'?usesSheetInfo(node.expr):node.type==='binary'?usesSheetInfo(node.left)||usesSheetInfo(node.right):false;
    const formulas=[...this.ast].filter(([,node])=>usesSheetInfo(node)).map(([ref])=>ref);if(!formulas.length)return;const affected=this.invalidate(formulas);this.valueRevisionCounter++;if(!deferRecalculation)for(const ref of affected)if(this.raw.has(ref))this.getValue(ref);if(!deferRecalculation)this.onCellsChanged?.(affected);
  }
  getHiddenRowsForChart():readonly number[]{return this.hiddenRowsForSheet(this.sheetName,true);}
  getHiddenColumnsForChart():readonly number[]{return this.hiddenColumnsForSheet(this.sheetName);}
  hasNestedAggregateFormula(ref:string):boolean {
    const root=this.ast.get(normalizeRef(ref));if(!root)return false;const visit=(node:Node):boolean=>node.type==='call'?(node.name==='SUBTOTAL'||node.name==='AGGREGATE'||node.args.some(visit)):node.type==='unary'?visit(node.expr):node.type==='binary'?visit(node.left)||visit(node.right):false;return visit(root);
  }
  /** Refresh SUBTOTAL/AGGREGATE formulas after row visibility changes, which are not cell edits. */
  recalculateAggregatesForRowVisibilityChange():void {
    const usesAggregate=(node:Node):boolean=>node.type==='call'?(node.name==='AGGREGATE'||node.args.some(usesAggregate)):node.type==='unary'?usesAggregate(node.expr):node.type==='binary'?usesAggregate(node.left)||usesAggregate(node.right):false;
    const formulas=[...this.ast].filter(([,node])=>usesAggregate(node)).map(([ref])=>ref);if(!formulas.length)return;const affected=this.invalidate(formulas);this.valueRevisionCounter++;for(const ref of affected)if(this.raw.has(ref))this.getValue(ref);this.onCellsChanged?.(affected);
  }
  /** Attach the containing document's PivotTable lookup after worksheet contexts are bound. */
  setPivotDataContext(resolver:PivotDataResolver|null,deferRecalculation=false):void {
    this.pivotDataResolver=resolver;
    const usesPivotData=(node:Node):boolean=>node.type==='call'?(node.name==='GETPIVOTDATA'||node.args.some(usesPivotData)):node.type==='unary'?usesPivotData(node.expr):node.type==='binary'?usesPivotData(node.left)||usesPivotData(node.right):false;
    const formulas=[...this.ast].filter(([,node])=>usesPivotData(node)).map(([ref])=>ref);
    if(!formulas.length)return;
    const affected=this.invalidate(formulas);this.valueRevisionCounter++;if(!deferRecalculation)for(const ref of affected)if(this.raw.has(ref))this.getValue(ref);
    if(!deferRecalculation)this.onCellsChanged?.(affected);
  }
  getDefinedNames():DefinedName[]{return [...this.definedNames.values()].map(item=>({...item}));}
  getWorkbookDefinedNames():DefinedName[]{return [...this.workbookDefinedNames.values()].map(item=>({...item}));}
  private cleanDefinedNames(items:unknown):Map<string,DefinedName>{
    if(!Array.isArray(items))throw Error('Defined names must be a list.');
    if(items.length>500)throw Error('A worksheet can contain at most 500 defined names.');
    const result=new Map<string,DefinedName>();
    for(const item of items){
      if(!item||typeof item!=='object')throw Error('Invalid defined name.');
      const raw=item as Record<string,unknown>,name=typeof raw.name==='string'?raw.name.trim():'';
      const rawRef=typeof raw.ref==='string'?raw.ref.trim():'';
      const ref=rawRef.startsWith('=')?rawRef:rawRef.replace(/\$/g,'').toUpperCase();
      if(name.length>255||!/^[_A-Za-z\\][A-Za-z0-9_.]*$/.test(name)||parseRef(name)||/^[RC]\d+$/i.test(name)||/^(TRUE|FALSE)$/i.test(name)||functions[name.toUpperCase()])throw Error(`'${name||'Name'}' is not a valid defined name.`);
      const key=name.toUpperCase();if(result.has(key))throw Error(`The name '${name}' is already defined on this worksheet.`);
      if(ref.startsWith('=')){
        let formula:Node;try{formula=parseFormula(ref);}catch(cause){throw Error(`'${name}' has an invalid formula: ${cause instanceof Error?cause.message:'formula parse failed'}.`);}
        if(formula.type==='call'&&formula.name==='LAMBDA'){
          if(formula.args.length<2||formula.args.length>254)throw Error(`'${name}' must define a LAMBDA with a calculation.`);
          const parameters=formula.args.slice(0,-1),seen=new Set<string>();for(const parameter of parameters){if(parameter.type!=='definedName'||parseRef(parameter.name)||!/^[_A-Za-z][A-Za-z0-9_.]*$/.test(parameter.name)||functions[parameter.name.toUpperCase()]||seen.has(parameter.name.toUpperCase()))throw Error(`'${name}' contains an invalid or repeated LAMBDA parameter.`);seen.add(parameter.name.toUpperCase());}
        }
        const clean:DefinedName={name,ref};if(typeof raw.comment==='string'&&raw.comment.trim())clean.comment=raw.comment.trim().slice(0,255);if(raw.hidden===true)clean.hidden=true;result.set(key,clean);continue;
      }
      const pieces=ref.split(':');if(pieces.length>2||pieces.some(part=>!parseRef(part)))throw Error(`'${ref}' must refer to a cell or rectangular range such as A1:B5.`);
      for(const part of pieces){const p=parseRef(part)!;if(p.row>=this.rows||p.col>=this.cols)throw Error(`'${ref}' is outside this worksheet.`);}
      const a=parseRef(pieces[0]!)!,b=parseRef(pieces[1]??pieces[0]!)!;
      const first=`${colToName(Math.min(a.col,b.col))}${Math.min(a.row,b.row)+1}`,last=`${colToName(Math.max(a.col,b.col))}${Math.max(a.row,b.row)+1}`;
      const clean:DefinedName={name,ref:first===last?first:`${first}:${last}`};
      if(typeof raw.comment==='string'&&raw.comment.trim())clean.comment=raw.comment.trim().slice(0,255);
      if(raw.hidden===true)clean.hidden=true;
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
  /** Apply workbook-scoped names after local names, so worksheet names can override them. */
  setWorkbookDefinedNames(items:unknown,deferRecalculation=false):void{
    if(!Array.isArray(items)||items.length>500)throw Error('A workbook can contain at most 500 defined names.');
    const next=new Map<string,DefinedName>();
    for(const item of items){
      if(!item||typeof item!=='object')throw Error('Invalid workbook defined name.');
      const raw=item as Record<string,unknown>,name=typeof raw.name==='string'?raw.name.trim():'',source=typeof raw.ref==='string'?raw.ref.trim():'';
      if(name.length>255||!/^[_A-Za-z\\][A-Za-z0-9_.]*$/.test(name)||parseRef(name)||/^[RC]\d+$/i.test(name)||/^(TRUE|FALSE)$/i.test(name)||functions[name.toUpperCase()])throw Error(`'${name||'Name'}' is not a valid workbook defined name.`);
      const key=name.toUpperCase();if(next.has(key))throw Error(`The workbook name '${name}' is defined more than once.`);
      const ref=source.startsWith('=')?source:`=${source}`;let formula:Node;try{formula=parseFormula(ref);}catch(cause){throw Error(`'${name}' has an invalid workbook formula: ${cause instanceof Error?cause.message:'formula parse failed'}.`);}
      if(formula.type==='call'&&formula.name==='LAMBDA'){if(formula.args.length<2||formula.args.length>254)throw Error(`'${name}' must define a LAMBDA with a calculation.`);const parameters=formula.args.slice(0,-1),seen=new Set<string>();for(const parameter of parameters){if(parameter.type!=='definedName'||parseRef(parameter.name)||!/^[_A-Za-z][A-Za-z0-9_.]*$/.test(parameter.name)||functions[parameter.name.toUpperCase()]||seen.has(parameter.name.toUpperCase()))throw Error(`'${name}' contains an invalid or repeated LAMBDA parameter.`);seen.add(parameter.name.toUpperCase());}}
      const clean:DefinedName={name,ref,...(typeof raw.comment==='string'&&raw.comment.trim()?{comment:raw.comment.trim().slice(0,255)}:{}),...(raw.hidden===true?{hidden:true}:{})};next.set(key,clean);
    }
    const before=this.getWorkbookDefinedNames(),after=[...next.values()];if(JSON.stringify(before)===JSON.stringify(after))return;
    this.workbookDefinedNames=next;this.rebuildNamedFormulas(deferRecalculation);this.valueRevisionCounter++;
  }
  private rebuildNamedFormulas(deferRecalculation=false):void{
    this.ast.clear();this.parseErrors.clear();this.deps.clear();this.rev.clear();this.dynamicDeps.clear();this.externalDeps.clear();this.staticRangeDeps.clear();this.staticRangeColumns.clear();
    for(const {ref,raw} of this.allEntries())if(raw.startsWith('='))this.update(ref,raw);
    this.cache.clear();this.dirty=new Set(this.raw.keys());
    const t0=performance.now(),affected=[...this.raw.keys()];let recalculated=0;
    if(!deferRecalculation)for(const ref of affected){if(this.raw.has(ref)){this.getValue(ref);recalculated++;}}
    this.lastStats={updated:0,recalculated,affected,durationMs:Math.round((performance.now()-t0)*1000)/1000,totalPopulated:this.raw.size};if(!deferRecalculation&&affected.length)this.onCellsChanged?.(affected);
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
      if(destination.startsWith('#'))info={location:destination.slice(1)};else if(/^[a-z][a-z0-9+.-]*:/i.test(destination)){if(!/\s/.test(destination)){const url=new URL(destination);if((url.protocol==='https:'||url.protocol==='http:')&&!!url.hostname)info={target:destination};else if(url.protocol==='mailto:'&&/^mailto:[^\s@]+@[^\s@]+$/i.test(destination))info={target:destination};}}else if(parseRef(destination.replace(/\$/g,''))||destination.includes('!')||this.resolveHyperlinkLocation(destination)!==null)info={location:destination};
      if(!info)return null;const label=node.args.length===2?text(this.evaluateScalar(node.args[1]!,[key])):text(this.getValue(key).value);return {...info,...(label?{display:label}:{})};
    }catch{return null;}
  }
  /** Resolve an internal hyperlink to a cell or range, including workbook and sheet-scoped names. */
  resolveHyperlinkLocation(location:string):string|null {
    const address=location.replace(/^#/,'').trim();if(!address||address.length>2048)return null;
    try{
      const target=this.resolveDefinedNames(parseFormula(`=${address}`));
      if(target.type!=='ref'&&target.type!=='range')return null;
      const first=parseRef(target.type==='range'?target.from:target.ref),last=target.type==='range'?parseRef(target.to):first;
      if(!first||!last||first.row<0||first.col<0||last.row<first.row||last.col<first.col||last.row>=this.rows||last.col>=this.cols)return null;
      const reference=target.type==='range'?`${colToName(first.col)}${first.row+1}:${colToName(last.col)}${last.row+1}`:`${colToName(first.col)}${first.row+1}`;
      return target.sheet?`'${target.sheet.replace(/'/g,"''")}'!${reference}`:reference;
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
    const projected=this.styles.size+changed.filter(([ref,style])=>Object.keys(style).length>0&&!this.styles.has(ref)).length-changed.filter(([ref,style])=>Object.keys(style).length===0&&this.styles.has(ref)).length;
    if(projected>MAX_WORKSHEET_POPULATED_CELLS)throw Error(`Worksheet exceeds GridForge's ${MAX_WORKSHEET_POPULATED_CELLS.toLocaleString()} styled-cell limit.`);
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
    this.deps.delete(ref);this.externalDeps.delete(ref);this.clearDynamic(ref);this.clearStaticRanges(ref);this.ast.delete(ref);this.parseErrors.delete(ref);
    if (raw) {this.raw.set(ref,raw);this.indexPopulatedCell(ref);} else {this.raw.delete(ref);this.unindexPopulatedCell(ref);}
    if (raw.startsWith('=')) {
      try {
        const parsed=this.resolveDefinedNames(parseFormula(this.tableContext.length?resolveStructuredFormula(raw,this,this.tableContext,ref):raw));
        const {cells:refs,ranges}=collectLocalFormulaDependencies(parsed,this.rows,this.cols,allowsSparseLargeLocalRanges(parsed,this.rows,this.cols,this.sheetName));
        const crossRefs=externalReferences(parsed,this.rows,this.cols);if(crossRefs.length)this.externalDeps.set(ref,crossRefs);
        for(const range of ranges)this.indexStaticRange(ref,range);
        this.ast.set(ref,parsed);this.deps.set(ref,refs);
        for(const dep of refs) {if(!this.rev.has(dep))this.rev.set(dep,new Set());this.rev.get(dep)!.add(ref);}
      } catch(e) {
        const d=e instanceof FormulaSyntaxError?e.diagnostic:diag('#PARSE!',e instanceof Error?e.message:String(e));
        this.parseErrors.set(ref,d);
      }
    }
  }
  private resolveDefinedNames(node:Node,locals:ReadonlySet<string>=new Set(),resolving:ReadonlySet<string>=new Set()):Node{
    if(node.type==='definedName'){
      if(locals.has(node.name.toUpperCase()))return node;
      const def=this.definedNames.get(node.name.toUpperCase())??this.workbookDefinedNames.get(node.name.toUpperCase());
      if(!def)return node;
      if(def.ref.startsWith('=')){
        const key=node.name.toUpperCase();if(resolving.has(key))throw new FormulaSyntaxError(`Defined LAMBDA '${node.name}' is recursive.`,node.start,'Use a non-recursive LAMBDA definition.',node.end-node.start);
        const nested=new Set(resolving);nested.add(key);return this.resolveDefinedNames(parseFormula(def.ref),locals,nested);
      }
      const parts=def.ref.split(':');
      if(parts.length===1)return {type:'ref',ref:parts[0]!,start:node.start,end:node.end};
      return {type:'range',from:parts[0]!,to:parts[1]!,start:node.start,end:node.end};
    }
    if(node.type==='unary')return {...node,expr:this.resolveDefinedNames(node.expr,locals,resolving)};
    if(node.type==='binary')return {...node,left:this.resolveDefinedNames(node.left,locals,resolving),right:this.resolveDefinedNames(node.right,locals,resolving)};
    if(node.type==='call'&&node.name==='LAMBDA'){
      const scope=new Set(locals);for(const parameter of node.args.slice(0,-1)){if(parameter.type==='definedName')scope.add(parameter.name.toUpperCase());}
      return {...node,args:[...node.args.slice(0,-1),this.resolveDefinedNames(node.args.at(-1)!,scope,resolving)]};
    }
    if(node.type==='call'&&node.name==='LET'){
      if(node.args.length<3||node.args.length>253||node.args.length%2!==1)throw new FormulaSyntaxError('LET needs 1 to 126 name/value pairs and a final expression.',node.start,'Use LET(name, value, calculation).',node.end-node.start);
      const scope=new Set(locals),args:Node[]=[],bound=new Set<string>();
      for(let index=0;index<node.args.length-1;index+=2){
        const name=node.args[index]!,value=node.args[index+1]!,key=name.type==='definedName'?name.name.toUpperCase():'';
        if(!key||!/^[A-Z_][A-Z0-9_.]*$/.test(key)||parseRef(key))throw new FormulaSyntaxError('LET variable names must be valid identifiers, not cell references.',name.start,'Use a name such as total or _rate.',name.end-name.start);
        if(bound.has(key))throw new FormulaSyntaxError(`LET variable '${key}' is defined more than once.`,name.start,'Give each LET binding a different name.',name.end-name.start);
        bound.add(key);args.push(name,this.resolveDefinedNames(value,scope,resolving));scope.add(key);
      }
      args.push(this.resolveDefinedNames(node.args.at(-1)!,scope,resolving));return {...node,args};
    }
    if(node.type==='call'){
      const def=this.definedNames.get(node.name.toUpperCase())??this.workbookDefinedNames.get(node.name.toUpperCase());if(def?.ref.startsWith('=')){
        const key=node.name.toUpperCase();if(resolving.has(key))throw new FormulaSyntaxError(`Defined LAMBDA '${node.name}' is recursive.`,node.start,'Use a non-recursive LAMBDA definition.',node.end-node.start);
        const nested=new Set(resolving);nested.add(key);const lambda=this.resolveDefinedNames(parseFormula(def.ref),new Set(),nested);
        return {type:'call',name:'__INVOKE_LAMBDA__',args:[lambda,...node.args.map(arg=>this.resolveDefinedNames(arg,locals,resolving))],start:node.start,end:node.end};
      }
      return {...node,args:node.args.map(arg=>this.resolveDefinedNames(arg,locals,resolving))};
    }
    return node;
  }
  private clearDynamic(owner:string):void {
    for(const dep of this.dynamicDeps.get(owner)??[]) {
      this.rev.get(dep)?.delete(owner);
      if(this.rev.get(dep)?.size===0)this.rev.delete(dep);
    }
    this.dynamicDeps.delete(owner);
    for(const range of this.dynamicRanges.get(owner)??[]){const from=parseRef(range.from),to=parseRef(range.to);if(!from||!to)continue;for(let col=Math.min(from.col,to.col);col<=Math.max(from.col,to.col);col++){const key=`${range.sheet.toLocaleLowerCase()}!${col}`,bucket=this.dynamicRangeColumns.get(key);bucket?.delete(owner);if(bucket?.size===0)this.dynamicRangeColumns.delete(key);}}
    this.dynamicRanges.delete(owner);
  }
  private clearStaticRanges(owner:string):void {
    for(const range of this.staticRangeDeps.get(owner)??[]){
      const from=parseRef(range.from),to=parseRef(range.to);if(!from||!to)continue;
      for(let col=Math.min(from.col,to.col);col<=Math.max(from.col,to.col);col++){
        const key=`${this.sheetName.toLocaleLowerCase()}!${col}`,bucket=this.staticRangeColumns.get(key);bucket?.delete(owner);if(bucket?.size===0)this.staticRangeColumns.delete(key);
      }
    }
    this.staticRangeDeps.delete(owner);
  }
  private indexStaticRange(owner:string,range:LocalRangeDependency):void {
    if(!this.staticRangeDeps.has(owner))this.staticRangeDeps.set(owner,[]);this.staticRangeDeps.get(owner)!.push(range);
    const from=parseRef(range.from),to=parseRef(range.to);if(!from||!to)return;
    for(let col=Math.min(from.col,to.col);col<=Math.max(from.col,to.col);col++){
      const key=`${this.sheetName.toLocaleLowerCase()}!${col}`;if(!this.staticRangeColumns.has(key))this.staticRangeColumns.set(key,new Set());this.staticRangeColumns.get(key)!.add(owner);
    }
  }
  private staticRangeContains(owner:string,ref:string):boolean {
    const point=parseRef(ref);if(!point)return false;
    return (this.staticRangeDeps.get(owner)??[]).some(range=>{const from=parseRef(range.from),to=parseRef(range.to);return !!from&&!!to&&point.row>=Math.min(from.row,to.row)&&point.row<=Math.max(from.row,to.row)&&point.col>=Math.min(from.col,to.col)&&point.col<=Math.max(from.col,to.col);});
  }
  private indexDynamicRange(owner:string,sheet:string,fromRef:string,toRef:string):void {
    const from=parseRef(fromRef),to=parseRef(toRef);if(!from||!to)return;for(let col=Math.min(from.col,to.col);col<=Math.max(from.col,to.col);col++){const key=`${sheet.toLocaleLowerCase()}!${col}`;if(!this.dynamicRangeColumns.has(key))this.dynamicRangeColumns.set(key,new Set());this.dynamicRangeColumns.get(key)!.add(owner);}
  }
  private trackDynamic(owner:string,dependency:string):void {
    const ref=known(dependency,this.rows,this.cols);
    if(!this.dynamicDeps.has(owner))this.dynamicDeps.set(owner,new Set());
    if(this.deps.get(owner)?.has(ref)||this.staticRangeContains(owner,ref))return;
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
    if(!existing.some(item=>item.sheet.toLocaleLowerCase()===sheet.toLocaleLowerCase()&&item.from===fromRaw&&item.to===(toRaw??fromRaw))){existing.push({sheet,from:fromRaw!,to:toRaw??fromRaw!});this.indexDynamicRange(owner,sheet,fromRaw!,toRaw??fromRaw!);}
    this.dynamicRanges.set(owner,existing);
  }
  private trackDynamicRange(owner:string,sheet:string,range:string):void {
    const [fromRaw,toRaw]=range.replace(/\$/g,'').toUpperCase().split(':'),from=fromRaw?parseRef(fromRaw):null,to=parseRef(toRaw??fromRaw??'');
    if(!from||!to||from.row>=this.rows||to.row>=this.rows||from.col>=this.cols||to.col>=this.cols)error('#REF!','INDIRECT reference is outside the worksheet.');
    const existing=this.dynamicRanges.get(owner)??[];
    if(!existing.some(item=>item.sheet.toLocaleLowerCase()===sheet.toLocaleLowerCase()&&item.from===fromRaw&&item.to===(toRaw??fromRaw))){existing.push({sheet,from:fromRaw!,to:toRaw??fromRaw!});this.indexDynamicRange(owner,sheet,fromRaw!,toRaw??fromRaw!);}
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
      if(changed){
        for(const owner of this.dynamicRangeColumns.get(`${this.sheetName.toLocaleLowerCase()}!${changed.col}`)??[]){const ranges=this.dynamicRanges.get(owner)??[];if(ranges.some(range=>range.sheet.toLocaleLowerCase()===this.sheetName.toLocaleLowerCase()&&(()=>{const from=parseRef(range.from),to=parseRef(range.to);return !!from&&!!to&&changed.row>=Math.min(from.row,to.row)&&changed.row<=Math.max(from.row,to.row)&&changed.col>=Math.min(from.col,to.col)&&changed.col<=Math.max(from.col,to.col);})())&&!affected.has(owner)){affected.add(owner);queue.push(owner);}}
        for(const owner of this.staticRangeColumns.get(`${this.sheetName.toLocaleLowerCase()}!${changed.col}`)??[]){const ranges=this.staticRangeDeps.get(owner)??[];if(ranges.some(range=>{const from=parseRef(range.from),to=parseRef(range.to);return !!from&&!!to&&changed.row>=Math.min(from.row,to.row)&&changed.row<=Math.max(from.row,to.row)&&changed.col>=Math.min(from.col,to.col)&&changed.col<=Math.max(from.col,to.col);})&&!affected.has(owner)){affected.add(owner);queue.push(owner);}}
      }
    }
    for(const ref of affected) {this.cache.delete(ref);this.dirty.add(ref);}
    return [...affected];
  }
  applyEdits(edits: Edit[], recordHistory = true, recalculate = true): RecalcStats {
    const sanitized=new Map<string,string>();
    for(const e of edits) sanitized.set(known(e.ref,this.rows,this.cols),e.raw);
    const changed=[...sanitized].filter(([ref,raw])=>this.getRaw(ref)!==raw);
    if(!changed.length) return {...this.lastStats,updated:0,recalculated:0,affected:[],durationMs:0};
    const projected=this.raw.size+changed.filter(([ref,raw])=>raw!==''&&!this.raw.has(ref)).length-changed.filter(([ref,raw])=>raw===''&&this.raw.has(ref)).length;
    if(projected>MAX_WORKSHEET_POPULATED_CELLS)throw Error(`Worksheet exceeds GridForge's ${MAX_WORKSHEET_POPULATED_CELLS.toLocaleString()} populated-cell limit.`);
    this.valueRevisionCounter++;
    const blocked=changed.find(([ref])=>this.spillChildren.has(ref));if(blocked)throw Error(`Cannot edit ${blocked[0]} because it belongs to a spilled array. Clear or edit the formula in ${this.spillChildren.get(blocked[0])!.owner}.`);
    const before=changed.map(([ref])=>({ref,raw:this.getRaw(ref)})),spillAffected=changed.flatMap(([ref])=>this.clearSpill(ref));
    for(const [ref,raw] of changed)this.update(ref,raw);
    const affected=this.invalidate([...changed.map(([ref])=>ref),...spillAffected]);
    const t0=performance.now();let recalculated=0;
    if(recalculate)for(const ref of affected) if(this.raw.has(ref)) {this.getValue(ref);recalculated++;}
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
    const candidates=new Set<string>();for(const ref of targets){const cell=parseRef(ref);if(cell)for(const owner of this.dynamicRangeColumns.get(`${sheet.toLocaleLowerCase()}!${cell.col}`)??[])candidates.add(owner);}
    for(const owner of candidates){const ranges=this.dynamicRanges.get(owner)??[];if(ranges.some(range=>range.sheet.toLocaleLowerCase()===sheet.toLocaleLowerCase()&&[...targets].some(ref=>{const cell=parseRef(ref),from=parseRef(range.from),to=parseRef(range.to);return !!cell&&!!from&&!!to&&cell.row>=Math.min(from.row,to.row)&&cell.row<=Math.max(from.row,to.row)&&cell.col>=Math.min(from.col,to.col)&&cell.col<=Math.max(from.col,to.col);}))&&!owners.includes(owner))owners.push(owner);}
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
    for(const ref of refs){if(this.spillChildren.get(ref)?.owner===owner){this.spillChildren.delete(ref);this.unindexPopulatedCell(ref);}this.cache.delete(ref);this.dirty.add(ref);}
    return refs;
  }
  private invokeLambda(lambda:Node,args:{value:Datum;omitted:boolean}[],visiting:string[]):Datum{
    if(lambda.type!=='call'||lambda.name!=='LAMBDA'||lambda.args.length<2)return error('#VALUE!','A LAMBDA needs at least one parameter and a calculation.');
    if(this.localFormulaNames.length>=64)return error('#NUM!','LAMBDA nesting exceeds 64 levels.');
    const parameters=lambda.args.slice(0,-1);
    if(parameters.length>253||args.length>parameters.length)return error('#VALUE!','The LAMBDA argument count does not match its parameters.');
    const values=new Map<string,Datum>(),omitted=new Set<string>();
    for(let index=0;index<parameters.length;index++){
      const parameter=parameters[index]!;
      if(parameter.type!=='definedName'||parseRef(parameter.name))return error('#VALUE!','LAMBDA parameters must be valid names, not cell references.');
      const key=parameter.name.toUpperCase();
      if(values.has(key))return error('#VALUE!',`LAMBDA parameter '${parameter.name}' is repeated.`);
      const supplied=args[index];values.set(key,supplied?.value??null);if(!supplied||supplied.omitted)omitted.add(key);
    }
    const scope={owner:visiting.at(-1)??'',values,omitted};this.localFormulaNames.push(scope);
    try{return this.evaluate(lambda.args.at(-1)!,visiting);}finally{this.localFormulaNames.pop();}
  }
  private formulaMatrix(node:Node,visiting:string[]):Datum[][]{const value=this.evaluate(node,visiting);return Array.isArray(value)?value:[[value]];}
  private summaryAggregate(fn:Node,subset:Datum[][],total:Datum[][],visiting:string[]):Value{
    let result:Datum;
    if(fn.type==='call'&&fn.name==='LAMBDA'){
      const parameterCount=fn.args.length-1;if(parameterCount<1||parameterCount>2)return error('#VALUE!','Summary LAMBDA aggregations must accept one values argument or a values/total pair.');
      result=this.invokeLambda(fn,parameterCount===1?[{value:subset,omitted:false}]:[{value:subset,omitted:false},{value:total,omitted:false}],visiting);
    }
    else if(fn.type==='definedName'){
      const definition=functions[fn.name.toUpperCase()];if(!definition)return error('#NAME?',`Unsupported summary function '${fn.name}'.`);
      result=definition.run(fn.name.toUpperCase()==='PERCENTOF'?[subset,total]:[subset]);
    }else return error('#VALUE!','GROUPBY and PIVOTBY need an aggregate function name or LAMBDA.');
    return Array.isArray(result)?error('#CALC!','GROUPBY and PIVOTBY aggregation must return one value.'):result;
  }
  private groupingKey(values:Datum[]):string{return JSON.stringify(values.map(value=>Array.isArray(value)?value:typeof value==='string'?['text',value.toLocaleLowerCase()]:value===null?['blank']:value));}
  private compareSummaryValue(left:Datum,right:Datum):number{
    const a=Array.isArray(left)?left[0]?.[0]??null:left,b=Array.isArray(right)?right[0]?.[0]??null:right;
    if(typeof a==='number'&&typeof b==='number')return a<b?-1:a>b?1:0;
    if(a===null||a==='')return b===null||b===''?0:-1;if(b===null||b==='')return 1;
    return String(a).toLocaleLowerCase().localeCompare(String(b).toLocaleLowerCase());
  }
  private evaluateGroupFormula(node:Extract<Node,{type:'call'}>,visiting:string[]):Datum[][]{
    const pivot=node.name==='PIVOTBY',minimum=pivot?4:3,maximum=pivot?11:8;
    if(node.args.length<minimum||node.args.length>maximum)return error('#ARG!',`${node.name} has an invalid number of arguments.`);
    const optional=(argument:Node|undefined):Node|undefined=>argument?.type==='literal'&&argument.value===null&&argument.start===argument.end?undefined:argument;
    const rowFields=this.formulaMatrix(node.args[0]!,visiting),columnFields=pivot?this.formulaMatrix(node.args[1]!,visiting):null;
    const values=this.formulaMatrix(node.args[pivot?2:1]!,visiting),fn=node.args[pivot?3:2]!,fieldHeaders=optional(node.args[pivot?4:3]),rowDepthNode=optional(node.args[pivot?5:4]),rowSortNode=optional(node.args[pivot?6:5]),columnDepthNode=pivot?optional(node.args[7]):undefined,columnSortNode=pivot?optional(node.args[8]):undefined,filterNode=optional(node.args[pivot?9:6]),relationshipNode=pivot?undefined:optional(node.args[7]),relativeNode=pivot?optional(node.args[10]):undefined;
    const rowCount=rowFields.length,valueCount=values.length;
    if(!rowCount||values.some(row=>row.length!==values[0]!.length)||rowFields.some(row=>row.length!==rowFields[0]!.length)||rowCount!==valueCount||columnFields&&(columnFields.length!==rowCount||columnFields.some(row=>row.length!==columnFields[0]!.length)))return error('#VALUE!','Grouping fields and values must have matching rectangular dimensions.');
    const rowFieldWidth=rowFields[0]!.length,valueWidth=values[0]!.length,columnFieldWidth=columnFields?.[0]?.length??0;
    if(pivot&&(rowFieldWidth>8||columnFieldWidth>8))return error('#VALUE!','PIVOTBY supports up to eight row fields and eight column fields in this subset.');
    if(!pivot&&rowFieldWidth>8)return error('#VALUE!','GROUPBY supports up to eight row-field levels.');
    if(relationshipNode&&number(this.evaluateScalar(relationshipNode,visiting))!==0)return error('#VALUE!','GROUPBY Table field relationships are not supported; use the default Hierarchy relationship.');
    let hasHeaders=false,showHeaders=false;
    if(fieldHeaders){const flag=number(this.evaluateScalar(fieldHeaders,visiting));if(!Number.isInteger(flag)||flag<0||flag>3)return error('#VALUE!','field_headers must be 0, 1, 2, or 3.');hasHeaders=flag===1||flag===3;showHeaders=flag===2||flag===3;}
    else hasHeaders=typeof values[0]?.[0]==='string'&&typeof values[1]?.[0]==='number';
    const start=hasHeaders?1:0;if(rowCount<=start)return error('#VALUE!','Summary inputs do not contain data rows.');
    const filters=filterNode?this.formulaMatrix(filterNode,visiting):null;
    if(filters&&(filters.length!==rowCount||filters.some(row=>row.length!==1)))return error('#VALUE!','filter_array must have one Boolean value per source row.');
    const validIndexes:number[]=[];for(let index=start;index<rowCount;index++){const include=filters?.[index]?.[0];if(include===undefined||bool(include as Value))validIndexes.push(index);}
    if(!validIndexes.length)return error('#CALC!','No rows remain after applying filter_array.');
    const groupMap=new Map<string,{row:Datum[];column:Datum[];indexes:number[]}>();
    for(const index of validIndexes){const row=rowFields[index]!,column=columnFields?.[index]??[],key=this.groupingKey([...row,...column]);let group=groupMap.get(key);if(!group){group={row:[...row],column:[...column],indexes:[]};groupMap.set(key,group);}group.indexes.push(index);}
    const relative=relativeNode?number(this.evaluateScalar(relativeNode,visiting)):0;
    if(pivot&&(!Number.isInteger(relative)||relative<0||relative>4))return error('#VALUE!','PIVOTBY relative_to must be 0 (column), 1 (row), 2 (grand), 3 (parent column), or 4 (parent row) total.');
    const groups=[...groupMap.values()];
    const resolveDepth=(nodeDepth:Node|undefined)=>nodeDepth?number(this.evaluateScalar(nodeDepth,visiting)):1;
    const rowDepth=resolveDepth(rowDepthNode),columnDepth=pivot?resolveDepth(columnDepthNode):0;
    const rowDepthLimit=rowFieldWidth,columnDepthLimit=pivot?columnFieldWidth:0;
    if(!Number.isInteger(rowDepth)||Math.abs(rowDepth)>rowDepthLimit||pivot&&(!Number.isInteger(columnDepth)||Math.abs(columnDepth)>columnDepthLimit))return error('#VALUE!','Subtotal depth must not exceed the available row or column field levels.');
    const sortGroups=(items:typeof groups,sortNode:Node|undefined,fieldWidth:number,subtotalLevels=0)=>{
      const sortResult=sortNode?this.evaluate(sortNode,visiting):1;
      let vectorSort=false,sorts:number[];
      if(Array.isArray(sortResult)){
        vectorSort=true;const matrix=sortResult as Datum[][];
        if(matrix.length===1)sorts=(matrix[0]??[]).map(value=>number(value as Value));
        else if(matrix.every(row=>row.length===1))sorts=matrix.map(row=>number(row[0] as Value));
        else throw new SheetError('#VALUE!','sort_order arrays must be a one-dimensional row or column vector.');
        if(!sorts.length||sorts.length>fieldWidth||sorts.some(value=>!Number.isInteger(value)||Math.abs(value)>fieldWidth))throw new SheetError('#VALUE!','Vector sort_order must list whole-number row/column field positions only.');
      }else sorts=[number(sortResult as Value)];
      if(sorts.some(value=>!Number.isInteger(value)))throw new SheetError('#VALUE!','sort_order must contain whole numbers.');
      if(subtotalLevels){
        const directions=Array.from({length:fieldWidth},(_,index)=>vectorSort?sorts[index]??1:index===0&&sorts[0]!==0?Math.sign(sorts[0]!):1);
        if(vectorSort&&sorts.some((sort,index)=>Math.abs(sort)!==index+1))throw new SheetError('#VALUE!','Nested subtotals need hierarchical sort vectors starting with field 1 and continuing in field order.');
        if(!vectorSort&&Math.abs(sorts[0]??0)>1)throw new SheetError('#VALUE!','Nested subtotals require sorting by the outer field first; use a field-order vector for nested sorting.');
        return items.sort((left,right)=>{for(let index=0;index<fieldWidth;index++){const comparison=this.compareSummaryValue(left.row[index]??null,right.row[index]??null);if(comparison)return directions[index]!<0?-comparison:comparison;}return this.groupingKey(left.row).localeCompare(this.groupingKey(right.row));});
      }
      if(!sorts.some(Boolean))return items;
      for(const sort of sorts){const position=Math.abs(sort);if(position>(vectorSort?fieldWidth:fieldWidth+valueWidth))throw new SheetError('#VALUE!','sort_order references an unavailable field or value column.');}
      return items.sort((left,right)=>{for(const sort of sorts){const position=Math.abs(sort);if(!position)continue;let comparison:number;if(position<=fieldWidth)comparison=this.compareSummaryValue(left.row[position-1]??left.column[position-1]??null,right.row[position-1]??right.column[position-1]??null);else{const measure=position-fieldWidth-1;comparison=this.compareSummaryValue(this.summaryAggregate(fn,left.indexes.map(index=>[values[index]![measure]??null]),validIndexes.map(index=>[values[index]![measure]??null]),visiting),this.summaryAggregate(fn,right.indexes.map(index=>[values[index]![measure]??null]),validIndexes.map(index=>[values[index]![measure]??null]),visiting));}if(comparison)return sort<0?-comparison:comparison;}return this.groupingKey(left.row).localeCompare(this.groupingKey(right.row));});
    };
    try{
      if(!pivot){
        const subtotalLevels=Math.max(0,Math.abs(rowDepth)-1),sortedRows=sortGroups(groups,rowSortNode,rowFieldWidth,subtotalLevels),tupleEqual=(left:Datum[],right:Datum[])=>this.groupingKey(left)===this.groupingKey(right),prefixMatches=(row:Datum[],prefix:Datum[])=>this.groupingKey(row.slice(0,prefix.length))===this.groupingKey(prefix);
        const subtotalCount=Array.from({length:subtotalLevels},(_,index)=>index+1).reduce((total,level)=>total+new Set(sortedRows.map(group=>this.groupingKey(group.row.slice(0,level)))).size,0),outputWidth=rowFieldWidth+valueWidth,outputHeight=sortedRows.length+subtotalCount+(rowDepth!==0?1:0)+(showHeaders?1:0);
        if(outputHeight*outputWidth>20000)return error('#NUM!','Summary output exceeds 20,000 cells. Reduce the number of groups or subtotal levels.');
        const aggregate=(indexes:number[],column:number)=>this.summaryAggregate(fn,indexes.map(index=>[values[index]![column]??null]),validIndexes.map(index=>[values[index]![column]??null]),visiting);
        const output:Datum[][]=[];if(showHeaders)output.push([...Array.from({length:rowFieldWidth},(_,index)=>hasHeaders?rowFields[0]![index]??`Row ${index+1}`:`Row ${index+1}`),...Array.from({length:valueWidth},(_,index)=>hasHeaders?values[0]![index]??`Value ${index+1}`:`Value ${index+1}`)]);
        const grandTotal:Datum[]=['Grand Total',...Array.from({length:rowFieldWidth-1},()=>null),...Array.from({length:valueWidth},(_,column)=>aggregate(validIndexes,column))];
        const subtotal=(prefix:Datum[]):Datum[]=>[...prefix,'Subtotal',...Array.from({length:Math.max(0,rowFieldWidth-prefix.length-1)},()=>null),...Array.from({length:valueWidth},(_,column)=>aggregate(sortedRows.filter(group=>prefixMatches(group.row,prefix)).flatMap(group=>group.indexes),column))];
        if(rowDepth<0)output.push(grandTotal);
        for(let index=0;index<sortedRows.length;index++){
          const group=sortedRows[index]!;
          if(rowDepth<0)for(let level=1;level<=subtotalLevels;level++)if(index===0||!tupleEqual(sortedRows[index-1]!.row.slice(0,level),group.row.slice(0,level)))output.push(subtotal(group.row.slice(0,level)));
          output.push([...group.row,...Array.from({length:valueWidth},(_,column)=>aggregate(group.indexes,column))]);
          if(rowDepth>0)for(let level=subtotalLevels;level>=1;level--)if(index===sortedRows.length-1||!tupleEqual(group.row.slice(0,level),sortedRows[index+1]!.row.slice(0,level)))output.push(subtotal(group.row.slice(0,level)));
        }
        if(rowDepth>0)output.push(grandTotal);
        return output;
      }
      const rowGroups=[...new Map(groups.map(group=>[this.groupingKey(group.row),group.row])).entries()].map(([key,row])=>({key,row,indexes:groups.filter(group=>this.groupingKey(group.row)===key).flatMap(group=>group.indexes)}));
      const columnGroups=[...new Map(groups.map(group=>[this.groupingKey(group.column),group.column])).entries()].map(([key,column])=>({key,column,indexes:groups.filter(group=>this.groupingKey(group.column)===key).flatMap(group=>group.indexes)}));
      const rowSubtotalLevels=Math.max(0,Math.abs(rowDepth)-1),columnSubtotalLevels=Math.max(0,Math.abs(columnDepth)-1);
      const sortedRowGroups=sortGroups(rowGroups.map(group=>({row:group.row,column:[],indexes:group.indexes})),rowSortNode,rowFieldWidth,rowSubtotalLevels),sortedColumnGroups=sortGroups(columnGroups.map(group=>({row:group.column,column:[],indexes:group.indexes})),columnSortNode,columnFieldWidth,columnSubtotalLevels);
      const sameTuple=(left:Datum[],right:Datum[])=>this.groupingKey(left)===this.groupingKey(right),matchesPrefix=(value:Datum[],prefix:Datum[])=>this.groupingKey(value.slice(0,prefix.length))===this.groupingKey(prefix);
      const prefixKey=(prefix:Datum[])=>`${prefix.length}:${this.groupingKey(prefix)}`,rowPrefixGroups=new Map<string,typeof groups>(),columnPrefixGroups=new Map<string,typeof groups>();
      for(const group of groups){for(let level=0;level<=rowFieldWidth;level++){const key=prefixKey(group.row.slice(0,level)),items=rowPrefixGroups.get(key)??[];items.push(group);rowPrefixGroups.set(key,items);}for(let level=0;level<=columnFieldWidth;level++){const key=prefixKey(group.column.slice(0,level)),items=columnPrefixGroups.get(key)??[];items.push(group);columnPrefixGroups.set(key,items);}}
      const sourceGroups=(rowPrefix:Datum[]|null,columnPrefix?:Datum[])=>{const rowCandidates=rowPrefix===null?groups:rowPrefixGroups.get(prefixKey(rowPrefix))??[],columnCandidates=columnPrefix===undefined?groups:columnPrefixGroups.get(prefixKey(columnPrefix))??[],scanRows=rowCandidates.length<=columnCandidates.length,candidates=scanRows?rowCandidates:columnCandidates;return candidates.filter(group=>(rowPrefix===null||matchesPrefix(group.row,rowPrefix))&&(columnPrefix===undefined||matchesPrefix(group.column,columnPrefix)));};
      const sourceIndexes=(rowPrefix:Datum[]|null,columnPrefix?:Datum[])=>sourceGroups(rowPrefix,columnPrefix).flatMap(group=>group.indexes);
      type ColumnSpec={total:boolean;subtotal:boolean;row:Datum[];indexes:number[]};
      const columnSpecs:ColumnSpec[]=[];
      const columnSubtotal=(prefix:Datum[]):ColumnSpec=>({total:false,subtotal:true,row:prefix,indexes:sourceIndexes(null,prefix)});
      if(columnDepth<0)columnSpecs.push({total:true,subtotal:false,row:[],indexes:validIndexes});
      for(let index=0;index<sortedColumnGroups.length;index++){
        const group=sortedColumnGroups[index]!;
        if(columnDepth<0)for(let level=1;level<=columnSubtotalLevels;level++)if(index===0||!sameTuple(sortedColumnGroups[index-1]!.row.slice(0,level),group.row.slice(0,level)))columnSpecs.push(columnSubtotal(group.row.slice(0,level)));
        columnSpecs.push({total:false,subtotal:false,row:group.row,indexes:group.indexes});
        if(columnDepth>0)for(let level=columnSubtotalLevels;level>=1;level--)if(index===sortedColumnGroups.length-1||!sameTuple(group.row.slice(0,level),sortedColumnGroups[index+1]!.row.slice(0,level)))columnSpecs.push(columnSubtotal(group.row.slice(0,level)));
      }
      if(columnDepth>0)columnSpecs.push({total:true,subtotal:false,row:[],indexes:validIndexes});
      const subtotalRowCount=(levels:number,items:typeof sortedRowGroups):number=>{let count=0;for(let level=1;level<=levels;level++)count+=new Set(items.map(item=>this.groupingKey(item.row.slice(0,level)))).size;return count;};
      const outputWidth=rowFieldWidth+columnSpecs.length*valueWidth,outputHeight=Math.max(1,columnFieldWidth)+(showHeaders?1+(valueWidth>1?1:0):0)+sortedRowGroups.length+subtotalRowCount(rowSubtotalLevels,sortedRowGroups)+(rowDepth!==0?1:0);
      if(outputHeight*outputWidth>20000)return error('#NUM!','Summary output exceeds 20,000 cells. Reduce the number of groups, measures, or subtotal levels.');
      const measureLabel=(measure:number):Datum=>hasHeaders?values[0]![measure]??`Value ${measure+1}`:`Value ${measure+1}`;
      const rowFieldLabels:Datum[]=Array.from({length:rowFieldWidth},(_,index)=>hasHeaders?rowFields[0]![index]??`Row ${index+1}`:`Row ${index+1}`),columnFieldLabels:Datum[]=Array.from({length:columnFieldWidth},(_,index)=>hasHeaders?columnFields![0]![index]??`Column ${index+1}`:`Column ${index+1}`);
      const hierarchyHeadings:Datum[][]=Array.from({length:Math.max(1,columnFieldWidth)},(_,level)=>[...Array.from({length:rowFieldWidth},()=>null),...columnSpecs.flatMap(spec=>Array.from({length:valueWidth},()=>spec.total?(level===0?'Grand Total':null):spec.subtotal?(level<spec.row.length?spec.row[level]??null:level===spec.row.length?'Subtotal':null):spec.row[level]??null))]);
      const measureHeading:Datum[]=[...(rowFieldWidth?[hasHeaders?'Values':'Values',...Array.from({length:rowFieldWidth-1},()=>null)]:[]),...columnSpecs.flatMap(()=>Array.from({length:valueWidth},(_,measure)=>measureLabel(measure)))];
      const parentColumnIndexes=(row:Datum[]|null,column:Datum[])=>column.length?sourceIndexes(row,column.slice(0,-1)):validIndexes;
      const parentRowIndexes=(row:Datum[],column?:Datum[])=>sourceIndexes(row.slice(0,-1),column);
      const aggregateRow=(prefix:Datum[],subtotal:boolean,label?:Datum):Datum[]=>{
        const row:Datum[]=subtotal?[...prefix,label??'Subtotal',...Array.from({length:Math.max(0,rowFieldWidth-prefix.length-1)},()=>null)]:[...prefix];
        for(const columnSpec of columnSpecs){
          const indices=sourceIndexes(prefix,columnSpec.total?undefined:columnSpec.row);
          for(let measure=0;measure<valueWidth;measure++){
            const relativeIndexes=relative===0?(columnSpec.total?validIndexes:columnSpec.indexes):relative===1?sourceIndexes(prefix):relative===2?validIndexes:relative===3?parentColumnIndexes(prefix,columnSpec.row):parentRowIndexes(prefix,columnSpec.total?undefined:columnSpec.row);
            row.push(this.summaryAggregate(fn,indices.map(index=>[values[index]![measure]??null]),relativeIndexes.map(index=>[values[index]![measure]??null]),visiting));
          }
        }
        return row;
      };
      const body:Datum[][]=[];
      for(let index=0;index<sortedRowGroups.length;index++){
        const group=sortedRowGroups[index]!;
        if(rowDepth<0)for(let level=1;level<=rowSubtotalLevels;level++)if(index===0||!sameTuple(sortedRowGroups[index-1]!.row.slice(0,level),group.row.slice(0,level)))body.push(aggregateRow(group.row.slice(0,level),true));
        body.push(aggregateRow(group.row,false));
        if(rowDepth>0)for(let level=rowSubtotalLevels;level>=1;level--)if(index===sortedRowGroups.length-1||!sameTuple(group.row.slice(0,level),sortedRowGroups[index+1]!.row.slice(0,level)))body.push(aggregateRow(group.row.slice(0,level),true));
      }
      const totalRow=aggregateRow([],true,'Grand Total');
      const output:Datum[][]=[];if(showHeaders){const labels:Datum[]=[...rowFieldLabels,...columnSpecs.flatMap(spec=>Array.from({length:valueWidth},()=>spec.total?'Grand Total':spec.subtotal?`${columnFieldLabels.slice(0,spec.row.length).join(' / ')} Subtotal`:columnFieldLabels.join(' / ')))];output.push(labels);if(valueWidth>1)output.push(measureHeading);}output.push(...hierarchyHeadings);if(rowDepth<0)output.push(totalRow);output.push(...body);if(rowDepth>0)output.push(totalRow);
      if(output.length*Math.max(1,...output.map(row=>row.length))>20000)return error('#NUM!','Summary output exceeds 20,000 cells.');return output;
    }catch(cause){if(cause instanceof SheetError)throw cause;return error('#VALUE!',cause instanceof Error?cause.message:'Invalid summary options.');}
  }
  private placeSpill(owner:string,value:Datum[][],ast:Node):CellResult {
    const rows=value.length,cols=Math.max(0,...value.map(row=>row.length));
    if(!rows||!cols)return {value:null,error:diag('#CALC!','Array formula returned an empty result.')};
    if(rows*cols>MAX_SPILL_CELLS)return {value:null,error:diag('#NUM!',`Array result exceeds the ${MAX_SPILL_CELLS.toLocaleString()}-cell spill limit.`)};
    const origin=parseRef(owner)!;
    if(origin.row+rows>this.rows||origin.col+cols>this.cols)return {value:null,error:diag('#SPILL!','The array result extends beyond the worksheet boundary.')};
    const refs:string[]=[];
    for(let r=0;r<rows;r++)for(let c=0;c<cols;c++){
      const ref=`${colToName(origin.col+c)}${origin.row+r+1}`;
      if(ref!==owner){
        if(this.raw.has(ref)||this.spillChildren.has(ref)){this.trackDynamic(owner,ref);return {value:null,error:diag('#SPILL!',`The spill range is blocked at ${ref}. Clear that cell and recalculate.`)};}
        if(this.deps.get(owner)?.has(ref)||this.staticRangeContains(owner,ref)||this.dynamicDeps.get(owner)?.has(ref))return {value:null,error:diag('#CYCLE!',`Array formula in ${owner} depends on its own spill cell ${ref}.`)};
      }
      refs.push(ref);
    }
    const children:string[]=[];
    for(let r=0;r<rows;r++)for(let c=0;c<cols;c++){
      const ref=`${colToName(origin.col+c)}${origin.row+r+1}`;if(ref===owner)continue;
      const cell=value[r]![c]??null;this.spillChildren.set(ref,{owner,value:cell as Value});this.indexPopulatedCell(ref);children.push(ref);
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
  private binaryValue(op:string,left:Value,right:Value):Value {
    if(op==='&')return text(left)+text(right);
    if(['=','<>','!=','<','>','<=','>='].includes(op)){
      const comparable=typeof left==='number'&&typeof right==='number'?(left===right?0:left<right?-1:1):text(left).toLowerCase().localeCompare(text(right).toLowerCase());
      switch(op){case '=':return comparable===0;case '<>':case '!=':return comparable!==0;case '<':return comparable<0;case '>':return comparable>0;case '<=':return comparable<=0;default:return comparable>=0;}
    }
    const a=number(left),b=number(right);let value:number;
    switch(op){case '+':value=a+b;break;case '-':value=a-b;break;case '*':value=a*b;break;case '/':if(b===0)error('#DIV/0!','Division by zero.');value=a/b;break;case '^':value=a**b;break;default:return error('#PARSE!',`Unsupported operator ${op}.`);}
    if(!Number.isFinite(value))error('#NUM!','Arithmetic result is not finite.');return value;
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
        const value=this.evaluate(n.expr,visiting);if(!Array.isArray(value))return value;
        // Excel's @ operator intersects a range against the formula cell, but
        // for an already-computed array it returns the top-left array item.
        // Keep this separate from the range branch above: computed arrays have
        // no worksheet origin from which a row/column intersection can be
        // calculated.
        return value[0]?.[0]??null;
      }
      const value=this.evaluate(n.expr,visiting),apply=(item:Value):Value=>{const scalar=number(item);return n.op==='-'?-scalar:n.op==='%'?scalar/100:scalar;};
      if(Array.isArray(value)){const matrix=value as Value[][];if(matrix.length*Math.max(0,...matrix.map(row=>row.length))>MAX_SPILL_CELLS)return error('#NUM!',`Array operation exceeds the ${MAX_SPILL_CELLS.toLocaleString()}-cell limit.`);return matrix.map(row=>row.map(item=>apply(item??null)));}
      return apply(value as Value);
    }
    if(n.type==='binary') {
      const rangeCells=(node:Node):number|undefined=>{if(node.type!=='range')return;const from=parseRef(node.from),to=parseRef(node.to);return from&&to?(Math.abs(to.row-from.row)+1)*(Math.abs(to.col-from.col)+1):undefined;};
      for(const node of [n.left,n.right]){const cells=rangeCells(node);if(cells!==undefined&&cells>MAX_SPILL_CELLS)return error('#NUM!',`Array operation exceeds the ${MAX_SPILL_CELLS.toLocaleString()}-cell limit.`);}
      const left=this.evaluate(n.left,visiting),right=this.evaluate(n.right,visiting),leftArray=Array.isArray(left),rightArray=Array.isArray(right);
      if(!leftArray&&!rightArray)return this.binaryValue(n.op,left as Value,right as Value);
      const leftMatrix=leftArray?left as Value[][]:[[left as Value]],rightMatrix=rightArray?right as Value[][]:[[right as Value]],leftRows=leftMatrix.length,leftCols=Math.max(0,...leftMatrix.map(row=>row.length)),rightRows=rightMatrix.length,rightCols=Math.max(0,...rightMatrix.map(row=>row.length));
      const rows=leftArray?leftRows:rightRows,cols=leftArray?leftCols:rightCols;
      if(leftArray&&rightArray&&(leftRows!==rightRows||leftCols!==rightCols))return error('#VALUE!','Array operator ranges must have matching row and column counts.');
      if(rows*cols>MAX_SPILL_CELLS)return error('#NUM!',`Array operation exceeds the ${MAX_SPILL_CELLS.toLocaleString()}-cell limit.`);
      return Array.from({length:rows},(_,row)=>Array.from({length:cols},(_,col)=>this.binaryValue(n.op,leftArray?(leftMatrix[row]?.[col]??null):(leftMatrix[0]?.[0]??null),rightArray?(rightMatrix[row]?.[col]??null):(rightMatrix[0]?.[0]??null))));
    }
    if(n.type==='call') {
      if(n.name==='__INVOKE_LAMBDA__'){
        const lambda=n.args[0]!;
        const args=n.args.slice(1).map(argument=>({omitted:argument.type==='literal'&&argument.value===null&&argument.start===argument.end,value:argument.type==='literal'&&argument.value===null&&argument.start===argument.end?null:this.evaluate(argument,visiting)}));
        return this.invokeLambda(lambda,args,visiting);
      }
      if(['SUM','AVERAGE','MIN','MAX','PRODUCT','STDEV','VAR','STDEV.S','VAR.S','STDEV.P','VAR.P'].includes(n.name)&&n.args.length>=1&&n.args.length<=255&&n.args.every(argument=>(argument.type==='ref'||argument.type==='range')&&(!argument.sheet||argument.sheet.toLocaleLowerCase()===this.sheetName.toLocaleLowerCase()))){
        let total=0,count=0,extreme:number|undefined;const values:number[]=[];
        for(const argument of n.args){
          if(argument.type!=='ref'&&argument.type!=='range')return error('#VALUE!',`${n.name} sparse aggregation needs cell or range references.`);
          const first=parseRef(argument.type==='range'?argument.from:argument.ref),last=parseRef(argument.type==='range'?argument.to:argument.ref);if(!first||!last)return error('#REF!',`${n.name} contains an invalid reference.`);
          const top=Math.min(first.row,last.row),bottom=Math.max(first.row,last.row),left=Math.min(first.col,last.col),right=Math.max(first.col,last.col);if(bottom>=this.rows||right>=this.cols)return error('#REF!',`${n.name} reference is outside worksheet boundaries.`);
          for(let col=left;col<=right;col++)for(const row of this.populatedRowsByColumn.get(col)??[]){if(row<top||row>bottom)continue;const ref=`${colToName(col)}${row+1}`,result=this.getValue(ref,visiting);if(result.error)return error(result.error.code,`${ref}: ${result.error.message}`);if(typeof result.value!=='number'||!Number.isFinite(result.value))continue;total+=result.value;count++;values.push(result.value);extreme=extreme===undefined?result.value:n.name==='MIN'?Math.min(extreme,result.value):Math.max(extreme,result.value);}
        }
        if(n.name==='SUM')return Number.isFinite(total)?total:error('#NUM!','The result is not a finite number.');
        if(n.name==='AVERAGE')return count?Number.isFinite(total/count)?total/count:error('#NUM!','The result is not a finite number.'):error('#DIV/0!','AVERAGE has no numeric values.');
        if(n.name==='PRODUCT'||['STDEV','VAR','STDEV.S','VAR.S','STDEV.P','VAR.P'].includes(n.name))return functions[n.name]!.run([[values]]);
        return extreme??0;
      }
      if((n.name==='MEDIAN'&&n.args.length>=1&&n.args.length<=255&&n.args.every(argument=>(argument.type==='ref'||argument.type==='range')&&(!argument.sheet||argument.sheet.toLocaleLowerCase()===this.sheetName.toLocaleLowerCase())))||(['LARGE','SMALL'].includes(n.name)&&n.args.length===2&&(n.args[0]!.type==='ref'||n.args[0]!.type==='range')&&(!n.args[0]!.sheet||n.args[0]!.sheet.toLocaleLowerCase()===this.sheetName.toLocaleLowerCase())&&n.args[1]!.type!=='range'&&!(n.args[1]!.type==='ref'&&n.args[1]!.sheet&&n.args[1]!.sheet.toLocaleLowerCase()!==this.sheetName.toLocaleLowerCase()))){
        const ranges=n.name==='MEDIAN'?n.args:n.args.slice(0,1),values:number[]=[];
        for(const argument of ranges){
          if(argument.type!=='ref'&&argument.type!=='range')return error('#VALUE!',`${n.name} sparse aggregation needs a cell or range reference.`);
          const first=parseRef(argument.type==='range'?argument.from:argument.ref),last=parseRef(argument.type==='range'?argument.to:argument.ref);if(!first||!last)return error('#REF!',`${n.name} contains an invalid reference.`);
          const top=Math.min(first.row,last.row),bottom=Math.max(first.row,last.row),left=Math.min(first.col,last.col),right=Math.max(first.col,last.col);if(bottom>=this.rows||right>=this.cols)return error('#REF!',`${n.name} reference is outside worksheet boundaries.`);
          for(let col=left;col<=right;col++)for(const row of this.populatedRowsByColumn.get(col)??[]){if(row<top||row>bottom)continue;const ref=`${colToName(col)}${row+1}`,result=this.getValue(ref,visiting);if(result.error)return error(result.error.code,`${ref}: ${result.error.message}`);if(typeof result.value==='number'&&Number.isFinite(result.value))values.push(result.value);}
        }
        const matrix:Datum[][]=[values];if(n.name==='MEDIAN')return functions.MEDIAN!.run([matrix]);
        return functions[n.name]!.run([matrix,this.evaluate(n.args[1]!,visiting)]);
      }
      if(n.name==='ISOMITTED'){
        if(n.args.length!==1)return error('#ARG!','ISOMITTED expects one LAMBDA parameter.');
        const argument=n.args[0]!;if(argument.type!=='definedName')return error('#VALUE!','ISOMITTED expects a LAMBDA parameter name.');
        const key=argument.name.toUpperCase(),owner=visiting.at(-1)??'';
        for(let index=this.localFormulaNames.length-1;index>=0;index--){const scope=this.localFormulaNames[index]!;if(scope.owner===owner&&scope.values.has(key))return scope.omitted.has(key);}
        return error('#VALUE!','ISOMITTED can only inspect a LAMBDA parameter.');
      }
      const lambdaArgument=(node:Node):boolean=>node.type==='call'&&node.name==='LAMBDA';
      const invoke=(lambda:Node,nodes:Node[])=>this.invokeLambda(lambda,nodes.map(argument=>({omitted:argument.type==='literal'&&argument.value===null&&argument.start===argument.end,value:argument.type==='literal'&&argument.value===null&&argument.start===argument.end?null:this.evaluate(argument,visiting)})),visiting);
      const invokeValues=(lambda:Node,values:Datum[])=>this.invokeLambda(lambda,values.map(value=>({value,omitted:false})),visiting);
      const asMatrix=(value:Datum):Datum[][]=>Array.isArray(value)?value:[[value]];
      const scalarLambdaResult=(value:Datum):Value=>Array.isArray(value)?error('#CALC!','A lambda used by MAP, REDUCE, or SCAN cannot return a nested array.'):value;
      if(n.name==='GROUPBY'||n.name==='PIVOTBY')return this.evaluateGroupFormula(n,visiting);
      if(n.name==='__INVOKE_LAMBDA__'){
        const lambda=n.args[0];if(!lambda||lambda.type!=='call'||lambda.name!=='LAMBDA')return error('#NAME?','Named formula does not resolve to a LAMBDA.');
        return invoke(lambda,n.args.slice(1));
      }
      if(n.name==='LAMBDA')return error('#CALC!','A LAMBDA must be invoked or passed to a lambda-aware function.');
      if(n.name==='MAKEARRAY'&&n.args.length===3&&lambdaArgument(n.args[2]!)){
        const rowValue=number(this.evaluateScalar(n.args[0]!,visiting)),columnValue=number(this.evaluateScalar(n.args[1]!,visiting));
        if(!Number.isInteger(rowValue)||!Number.isInteger(columnValue)||rowValue<1||columnValue<1)return error('#VALUE!','MAKEARRAY dimensions must be positive whole numbers.');
        if(rowValue*columnValue>20000)return error('#NUM!','Array results are limited to 20,000 elements.');
        return Array.from({length:rowValue},(_,row)=>Array.from({length:columnValue},(_,column)=>scalarLambdaResult(invoke(n.args[2]!,[{type:'literal',value:row+1,start:n.start,end:n.end},{type:'literal',value:column+1,start:n.start,end:n.end}]))));
      }
      if((n.name==='BYROW'||n.name==='BYCOL')&&n.args.length===2&&lambdaArgument(n.args[1]!)){
        const source=asMatrix(this.evaluate(n.args[0]!,visiting)),height=source.length,width=Math.max(0,...source.map(row=>row.length));
        if(height*width>20000)return error('#NUM!','Array results are limited to 20,000 elements.');
        return n.name==='BYROW'?source.map(row=>[scalarLambdaResult(invokeValues(n.args[1]!,[[[...row]]]))]):[Array.from({length:width},(_,column)=>scalarLambdaResult(invokeValues(n.args[1]!,[[[...source.map(row=>row[column]??null)]]])) )];
      }
      if(n.name==='MAP'&&n.args.length>=2&&lambdaArgument(n.args.at(-1)!)){
        const lambda=n.args.at(-1)!,arrays=n.args.slice(0,-1).map(argument=>asMatrix(this.evaluate(argument,visiting))),height=arrays[0]!.length,width=Math.max(0,...arrays[0]!.map(row=>row.length));
        if(arrays.some(matrix=>matrix.length!==height||Math.max(0,...matrix.map(row=>row.length))!==width))return error('#VALUE!','All MAP arrays must have matching dimensions.');
        if(height*width>20000)return error('#NUM!','Array results are limited to 20,000 elements.');
        return Array.from({length:height},(_,row)=>Array.from({length:width},(_,column)=>scalarLambdaResult(invokeValues(lambda,arrays.map(matrix=>matrix[row]?.[column]??null)))));
      }
      if((n.name==='REDUCE'||n.name==='SCAN')&&n.args.length===3&&lambdaArgument(n.args[2]!)){
        let accumulator=this.evaluate(n.args[0]!,visiting);const source=asMatrix(this.evaluate(n.args[1]!,visiting)),height=source.length,width=Math.max(0,...source.map(row=>row.length));
        if(height*width>20000)return error('#NUM!','Array results are limited to 20,000 elements.');
        if(!height||!width)return n.name==='REDUCE'?accumulator:error('#CALC!','SCAN cannot return an empty array.');
        const result:Datum[][]=Array.from({length:height},()=>Array.from({length:width},()=>null));
        for(let row=0;row<height;row++)for(let column=0;column<width;column++){accumulator=scalarLambdaResult(invokeValues(n.args[2]!,[accumulator,source[row]?.[column]??null]));result[row]![column]=accumulator;}
        return n.name==='REDUCE'?accumulator:result;
      }
      if(n.name==='LET'){
        const scope:{owner:string;values:Map<string,Datum>;omitted:Set<string>}={owner:visiting.at(-1)??'',values:new Map(),omitted:new Set()};this.localFormulaNames.push(scope);
        try{for(let index=0;index<n.args.length-1;index+=2){const binding=n.args[index];if(!binding||binding.type!=='definedName')return error('#NAME?','LET variable name is invalid.');scope.values.set(binding.name.toUpperCase(),this.evaluate(n.args[index+1]!,visiting));}return this.evaluate(n.args.at(-1)!,visiting);}
        finally{this.localFormulaNames.pop();}
      }
      if(n.name==='INFO'&&n.args.length===1&&text(this.evaluateScalar(n.args[0]!,visiting)).trim().toLowerCase()==='numfile')return this.documentSheetNames.length||1;
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
      if((n.name==='SUMIF'||n.name==='AVERAGEIF')&&n.args.length===3){
        const criteriaNode=n.args[0]!,sumNode=n.args[2]!,isReference=(node:Node):node is Extract<Node,{type:'ref'|'range'}>=>node.type==='ref'||node.type==='range';
        if(isReference(criteriaNode)&&isReference(sumNode)){
          const criteriaFrom=parseRef(criteriaNode.type==='range'?criteriaNode.from:criteriaNode.ref),criteriaTo=parseRef(criteriaNode.type==='range'?criteriaNode.to:criteriaNode.ref),sumStart=parseRef(sumNode.type==='range'?sumNode.from:sumNode.ref);
          if(!criteriaFrom||!criteriaTo||!sumStart)return error('#REF!','Conditional aggregation contains an invalid reference.');
          const topRow=Math.min(criteriaFrom.row,criteriaTo.row),bottomRow=Math.max(criteriaFrom.row,criteriaTo.row),leftCol=Math.min(criteriaFrom.col,criteriaTo.col),rightCol=Math.max(criteriaFrom.col,criteriaTo.col),height=bottomRow-topRow+1,width=rightCol-leftCol+1;
          if(bottomRow>=this.rows||rightCol>=this.cols)return error('#REF!','Criteria range extends beyond worksheet boundaries.');
          if(sumStart.row+height>this.rows||sumStart.col+width>this.cols)return error('#REF!','Aligned aggregation range extends beyond worksheet boundaries.');
          const criteria=this.evaluate(criteriaNode,visiting),criterion=this.evaluate(n.args[1]!,visiting),aligned:Datum[][]=[];
          for(let row=0;row<height;row++){const line:Datum[]=[];for(let col=0;col<width;col++)line.push(this.evaluate({type:'ref',ref:`${colToName(sumStart.col+col)}${sumStart.row+row+1}`,...(sumNode.sheet?{sheet:sumNode.sheet}:{}),start:sumNode.start,end:sumNode.end},visiting));aligned.push(line);}
          return f.run([criteria,criterion,aligned]);
        }
      }
      if(n.name==='NPV'&&n.args.length>=2&&n.args.length<=255){
        const evaluated:Datum[]=[this.evaluate(n.args[0]!,visiting)];
        for(const argument of n.args.slice(1)){
          if(argument.type==='ref'||argument.type==='range'){
            const first=parseRef(argument.type==='range'?argument.from:argument.ref),last=parseRef(argument.type==='range'?argument.to:argument.ref);if(!first||!last||Math.max(first.row,last.row)>=this.rows||Math.max(first.col,last.col)>=this.cols)return error('#REF!','NPV reference is outside worksheet boundaries.');
            const numeric:number[]=[];for(let row=Math.min(first.row,last.row);row<=Math.max(first.row,last.row);row++)for(let col=Math.min(first.col,last.col);col<=Math.max(first.col,last.col);col++){
              const ref=`${colToName(col)}${row+1}`,result=argument.sheet?this.externalSheetResolver?.(argument.sheet,ref,visiting):this.getValue(ref,visiting);if(!result)return error('#REF!',`Worksheet '${argument.sheet}' does not exist.`);if(!result.error&&typeof result.value==='number'&&Number.isFinite(result.value))numeric.push(result.value);
            }
            // NPV ignores text, logical values, blanks, and errors inside a
            // reference; only numeric cells advance the cash-flow period.
            evaluated.push([numeric]);
          }else{
            let value:Datum;try{value=this.evaluate(argument,visiting);}catch(cause){if(cause instanceof SheetError)continue;throw cause;}if(Array.isArray(value))evaluated.push(value.map(row=>row.filter((cell):cell is number=>typeof cell==='number'&&Number.isFinite(cell))));else if(typeof value==='number'&&Number.isFinite(value))evaluated.push(value);
          }
        }
        return f.run(evaluated);
      }
      if(['COUNT','COUNTA','COUNTBLANK'].includes(n.name)&&n.args.length>=1&&n.args.length<=255){
        const isCellReference=(argument:Node):argument is Extract<Node,{type:'ref'|'range'}>=>argument.type==='ref'||argument.type==='range';
        const localReferences=n.args.every(argument=>isCellReference(argument)&&(!argument.sheet||argument.sheet.toLocaleLowerCase()===this.sheetName.toLocaleLowerCase()));
        if(localReferences){
          let total=0;
          for(const argument of n.args){
            if(!isCellReference(argument))return error('#INTERNAL!','COUNT range aggregation received a non-reference argument.');
            const first=parseRef(argument.type==='range'?argument.from:argument.ref),last=argument.type==='range'?parseRef(argument.to):first;if(!first||!last)return error('#REF!',`${n.name} contains an invalid reference.`);
            const top=Math.min(first.row,last.row),bottom=Math.max(first.row,last.row),left=Math.min(first.col,last.col),right=Math.max(first.col,last.col);if(bottom>=this.rows||right>=this.cols)return error('#REF!',`${n.name} reference is outside worksheet boundaries.`);
            const stored=new Set<string>();for(const ref of this.raw.keys()){const point=parseRef(ref);if(point&&point.row>=top&&point.row<=bottom&&point.col>=left&&point.col<=right)stored.add(ref);}for(const ref of this.spillChildren.keys()){const point=parseRef(ref);if(point&&point.row>=top&&point.row<=bottom&&point.col>=left&&point.col<=right)stored.add(ref);}
            if(n.name==='COUNTBLANK')total+=(bottom-top+1)*(right-left+1)-stored.size;
            for(const ref of stored){const result=this.getValue(ref,visiting);if(n.name==='COUNT'){if(!result.error&&typeof result.value==='number'&&Number.isFinite(result.value))total++;}else if(n.name==='COUNTA'){if(result.error||result.value!==null)total++;}else if(!result.error&&(result.value===null||result.value===''))total++;}
          }
          return total;
        }
        const evaluated:Datum[]=[];
        for(const argument of n.args){
          if(argument.type!=='ref'&&argument.type!=='range'){try{evaluated.push(this.evaluate(argument,visiting));}catch(cause){if(!(cause instanceof SheetError))throw cause;if(n.name!=='COUNT')evaluated.push('#ERROR.VALUE');}continue;}
          const first=parseRef(argument.type==='range'?argument.from:argument.ref),last=parseRef(argument.type==='range'?argument.to:argument.ref);if(!first||!last||Math.max(first.row,last.row)>=this.rows||Math.max(first.col,last.col)>=this.cols)return error('#REF!',`${n.name} reference is outside worksheet boundaries.`);
          const matrix:Datum[][]=[];for(let row=Math.min(first.row,last.row);row<=Math.max(first.row,last.row);row++){const line:Datum[]=[];for(let col=Math.min(first.col,last.col);col<=Math.max(first.col,last.col);col++){
            const ref=`${colToName(col)}${row+1}`,result=argument.sheet?this.externalSheetResolver?.(argument.sheet,ref,visiting):this.getValue(ref,visiting);if(!result)return error('#REF!',`Worksheet '${argument.sheet}' does not exist.`);
            // COUNTA and COUNTBLANK treat referenced errors as nonblank cells;
            // COUNT ignores them. A nonempty marker preserves the first two
            // functions' range semantics without passing an error to the scalar evaluator.
            line.push(result.error?(n.name==='COUNT'?null:'#ERROR.VALUE'):(result.value??null));
          }matrix.push(line);}evaluated.push(matrix);
        }
        return f.run(evaluated);
      }
      if(n.name==='COUNTIF'&&n.args.length===2){
        const range=n.args[0]!;if((range.type==='ref'||range.type==='range')&&(!range.sheet||range.sheet.toLocaleLowerCase()===this.sheetName.toLocaleLowerCase())){
          const criterion=this.evaluate(n.args[1]!,visiting);if(!Array.isArray(criterion)){
            const first=parseRef(range.type==='range'?range.from:range.ref),last=range.type==='range'?parseRef(range.to):first;if(!first||!last)return error('#REF!','COUNTIF contains an invalid reference.');
            const top=Math.min(first.row,last.row),bottom=Math.max(first.row,last.row),left=Math.min(first.col,last.col),right=Math.max(first.col,last.col);if(bottom>=this.rows||right>=this.cols)return error('#REF!','COUNTIF reference is outside worksheet boundaries.');
            const stored=new Set<string>();for(const ref of this.raw.keys()){const point=parseRef(ref);if(point&&point.row>=top&&point.row<=bottom&&point.col>=left&&point.col<=right)stored.add(ref);}for(const ref of this.spillChildren.keys()){const point=parseRef(ref);if(point&&point.row>=top&&point.row<=bottom&&point.col>=left&&point.col<=right)stored.add(ref);}
            const normalizedCriterion=criterion===null?0:criterion;let count=matchCriterion(null,normalizedCriterion)?(bottom-top+1)*(right-left+1)-stored.size:0;
            for(const ref of stored){const result=this.getValue(ref,visiting);if(result.error)return error(result.error.code,`${ref}: ${result.error.message}`);if(matchCriterion(result.value,normalizedCriterion))count++;}
            return count;
          }
        }
      }
      if(n.name==='COUNTIFS'&&n.args.length>=2&&n.args.length%2===0){
        const rangeNodes=n.args.filter((_,index)=>index%2===0);
        const localReferences=rangeNodes.every((range):range is Extract<Node,{type:'ref'|'range'}>=>(range.type==='ref'||range.type==='range')&&(!range.sheet||range.sheet.toLocaleLowerCase()===this.sheetName.toLocaleLowerCase()));
        if(localReferences){
          if(rangeNodes.length>127)return error('#ARG!','Excel supports at most 127 criteria range/criterion pairs.');
          const bounds:{top:number;left:number;height:number;width:number}[]=[];
          for(const range of rangeNodes){
            const first=parseRef(range.type==='range'?range.from:range.ref),last=range.type==='range'?parseRef(range.to):first;
            if(!first||!last)return error('#REF!','COUNTIFS contains an invalid reference.');
            const top=Math.min(first.row,last.row),bottom=Math.max(first.row,last.row),left=Math.min(first.col,last.col),right=Math.max(first.col,last.col);
            if(bottom>=this.rows||right>=this.cols)return error('#REF!','COUNTIFS reference is outside worksheet boundaries.');
            bounds.push({top,left,height:bottom-top+1,width:right-left+1});
          }
          const shape=bounds[0]!;
          if(bounds.some(current=>current.height!==shape.height||current.width!==shape.width))return error('#VALUE!','COUNTIFS criteria ranges must have the same dimensions.');
          for(const area of bounds)for(const ref of this.raw.keys()){
            const point=parseRef(ref);if(point&&point.row>=area.top&&point.row<area.top+area.height&&point.col>=area.left&&point.col<area.left+area.width){const result=this.getValue(ref,visiting);if(result.error)return error(result.error.code,`${ref}: ${result.error.message}`);}
          }
          const criteria:Value[]=[];
          for(let index=1;index<n.args.length;index+=2){const criterion=this.evaluateScalar(n.args[index]!,visiting);if(Array.isArray(criterion))return error('#VALUE!','COUNTIFS criteria must be single values.');criteria.push(criterion??0);}
          const populated=new Set<number>();
          for(let rangeIndex=0;rangeIndex<rangeNodes.length;rangeIndex++){
            const area=bounds[rangeIndex]!,range=rangeNodes[rangeIndex]!;
            for(const ref of this.raw.keys()){
              const point=parseRef(ref);if(!point||point.row<area.top||point.row>=area.top+area.height||point.col<area.left||point.col>=area.left+area.width)continue;
              populated.add((point.row-area.top)*shape.width+(point.col-area.left));
            }
            for(const ref of this.spillChildren.keys()){
              const point=parseRef(ref);if(!point||point.row<area.top||point.row>=area.top+area.height||point.col<area.left||point.col>=area.left+area.width)continue;
              populated.add((point.row-area.top)*shape.width+(point.col-area.left));
            }
          }
          const matches=(offset:number):boolean=>{
            for(let index=0;index<rangeNodes.length;index++){
              const area=bounds[index]!,range=rangeNodes[index]!,row=area.top+Math.floor(offset/shape.width),col=area.left+offset%shape.width,ref=`${colToName(col)}${row+1}`;
              const value=this.getValue(ref,visiting);if(value.error||!matchCriterion(value.value,criteria[index]!))return false;
            }
            return true;
          };
          let count=0;
          for(const offset of populated)if(matches(offset))count++;
          if(criteria.every(criterion=>matchCriterion(null,criterion)))count+=shape.height*shape.width-populated.size;
          return count;
        }
      }
      if(['SUMIFS','AVERAGEIFS','MAXIFS','MINIFS'].includes(n.name)&&n.args.length>=3&&n.args.length%2===1){
        const resultNode=n.args[0]!,rangeNodes=n.args.filter((_,index)=>index>=1&&index%2===1);
        const isLocalReference=(range:Node):range is Extract<Node,{type:'ref'|'range'}>=>(range.type==='ref'||range.type==='range')&&(!range.sheet||range.sheet.toLocaleLowerCase()===this.sheetName.toLocaleLowerCase());
        if(isLocalReference(resultNode)&&rangeNodes.every(isLocalReference)){
          if(rangeNodes.length>127)return error('#ARG!','Excel supports at most 127 criteria range/criterion pairs.');
          const getBounds=(range:Extract<Node,{type:'ref'|'range'}>,label:string):{top:number;left:number;height:number;width:number}=>{
            const first=parseRef(range.type==='range'?range.from:range.ref),last=range.type==='range'?parseRef(range.to):first;
            if(!first||!last)return error('#REF!',`${label} contains an invalid reference.`);
            const top=Math.min(first.row,last.row),bottom=Math.max(first.row,last.row),left=Math.min(first.col,last.col),right=Math.max(first.col,last.col);
            if(bottom>=this.rows||right>=this.cols)return error('#REF!',`${label} reference is outside worksheet boundaries.`);
            return {top,left,height:bottom-top+1,width:right-left+1};
          };
          const resultBounds=getBounds(resultNode,n.name);
          const criteriaBounds:{top:number;left:number;height:number;width:number}[]=[];
          for(const range of rangeNodes)criteriaBounds.push(getBounds(range,n.name));
          const shape=resultBounds;
          if(criteriaBounds.some(current=>current.height!==shape.height||current.width!==shape.width))return error('#VALUE!',`${n.name} result and criteria ranges must have the same dimensions.`);
          for(const area of [shape,...criteriaBounds])for(const ref of this.raw.keys()){
            const point=parseRef(ref);if(point&&point.row>=area.top&&point.row<area.top+area.height&&point.col>=area.left&&point.col<area.left+area.width){const result=this.getValue(ref,visiting);if(result.error)return error(result.error.code,`${ref}: ${result.error.message}`);}
          }
          const criteria:Value[]=[];
          for(let index=2;index<n.args.length;index+=2){const criterion=this.evaluateScalar(n.args[index]!,visiting);if(Array.isArray(criterion))return error('#VALUE!',`${n.name} criteria must be single values.`);criteria.push(criterion??0);}
          const resultRefs=new Set<string>();for(const ref of this.raw.keys())resultRefs.add(ref);for(const ref of this.spillChildren.keys())resultRefs.add(ref);
          let total=0,matched=0,extreme:number|undefined;
          for(const ref of resultRefs){
            const point=parseRef(ref);if(!point||point.row<shape.top||point.row>=shape.top+shape.height||point.col<shape.left||point.col>=shape.left+shape.width)continue;
            const offset=(point.row-shape.top)*shape.width+(point.col-shape.left),rowOffset=Math.floor(offset/shape.width),colOffset=offset%shape.width;
            let satisfies=true;
            for(let index=0;index<rangeNodes.length;index++){
              const bounds=criteriaBounds[index]!,criterionRange=rangeNodes[index]!,cellRef=`${colToName(bounds.left+colOffset)}${bounds.top+rowOffset+1}`,value=this.getValue(cellRef,visiting);
              if(value.error||!matchCriterion(value.value,criteria[index]!)){satisfies=false;break;}
            }
            if(!satisfies)continue;
            const value=this.getValue(ref,visiting);if(value.error)return error(value.error.code,`${ref}: ${value.error.message}`);
            if(typeof value.value==='number'&&Number.isFinite(value.value)){
              total+=value.value;matched++;
              extreme=extreme===undefined?value.value:n.name==='MINIFS'?Math.min(extreme,value.value):Math.max(extreme,value.value);
            }
          }
          if(n.name==='AVERAGEIFS')return matched?total/matched:error('#DIV/0!','No matching numeric values.');
          if(n.name==='MAXIFS'||n.name==='MINIFS')return extreme??0;
          return Number.isFinite(total)?total:error('#NUM!','The result is not a finite number.');
        }
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
      if((n.name==='TOCOL'||n.name==='TOROW')&&n.args[0]&&(n.args[0].type==='ref'||n.args[0].type==='range')){
        // TOCOL/TOROW can explicitly discard errors. Ordinary range evaluation
        // propagates the first cell error, so build this argument cell-by-cell
        // and suppress only errors when ignore=2 or ignore=3.
        const source=n.args[0],ignore=n.args[1]?number(this.evaluateScalar(n.args[1],visiting)):0;
        if(!Number.isInteger(ignore)||ignore<0||ignore>3)return error('#VALUE!','Ignore must be 0, 1, 2, or 3.');
        const ignoreErrors=ignore===2||ignore===3;
        const first=parseRef(source.type==='range'?source.from:source.ref),last=parseRef(source.type==='range'?source.to:source.ref);
        if(!first||!last)return error('#REF!','TOCOL/TOROW source reference is invalid.');
        if(Math.max(first.row,last.row)>=this.rows||Math.max(first.col,last.col)>=this.cols)return error('#REF!','TOCOL/TOROW source reference is outside worksheet boundaries.');
        const matrix:Datum[][]=[];
        for(let row=Math.min(first.row,last.row);row<=Math.max(first.row,last.row);row++){
          const line:Datum[]=[];
          for(let col=Math.min(first.col,last.col);col<=Math.max(first.col,last.col);col++){
            const ref=`${colToName(col)}${row+1}`;
            const result=source.sheet?this.externalSheetResolver?.(source.sheet,ref,visiting):this.getValue(ref,visiting);
            if(!result)return error('#REF!',`Worksheet '${source.sheet}' does not exist.`);
            if(result.error){if(ignoreErrors){line.push(null);continue;}return error(result.error.code,`${source.sheet?`${source.sheet}!`:''}${ref}: ${result.error.message}`);}
            line.push(result.value);
          }
          matrix.push(line);
        }
        return f.run([matrix,...n.args.slice(1).map(arg=>this.evaluate(arg,visiting))]);
      }
      return f.run(n.args.map(arg=>this.evaluate(arg,visiting)));
    }
    return error('#ERROR!','Unexpected AST node.');
  }
  private evaluateScalar(n:Node,visiting:string[]):Value {
    const v=this.evaluate(n,visiting);return Array.isArray(v)?error('#VALUE!','Cannot perform scalar arithmetic on a whole range.'):v;
  }
  inspect(ref:string):Inspect {
    ref=known(ref,this.rows,this.cols);const ds=[...(this.deps.get(ref)??[])],ranges=this.staticRangeDeps.get(ref)??[],rvs=[...(this.rev.get(ref)??[])],referencedRefs=new Set(ds);
    for(const range of ranges){const from=parseRef(range.from),to=parseRef(range.to);if(!from||!to)continue;for(let row=Math.min(from.row,to.row);row<=Math.max(from.row,to.row)&&referencedRefs.size<100;row++)for(let col=Math.min(from.col,to.col);col<=Math.max(from.col,to.col)&&referencedRefs.size<100;col++)referencedRefs.add(`${colToName(col)}${row+1}`);}
    return {ref,raw:this.getRaw(ref),result:this.getValue(ref),dependencies:[...ds,...ranges.map(range=>`${range.from}:${range.to}`)],dependents:rvs,
      expressionTree:this.ast.has(ref)?tree(this.ast.get(ref)!):this.parseErrors.has(ref)?'Unable to parse formula':'Literal / empty cell',
      referencedValues:[...referencedRefs].slice(0,100).map(r=>({ref:r,raw:this.getRaw(r),display:display(this.getValue(r))}))};
  }
  preview(ref:string,proposed:string):{changes:{ref:string;before:string;after:string}[];stats:RecalcStats} {
    const target=known(ref,this.rows,this.cols);const impacted=this.transitiveDependents(target);
    const before=new Map(impacted.map(id=>[id,display(this.getValue(id))]));
    const cloned=Workbook.fromJSON(this.toJSON());if(this.tableContext.length)cloned.setTableContext(this.tableContext);const stats=cloned.applyEdits([{ref:target,raw:proposed}],false);
    return {changes:impacted.map(id=>({ref:id,before:before.get(id)??'',after:display(cloned.getValue(id))})).filter(c=>c.before!==c.after),stats};
  }
  transitiveDependents(ref:string):string[] {
    const queue=[known(ref,this.rows,this.cols)], seen=new Set(queue);
    for(let i=0;i<queue.length;i++){
      const current=queue[i]!,point=parseRef(current),owners=new Set(this.rev.get(current)??[]);
      if(point)for(const owner of this.staticRangeColumns.get(`${this.sheetName.toLocaleLowerCase()}!${point.col}`)??[]){if((this.staticRangeDeps.get(owner)??[]).some(range=>{const from=parseRef(range.from),to=parseRef(range.to);return !!from&&!!to&&point.row>=Math.min(from.row,to.row)&&point.row<=Math.max(from.row,to.row)&&point.col>=Math.min(from.col,to.col)&&point.col<=Math.max(from.col,to.col);}))owners.add(owner);}
      for(const owner of owners)if(!seen.has(owner)){seen.add(owner);queue.push(owner);}
    }
    return queue;
  }
  getLastUsedCell():{row:number;col:number}|null {
    let row=-1,col=-1;
    for(const ref of this.raw.keys()){const p=parseRef(ref);if(p){row=Math.max(row,p.row);col=Math.max(col,p.col);}}
    for(const ref of this.styles.keys()){const p=parseRef(ref);if(p){row=Math.max(row,p.row);col=Math.max(col,p.col);}}
    return row<0||col<0?null:{row,col};
  }
  toJSON():{version:2;rows:number;cols:number;cells:Record<string,string>;styles:Record<string,CellStyle>;definedNames:DefinedName[]} {return {version:2,rows:this.rows,cols:this.cols,cells:Object.fromEntries(this.raw),styles:Object.fromEntries(this.styles),definedNames:this.getDefinedNames()};}
  static fromJSON(data:unknown,options:{deferCalculation?:boolean}={}):Workbook {
    if(!data||typeof data!=='object')throw Error('Invalid workbook data.');
    const d=data as {rows?:number;cols?:number;cells?:Record<string,unknown>;styles?:Record<string,unknown>;definedNames?:unknown};
    if(!Number.isInteger(d.rows)||!Number.isInteger(d.cols)||(d.rows??0)<1||(d.rows??0)>MAX_WORKSHEET_ROWS||(d.cols??0)<1||(d.cols??0)>MAX_WORKSHEET_COLUMNS)throw Error(`Workbook dimensions must be 1–${MAX_WORKSHEET_ROWS.toLocaleString()} rows and 1–${MAX_WORKSHEET_COLUMNS} columns.`);
    if(!d.cells||typeof d.cells!=='object'||Array.isArray(d.cells))throw Error('Workbook cells must be an object.');
    const book=new Workbook(d.rows!,d.cols!);if(d.definedNames!==undefined)book.setDefinedNames(d.definedNames,false);
    const edits=Object.entries(d.cells).map(([ref,raw])=>{if(typeof raw!=='string')throw Error(`Invalid content in ${ref}.`);return {ref,raw};});
    if(edits.length>MAX_WORKSHEET_POPULATED_CELLS)throw Error(`Workbook exceeds GridForge's ${MAX_WORKSHEET_POPULATED_CELLS.toLocaleString()} populated-cell limit.`);
    book.applyEdits(edits,false,!options.deferCalculation);
    if(d.styles!==undefined){if(!d.styles||typeof d.styles!=='object'||Array.isArray(d.styles))throw Error('Invalid workbook formatting.');
      if(Object.keys(d.styles).length>MAX_WORKSHEET_POPULATED_CELLS)throw Error(`Workbook exceeds GridForge's ${MAX_WORKSHEET_POPULATED_CELLS.toLocaleString()} styled-cell limit.`);
      book.applyStyles(Object.entries(d.styles).map(([ref,style])=>({ref,style:cleanStyle(style)})),false);
    }
    return book;
  }
}
