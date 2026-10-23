/** GridForge automation is a validated declarative action format, not VBA/JavaScript.
 * v1 action macros remain supported. v2 adds safe cross-sheet operations, sorting,
 * fills, copying, find/replace, previews and preflight validation.
 * No networking, filesystem access, code execution or user-defined scripting.
 */
import {SpreadsheetDocument} from './document.js';
import {type Edit} from './workbook.js';
import {parseRef,colToName,shiftFormula} from './parser.js';
import {cleanStyle,type StyleEdit} from './format.js';
import {snapshotCells,pasteSnapshot} from './clipboard.js';
import {sortRangeTransaction,replaceCells,type Rect} from './operations.js';
export type LegacyStep={kind:'cells';edits:Edit[]}|{kind:'styles';edits:StyleEdit[]};
export type StepV2=
 | {kind:'cells';sheetName?:string;edits:Edit[]}
 | {kind:'styles';sheetName?:string;edits:StyleEdit[]}
 | {kind:'clear';sheetName?:string;range:string;clearFormatting?:boolean}
 | {kind:'fillDown';sheetName?:string;source:string;range:string}
 | {kind:'copy';sheetName?:string;sourceRange:string;destination:string;valuesOnly?:boolean}
 | {kind:'sort';sheetName?:string;range:string;byColumn:number;descending?:boolean;hasHeader?:boolean}
 | {kind:'replace';sheetName?:string;find:string;replace:string;caseSensitive?:boolean;includeFormulas?:boolean};
export type ActionMacro={version:1;name:string;sheetName:string;steps:LegacyStep[]}|{version:2;name:string;sheetName:string;steps:StepV2[]};
const address=(r:number,c:number)=>`${colToName(c)}${r+1}`;
function rectangle(source:string):Rect{
  if(typeof source!=='string'||source.length>30)throw Error('Invalid macro range.');
  const xs=source.toUpperCase().split(':');if(xs.length>2)throw Error('Invalid rectangular range.');
  const a=parseRef(xs[0]!),b=parseRef(xs[1]??xs[0]!);if(!a||!b)throw Error(`Invalid macro address ${source}.`);
  if(Math.abs(a.row-b.row)+1>1000||Math.abs(a.col-b.col)+1>104)throw Error('Range too large.');
  return {r1:Math.min(a.row,b.row),r2:Math.max(a.row,b.row),c1:Math.min(a.col,b.col),c2:Math.max(a.col,b.col)};
}
const isObj=(x:unknown):x is Record<string,unknown>=>!!x&&typeof x==='object'&&!Array.isArray(x);
const cleanName=(s:unknown,kind:string):string=>{if(typeof s!=='string'||!s.trim()||s.length>(kind==='name'?64:31))throw Error(`Invalid ${kind}.`);return s.trim();};
function parseEdits(v:unknown,kind:'cells'|'styles'){
  if(!Array.isArray(v)||v.length>5000)throw Error('Invalid macro edit list.');
  return v.map(item=>{if(!isObj(item)||typeof item.ref!=='string'||!parseRef(item.ref))throw Error('Invalid edit address.');
    const ref=item.ref.toUpperCase().replace(/\$/g,'');
    if(kind==='cells'){
      if(typeof item.raw!=='string'||item.raw.length>30000)throw Error('Invalid macro cell contents.');
      return {ref,raw:item.raw};
    }
    return {ref,style:cleanStyle(item.style)};
  });
}
function parseStep(step:unknown,version:1|2):LegacyStep|StepV2{
  if(!isObj(step)||typeof step.kind!=='string')throw Error('Invalid macro action.');
  const kind=step.kind;if(version===1&&!['cells','styles'].includes(kind))throw Error('Legacy macro has unsupported action.');
  const sheetName=version===2&&step.sheetName!==undefined?cleanName(step.sheetName,'sheet name'):undefined;
  if(kind==='cells'||kind==='styles')return {kind,sheetName,edits:parseEdits(step.edits,kind)} as StepV2;
  if(version===1)throw Error('Unsupported v1 macro step.');
  if(kind==='clear')return {kind,sheetName,range:cleanRange(step.range),clearFormatting:step.clearFormatting===true};
  if(kind==='fillDown')return {kind,sheetName,source:cleanRef(step.source),range:cleanRange(step.range)};
  if(kind==='copy')return {kind,sheetName,sourceRange:cleanRange(step.sourceRange),destination:cleanRef(step.destination),valuesOnly:step.valuesOnly===true};
  if(kind==='sort'){
    const range=cleanRange(step.range),byColumn=step.byColumn,selected=rectangle(range);
    if(typeof byColumn!=='number'||!Number.isInteger(byColumn)||byColumn<0||byColumn>=104||byColumn<selected.c1||byColumn>selected.c2)throw Error('Invalid sort column outside selection.');
    return {kind,sheetName,range,byColumn,descending:step.descending===true,hasHeader:step.hasHeader===true};
  }
  if(kind==='replace'){
    if(typeof step.find!=='string'||!step.find||step.find.length>200||typeof step.replace!=='string'||step.replace.length>2000)throw Error('Invalid replace expression.');
    return {kind,sheetName,find:step.find,replace:step.replace,caseSensitive:step.caseSensitive===true,includeFormulas:step.includeFormulas===true};
  }
  throw Error(`Unsupported macro action '${kind}'. No custom code is executed.`);
}
function cleanRef(x:unknown):string{if(typeof x!=='string'||!parseRef(x))throw Error('Invalid macro address.');return x.toUpperCase().replace(/\$/g,'');}
function cleanRange(x:unknown):string{if(typeof x!=='string')throw Error('Invalid macro range.');rectangle(x);return x.toUpperCase().replace(/\$/g,'');}
export function validateMacro(data:unknown):ActionMacro{
  if(!isObj(data)||(data.version!==1&&data.version!==2)||!Array.isArray(data.steps)||data.steps.length>200)throw Error('Invalid GridForge macro header.');
  const version=data.version,name=cleanName(data.name,'name'),sheetName=cleanName(data.sheetName,'sheet name');
  const steps=data.steps.map(x=>parseStep(x,version));let operations=0;
  for(const step of steps){
    if(step.kind==='cells'||step.kind==='styles')operations+=step.edits.length;
    else if(step.kind==='clear'||step.kind==='fillDown'||step.kind==='sort'){
      const b=rectangle(step.range);operations+=(b.r2-b.r1+1)*(b.c2-b.c1+1);
    }else if(step.kind==='copy'){
      const b=rectangle(step.sourceRange);operations+=(b.r2-b.r1+1)*(b.c2-b.c1+1);
    }else if(step.kind==='replace')operations+=1000;
    if(operations>15000)throw Error('Macro operation limit is 15,000 cell touches.');
  }
  return version===1?{version,name,sheetName,steps:steps as LegacyStep[]}:{version,name,sheetName,steps:steps as StepV2[]};
}
export interface MacroReport {values:number;styles:number;sheets:string[];changes:{sheet:string;ref:string;before:string;after:string}[];dryRun:boolean}
function checkRect(book:{rows:number;cols:number},b:Rect){if(b.r1<0||b.c1<0||b.r2>=book.rows||b.c2>=book.cols)throw Error('Macro range exceeds worksheet boundaries.');}
const refs=(b:Rect)=>{const out:string[]=[];for(let r=b.r1;r<=b.r2;r++)for(let c=b.c1;c<=b.c2;c++)out.push(address(r,c));return out;};
function simulate(doc:SpreadsheetDocument,macro:ActionMacro):SpreadsheetDocument{
  const clone=SpreadsheetDocument.fromJSON(doc.toJSON());
  for(const step of macro.steps){
    const name=('sheetName'in step&&step.sheetName)||macro.sheetName;
    const entry=clone.sheets.find(x=>x.name===name);if(!entry)throw Error(`Worksheet '${name}' does not exist.`);
    const book=entry.book;
    if(step.kind==='cells'){for(const edit of step.edits){const p=parseRef(edit.ref)!;checkRect(book,{r1:p.row,r2:p.row,c1:p.col,c2:p.col});}book.applyEdits(step.edits,false);}
    if(step.kind==='styles'){for(const edit of step.edits){const p=parseRef(edit.ref)!;checkRect(book,{r1:p.row,r2:p.row,c1:p.col,c2:p.col});}book.applyStyles(step.edits,false);}
    if(step.kind==='clear'){
      const b=rectangle(step.range);checkRect(book,b);const cells=refs(b);book.applyEdits(cells.map(ref=>({ref,raw:''})),false);
      if(step.clearFormatting)book.applyStyles(cells.map(ref=>({ref,style:{}})),false);
    }
    if(step.kind==='fillDown'){
      const b=rectangle(step.range),source=parseRef(step.source)!;checkRect(book,b);checkRect(book,{r1:source.row,r2:source.row,c1:source.col,c2:source.col});
      const raw=book.getRaw(step.source),style=book.getStyle(step.source);
      book.applyTransaction(refs(b).map(ref=>{const p=parseRef(ref)!;return {ref,raw:raw.startsWith('=')?shiftFormula(raw,p.row-source.row,p.col-source.col):raw};}),refs(b).map(ref=>({ref,style})));
    }
    if(step.kind==='copy'){
      const b=rectangle(step.sourceRange),destination=parseRef(step.destination)!;checkRect(book,b);
      const snapshot=snapshotCells(book,b,entry.id),tx=pasteSnapshot(book,snapshot,{row:destination.row,col:destination.col},entry.id,step.valuesOnly);
      book.applyTransaction(tx.edits,tx.styles);
    }
    if(step.kind==='sort'){
      const b=rectangle(step.range);checkRect(book,b);
      const tx=sortRangeTransaction(book,b,step.byColumn,step.descending,step.hasHeader);
      book.applyTransaction(tx.edits,tx.styles);
    }
    if(step.kind==='replace')book.applyEdits(replaceCells(book,step.find,step.replace,step.caseSensitive,step.includeFormulas),false);
  }
  return clone;
}
/** Complete dry-run before applying any changes; safe even for cross-sheet macros.
 * Changes to each sheet become one undoable transaction. */
export function runActionMacro(doc:SpreadsheetDocument,data:unknown,options?:{dryRun?:boolean}):MacroReport {
  const macro=validateMacro(data),cloned=simulate(doc,macro),changes:MacroReport['changes']=[];
  let values=0,styles=0;const changedSheets:string[]=[];
  const transactions:{name:string;edits:Edit[];styleEdits:StyleEdit[]}[]=[];
  for(const updated of cloned.sheets){
    const original=doc.sheets.find(x=>x.name===updated.name)!;
    const oldRaw=new Map(original.book.allEntries().map(x=>[x.ref,x.raw]));
    const newRaw=new Map(updated.book.allEntries().map(x=>[x.ref,x.raw]));
    const oldStyles=new Map(original.book.allStyles().map(x=>[x.ref,x.style]));
    const newStyles=new Map(updated.book.allStyles().map(x=>[x.ref,x.style]));
    const edits:Edit[]=[],styleEdits:StyleEdit[]=[];
    for(const ref of new Set([...oldRaw.keys(),...newRaw.keys()])){
      const before=oldRaw.get(ref)??'',after=newRaw.get(ref)??'';
      if(before!==after){edits.push({ref,raw:after});if(changes.length<100)changes.push({sheet:updated.name,ref,before,after});}
    }
    for(const ref of new Set([...oldStyles.keys(),...newStyles.keys()])){
      const before=oldStyles.get(ref)??{},after=newStyles.get(ref)??{};
      if(JSON.stringify(before)!==JSON.stringify(after))styleEdits.push({ref,style:after});
    }
    if(edits.length||styleEdits.length){changedSheets.push(updated.name);transactions.push({name:updated.name,edits,styleEdits});values+=edits.length;styles+=styleEdits.length;}
  }
  if(values+styles>15000)throw Error('Macro would modify too many cells.');
  if(!options?.dryRun)for(const tx of transactions){const sheet=doc.sheets.find(s=>s.name===tx.name)!;sheet.book.applyTransaction(tx.edits,tx.styleEdits);}
  if(macro.version===1&&!options?.dryRun)return {values,styles} as MacroReport;
  return {values,styles,sheets:changedSheets,changes,dryRun:!!options?.dryRun};
}
