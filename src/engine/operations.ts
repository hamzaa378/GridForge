import {type Edit,Workbook} from './workbook.js';
import {type StyleEdit} from './format.js';
import {colToName,parseRef,shiftFormula} from './parser.js';
import {display} from './workbook.js';
export interface Rect {r1:number;r2:number;c1:number;c2:number}
export interface SelectionAnalysis {count:number;numbers:number;sum:number;average:number|null;min:number|null;max:number|null;blanks:number;errors:number;duplicates:number;chart:{label:string;value:number}[]}
const address=(row:number,col:number)=>`${colToName(col)}${row+1}`;
export function analyzeSelection(book:Workbook,b:Rect):SelectionAnalysis {
  let count=0,numbers=0,sum=0,min=Infinity,max=-Infinity,blanks=0,errors=0,duplicates=0;
  const seen=new Set<string>();const chart:{label:string;value:number}[]=[];
  for(let r=b.r1;r<=b.r2;r++)for(let c=b.c1;c<=b.c2;c++) {
    const result=book.getValue(address(r,c));count++;
    if(result.error){errors++;continue;}
    const v=result.value;if(v===null||v===''){blanks++;continue;}
    const key=JSON.stringify([typeof v,v]);if(seen.has(key))duplicates++;else seen.add(key);
    if(typeof v==='number'&&Number.isFinite(v)){numbers++;sum+=v;min=Math.min(min,v);max=Math.max(max,v);}
  }
  for(let r=b.r1;r<=Math.min(b.r2,b.r1+24);r++){
    const label=display(book.getValue(address(r,b.c1)))||address(r,b.c1);
    const target=b.c2>b.c1?b.c1+1:b.c1;
    const v=book.getValue(address(r,target));if(!v.error&&typeof v.value==='number'&&Number.isFinite(v.value))chart.push({label,value:v.value});
  }
  return {count,numbers,sum,average:numbers?sum/numbers:null,min:numbers?min:null,max:numbers?max:null,blanks,errors,duplicates,chart};
}
export interface SearchMatch {ref:string;raw:string}
export function findCells(book:Workbook,needle:string,caseSensitive=false,includeFormulas=true):SearchMatch[]{
  if(!needle)return[];const query=caseSensitive?needle:needle.toLowerCase();const items=book.allEntries().filter(({raw})=>{
    if(!includeFormulas&&raw.startsWith('='))return false;
    return (caseSensitive?raw:raw.toLowerCase()).includes(query);
  });
  return items.sort((a,b)=>{const x=parseRef(a.ref)!,y=parseRef(b.ref)!;return x.row-y.row||x.col-y.col;});
}
export function replaceCells(book:Workbook,needle:string,replacement:string,caseSensitive=false,includeFormulas=false):Edit[]{
  if(!needle)throw Error('Enter text to find.');
  return findCells(book,needle,caseSensitive,includeFormulas).map(({ref,raw})=>{
    let next='';let start=0;const original=caseSensitive?raw:raw.toLowerCase(),query=caseSensitive?needle:needle.toLowerCase();
    for(let i=original.indexOf(query,start);i!==-1;i=original.indexOf(query,start)) {
      next+=raw.slice(start,i)+replacement;start=i+needle.length;
    }
    return {ref,raw:next+raw.slice(start)};
  });
}
/** A single atomic, undoable data-and-format transaction for a selected rectangle.
 * Other columns are intentionally untouched; sorting an entire table is the caller's job.
 */
export function sortRangeTransaction(book:Workbook,b:Rect,byCol:number,descending=false,firstRowHeader=false):{edits:Edit[];styles:StyleEdit[]}{
  if(byCol<b.c1||byCol>b.c2)throw Error('Sort key is outside the selection.');
  const from=b.r1+(firstRowHeader?1:0);if(b.r2<=from)return {edits:[],styles:[]};
  const entries=[];
  for(let r=from;r<=b.r2;r++)entries.push({original:r,key:book.getValue(address(r,byCol)),raw:Array.from({length:b.c2-b.c1+1},(_,j)=>book.getRaw(address(r,b.c1+j))),styles:Array.from({length:b.c2-b.c1+1},(_,j)=>book.getStyle(address(r,b.c1+j)))});
  entries.sort((a,b)=>{
    const x=a.key.value,y=b.key.value;if(a.key.error&&!b.key.error)return 1;if(b.key.error&&!a.key.error)return -1;
    if(x===null||x==='')return y===null||y===''?a.original-b.original:1;
    if(y===null||y==='')return -1;
    const comparison=typeof x==='number'&&typeof y==='number'?x-y:String(x).localeCompare(String(y),undefined,{numeric:true,sensitivity:'base'});
    return (descending?-comparison:comparison)||a.original-b.original;
  });
  const edits:Edit[]=[],styles:StyleEdit[]=[];entries.forEach((item,index)=>{const destination=from+index;
    item.raw.forEach((raw,delta)=>{const adjusted=raw.startsWith('=')?shiftFormula(raw,destination-item.original,0):raw;
      if(adjusted!==book.getRaw(address(destination,b.c1+delta)))edits.push({ref:address(destination,b.c1+delta),raw:adjusted});
      const ref=address(destination,b.c1+delta),style=item.styles[delta]!;
      if(JSON.stringify(style)!==JSON.stringify(book.getStyle(ref)))styles.push({ref,style});
    });
  });
  return {edits,styles};
}
/** Backwards-compatible values-only helper retained for v2 tests and integrations. */
export function sortSelection(book:Workbook,b:Rect,byCol:number,descending=false,firstRowHeader=false):Edit[]{
  return sortRangeTransaction(book,b,byCol,descending,firstRowHeader).edits;
}
