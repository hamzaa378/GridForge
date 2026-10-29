/** Calculated-column propagation for native table formulas.
 *  Policy: formula edits become the column template. Blank cells and prior formula cells
 *  are filled; occupied literal cells are preserved as explicit manual overrides.
 */
import type {Edit,Workbook} from './workbook.js';
import {colToName,parseRef} from './parser.js';
import {tableRange,type TableSpec} from './tables.js';

const ref=(row:number,col:number)=>`${colToName(col)}${row+1}`;
function translate(formula:string,delta:number):string{
 let out='',quoted=false,bracket=0;
 for(let i=0;i<formula.length;){const ch=formula[i]!;
  if(ch==='"'){out+=ch;if(quoted&&formula[i+1]==='"'){out+='"';i+=2;continue;}quoted=!quoted;i++;continue;}
  if(!quoted&&ch==='['){bracket++;out+=ch;i++;continue;}if(!quoted&&ch===']'&&bracket){bracket--;out+=ch;i++;continue;}
  if(!quoted&&!bracket){const m=/^\$?[A-Z]{1,3}\$?\d+/i.exec(formula.slice(i)),before=formula[i-1]??'',after=m?formula[i+m[0].length]??'':'';if(m&&!/[A-Z0-9_.]/i.test(before)&&!/[A-Z0-9_.]/i.test(after)){const p=parseRef(m[0].replace(/\$/g,''));if(p){const row=p.row+(/\$\d/.test(m[0])?0:delta);out+=row<0?'#REF!':`${m[0].match(/^\$?[A-Z]+/)?.[0]?.startsWith('$')?'$':''}${colToName(p.col)}${/\$\d/.test(m[0])?'$':''}${row+1}`;i+=m[0].length;continue;}}}
  out+=ch;i++;
 }
 return out;
}
export function propagateCalculatedColumns(book:Workbook,tables:TableSpec[],edits:Edit[]):Edit[]{
 const result=new Map<string,string>(edits.map(e=>[e.ref.toUpperCase(),e.raw]));
 for(const edit of edits){
  const p=parseRef(edit.ref);if(!p)continue;
  const table=tables.find(t=>{const b=tableRange(t.range,book.rows,book.cols);return p.row>b.r1&&p.row<=b.r2&&p.col>=b.c1&&p.col<=b.c2;});
  if(!table)continue;
  const b=tableRange(table.range,book.rows,book.cols),column=p.col-b.c1,existing=table.calculatedColumns?.find(c=>c.column===column);
  if(edit.raw.startsWith('=')){
   const overrides:number[]=[];
   for(let r=b.r1+1;r<=b.r2;r++){
    const cell=ref(r,p.col),raw=result.get(cell)??book.getRaw(cell);
    if(r===p.row){result.set(cell,edit.raw);continue;}
    if(!raw||raw.startsWith('='))result.set(cell,translate(edit.raw,r-p.row));
    else overrides.push(r-(b.r1+1));
   }
   const next={column,formula:translate(edit.raw,b.r1+1-p.row),overrides};
   table.calculatedColumns=[...(table.calculatedColumns??[]).filter(c=>c.column!==column),next];
  }else if(existing){
   const offset=p.row-(b.r1+1),overrides=new Set(existing.overrides??[]);
   if(edit.raw===translate(existing.formula,p.row-(b.r1+1)))overrides.delete(offset);else overrides.add(offset);
   existing.overrides=[...overrides].sort((a,b)=>a-b);
  }
 }
 return [...result].map(([ref,raw])=>({ref,raw}));
}

export function fillCalculatedRows(book:Workbook,table:TableSpec,firstRow:number,lastRow:number):Edit[]{
 const b=tableRange(table.range,book.rows,book.cols),edits:Edit[]=[];
 for(const calculated of table.calculatedColumns??[]){
  const col=b.c1+calculated.column;
  for(let row=Math.max(firstRow,b.r1+1);row<=Math.min(lastRow,b.r2);row++){
   if((calculated.overrides??[]).includes(row-(b.r1+1)))continue;
   const cell=ref(row,col),raw=book.getRaw(cell);if(!raw)edits.push({ref:cell,raw:translate(calculated.formula,row-(b.r1+1))});else if(!raw.startsWith('='))calculated.overrides=[...new Set([...(calculated.overrides??[]),row-(b.r1+1)])].sort((a,b)=>a-b);
  }
 }
 return edits;
}

/** Keep calculated-column metadata aligned when the workbook's undo/redo stack is used. */
export function reconcileCalculatedColumns(book:Workbook,tables:TableSpec[]):void{
 for(const table of tables){const b=tableRange(table.range,book.rows,book.cols),updated:NonNullable<TableSpec['calculatedColumns']>=[];
  for(let col=b.c1;col<=b.c2;col++){const formulas:Array<{row:number;raw:string}>=[];const overrides:number[]=[];
   for(let row=b.r1+1;row<=b.r2;row++){const raw=book.getRaw(ref(row,col));if(raw.startsWith('='))formulas.push({row,raw});else if(raw)overrides.push(row-(b.r1+1));}
   if(!formulas.length)continue;const first=formulas[0]!,formula=translate(first.raw,b.r1+1-first.row);
   if(formulas.every(item=>translate(item.raw,b.r1+1-item.row)===formula))updated.push({column:col-b.c1,formula,overrides});
  }
  table.calculatedColumns=updated;
 }
}
