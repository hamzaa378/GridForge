import { Workbook,display } from './workbook.js';
import {parseRef,colToName} from './parser.js';
export interface SheetTools { freezeRows:number; freezeCols:number; filter:FilterSpec|null; conditional:ConditionalRule[] }
export interface FilterSpec { startRow:number; endRow:number; startCol:number; endCol:number; column:number; query:string; }
export interface ConditionalRule { start:string; end:string; operator:'gt'|'lt'|'eq'; threshold:number; color:string; }
export const freshTools=():SheetTools=>({freezeRows:0,freezeCols:0,filter:null,conditional:[]});
export function validateTools(data:unknown,rows:number,cols:number):SheetTools{
  const t=freshTools();if(!data||typeof data!=='object')return t;
  const obj=data as Partial<SheetTools>;
  for(const key of ['freezeRows','freezeCols'] as const){const v=obj[key];if(v!==undefined){if(!Number.isInteger(v)||v<0||v>(key==='freezeRows'?rows-1:cols-1))throw Error(`Invalid ${key}`);t[key]=v;}}
  if(obj.filter){const f=obj.filter;
    if(![f.startRow,f.endRow,f.startCol,f.endCol,f.column].every(Number.isInteger)||f.startRow<0||f.endRow>=rows||f.endRow<=f.startRow||f.startCol<0||f.endCol>=cols||f.column<f.startCol||f.column>f.endCol||typeof f.query!=='string'||f.query.length>100)throw Error('Invalid table filter');
    t.filter={...f};
  }
  if(obj.conditional){if(!Array.isArray(obj.conditional)||obj.conditional.length>30)throw Error('Too many conditional rules');
    for(const rule of obj.conditional){const a=parseRef(rule.start),b=parseRef(rule.end);
      if(!a||!b||a.row>=rows||b.row>=rows||a.col>=cols||b.col>=cols||!['gt','lt','eq'].includes(rule.operator)||!Number.isFinite(rule.threshold)||!/^#[0-9a-f]{6}$/i.test(rule.color))throw Error('Invalid conditional formatting rule');
      t.conditional.push({...rule});
    }
  }
  return t;
}
export function filterVisibleRows(book:Workbook,spec:FilterSpec|null):Set<number> {
  const shown=new Set<number>();if(!spec)return shown;
  for(let r=spec.startRow+1;r<=spec.endRow;r++){
    const ref=`${colToName(spec.column)}${r+1}`;
    const value=display(book.getValue(ref));
    if(!spec.query||value.toLowerCase().includes(spec.query.toLowerCase()))shown.add(r);
  }
  return shown;
}
export function ruleColor(ref:string,value:number|string|boolean|null,rules:ConditionalRule[]):string|null {
  if(typeof value!=='number')return null;
  const p=parseRef(ref);if(!p)return null;
  for(const rule of rules){const a=parseRef(rule.start)!,b=parseRef(rule.end)!;
    if(p.row<Math.min(a.row,b.row)||p.row>Math.max(a.row,b.row)||p.col<Math.min(a.col,b.col)||p.col>Math.max(a.col,b.col))continue;
    if(rule.operator==='gt'&&value>rule.threshold||rule.operator==='lt'&&value<rule.threshold||rule.operator==='eq'&&value===rule.threshold)return rule.color;
  }
  return null;
}
