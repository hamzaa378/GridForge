/** Built-ins are intentionally small, explicit, and testable. */
import type { Value } from './parser.js';
import { moreFunctions } from './functions_extra.js';
import {v5Functions,v5FunctionNames,v5Categories} from './functions_v5.js';
import {v6Functions,v6FunctionNames,v6Categories} from './functions_v6.js';
import {v9Functions,v9Aliases,v9FunctionNames,v9Categories} from './functions_v9.js';
import {buildV10Functions,v10Categories} from './functions_v10.js';
import {connectedFunctionDefinitions,connectedFunctionNames,connectedFunctionRequirements} from './functions_connected.js';
export {v6FunctionNames,v6Categories,v9FunctionNames,v9Categories};
export type Datum = Value | Datum[][];
export class SheetError extends Error {
  constructor(readonly code: string, message: string) { super(message); this.name = 'SheetError'; }
}
export const error = (code: string, message: string): never => { throw new SheetError(code, message); };
export const flatten = (xs: Datum[]): Value[] => {
  const all: Value[] = [];
  const visit = (x: Datum): void => { if (Array.isArray(x)) x.forEach(row => row.forEach(visit)); else all.push(x); };
  xs.forEach(visit); return all;
};
export function number(x: Value): number {
  if (x === null || x === '') return 0;
  if (typeof x === 'boolean') return x ? 1 : 0;
  const n = Number(x);
  if (!Number.isFinite(n)) return error('#VALUE!', `Expected a number, received '${x}'.`);
  return n;
}
export function text(x: Value): string { return x === null ? '' : typeof x === 'boolean' ? (x ? 'TRUE' : 'FALSE') : String(x); }
export function bool(x: Value): boolean { return x === null || x === '' ? false : typeof x === 'string' ? x.toUpperCase() !== 'FALSE' : Boolean(x); }
const scalar = (x: Datum | undefined): Value => {
  if (Array.isArray(x)) return error('#VALUE!', 'Expected a single cell, not a range.');
  return x ?? null;
};
const at = (xs: Datum[], idx: number): Value => scalar(xs[idx]);
const nums = (xs: Datum[]): number[] => flatten(xs).filter((x): x is number => typeof x === 'number' && Number.isFinite(x));
const want = (xs: Datum[], min: number, max = Infinity): void => { if (xs.length < min || xs.length > max) error('#ARG!', `Expected ${min}${max === min ? '' : `–${Number.isFinite(max) ? max : 'more'}`} arguments; got ${xs.length}.`); };
const result = (n: number): number => Number.isFinite(n) ? n : error('#NUM!', 'Calculation produced an invalid or infinite number.');
const numericArg = (xs: Datum[], i: number): number => number(at(xs, i));
const nonEmpty = (x: Value): boolean => x !== null && x !== '';
const sortNums = (xs: Datum[]): number[] => nums(xs).sort((a,b)=>a-b);
export function matchCriterion(v: Value, criterion: Value): boolean {
  const c = text(criterion); const m = /^(<=|>=|<>|!=|=|<|>)(.*)$/.exec(c);
  const operand = m ? m[2]! : c; const op = m ? m[1]! : '=';
  // Excel criteria can contain * / ? wildcards and ~ escapes. Only apply
  // wildcard semantics to text comparison, not numeric inequalities.
  if ((op === '=' || op === '<>' || op === '!=') && /[*?~]/.test(operand) && typeof v === 'string') {
    let pattern = '^';
    for(let i=0;i<operand.length;i++) {
      let ch=operand[i]!;
      if(ch==='~' && i+1<operand.length){ch=operand[++i]!; pattern+=ch.replace(/[.*+?^${}()|[\]\\]/g,'\\$&');}
      else if(ch==='*')pattern+='.*';
      else if(ch==='?')pattern+='.';
      else pattern+=ch.replace(/[.*+?^${}()|[\]\\]/g,'\\$&');
    }
    const hit=new RegExp(pattern+'$','i').test(v);
    return op==='=' ? hit : !hit;
  }
  const aNum = Number(v), bNum = Number(operand);
  const numeric = v !== null && v !== '' && operand !== '' && Number.isFinite(aNum) && Number.isFinite(bNum);
  const cmp = numeric ? Math.sign(aNum - bNum) : text(v).toLowerCase().localeCompare(operand.toLowerCase());
  switch (op) { case '=': return cmp === 0; case '<>': case '!=': return cmp !== 0; case '>': return cmp > 0; case '<': return cmp < 0; case '>=': return cmp >= 0; case '<=': return cmp <= 0; default: return false; }
}
const matchOrder=(a:Value,b:Value):number=>{
  const rank=(v:Value):number=>v===null?0:typeof v==='number'?1:typeof v==='string'?2:3;
  const ar=rank(a),br=rank(b);
  if(ar!==br)return ar-br;
  if(typeof a==='number'&&typeof b==='number')return a<b?-1:a>b?1:0;
  if(typeof a==='boolean'&&typeof b==='boolean')return Number(a)-Number(b);
  const left=text(a).toLowerCase(),right=text(b).toLowerCase();
  return left<right?-1:left>right?1:0;
};
export type FunctionDefinition = { description: string; example: string; run: (args: Datum[]) => Datum };
const D = (description: string, example: string, run: (a: Datum[]) => Datum): FunctionDefinition => ({description, example, run});
export const functions: Record<string, FunctionDefinition> = {
 SUM: D('Sum numeric values, ignoring text and blanks.', '=SUM(A1:A10)', a => result(nums(a).reduce((s,n)=>s+n,0))),
 AVERAGE: D('Arithmetic mean of numeric values.', '=AVERAGE(B1:B8)', a => {const n=nums(a); return n.length ? result(n.reduce((s,x)=>s+x,0)/n.length) : error('#DIV/0!', 'AVERAGE has no numeric values.');}),
 MIN: D('Smallest numeric value.', '=MIN(A1:A10)', a => {const n=nums(a);return n.length ? Math.min(...n) : 0;}),
 MAX: D('Largest numeric value.', '=MAX(A1:A10)', a => {const n=nums(a);return n.length ? Math.max(...n) : 0;}),
 COUNT: D('Count numeric cells.', '=COUNT(A1:A10)', a => nums(a).length),
 COUNTA: D('Count nonempty values, including empty text and errors.', '=COUNTA(A1:A10)', a => flatten(a).filter(x=>x!==null).length),
 COUNTBLANK: D('Count blank cells in a range.', '=COUNTBLANK(A1:A10)', a => flatten(a).filter(x=>!nonEmpty(x)).length),
 MEDIAN: D('Median of numeric values.', '=MEDIAN(A1:A10)', a => {const n=sortNums(a);return n.length ? n.length%2 ? n[(n.length-1)/2]! : (n[n.length/2-1]!+n[n.length/2]!)/2 : error('#NUM!','MEDIAN has no numeric values.');}),
 PRODUCT: D('Multiply numeric values.', '=PRODUCT(A1:A3)', a => result(nums(a).reduce((p,n)=>p*n,1))),
 ABS: D('Absolute value.', '=ABS(-8)', a => {want(a,1,1);return Math.abs(numericArg(a,0));}),
 ROUND: D('Round to a number of decimal places.', '=ROUND(3.14159,2)', a => {want(a,1,2);const d=Math.trunc(numericArg(a,1));const p=10**d; const scaled=numericArg(a,0)*p;return result((scaled<0?-Math.round(-scaled):Math.round(scaled))/p);}),
 ROUNDUP: D('Round away from zero.', '=ROUNDUP(2.12,1)', a => {want(a,1,2);const p=10**Math.trunc(numericArg(a,1));const v=numericArg(a,0)*p;return result((v<0?Math.floor(v):Math.ceil(v))/p);}),
 ROUNDDOWN: D('Round toward zero.', '=ROUNDDOWN(2.19,1)', a => {want(a,1,2);const p=10**Math.trunc(numericArg(a,1));return result(Math.trunc(numericArg(a,0)*p)/p);}),
 SQRT: D('Square root.', '=SQRT(81)', a => {want(a,1,1);return result(Math.sqrt(numericArg(a,0)));}),
 POWER: D('Exponentiation.', '=POWER(2,8)', a => {want(a,2,2);return result(numericArg(a,0)**numericArg(a,1));}),
 MOD: D('Remainder after division.', '=MOD(17,5)', a => {want(a,2,2);const d=numericArg(a,1);return d === 0 ? error('#DIV/0!','Cannot divide by zero.') : result(((numericArg(a,0)%d)+d)%d);}),
 INT: D('Round down to an integer.', '=INT(3.9)', a => {want(a,1,1);return Math.floor(numericArg(a,0));}),
 CEILING: D('Round up to a multiple.', '=CEILING(13,5)', a => {want(a,1,2);const s=a.length===1?1:numericArg(a,1);return s===0?0:result(Math.ceil(numericArg(a,0)/s)*s);}),
 FLOOR: D('Round down to a multiple.', '=FLOOR(13,5)', a => {want(a,1,2);const s=a.length===1?1:numericArg(a,1);return s===0?0:result(Math.floor(numericArg(a,0)/s)*s);}),
 SIGN: D('Sign of a number.', '=SIGN(-4)', a => {want(a,1,1);return Math.sign(numericArg(a,0));}),
 EXP: D('Exponential e raised to x.', '=EXP(1)', a => {want(a,1,1);return result(Math.exp(numericArg(a,0)));}),
 LN: D('Natural logarithm.', '=LN(10)', a => {want(a,1,1);return result(Math.log(numericArg(a,0)));}),
 LOG: D('Logarithm with optional base.', '=LOG(100,10)', a => {want(a,1,2);const x=numericArg(a,0), b=a.length===1?10:numericArg(a,1);return x>0 && b>0 && b!==1 ? result(Math.log(x)/Math.log(b)) : error('#NUM!','Log input must be positive, and base cannot equal 1.');}),
 PI: D('The mathematical constant pi.', '=PI()', a => {want(a,0,0);return Math.PI;}),
 LEN: D('Number of characters.', '=LEN(A1)', a => {want(a,1,1);return text(at(a,0)).length;}),
 UPPER: D('Convert text to uppercase.', '=UPPER(A1)', a => {want(a,1,1);return text(at(a,0)).toUpperCase();}),
 LOWER: D('Convert text to lowercase.', '=LOWER(A1)', a => {want(a,1,1);return text(at(a,0)).toLowerCase();}),
 TRIM: D('Remove extra ASCII spaces, leaving one space between words.', '=TRIM(A1)', a => {want(a,1,1);return text(at(a,0)).replace(/ +/g,' ').replace(/^ | $/g,'');}),
 CONCAT: D('Join text or ranges.', '=CONCAT(A1," ",B1)', a => flatten(a).map(text).join('')),
 LEFT: D('Take the leftmost characters.', '=LEFT(A1,3)', a => {want(a,1,2);return text(at(a,0)).slice(0,a.length===1?1:Math.max(0,numericArg(a,1)));}),
 RIGHT: D('Take the rightmost characters.', '=RIGHT(A1,3)', a => {want(a,1,2);const s=text(at(a,0)),n=a.length===1?1:Math.max(0,numericArg(a,1));return n===0?'':s.slice(-n);}),
 MID: D('Take characters starting at a 1-based position.', '=MID(A1,2,4)', a => {want(a,3,3);return text(at(a,0)).substr(Math.max(0,numericArg(a,1)-1),Math.max(0,numericArg(a,2)));}),
 REPLACE: D('Replace part of a text string.', '=REPLACE("hello",2,2,"a")', a => {want(a,4,4);const s=text(at(a,0)),i=Math.max(0,numericArg(a,1)-1),n=Math.max(0,numericArg(a,2));return s.slice(0,i)+text(at(a,3))+s.slice(i+n);}),
 FIND: D('Find substring, case-sensitive, 1-based.', '=FIND("a","cable")', a => {want(a,2,3);const atIndex=text(at(a,1)).indexOf(text(at(a,0)),a.length===3?Math.max(0,numericArg(a,2)-1):0);return atIndex>=0?atIndex+1:error('#VALUE!','Text not found.');}),
 SEARCH: D('Find substring, case-insensitive.', '=SEARCH("A","cable")', a => {want(a,2,3);const atIndex=text(at(a,1)).toLowerCase().indexOf(text(at(a,0)).toLowerCase(),a.length===3?Math.max(0,numericArg(a,2)-1):0);return atIndex>=0?atIndex+1:error('#VALUE!','Text not found.');}),
 SUBSTITUTE: D('Replace all matching substrings.', '=SUBSTITUTE(A1,"old","new")', a => {want(a,3,3);const s=text(at(a,0)),find=text(at(a,1));return find?s.split(find).join(text(at(a,2))):s;}),
 EXACT: D('Compare two strings with case sensitivity.', '=EXACT(A1,B1)', a => {want(a,2,2);return text(at(a,0))===text(at(a,1));}),
 VALUE: D('Convert numeric text to a number.', '=VALUE("4.5")', a => {want(a,1,1);return number(at(a,0));}),
 IF: D('Choose one of two branches lazily.', '=IF(A1>10,"high","low")', () => error('#INTERNAL!','IF is evaluated lazily by the evaluator.')),
 IFERROR: D('Fallback on a formula evaluation error.', '=IFERROR(A1/B1,0)', () => error('#INTERNAL!','IFERROR is evaluated lazily by the evaluator.')),
 AND: D('True if all conditions are true.', '=AND(A1>0,B1>0)', a => flatten(a).every(bool)),
 OR: D('True if any condition is true.', '=OR(A1>0,B1>0)', a => flatten(a).some(bool)),
 NOT: D('Negate a boolean.', '=NOT(A1>10)', a => {want(a,1,1);return !bool(at(a,0));}),
 ISNUMBER: D('Check if a value is numeric.', '=ISNUMBER(A1)', a => {want(a,1,1);return typeof at(a,0)==='number';}),
 ISTEXT: D('Check if a value contains text.', '=ISTEXT(A1)', a => {want(a,1,1);return typeof at(a,0)==='string';}),
 ISBLANK: D('Check if a cell is truly empty (a formula returning "" is not blank).', '=ISBLANK(A1)', a => {want(a,1,1);return at(a,0)===null;}),
 SUMIF: D('Sum values matching a criterion.', '=SUMIF(A1:A5,">10",B1:B5)', a => {want(a,2,3);const range=flatten([a[0]!]), sumRange=a.length===3?flatten([a[2]!]):range; if(range.length!==sumRange.length)return error('#VALUE!','SUMIF ranges must have the same size.');return result(range.reduce<number>((s,x,i)=>s+(matchCriterion(x,at(a,1)) && typeof sumRange[i]==='number'?sumRange[i] as number:0),0));}),
 COUNTIF: D('Count cells matching a criterion.', '=COUNTIF(A1:A10,">10")', a => {want(a,2,2);return flatten([a[0]!]).filter(x=>matchCriterion(x,at(a,1))).length;}),
 AVERAGEIF: D('Average numeric matches.', '=AVERAGEIF(A1:A10,">5",B1:B10)', a => {want(a,2,3);const range=flatten([a[0]!]), values=a.length===3?flatten([a[2]!]):range;if(range.length!==values.length)return error('#VALUE!','AVERAGEIF ranges must have the same size.');const n=values.filter((x,i)=>matchCriterion(range[i]!,at(a,1))&&typeof x==='number') as number[];return n.length?n.reduce((s,x)=>s+x,0)/n.length:error('#DIV/0!','No numeric values matched.');}),
 VLOOKUP: D('Find an exact match or an approximate next-smaller match in the first table column.', '=VLOOKUP(E2,A2:C8,3,FALSE)', a => {
    want(a,3,4);const table=a[1];if(!Array.isArray(table))return error('#VALUE!','VLOOKUP needs a 2D range.');
    const col=Math.trunc(numericArg(a,2))-1;if(col<0 || col>=(table[0]?.length??0))return error('#REF!','VLOOKUP return column is outside the table.');
    const key=at(a,0), approx=a.length<4||bool(at(a,3));
    let row:Datum[]|undefined;
    if(approx){
      for(let i=1;i<table.length;i++)if(matchOrder(scalar(table[i-1]![0]),scalar(table[i]![0]))>0)return error('#N/A','VLOOKUP approximate lookup requires ascending sorted first-column data.');
      for(const cells of table){const current=scalar(cells[0]),order=matchOrder(current,key);if(order<=0)row=cells;else break;}
    }else{
      for(const cells of table){const current=scalar(cells[0]);if(matchCriterion(current,key)){row=cells;break;}}
    }
    return row?scalar(row[col]):error('#N/A',`VLOOKUP could not find '${text(key)}'.`);
  }),
 MATCH: D('Return the relative position of an exact or approximate match.', '=MATCH("pear",A1:A10,0)', a => {
   want(a,2,3);
   const lookup=at(a,0),items=flatten([a[1]!]),mode=a.length===3?numericArg(a,2):1;
   if(![-1,0,1].includes(mode))return error('#N/A','MATCH match_type must be -1, 0 or 1.');
   if(!items.length)return error('#N/A','MATCH lookup array is empty.');
   if(mode===0){const pos=items.findIndex(x=>matchCriterion(x,lookup));return pos>=0?pos+1:error('#N/A','MATCH found no matching value.');}
   for(let i=1;i<items.length;i++){
     const order=matchOrder(items[i-1]!,items[i]!);
     if(mode===1?order>0:order<0)return error('#N/A',`MATCH lookup array is not sorted ${mode===1?'ascending':'descending'}.`);
   }
   let candidate=-1;
   for(let i=0;i<items.length;i++){
     const order=matchOrder(items[i]!,lookup);
     if(mode===1?order<=0:order>=0)candidate=i;
     else if(mode===-1)break;
   }
   return candidate>=0?candidate+1:error('#N/A','MATCH found no eligible value.');
 }),
 INDEX: D('Return a value, row, column, or array at 1-based positions; zero selects a whole row or column.', '=INDEX(A1:C5,0,2)', a => {want(a,2,3);const data=a[0];if(!Array.isArray(data))return error('#VALUE!','INDEX expects a range.');const row=Math.trunc(numericArg(a,1)),column=a.length===3?Math.trunc(numericArg(a,2)):1;if(row<0||column<0)return error('#VALUE!','INDEX row and column numbers cannot be negative.');if(row===0&&column===0)return data.map(line=>line.map(value=>value??null));if(row===0){const c=column-1;if(c>=(data[0]?.length??0))return error('#REF!','INDEX column is outside the range.');return data.map(line=>[line[c]??null]);}if(row>data.length)return error('#REF!','INDEX row is outside the range.');if(column===0)return [(data[row-1]??[]).map(value=>value??null)];const c=column-1;if(c>=(data[row-1]?.length??0))return error('#REF!','INDEX column is outside the range.');return scalar(data[row-1]?.[c]);}),
 LARGE: D('k-th largest numeric value.', '=LARGE(A1:A10,2)', a => {want(a,2,2);const n=sortNums([a[0]!]).reverse(),k=Math.trunc(numericArg(a,1));return k<1||k>n.length?error('#NUM!','LARGE rank is out of range.'):n[k-1]!;}),
 SMALL: D('k-th smallest numeric value.', '=SMALL(A1:A10,2)', a => {want(a,2,2);const n=sortNums([a[0]!]),k=Math.trunc(numericArg(a,1));return k<1||k>n.length?error('#NUM!','SMALL rank is out of range.'):n[k-1]!;}),
 STDEV: D('Sample standard deviation.', '=STDEV(A1:A10)', a => {const n=nums(a);if(n.length<2)return error('#DIV/0!','STDEV requires two numeric values.');const avg=n.reduce((s,x)=>s+x,0)/n.length;return Math.sqrt(n.reduce((s,x)=>s+(x-avg)**2,0)/(n.length-1));}),
 VAR: D('Sample variance.', '=VAR(A1:A10)', a => {const n=nums(a);if(n.length<2)return error('#DIV/0!','VAR requires two numeric values.');const avg=n.reduce((s,x)=>s+x,0)/n.length;return n.reduce((s,x)=>s+(x-avg)**2,0)/(n.length-1);}),
};
Object.assign(functions, moreFunctions, v5Functions, v6Functions, v9Functions);
Object.assign(functions, v9Aliases(functions));
const v10Functions=buildV10Functions(functions);
Object.assign(functions,v10Functions,connectedFunctionDefinitions);
export const v10FunctionNames=Object.keys(v10Functions);
export {v10Categories,connectedFunctionNames,connectedFunctionRequirements};
export {v5FunctionNames,v5Categories};
export const functionNames = Object.keys(functions).sort();
