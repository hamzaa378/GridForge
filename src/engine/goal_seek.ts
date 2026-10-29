import {parseRef} from './parser.js';
import type {CellResult} from './workbook.js';
import {SpreadsheetDocument} from './document.js';

export interface GoalSeekSpec {sheetId:string;targetCell:string;targetValue:number;changingCell:string;tolerance?:number;maxIterations?:number}
export interface GoalSeekResult {changingValue:number;achievedValue:number;iterations:number;beforeValue:number}

/** Secant-based, non-destructive Goal Seek. Trials reuse a private document clone. */
export function solveGoalSeek(document:SpreadsheetDocument,raw:GoalSeekSpec):GoalSeekResult{
 const sheetIndex=document.sheets.findIndex(item=>item.id===raw.sheetId),sheet=document.sheets[sheetIndex];if(!sheet)throw Error('Goal Seek worksheet does not exist.');
 const target=parseRef(raw.targetCell.replace(/\$/g,'').toUpperCase()),changing=parseRef(raw.changingCell.replace(/\$/g,'').toUpperCase());
 if(!target||!changing||target.row>=sheet.book.rows||target.col>=sheet.book.cols||changing.row>=sheet.book.rows||changing.col>=sheet.book.cols)throw Error('Enter valid target and changing cells on the selected worksheet.');
 const targetCell=raw.targetCell.replace(/\$/g,'').toUpperCase(),changingCell=raw.changingCell.replace(/\$/g,'').toUpperCase();if(targetCell===changingCell)throw Error('The target formula cell and changing cell must be different.');
 if(!sheet.book.getRaw(targetCell).startsWith('='))throw Error('The target cell must contain a formula.');
 if(sheet.book.getRaw(changingCell).startsWith('='))throw Error('The changing cell must contain a value, not another formula.');
 if(!Number.isFinite(raw.targetValue))throw Error('Enter a finite target value.');
 const tolerance=raw.tolerance??1e-7,maxIterations=raw.maxIterations??60;if(!Number.isFinite(tolerance)||tolerance<=0||tolerance>1||!Number.isInteger(maxIterations)||maxIterations<1||maxIterations>100)throw Error('Goal Seek tolerance or iteration limit is invalid.');
 const before=sheet.book.getValue(changingCell);if(before.error||before.value!==null&&typeof before.value!=='number')throw Error('The changing cell must be blank or numeric.');const beforeValue=typeof before.value==='number'?before.value:0;
 // Clone the workbook once, then reuse its dependency graph for each bounded trial.
 // Re-cloning for every secant step made Goal Seek needlessly expensive on larger sheets.
 const trialDocument=SpreadsheetDocument.fromJSON(document.toJSON()),trialSheet=trialDocument.sheets[sheetIndex]!;
 const evaluate=(candidate:number):{x:number;f:number;result:CellResult}=>{
   if(!Number.isFinite(candidate)||Math.abs(candidate)>1e12)throw Error('Goal Seek diverged beyond the supported numeric range.');
   trialSheet.book.setCell(changingCell,String(candidate));const result=trialSheet.book.getValue(targetCell);
   if(result.error||typeof result.value!=='number'||!Number.isFinite(result.value))throw Error(`Goal Seek cannot evaluate ${targetCell} as a finite number${result.error?` (${result.error.code})`:''}.`);
   return {x:candidate,f:result.value-raw.targetValue,result};
 };
 let left=evaluate(beforeValue),right=evaluate(beforeValue+(Math.abs(beforeValue)*0.01||0.1));
 const isSolved=(point:{x:number;f:number})=>Math.abs(point.f)<=tolerance*Math.max(1,Math.abs(raw.targetValue));
 if(isSolved(left))return {changingValue:left.x,achievedValue:left.f+raw.targetValue,iterations:0,beforeValue};
 if(isSolved(right))return {changingValue:right.x,achievedValue:right.f+raw.targetValue,iterations:1,beforeValue};
 for(let iteration=1;iteration<=maxIterations;iteration++){
   const denominator=right.f-left.f,step=denominator===0?(right.x-left.x||1):right.f*(right.x-left.x)/denominator,nextX=right.x-step;
   const next=evaluate(nextX);if(isSolved(next))return {changingValue:next.x,achievedValue:raw.targetValue+next.f,iterations:iteration+1,beforeValue};
   left=right;right=next;
 }
 throw Error(`Goal Seek did not converge after ${maxIterations} iterations; the workbook was not changed.`);
}
