import type {CellResult} from './workbook.js';
import {display} from './workbook.js';
export type NumberFormat='general'|'number'|'integer'|'currency'|'percent';
export type CellAlign='left'|'center'|'right';
export interface CellStyle { bold?:boolean; italic?:boolean; underline?:boolean; color?:string; background?:string; align?:CellAlign; numberFormat?:NumberFormat }
export type StyleEdit={ref:string;style:CellStyle};
const validNumbers=new Set(['general','number','integer','currency','percent']);
const validAlign=new Set(['left','center','right']);
export function cleanStyle(value:unknown):CellStyle {
  if(!value||typeof value!=='object'||Array.isArray(value))throw new Error('Cell formatting must be an object.');
  const source=value as Record<string,unknown>, out:CellStyle={};
  for(const key of ['bold','italic','underline'] as const)if(source[key]!==undefined){if(typeof source[key]!=='boolean')throw new Error(`Invalid ${key} formatting.`);if(source[key])out[key]=true;}
  for(const key of ['color','background'] as const)if(source[key]!==undefined){if(typeof source[key]!=='string'||!/^#[0-9a-fA-F]{6}$/.test(source[key]))throw new Error(`Invalid ${key}: use a six-digit hex color.`);out[key]=source[key];}
  if(source.align!==undefined){if(typeof source.align!=='string'||!validAlign.has(source.align))throw new Error('Invalid alignment.');if(source.align!=='left')out.align=source.align as CellAlign;}
  if(source.numberFormat!==undefined){if(typeof source.numberFormat!=='string'||!validNumbers.has(source.numberFormat))throw new Error('Invalid numeric format.');if(source.numberFormat!=='general')out.numberFormat=source.numberFormat as NumberFormat;}
  return out;
}
export function styledDisplay(value:CellResult,style:CellStyle):string {
  if(value.error||typeof value.value!=='number'||!Number.isFinite(value.value))return display(value);
  const n=value.value;
  switch(style.numberFormat){
    case 'integer':return new Intl.NumberFormat('en-US',{maximumFractionDigits:0}).format(n);
    case 'number':return new Intl.NumberFormat('en-US',{minimumFractionDigits:2,maximumFractionDigits:2}).format(n);
    case 'currency':return new Intl.NumberFormat('en-US',{style:'currency',currency:'USD'}).format(n);
    case 'percent':return new Intl.NumberFormat('en-US',{style:'percent',minimumFractionDigits:1,maximumFractionDigits:1}).format(n);
    default:return display(value);
  }
}
