import type {Workbook} from './workbook.js';
import {parseRef,colToName} from './parser.js';
import type {PageSetupSpec} from './sheettools.js';
import {cleanStyle,styledDisplay,type CellStyle} from './format.js';

const escape=(value:string)=>value.replace(/[&<>"']/g,ch=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[ch]!));
function printCellCss(style:CellStyle):string {
 const rules:string[]=[];if(style.fontFamily)rules.push(`font-family:"${style.fontFamily.replace(/[\\"]/g,'\\$&')}"`);if(style.fontSize)rules.push(`font-size:${style.fontSize}pt`);if(style.bold)rules.push('font-weight:700');if(style.italic)rules.push('font-style:italic');const decorations=[style.underline?'underline':'',style.strike?'line-through':''].filter(Boolean);if(decorations.length)rules.push(`text-decoration:${decorations.join(' ')}`);if(style.color)rules.push(`color:${style.color}`);if(style.background)rules.push(`background-color:${style.background}`);if(style.align)rules.push(`text-align:${style.align}`);if(style.verticalAlign)rules.push(`vertical-align:${style.verticalAlign}`);rules.push(`white-space:${style.wrap?'pre-wrap':'pre'}`);if(style.indent){const amount=Math.min(25,style.indent)*0.1;rules.push(style.align==='right'?`padding-right:${amount}in`:`padding-left:${amount}in`);}
 const width=(weight:string)=>weight==='thick'?'2.25pt':weight==='medium'?'1.25pt':'0.5pt',line=(edge:NonNullable<CellStyle['borderEdges']>[keyof NonNullable<CellStyle['borderEdges']>])=>edge?`${width(edge.weight)} ${edge.style} ${edge.color??'#000000'}`:undefined;
 if(style.borderEdges){for(const edge of ['left','right','top','bottom'] as const){const value=line(style.borderEdges[edge]);if(value)rules.push(`border-${edge}:${value}`);}}
 else if(style.border){const active=style.border==='all'?['left','right','top','bottom']: [style.border],value=`${width(style.borderWeight??'thin')} ${style.borderStyle??'solid'} ${style.borderColor??'#000000'}`;for(const edge of active)rules.push(`border-${edge}:${value}`);}
 if(style.borderDiagonal){const diagonal=style.borderDiagonal,color=diagonal.color??'#000000',lines:string[]=[];if(diagonal.up)lines.push(`linear-gradient(to top right,transparent 49%,${color} 50%,transparent 51%)`);if(diagonal.down)lines.push(`linear-gradient(to bottom right,transparent 49%,${color} 50%,transparent 51%)`);if(lines.length)rules.push(`background-image:${lines.join(',')}`);}
 return rules.join(';');
}
function printCell(book:Workbook,ref:string,header=false):string {const style=cleanStyle(book.getStyle(ref)),css=printCellCss(style),tag=header?'th':'td';return `<${tag}${css?` style="${escape(css)}"`:''}>${escape(styledDisplay(book.getValue(ref),style))}</${tag}>`;}
function headerText(value:string,sheetName:string):string{return escape(value).replace(/&amp;P/g,'<span class="page-number"></span>').replace(/&amp;N/g,'<span class="page-count"></span>').replace(/&amp;A/g,escape(sheetName)).replace(/&amp;D/g,new Intl.DateTimeFormat().format(new Date()));}

/** Build a self-contained, print-ready worksheet view for the browser print dialog. */
export function buildPrintHtml(book:Workbook,setup:PageSetupSpec,sheetName:string,autoPrint=true):string{
 let start=parseRef('A1')!,end={row:0,col:0};
 if(setup.printArea){const [a,b]=setup.printArea.split(':'),first=parseRef(a!),last=parseRef(b??a!);if(!first||!last)throw Error('Print area is invalid.');start=first;end=last;}
 else for(const {ref} of book.allEntries()){const point=parseRef(ref);if(point){end.row=Math.max(end.row,point.row);end.col=Math.max(end.col,point.col);}}
 let repeatStart=-1,repeatEnd=-1;
 if(setup.repeatRows){const match=/^([1-9]\d*):([1-9]\d*)$/.exec(setup.repeatRows);if(!match||Number(match[1])>Number(match[2])||Number(match[2])>1_048_576)throw Error('Print title rows are invalid. Use a whole-row range such as 1:3.');repeatStart=Number(match[1])-1;repeatEnd=Number(match[2])-1;}
 const columnCount=end.col-start.col+1,repeatedCount=repeatStart<0?0:repeatEnd-repeatStart+1,repeatedInsideCount=repeatStart<0?0:Math.max(0,Math.min(end.row,repeatEnd)-Math.max(start.row,repeatStart)+1),bodyRowCount=end.row-start.row+1-repeatedInsideCount;
 if((bodyRowCount+repeatedCount)*columnCount>12000)throw Error('The print area and repeated title rows are larger than 12,000 cells. Set a smaller print area or repeat fewer rows.');
 const repeatedRows=repeatStart<0?[]:Array.from({length:repeatedCount},(_,index)=>repeatStart+index);
 const paper=({letter:'Letter',legal:'Legal',a3:'A3',a4:'A4',tabloid:'Tabloid'} as const)[setup.paper];
 const headings=setup.printHeadings,grid=setup.printGridlines;
 const headerValues=[setup.headerLeft??'',setup.headerCenter??'',setup.headerRight??''],footerValues=[setup.footerLeft??'',setup.footerCenter??'',setup.footerRight??''];
 const header=headerValues.some(Boolean)?`<header class="print-band print-header">${headerValues.map(value=>`<span>${headerText(value,sheetName)}</span>`).join('')}</header>`:'',footer=footerValues.some(Boolean)?`<footer class="print-band print-footer">${footerValues.map(value=>`<span>${headerText(value,sheetName)}</span>`).join('')}</footer>`:'';
 const columnHeadingRow=headings?`<tr><th></th>${Array.from({length:columnCount},(_,i)=>`<th>${colToName(start.col+i)}</th>`).join('')}</tr>`:'';
 // Browser printing can repeat vertical table headers safely; horizontal repeated columns require page slicing and remain in XLSX output only.
 const repeatedHeaderRows=repeatedRows.map(row=>`<tr>${headings?`<th>${row+1}</th>`:''}${Array.from({length:columnCount},(_,c)=>printCell(book,`${colToName(start.col+c)}${row+1}`,true)).join('')}</tr>`).join('');
 const tableHeader=columnHeadingRow||repeatedHeaderRows?`<thead>${columnHeadingRow}${repeatedHeaderRows}</thead>`:'';
 const repeatedSet=new Set(repeatedRows),manualRowBreaks=new Set(setup.rowBreaks??[]),body=Array.from({length:end.row-start.row+1},(_,r)=>{const row=start.row+r;if(repeatedSet.has(row))return '';return `<tr${manualRowBreaks.has(row)?' class="manual-page-break"':''}>${headings?`<th>${row+1}</th>`:''}${Array.from({length:columnCount},(_,c)=>printCell(book,`${colToName(start.col+c)}${row+1}`)).join('')}</tr>`;}).join('');
 return `<!doctype html><html><head><meta charset="utf-8"><title>${escape(sheetName)} — Print</title><style>@page{size:${paper} ${setup.orientation};margin:${setup.marginTop}in ${setup.marginRight}in ${setup.marginBottom}in ${setup.marginLeft}in}*{box-sizing:border-box}body{font:10pt Calibri,Arial,sans-serif;color:#111;margin:0}h1{font-size:15pt;margin:0 0 10pt}table{border-collapse:collapse;width:100%;table-layout:auto}thead{display:table-header-group}td,th{padding:3pt 5pt;${grid?'border:1px solid #aab2bd;':'border:0;'}text-align:left;vertical-align:top;white-space:pre-wrap;overflow-wrap:anywhere}thead th,tbody th{background:#eef2f7;font-weight:600}tr.manual-page-break{break-before:page;page-break-before:always}.print-band{position:fixed;left:0;right:0;display:grid;grid-template-columns:1fr 1fr 1fr;gap:8pt;font-size:8pt;color:#444}.print-band span:nth-child(2){text-align:center}.print-band span:nth-child(3){text-align:right}.print-header{top:-0.45in}.print-footer{bottom:-0.45in}.page-number:after{content:counter(page)}.page-count:after{content:counter(pages)}${setup.fitWidth?`table{max-width:100%}body{zoom:${Math.max(0.25,Math.min(1,1/(setup.fitWidth||1)))} }`:''}</style></head><body>${header}<h1>${escape(sheetName)}</h1><table>${tableHeader}<tbody>${body}</tbody></table>${footer}${autoPrint?'<script>addEventListener("load",()=>setTimeout(()=>print(),200))<\\/script>':''}</body></html>`;
}
