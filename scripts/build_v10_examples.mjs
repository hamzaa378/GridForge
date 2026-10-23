#!/usr/bin/env node
/** v10 review workbook: all 100 registered new names and tested live examples. */
import {writeFileSync} from 'node:fs';
import {Workbook} from '../dist/engine/workbook.js';
import {SpreadsheetDocument} from '../dist/engine/document.js';
import {functionCategory} from '../dist/engine/categories.js';
import {functions,v10FunctionNames} from '../dist/engine/functions.js';
import {exportXlsx} from '../dist/engine/xlsx.js';
const examples=[
 ['BETA.DIST','=BETA.DIST(0.5,2,2,TRUE)'],
 ['BETA.INV','=BETA.INV(0.5,2,2)'],
 ['CHISQ.DIST','=CHISQ.DIST(2,2,TRUE)'],
 ['ERF.PRECISE','=ERF.PRECISE(0.5)'],
 ['F.DIST','=F.DIST(1,5,10,TRUE)'],
 ['GAMMA.DIST','=GAMMA.DIST(2,3,2,TRUE)'],
 ['T.DIST','=T.DIST(1,7,TRUE)'],
 ['SEQUENCE','=INDEX(SEQUENCE(3,2),2,2)'],
 ['RANDARRAY','=INDEX(RANDARRAY(2,2,5,5),1,1)'],
 ['MUNIT','=INDEX(MUNIT(3),2,2)'],
 ['TEXTSPLIT','=INDEX(TEXTSPLIT("red,blue",","),1,2)'],
 ['EXPAND','=INDEX(EXPAND(A2:B3,3,3,0),3,3)'],
 ['TAKE','=INDEX(TAKE(A2:B4,2),1,2)'],
 ['DROP','=INDEX(DROP(A2:B4,1),1,2)'],
 ['SORT','=INDEX(SORT(A2:B4,1,-1),1,1)'],
 ['UNIQUE','=INDEX(UNIQUE(A2:A4),1,1)'],
 ['HSTACK','=INDEX(HSTACK(A2:A4,B2:B4),1,2)'],
 ['VSTACK','=INDEX(VSTACK(A2:A4,B2:B4),4,1)'],
 ['MAP','=INDEX(MAP(J2:K4,"ABS"),1,2)'],
 ['BYROW','=INDEX(BYROW(J2:K4,"SUM"),1,1)'],
 ['MAKEARRAY','=INDEX(MAKEARRAY(2,2,"PRODUCT"),2,2)'],
 ['REDUCE','=REDUCE(0,J2:J4,"SUM")'],
 ['SCAN','=INDEX(SCAN(0,J2:J4,"SUM"),3,1)'],
 ['ISBETWEEN','=ISBETWEEN(5,1,10)'],
 ['DETECTLANGUAGE','=DETECTLANGUAGE("Hello")'],
 ['EUROCONVERT','=EUROCONVERT(100,"EUR","DEM")'],
 ['PERCENTOF','=PERCENTOF(J2:J3,J2:J4)'],
 ['INFO','=INFO("release")'],
 ['CUBESETCOUNT','=CUBESETCOUNT("A,B,C")'],
 ['FILTERXML','=FILTERXML("<item>ready</item>","item")'],
];
const book=new Workbook(400,52),edits=[];
for(const [r,vals]of [[1,['New v10 name','Category','Description','Syntax / example']],[1,[]]]){
 vals.forEach((v,i)=>edits.push({ref:`${'ABCD'[i]}${r}`,raw:v}));
}
for(const [i,name]of v10FunctionNames.entries()){
 const r=i+2,info=functions[name];
 edits.push({ref:`A${r}`,raw:name},{ref:`B${r}`,raw:functionCategory(name)},{ref:`C${r}`,raw:info.description},{ref:`D${r}`,raw:"'"+info.example});
}
for(const [r,v]of [[1,'Function'],[2,'Formula source'],[3,'Evaluated result']])edits.push({ref:`${'FGH'[r-1]}1`,raw:v});
for(let i=0;i<3;i++){edits.push({ref:`J${i+2}`,raw:String(i+2)},{ref:`K${i+2}`,raw:String((i+2)*10)});}
for(const [i,[name,formula]]of examples.entries()){
 const row=i+2;edits.push({ref:`F${row}`,raw:name},{ref:`G${row}`,raw:"'"+formula},{ref:`H${row}`,raw:formula});
}
book.applyEdits(edits,false);
for(const [i,[name,formula]]of examples.entries()){
 const ref=`H${i+2}`,result=book.getValue(ref);
 if(result.error)throw Error(`Example ${name} ${formula}: ${result.error.code} ${result.error.message}`);
}
book.patchStyles(['A1','B1','C1','D1','F1','G1','H1'],{bold:true,background:'#1d536b',color:'#ffffff'});
const doc=new SpreadsheetDocument(book);doc.renameSheet(doc.activeId,'100 New Functions v10');doc.active.tools.freezeRows=1;
writeFileSync(new URL('../examples/v10_100_function_reference.json',import.meta.url),JSON.stringify(doc.toJSON(),null,2)+'\n');
writeFileSync(new URL('../examples/v10_100_function_reference.xlsx',import.meta.url),exportXlsx(doc));
console.log(`Generated v10 workbook: ${v10FunctionNames.length} function descriptions and ${examples.length} verified live formula demonstrations.`);
