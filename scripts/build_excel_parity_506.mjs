#!/usr/bin/env node
/** Complete registry manifest and reproducible per-function fixture generation.
 * A worksheet-level smoke test per function is NOT equivalent to Excel parity.
 * No Python, Excel or external formula engine is used here.
 */
import fs from 'node:fs';
import {functions,functionNames,connectedFunctionNames} from '../dist/engine/functions.js';
import {functionCategory} from '../dist/engine/categories.js';
import {Workbook} from '../dist/engine/workbook.js';
const ROOT=new URL('../',import.meta.url);
const sources=JSON.parse(fs.readFileSync(new URL('../tests/parity_506_recipe_sources.json',import.meta.url)));
const sourceByName=new Map(sources.map(x=>[x.name,x]));
const offlineFunctionNames=functionNames.filter(name=>!connectedFunctionNames.includes(name));
if(offlineFunctionNames.length!==506||new Set(offlineFunctionNames).size!==506||sourceByName.size!==506||offlineFunctionNames.some(name=>!sourceByName.has(name)))throw Error('Offline registry/fixture coverage is not exactly 506 unique names');
const volatile=new Set(['RAND','RANDBETWEEN','RANDARRAY','NOW','TODAY']);
const nonExcelOrLimited=new Set(['BYROW','BYCOL','MAKEARRAY','MAP','REDUCE','SCAN','LET','LAMBDA','ISOMITTED','GROUPBY','PIVOTBY','FLATTEN','TRIMRANGE','PERCENTOF','DETECTLANGUAGE','TRANSLATE','IMAGE','FIELDVALUE','FILTERXML','FORECAST.ETS','FORECAST.ETS.CONFINT','FORECAST.ETS.SEASONALITY','FORECAST.ETS.STAT','CUBESETCOUNT','INFO','SHEET','SHEETS','BAHTTEXT','PHONETIC']);
const arrays=new Set(['CHOOSECOLS','CHOOSEROWS','DROP','EXPAND','FILTER','FREQUENCY','GROWTH','HSTACK','LINEST','LOGEST','MAKEARRAY','MAP','MDETERM_____','MINVERSE','MMULT','MODE.MULT','MUNIT','RANDARRAY','SCAN','SEQUENCE','SORT','SORTBY','TAKE','TEXTSPLIT','TOCOL','TOROW','TRANSPOSE','TREND','TRIMRANGE','UNIQUE','VSTACK','WRAPCOLS','WRAPROWS','BYROW','BYCOL','GROUPBY','PIVOTBY','FLATTEN']);
const overrides={
 'AGGREGATE':'=AGGREGATE(9,4,A1:A5)', 'MATCH':'=MATCH(3,A1:A5,0)', 'VLOOKUP':'=VLOOKUP("k2",D1:E3,2,FALSE)',
 'AVERAGEIF':'=AVERAGEIF(A1:A5,">2",B1:B5)', 'SUMIF':'=SUMIF(A1:A5,">2",B1:B5)',
 'ISBLANK':'=ISBLANK(L40)', 'ISFORMULA':'=ISFORMULA(L39)', 'EXACT':'=EXACT("A","A")',
 'IF':'=IF(2>1,5,3)', 'IFERROR':'=IFERROR(1/0,7)', 'NA':'=NA()',
 'INDEX':'=INDEX(A1:B5,2,2)', 'COUNTIF':'=COUNTIF(A1:A5,">2")',
 'SUBTOTAL':'=SUBTOTAL(9,A1:A5)', 'CONCAT':'=CONCAT("A","B")',
 'LOWER':'=LOWER("HELLO")', 'UPPER':'=UPPER("hello")', 'LEN':'=LEN("Hello")',
 'LEFT':'=LEFT("Hello",2)', 'RIGHT':'=RIGHT("Hello",2)', 'MID':'=MID("Hello",2,2)',
 'NOT':'=NOT(FALSE)', 'AND':'=AND(TRUE,FALSE)', 'OR':'=OR(TRUE,FALSE)',
 'SHEET':'=SHEET()', 'SHEETS':'=SHEETS()', 'ISOMITTED':'=ISOMITTED(A1)',
 'FLATTEN':'=INDEX(FLATTEN(A1:B2),1,1)',
 'TODAY':'=TODAY()', 'NOW':'=NOW()'
};
const common=stage=>{
 const x={};
 for(let i=1;i<=15;i++){
  x[`A${i}`]=i;
  x[`B${i}`]=i*2+1;
  x[`C${i}`]=i%3===0?'South':'North';
 }
 x.L39='=SUM(1,2)';
 if(stage==='v4'){
  Object.assign(x,{D1:'k1',D2:'k2',D3:'k3',E1:10,E2:20,E3:30,H1:'Q1',I1:'Q2',J1:'Q3',H2:22,I2:33,J2:44});
 }else if(stage==='v9'){
  Object.assign(x,{D1:'Region',E1:'Sales',D2:'East',E2:20,D3:'West',E3:30,D4:'East',E4:40,G1:'Region',G2:'East',H1:1,I1:2,H2:3,I2:4});
  for(let i=1;i<=5;i++)x[`J${i}`]=0.2;
 }else{
  Object.assign(x,{D1:'k1',D2:'k2',D3:'k3',E1:10,E2:20,E3:30,H1:1,I1:2,H2:3,I2:4});
  for(let i=1;i<=5;i++)x[`J${i}`]=0.2;
 }
 if(stage==='v5'){
  Object.assign(x,{F1:-1000,F2:300,F3:400,F4:500,F5:200,G1:{formula:'=DATE(2020,1,1)'},G2:{formula:'=DATE(2021,1,1)'},G3:{formula:'=DATE(2022,1,1)'},G4:{formula:'=DATE(2023,1,1)'},G5:{formula:'=DATE(2024,1,1)'}});
 }
 if(stage==='v6'){
   for(let i=1;i<=5;i++)x[`D${i}`]=i%2?'East':'West';
   x.E2={formula:'=SUM(A1:A2)'};
 }
 return x;
};
function argFormula(arg, inputs, serial){
 if(arg===null)return 'L40';
 if(typeof arg==='string')return '"'+arg.replaceAll('"','""')+'"';
 if(typeof arg==='boolean')return arg?'TRUE':'FALSE';
 if(typeof arg==='number')return String(arg);
 if(Array.isArray(arg)){
  const col=13+serial*3; // M,P,S,V,Y,AB,AE,AH ... max index 52
  const toCol=n=>{let r=''; for(let c=n;c;c=Math.floor((c-1)/26))r=String.fromCharCode(65+(c-1)%26)+r;return r};
  const c0=toCol(col); const wide=Array.isArray(arg[0]) ? arg : [arg];
  const nrows=wide.length, ncols=Math.max(...wide.map(r=>Array.isArray(r)?r.length:1));
  if(col+ncols-1>48 || nrows>30)throw Error('Argument fixture out of bounds');
  for(let r=0;r<nrows;r++)for(let c=0;c<ncols;c++){
   const cell=wide[r]?.[c]; if(cell!==null && cell!==undefined) inputs[`${toCol(col+c)}${1+r}`]=cell;
  }
  return `${c0}1:${toCol(col+ncols-1)}${nrows}`;
 }
 throw Error('Unsupported fixture argument '+JSON.stringify(arg));
}
const records=[];
for(const name of offlineFunctionNames){
 const recipe=sourceByName.get(name);
 const inputs=common(recipe.stage);
 let formula=overrides[name]||recipe.formula;
 if(recipe.args){const args=recipe.args.map((x,i)=>argFormula(x,inputs,i));formula=`=${name}(${args.join(',')})`;}
 if(!formula || !formula.startsWith('='))throw Error(`Bad fixture ${name}`);
 if(!new RegExp('(?:^|[^A-Z0-9.])'+name.replaceAll('.','\\.')+'\\s*\\(','i').test(formula)){
   throw Error(`Function name ${name} missing in formula ${formula}`);
 }
 const matrixOutput=arrays.has(name);
 // Excel dynamic arrays display the top-left result in an ordinary grid cell.
 // Wrapping in INDEX makes that scalar observable without falsely claiming spill parity.
 if(matrixOutput&&!/^=INDEX\(/i.test(formula)){formula=`=INDEX(${formula.slice(1)},1,1)`;}
 const scope=volatile.has(name)?'volatile':nonExcelOrLimited.has(name)?'limited-semantics':matrixOutput?'array-index':'scalar';
 records.push({id:`F${String(records.length+1).padStart(4,'0')}`,function:name,category:functionCategory(name),stage:recipe.stage,scope,formula,inputs});
}
const output={schema:'gridforge-excel-full-registry/1',functionCount:506,comparisonCell:'Z20',description:'One auditable formula invocation per registered function. Not proof of behavior beyond each fixture.',cases:records};
fs.writeFileSync(new URL('../tests/excel_parity_506_cases.json',import.meta.url),JSON.stringify(output,null,2)+'\n');
const results=[];
for(const c of records){
 const book=new Workbook(400,52);
 const edits=Object.entries(c.inputs).map(([ref,val])=>({ref,raw:val?.formula??(val===null?'':typeof val==='boolean'?(val?'TRUE':'FALSE'):typeof val==='string'?"'"+val:String(val))}));
 edits.push({ref:output.comparisonCell,raw:c.formula});
 try{book.applyEdits(edits,false);const r=book.getValue(output.comparisonCell);
 const value=r.error?{type:'error',value:r.error.code,message:r.error.message}:r.value===null?{type:'blank',value:null}:{type:typeof r.value,value:r.value};
 results.push({...c,gridforge:value});
 }catch(e){results.push({...c,gridforge:{type:'exception',value:String(e)}});}
}
const metrics={registered:506,executed:results.length,successful:results.filter(x=>!['error','exception'].includes(x.gridforge.type)).length,errors:results.filter(x=>x.gridforge.type==='error').length,exceptions:results.filter(x=>x.gridforge.type==='exception').length,scope:Object.fromEntries([...new Set(records.map(x=>x.scope))].map(s=>[s,records.filter(x=>x.scope===s).length]))};
fs.writeFileSync(new URL('../docs/PARITY_506_GRIDFORGE_RESULTS.json',import.meta.url),JSON.stringify({schema:'gridforge-excel-full-result/1',engine:'GridForge current source',measuredExcel:false,metrics,cases:results},null,2)+'\n');
console.log(JSON.stringify(metrics,null,2));
for(const x of results.filter(x=>['error','exception'].includes(x.gridforge.type)))console.log(x.function,x.formula,'->',x.gridforge.type,x.gridforge.value);
