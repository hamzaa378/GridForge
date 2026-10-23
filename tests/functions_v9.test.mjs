import test from 'node:test';
import assert from 'node:assert/strict';
import {Workbook} from '../dist/engine/workbook.js';
import {functionNames,v9FunctionNames} from '../dist/engine/functions.js';
import {categorizedFunctions,functionCategory} from '../dist/engine/categories.js';
import {formulas} from './v9_examples.mjs';
const fixture=[...Array.from({length:5},(_,i)=>({ref:`A${i+1}`,raw:String(i+1)})),...Array.from({length:5},(_,i)=>({ref:`B${i+1}`,raw:String(i+2)})),...Array.from({length:5},(_,i)=>({ref:`J${i+1}`,raw:'.2'})),
 {ref:'H1',raw:'1'},{ref:'I1',raw:'2'},{ref:'H2',raw:'3'},{ref:'I2',raw:'4'},
 {ref:'D1',raw:'Region'},{ref:'E1',raw:'Sales'},{ref:'D2',raw:'East'},{ref:'E2',raw:'20'},{ref:'D3',raw:'West'},{ref:'E3',raw:'30'},{ref:'D4',raw:'East'},{ref:'E4',raw:'40'},{ref:'G1',raw:'Region'},{ref:'G2',raw:'East'}];

test('v9 registry has exactly 100 genuinely new function names and categories',()=>{
 assert.equal(v9FunctionNames.length,100);
 assert.equal(functionNames.length,506);
 assert.deepEqual(Object.keys(formulas).sort(),v9FunctionNames);
 assert.equal(Object.values(categorizedFunctions).flat().length,506);
 assert.equal(new Set(Object.values(categorizedFunctions).flat()).size,506);
 for(const name of v9FunctionNames)assert.notEqual(functionCategory(name),'',name);
});
for(const [name,formula] of Object.entries(formulas))test(`v9 ${name} executes with worksheet values`,()=>{
 const w=new Workbook(400,52);w.applyEdits([...fixture,{ref:'Z20',raw:formula}],false);
 const r=w.getValue('Z20');assert.equal(r.error,undefined,`${name}: ${formula}: ${r.error?.code} ${r.error?.message}`);
 assert.notEqual(r.value,null,`${name}: empty result`);
 if(typeof r.value==='number')assert.ok(Number.isFinite(r.value),`${name}: non-finite result`);
});
test('v9 selected numerical results agree with known examples',()=>{
 const expected={BIN2DEC:10,DEC2HEX:'FF',HEX2DEC:255,IMABS:5,IMREAL:3,IMAGINARY:4,COMPLEX:'3+4i',DSUM:60,DCOUNT:2,DCOUNTA:2,DPRODUCT:800,DMAX:40,DMIN:20,MDETERM:-2,MMULT:7,MINVERSE:-2,CONVERT:1.609344,CRITBINOM:5,SUBTOTAL:15,PROB:.6,BETADIST:.5,BETAINV:.5,CHITEST:null,TRUE:true,FALSE:false,'BINOM.INV':5};
 for(const [name,value] of Object.entries(expected)){
  if(value===null)continue;
  const w=new Workbook();w.applyEdits([...fixture,{ref:'Z20',raw:formulas[name]}],false);const result=w.getValue('Z20');assert.equal(result.error,undefined,name);
  if(typeof value==='number')assert.ok(Math.abs(result.value-value)<1e-6,`${name}: ${result.value} != ${value}`);else assert.equal(result.value,value);
 }
});
test('v9 error diagnostics reject invalid security, matrix and complex inputs',()=>{
 for(const [formula,code] of [['=MINVERSE(H1:I1)','#VALUE!'],['=IMDIV("1+i","0")','#NUM!'],['=BIN2DEC("201")','#NUM!'],['=CONVERT(1,"m","kg")','#N/A'],['=TINV(-1,4)','#NUM!'],['=DEC2BIN(512)','#NUM!'],['=SUBTOTAL(12,A1:A5)','#VALUE!']]){
  const w=new Workbook();w.applyEdits([...fixture,{ref:'Z20',raw:formula}],false);assert.equal(w.getValue('Z20').error?.code,code,formula);
 }
});
