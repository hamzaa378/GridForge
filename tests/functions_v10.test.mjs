import test from 'node:test';
import assert from 'node:assert/strict';
import {functions,functionNames,v10FunctionNames,v10Categories} from '../dist/engine/functions.js';
import {categorizedFunctions} from '../dist/engine/categories.js';
import {Workbook} from '../dist/engine/workbook.js';
const mat=[[1,2],[3,4],[5,6]],vec=[[1],[2],[2],[4]];
const sample = name => {
 if(name==='BETA.DIST')return [.5,2,2,true];if(name==='BETA.INV')return [.5,2,2];
 if(name==='BINOM.DIST.RANGE')return [10,.5,3,5];
 if(name==='CHISQ.TEST')return [[[10,20]],[[12,18]]];
 if(name.startsWith('CHISQ.INV'))return [.5,3];
 if(name==='CHISQ.DIST')return [2,3,true];
 if(name.startsWith('CHISQ.'))return [2,3];
 if(name==='CONFIDENCE.T')return [.05,3,20];
 if(name.startsWith('COVARIANCE.'))return [[[1,2,3]],[[3,4,5]]];
 if(name.startsWith('ERF'))return [.3];
 if(name==='F.TEST')return [[[1,2,3,4]],[[3,4,5,6]]];
 if(name==='F.DIST')return [1,4,8,true];
 if(name.startsWith('F.INV'))return [.6,4,8];
 if(name.startsWith('F.'))return [1,4,8];
 if(name==='GAMMA.DIST')return [2,3,2,true];
 if(name==='GAMMA.INV')return [.5,3,2];
 if(name==='HYPGEOM.DIST')return [1,3,4,12,false];
 if(name==='MODE.MULT')return [[[1,2,2,3,3]]];
 if(name==='T.TEST')return [[[1,2,3,4]],[[3,4,5,6]],2,2];
 if(name==='T.DIST')return [1,7,true];
 if(name.startsWith('T.INV'))return [.5,7];
 if(name.startsWith('T.DIST'))return [1,7];
 if(name.startsWith('FORECAST.ETS'))return name.endsWith('.STAT')||name.endsWith('.SEASONALITY')?[5,[[2],[4],[6],[8]]]:[5,[[2],[4],[6],[8]],[[1],[2],[3],[4]]];
 if(name==='SEQUENCE')return [3,2];if(name==='RANDARRAY')return [2,3];if(name==='MUNIT')return [3];
 if(name==='TAKE'||name==='DROP')return [mat,1];
 if(name==='CHOOSECOLS'||name==='CHOOSEROWS')return [mat,1];
 if(name==='HSTACK'||name==='VSTACK')return [mat,mat];
 if(name==='TOCOL'||name==='TOROW'||name==='FLATTEN'||name==='TRIMRANGE')return [mat];
 if(name==='WRAPCOLS'||name==='WRAPROWS')return [vec,2];
 if(name==='EXPAND')return [mat,5,3,0];
 if(name==='FILTER')return [mat,[[1],[0],[1]]];
 if(name==='SORT')return [mat,1,-1];
 if(name==='SORTBY')return [mat,[[2],[3],[1]]];
 if(name==='UNIQUE')return [vec];
 if(name==='TEXTSPLIT')return ['a,b,c',','];
 if(name==='BYROW'||name==='BYCOL')return [mat,'SUM'];
 if(name==='MAKEARRAY')return [3,3,'PRODUCT'];
 if(name==='MAP')return [mat,'ABS'];
 if(name==='REDUCE'||name==='SCAN')return [0,vec,'SUM'];
 if(name==='LET')return ['x',2,2];
 if(name==='LAMBDA')return ['x',2];
 if(name==='ISOMITTED')return [null];
 if(name==='GROUPBY')return [[['A'],['B'],['A']],[[2],[3],[5]],'SUM'];
 if(name==='PIVOTBY')return [[['A'],['B'],['A']],[['P'],['P'],['Q']],[[2],[3],[5]],'SUM'];
 if(name.startsWith('COUP'))return [45000,47000,2];
 if(name==='INTRATE')return [45000,46000,1000,1100];
 if(name==='ACCRINTM')return [45000,45300,.05];
 if(name==='ACCRINT')return [45000,45300,45500,.05];
 if(name.startsWith('AMOR'))return [1200,45000,45200,100,1,.1];
 if(name==='ODDFPRICE')return [45000,47000,44800,45200,.04,.05,100,2];if(name==='ODDFYIELD')return [45000,47000,44800,45200,.04,95,100,2];if(name==='ODDLPRICE')return [45000,47000,44800,.04,.05,100,2];if(name==='ODDLYIELD')return [45000,47000,44800,.04,95,100,2];
 if(name==='EUROCONVERT')return [100,'EUR','DEM'];
 if(name==='ASC')return ['Ｈｅｌｌｏ'];if(name==='JIS')return ['Hello'];
 if(name==='PHONETIC')return ['Good morning'];
 if(name==='BAHTTEXT')return [12];
 if(name==='FILTERXML')return ['<item>ok</item>','item'];
 if(name==='TRANSLATE')return ['hello','en','es'];
 if(name==='INFO')return ['release'];
 if(name==='SHEET'||name==='SHEETS')return [];
 if(name==='DETECTLANGUAGE')return ['Hello'];
 if(name==='FIELDVALUE')return ['cost=500;region=North','cost'];
 if(name==='IMAGE')return ['https://example.org/sample.png','Example'];
 if(name.startsWith('BESSEL'))return [1,0];
 if(name==='PERCENTOF')return [[[2],[3]],[[2],[3],[5]]];
 if(name==='AGGREGATE')return [9,4,[[1,2,3]]];
 if(name==='CUBESETCOUNT')return ['A,B,C'];
 if(name==='ISBETWEEN')return [5,1,10];
 throw Error('Missing v10 fixture for '+name);
};
test('v10: exactly 100 unique functions, 506 overall, categorized without duplicates',()=>{
 assert.equal(v10FunctionNames.length,100);assert.equal(new Set(v10FunctionNames).size,100);
 assert.equal(functionNames.length,506);
 assert.equal(Object.values(categorizedFunctions).flat().length,506);
 assert.equal(new Set(Object.values(categorizedFunctions).flat()).size,506);
 for(const name of v10FunctionNames)assert.ok(functions[name] && functions[name].run);
});
for(const name of v10FunctionNames){
 test(`v10 calculation: ${name}`,()=>{
  const result=functions[name].run(sample(name));
  if(Array.isArray(result))assert.ok(result.length>0,'Matrix must contain rows');
  else if(typeof result==='number')assert.ok(Number.isFinite(result),'Numeric result must be finite');
  else assert.ok(['string','boolean'].includes(typeof result));
 });
}
const expected=[
 ['BETA.DIST',[.5,2,2,true],.5,1e-8],
 ['BETA.INV',[.5,2,2],.5,1e-8],
 ['BINOM.DIST.RANGE',[4,.5,2,3],.625,1e-8],
 ['COVARIANCE.P',[[[1,2,3]],[[3,4,5]]],2/3,1e-10],
 ['F.DIST',[1,5,10,true],0.534880573,5e-5],
 ['GAMMA.DIST',[2,3,2,true],0.0803014,1e-4],
 ['CHISQ.DIST',[2,2,true],1-Math.exp(-1),1e-7],
 ['T.DIST',[0,5,true],.5,1e-8],
 ['COUPNUM',[45000,47000,2],11,1],
 ['EUROCONVERT',[100,'EUR','DEM'],195.583,1e-8],
 ['PERCENTOF',[[[2,3]],[[2,3,5]]],.5,1e-8],
 ['SEQUENCE',[2,2],[[1,2],[3,4]],0],
 ['CHOOSECOLS',[mat,2],[[2],[4],[6]],0],
 ['FILTER',[mat,[[1],[0],[1]]],[[1,2],[5,6]],0],
 ['UNIQUE',[vec],[[1],[2],[4]],0],
 ['TEXTSPLIT',['A|B|C','|'],[['A','B','C']],0],
 ['BYROW',[mat,'SUM'],[[3],[7],[11]],0],
 ['MAKEARRAY',[2,2,'PRODUCT'],[[1,2],[2,4]],0],
 ['ISBETWEEN',[3,1,5],true,0],
 ['DETECTLANGUAGE',['नमस्ते'],'hi',0],
];
for(const [name,args,wanted,tolerance] of expected){
 test(`v10 known result: ${name}`,()=>{
  const got=functions[name].run(args);
  if(typeof wanted==='number')assert.ok(typeof got==='number'&&Math.abs(got-wanted)<=tolerance,`${name} got ${got}; expected ${wanted}`);
  else assert.deepEqual(got,wanted);
 });
}
test('v10 formulas use the custom parser and INDEX for matrix access',()=>{
 const w=new Workbook(400,52);
 const examples={A1:'=INDEX(SEQUENCE(3,2),2,2)',A2:'=BETA.DIST(0.5,2,2,TRUE)',A3:'=INDEX(TEXTSPLIT("alpha,beta",","),1,2)',A4:'=SUM(INDEX(FILTER(A10:B11,A12:A13),1,2),1)',A10:'10',B10:'20',A11:'30',B11:'40',A12:'1',A13:'0'};
 w.applyEdits(Object.entries(examples).map(([ref,raw])=>({ref,raw})),false);
 assert.equal(w.getValue('A1').value,4);
 assert.ok(Math.abs(w.getValue('A2').value-.5)<1e-7);
 assert.equal(w.getValue('A3').value,'beta');
 assert.equal(w.getValue('A4').value,21);
});
test('v10 bounds and errors are reported, never crash the sheet',()=>{
 const w=new Workbook(400,52);
 for(const [ref,formula] of Object.entries({A1:'=SEQUENCE(999999,999999)',A2:'=CHOOSECOLS(A1:B2,999)',A3:'=FILTER(A10:B11,A10:A10)',A4:'=EUROCONVERT(5,"EUR","USD")'}))w.setCell(ref,formula);
 for(const ref of ['A1','A2','A3','A4'])assert.ok(w.getValue(ref).error,ref+' should return a structured error');
});

test('v10 bugfix: DROP only removes rows when columns argument is omitted',()=>{
 assert.deepEqual(functions.DROP.run([[[1,2],[3,4],[5,6]],1]),[[3,4],[5,6]]);
 assert.deepEqual(functions.DROP.run([[[1,2],[3,4]],1,1]),[[4]]);
});
test('v10 RANDARRAY accepts Excel five-argument whole-number flag',()=>{
 const result=functions.RANDARRAY.run([4,3,7,7,true]);
 assert.ok(Array.isArray(result));
 assert.equal(result.length,4);
 result.forEach(row=>assert.deepEqual(row,[7,7,7]));
});

