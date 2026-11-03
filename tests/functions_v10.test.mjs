import test from 'node:test';
import assert from 'node:assert/strict';
import {functions,functionNames,v10FunctionNames,v10Categories,connectedFunctionNames} from '../dist/engine/functions.js';
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
 if(name==='FORECAST.ETS.SEASONALITY')return [[[2],[4],[6],[8]],[[1],[2],[3],[4]]];
 if(name==='FORECAST.ETS.STAT')return [[[2],[4],[6],[8]],[[1],[2],[3],[4]],6];
 if(name.startsWith('FORECAST.ETS'))return [5,[[2],[4],[6],[8]],[[1],[2],[3],[4]]];
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
test('historical v10 functions remain registered and categorized after later expansions',()=>{
 assert.ok(v10FunctionNames.length>=99);assert.equal(new Set(v10FunctionNames).size,v10FunctionNames.length);
 assert.ok(functionNames.length>=516);
 assert.equal(Object.values(categorizedFunctions).flat().length,516);
 assert.equal(new Set(Object.values(categorizedFunctions).flat()).size,516);
 for(const name of v10FunctionNames)assert.ok(functions[name] && functions[name].run);
 assert.ok(connectedFunctionNames.includes('CUBESETCOUNT'));
});
const workbookRouted=new Set(['BYROW','BYCOL','MAKEARRAY','MAP','REDUCE','SCAN','LET','LAMBDA','ISOMITTED','GROUPBY','PIVOTBY','FIELDVALUE']);
for(const name of v10FunctionNames){
 if(workbookRouted.has(name))continue;
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
test('workbook-context dynamic array and LAMBDA functions calculate and spill',()=>{
 const w=new Workbook(100,20),entries={A1:'1',B1:'2',A2:'3',B2:'4',A3:'5',B3:'6',C1:'=INDEX(BYROW(A1:B3,LAMBDA(row,SUM(row))),2,1)',C2:'=INDEX(BYCOL(A1:B3,LAMBDA(col,SUM(col))),1,2)',C3:'=INDEX(MAKEARRAY(2,2,LAMBDA(r,c,r*c)),2,2)',C4:'=INDEX(MAP(A1:A3,LAMBDA(x,x*2)),3,1)',C5:'=REDUCE(0,A1:A3,LAMBDA(acc,x,acc+x))',C6:'=INDEX(SCAN(0,A1:A3,LAMBDA(acc,x,acc+x)),3,1)',C7:'=LET(x,2,x+2)',C8:'=LAMBDA(x,x+1)(2)',C9:'=LAMBDA(x,ISOMITTED(x))()'};
 w.applyEdits(Object.entries(entries).map(([ref,raw])=>({ref,raw})),false);
 assert.deepEqual(['C1','C2','C3','C4','C5','C6','C7','C8','C9'].map(ref=>w.getValue(ref).value),[7,12,4,10,9,9,4,3,true]);
});
test('HSTACK and VSTACK expose missing array elements as per-cell #N/A errors',()=>{
 const w=new Workbook(20,20);w.applyEdits([
  {ref:'A1',raw:'1'},{ref:'A2',raw:'2'},{ref:'A3',raw:'3'},
  {ref:'C1',raw:"'x"},{ref:'C2',raw:"'y"},
  {ref:'E1',raw:'=HSTACK(A1:A3,C1:C2)'},{ref:'H1',raw:'=VSTACK(A1:A1,C1:D2)'},
 ],false);
 assert.equal(w.getValue('F1').value,'x');assert.equal(w.getValue('F2').value,'y');
 assert.equal(w.getValue('F3').error?.code,'#N/A');
 assert.equal(w.getValue('H1').value,1);assert.equal(w.getValue('I1').error?.code,'#N/A');
 assert.equal(w.getValue('H2').value,'x');assert.equal(w.getValue('H3').value,'y');
});
test('EXPAND default padding is #N/A and IFERROR/IFNA replace padded spill cells',()=>{
 const w=new Workbook(20,20);w.applyEdits([{ref:'A1',raw:'7'},{ref:'C1',raw:'=EXPAND(A1,2,2)'},{ref:'F1',raw:'=IFERROR(EXPAND(A1,2,2),"pad")'},{ref:'I1',raw:'=IFNA(EXPAND(A1,2,2),"missing")'}],false);
 assert.equal(w.getValue('C1').value,7);assert.equal(w.getValue('D1').error?.code,'#N/A');assert.equal(w.getValue('C2').error?.code,'#N/A');
 assert.deepEqual(['F1','G1','F2','G2'].map(ref=>w.getValue(ref).value),[7,'pad','pad','pad']);
 assert.deepEqual(['I1','J1','I2','J2'].map(ref=>w.getValue(ref).value),[7,'missing','missing','missing']);
});
test('TOCOL ignore-errors removes HSTACK padding while no ignore option preserves cell errors',()=>{
 const w=new Workbook(20,16);w.applyEdits([{ref:'A1',raw:'1'},{ref:'A2',raw:'2'},{ref:'B1',raw:'3'},{ref:'D1',raw:'=TOCOL(HSTACK(A1:A2,B1:B1))'},{ref:'F1',raw:'=TOCOL(HSTACK(A1:A2,B1:B1),2)'}],false);
 assert.equal(w.getValue('D1').value,1);
 assert.equal(w.getValue('D4').error?.code,'#N/A');
 assert.deepEqual(['F1','F2','F3'].map(ref=>w.getValue(ref).value),[1,3,2]);
});
test('element-wise arithmetic preserves array-cell errors and IFERROR repairs them',()=>{
 const w=new Workbook(20,20);w.applyEdits([{ref:'A1',raw:'1'},{ref:'A2',raw:'2'},{ref:'B1',raw:'3'},{ref:'D1',raw:'=HSTACK(A1:A2,B1:B1)+1'},{ref:'G1',raw:'=IFERROR(HSTACK(A1:A2,B1:B1)+1,0)'}],false);
 assert.deepEqual(['D1','E1','D2'].map(ref=>w.getValue(ref).value),[2,4,3]);assert.equal(w.getValue('E2').error?.code,'#N/A');
 assert.deepEqual(['G1','H1','G2','H2'].map(ref=>w.getValue(ref).value),[2,4,3,0]);
});
test('counting functions treat padded #N/A cells as errors rather than blanks or numbers',()=>{
 const stacked=functions.HSTACK.run([[[1],[2]],[[3]]]);
 assert.equal(functions.COUNT.run([stacked]),3);assert.equal(functions.COUNTA.run([stacked]),4);assert.equal(functions.COUNTBLANK.run([stacked]),0);
});
test('ISERROR, ISERR, and ISNA classify errors independently inside spilled arrays',()=>{
 const w=new Workbook(20,20);w.applyEdits([
  {ref:'A1',raw:'1'},{ref:'A2',raw:'2'},{ref:'A3',raw:'3'},{ref:'B1',raw:'=1/0'},{ref:'B2',raw:"'text"},
  {ref:'D1',raw:'=ISERROR(HSTACK(A1:A3,B1:B2))'},
  {ref:'G1',raw:'=ISERR(HSTACK(A1:A3,B1:B2))'},
  {ref:'J1',raw:'=ISNA(HSTACK(A1:A3,B1:B2))'},
  {ref:'M1',raw:'=ISERROR(HSTACK(A1:A2,1/0))'},
  {ref:'O1',raw:'=ISERR(HSTACK(A1:A2,1/0))'},
  {ref:'Q1',raw:'=ISNA(HSTACK(A1:A2,1/0))'},
 ],false);
 assert.deepEqual(['D1','E1','D2','E2','D3','E3'].map(ref=>w.getValue(ref).value),[false,true,false,false,false,true]);
 assert.deepEqual(['G1','H1','G2','H2','G3','H3'].map(ref=>w.getValue(ref).value),[false,true,false,false,false,false]);
 assert.deepEqual(['J1','K1','J2','K2','J3','K3'].map(ref=>w.getValue(ref).value),[false,false,false,false,false,true]);
 assert.deepEqual(['M1','N1','M2','N2'].map(ref=>w.getValue(ref).value),[false,true,false,true]);
 assert.deepEqual(['O1','P1','O2','P2'].map(ref=>w.getValue(ref).value),[false,true,false,false]);
 assert.deepEqual(['Q1','R1','Q2','R2'].map(ref=>w.getValue(ref).value),[false,false,false,true]);
});
test('range arguments retain individual error cells for dynamic arrays while aggregations still propagate them',()=>{
 const w=new Workbook(20,12);w.applyEdits([
  {ref:'A1',raw:'4'},{ref:'A2',raw:'=1/0'},{ref:'A3',raw:'6'},
  {ref:'C1',raw:'=ISERROR(A1:A3)'},{ref:'F1',raw:'=SUM(A1:A3)'},
 ],false);
 assert.deepEqual(['C1','C2','C3'].map(ref=>w.getValue(ref).value),[false,true,false]);
 assert.equal(w.getValue('F1').error?.code,'#DIV/0!');
});
test('FILTER retains a source error only when its row is selected',()=>{
 const w=new Workbook(20,12);w.applyEdits([
  {ref:'A1',raw:'1'},{ref:'A2',raw:'=1/0'},{ref:'A3',raw:'3'},{ref:'A4',raw:'4'},
  {ref:'B1',raw:'=TRUE()'},{ref:'B2',raw:'=TRUE()'},{ref:'B3',raw:'=TRUE()'},{ref:'B4',raw:'=FALSE()'},
  {ref:'D1',raw:'=FILTER(A1:A4,B1:B4)'},
 ],false);
 assert.deepEqual(['D1','D2','D3'].map(ref=>{const cell=w.getValue(ref);return cell.error?.code??cell.value;}),[1,'#DIV/0!',3]);
});
test('IF evaluates range conditions elementwise, spills branch values, and skips unused errors',()=>{
 const w=new Workbook(20,20);w.applyEdits([
  {ref:'A1',raw:'1'},{ref:'A2',raw:'-1'},{ref:'A3',raw:'3'},{ref:'A4',raw:'=NA()'},
  {ref:'C1',raw:'=IF(A1:A4>0,A1:A4,1/0)'},
  {ref:'E1',raw:'=IF(A1:A4>0,"yes","no")'},
  {ref:'G1',raw:'=IF(A1:A4>0,1/0,"safe")'},
  {ref:'I1',raw:'=IF(A1:A4>0,SEQUENCE(2),0)'},
 ],false);
 assert.deepEqual(['C1','C2','C3','C4'].map(ref=>{const cell=w.getValue(ref);return cell.error?.code??cell.value;}),[1,'#DIV/0!',3,'#N/A']);
 assert.deepEqual(['E1','E2','E3','E4'].map(ref=>{const cell=w.getValue(ref);return cell.error?.code??cell.value;}),['yes','no','yes','#N/A']);
 assert.deepEqual(['G1','G2','G3','G4'].map(ref=>{const cell=w.getValue(ref);return cell.error?.code??cell.value;}),['#DIV/0!','safe','#DIV/0!','#N/A']);
 assert.equal(w.getValue('I1').error?.code,'#VALUE!');
});
test('IF rejects text logical tests while preserving numeric, blank, and per-cell error conditions',()=>{
 const w=new Workbook(12,12);w.applyEdits([
  {ref:'A1',raw:'TRUE'},{ref:'A2',raw:'note'},{ref:'A3',raw:'0'},{ref:'A4',raw:''},{ref:'A5',raw:'=NA()'},
  {ref:'B1',raw:'=IF("text",1,0)'},{ref:'B2',raw:'=IF(A2,1,0)'},{ref:'B3',raw:'=IF(A3,1,0)'},{ref:'B4',raw:'=IF(A4,1,0)'},{ref:'B5',raw:'=IFS("text",1,TRUE,2)'},{ref:'B6',raw:'=IFS(0,"zero",1,"one")'},
  {ref:'D1',raw:'=IF(A1:A5,1,0)'},
 ],false);
 assert.equal(w.getValue('B1').error?.code,'#VALUE!');
 assert.equal(w.getValue('B2').error?.code,'#VALUE!');
 assert.equal(w.getValue('B3').value,0);
 assert.equal(w.getValue('B4').value,0);
 assert.equal(w.getValue('B5').error?.code,'#VALUE!');
 assert.equal(w.getValue('B6').value,'one');
 assert.deepEqual(['D1','D2','D3','D4','D5'].map(ref=>{const cell=w.getValue(ref);return cell.error?.code??cell.value;}),[1,'#VALUE!',0,0,'#N/A']);
});
test('IFS evaluates array conditions in order, spills per-cell results, and only calculates selected branches',()=>{
 const w=new Workbook(16,16);w.applyEdits([
  {ref:'A1',raw:'1'},{ref:'A2',raw:'3'},{ref:'A3',raw:'-1'},{ref:'A4',raw:'=NA()'},
  {ref:'B1',raw:'=IFS(A1:A4>2,"high",A1:A4>0,"low",TRUE,"none")'},
  {ref:'D1',raw:'=IFS(A1:A3>0,1/0,TRUE,7)'},
  {ref:'F1',raw:'=IFS(A1:A3>0,SEQUENCE(2),TRUE,0)'},
  {ref:'H1',raw:'=IFS(FALSE,1/0,A1:A3>0,"yes",TRUE,"no")'},
  {ref:'K1',raw:'1'},{ref:'K2',raw:'=1/0'},{ref:'K3',raw:'3'},
  {ref:'L1',raw:'=IFS(A1:A3>0,K1:K3,TRUE,0)'},
  {ref:'M1',raw:'A'},{ref:'M2',raw:'A'},{ref:'M3',raw:'B'},
  {ref:'N1',raw:'=SWITCH(M1:M3,"A",K1:K3,0)'},
 ],false);
 assert.deepEqual(['B1','B2','B3','B4'].map(ref=>{const cell=w.getValue(ref);return cell.error?.code??cell.value;}),['low','high','none','#N/A']);
 assert.deepEqual(['D1','D2','D3'].map(ref=>{const cell=w.getValue(ref);return cell.error?.code??cell.value;}),['#DIV/0!','#DIV/0!',7]);
 assert.equal(w.getValue('F1').error?.code,'#VALUE!');
 assert.deepEqual(['H1','H2','H3'].map(ref=>{const cell=w.getValue(ref);return cell.error?.code??cell.value;}),['yes','yes','no']);
 assert.deepEqual(['L1','L2','L3'].map(ref=>{const cell=w.getValue(ref);return cell.error?.code??cell.value;}),[1,'#DIV/0!',0]);
 assert.deepEqual(['N1','N2','N3'].map(ref=>{const cell=w.getValue(ref);return cell.error?.code??cell.value;}),[1,'#DIV/0!',0]);
});
test('SWITCH spills matching array expressions, preserves expression errors, and evaluates selected results only',()=>{
 const w=new Workbook(16,16);w.applyEdits([
  {ref:'A1',raw:'A'},{ref:'A2',raw:'B'},{ref:'A3',raw:'C'},{ref:'A4',raw:'=NA()'},
  {ref:'B1',raw:'=SWITCH(A1:A4,"A",1,"B",2,99)'},
  {ref:'D1',raw:'=SWITCH(A1:A3,"A",1/0,"B",2,99)'},
  {ref:'F1',raw:'=SWITCH(A1:A3,"A",SEQUENCE(2),99)'},
 ],false);
 assert.deepEqual(['B1','B2','B3','B4'].map(ref=>{const cell=w.getValue(ref);return cell.error?.code??cell.value;}),[1,2,99,'#N/A']);
 assert.deepEqual(['D1','D2','D3'].map(ref=>{const cell=w.getValue(ref);return cell.error?.code??cell.value;}),['#DIV/0!',2,99]);
 assert.equal(w.getValue('F1').error?.code,'#VALUE!');
});
test('ISNUMBER, ISTEXT, ISLOGICAL, ISNONTEXT, NOT, N, T, and ISBLANK evaluate reference arrays and spilled cells positionally',()=>{
 const w=new Workbook(20,26);w.applyEdits([
  {ref:'A1',raw:'1'},{ref:'A2',raw:'note'},{ref:'A3',raw:'=1/0'},{ref:'A4',raw:'TRUE'},
  {ref:'C1',raw:'=ISNUMBER(A1:A4)'},{ref:'E1',raw:'=ISTEXT(A1:A4)'},{ref:'G1',raw:'=NOT(A1:A4)'},{ref:'I1',raw:'=ISNUMBER(1/0)'},
  {ref:'K1',raw:'=""'},{ref:'L1',raw:'=ISBLANK(K1:K2)'},{ref:'N1',raw:'=SEQUENCE(2)'},{ref:'P1',raw:'=ISBLANK(N1:N2)'},{ref:'R1',raw:'=ISBLANK(N2)'},{ref:'T1',raw:'=N(A1:A4)'},{ref:'V1',raw:'=T(A1:A4)'},{ref:'X1',raw:'=ISLOGICAL(A1:A4)'},{ref:'Z1',raw:'=ISNONTEXT(A1:A4)'},
 ],false);
 assert.deepEqual(['C1','C2','C3','C4'].map(ref=>w.getValue(ref).value),[true,false,false,false]);
 assert.deepEqual(['E1','E2','E3','E4'].map(ref=>w.getValue(ref).value),[false,true,false,false]);
 assert.deepEqual(['G1','G2','G3','G4'].map(ref=>{const cell=w.getValue(ref);return cell.error?.code??cell.value;}),[false,'#VALUE!','#DIV/0!',false]);
 assert.equal(w.getValue('I1').value,false);
 assert.deepEqual(['L1','L2'].map(ref=>w.getValue(ref).value),[false,true]);
 assert.deepEqual(['P1','P2'].map(ref=>w.getValue(ref).value),[false,false]);
 assert.equal(w.getValue('R1').value,false);
 assert.deepEqual(['T1','T2','T3','T4'].map(ref=>{const cell=w.getValue(ref);return cell.error?.code??cell.value;}),[1,0,'#DIV/0!',1]);
 assert.deepEqual(['V1','V2','V3','V4'].map(ref=>{const cell=w.getValue(ref);return cell.error?.code??cell.value;}),['','note','#DIV/0!','']);
 assert.deepEqual(['X1','X2','X3','X4'].map(ref=>w.getValue(ref).value),[false,false,false,true]);
 assert.deepEqual(['Z1','Z2','Z3','Z4'].map(ref=>w.getValue(ref).value),[true,false,true,true]);
});
test('MAP, BYROW, BYCOL, and SCAN keep calculation errors at the affected array positions',()=>{
 const w=new Workbook(20,20);w.applyEdits([
  {ref:'A1',raw:'1'},{ref:'A2',raw:'=1/0'},{ref:'A3',raw:'3'},
  {ref:'B1',raw:'1'},{ref:'B2',raw:'4'},{ref:'B3',raw:'5'},
  {ref:'D1',raw:'=MAP(A1:A3,LAMBDA(x,x+1))'},
  {ref:'F1',raw:'=BYROW(A1:B3,LAMBDA(row,SUM(row)))'},
  {ref:'I1',raw:'=BYCOL(A1:B3,LAMBDA(col,SUM(col)))'},
  {ref:'L1',raw:'=SCAN(0,A1:A3,LAMBDA(acc,x,acc+x))'},
  {ref:'O1',raw:'=REDUCE(0,A1:A3,LAMBDA(acc,x,acc+x))'},
  {ref:'R1',raw:'=MAKEARRAY(2,2,LAMBDA(r,c,1/(r-c)))'},
 ],false);
 assert.equal(w.getValue('D1').value,2);assert.equal(w.getValue('D2').error?.code,'#DIV/0!');assert.equal(w.getValue('D3').value,4);
 assert.equal(w.getValue('F1').value,2);assert.equal(w.getValue('F2').error?.code,'#DIV/0!');assert.equal(w.getValue('F3').value,8);
 assert.equal(w.getValue('I1').error?.code,'#DIV/0!');assert.equal(w.getValue('J1').value,10);
 assert.deepEqual(['L1','L2','L3'].map(ref=>w.getValue(ref).error?.code??w.getValue(ref).value),[1,'#DIV/0!','#DIV/0!']);
 assert.equal(w.getValue('O1').error?.code,'#DIV/0!');
 assert.deepEqual(['R1','S1','R2','S2'].map(ref=>w.getValue(ref).error?.code??w.getValue(ref).value),['#DIV/0!',-1,1,'#DIV/0!']);
});
test('FIELDVALUE reports that linked Excel data types are unavailable',()=>{
 assert.throws(()=>functions.FIELDVALUE.run(sample('FIELDVALUE')),error=>error.code==='#FIELD!');
});
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
test('SORT and SORTBY keep blank cells last in either sort direction',()=>{
 const column=[[3],[null],[1],[2]],row=[[3,null,1,2]];
 assert.deepEqual(functions.SORT.run([column,1,1]),[[1],[2],[3],[null]]);
 assert.deepEqual(functions.SORT.run([column,1,-1]),[[3],[2],[1],[null]]);
 assert.deepEqual(functions.SORT.run([row,1,1,true]),[[1,2,3,null]]);
 assert.deepEqual(functions.SORT.run([row,1,-1,true]),[[3,2,1,null]]);
 assert.deepEqual(functions.SORTBY.run([column,column]),[[1],[2],[3],[null]]);
 assert.deepEqual(functions.SORTBY.run([column,column,-1]),[[3],[2],[1],[null]]);
});
test('SORT orders numeric-looking text lexically and keeps it distinct from numbers',()=>{
 const textValues=[['2'],['10'],['1']],mixed=[[2],['10'],[1],['2']];
 assert.deepEqual(functions.SORT.run([textValues]),[['1'],['10'],['2']]);
 assert.deepEqual(functions.SORT.run([textValues,1,-1]),[['2'],['10'],['1']]);
 assert.deepEqual(functions.SORT.run([mixed]),[[1],[2],['10'],['2']]);
 assert.deepEqual(functions.SORTBY.run([textValues,[['2'],['10'],['1']]]),[['1'],['10'],['2']]);
});
test('SORT dynamic-array spills keep blank source cells at the end in descending order',()=>{
 const w=new Workbook(20,8);w.applyEdits([{ref:'A1',raw:'3'},{ref:'A3',raw:'1'},{ref:'A4',raw:'2'},{ref:'C1',raw:'=SORT(A1:A4,1,-1)'}],false);
 assert.deepEqual(['C1','C2','C3','C4'].map(ref=>w.getValue(ref).value),[3,2,1,null]);
});
test('SORT preserves error cells, treats different errors as tied, and leaves blanks last',()=>{
 const w=new Workbook(20,12);w.applyEdits([
  {ref:'A1',raw:'3'},{ref:'A2',raw:'=1/0'},{ref:'A3',raw:'1'},{ref:'A4',raw:'=NA()'},
  {ref:'B1',raw:"'x1"},{ref:'B2',raw:"'x2"},{ref:'B3',raw:"'x3"},{ref:'B4',raw:"'x4"},{ref:'B5',raw:"'x5"},
  {ref:'C1',raw:'=SORT(A1:A5)'},{ref:'E1',raw:'=SORT(A1:A5,1,-1)'},
  {ref:'G1',raw:'=SORTBY(B1:B5,A1:A5)'},{ref:'I1',raw:'=SORTBY(B1:B5,A1:A5,-1)'},
 ],false);
 assert.deepEqual(['C1','C2','C3','C4','C5'].map(ref=>{const cell=w.getValue(ref);return cell.error?.code??cell.value;}),[1,3,'#DIV/0!','#N/A',null]);
 assert.deepEqual(['E1','E2','E3','E4','E5'].map(ref=>{const cell=w.getValue(ref);return cell.error?.code??cell.value;}),['#DIV/0!','#N/A',3,1,null]);
 assert.deepEqual(['G1','G2','G3','G4','G5'].map(ref=>{const cell=w.getValue(ref);return cell.error?.code??cell.value;}),['x3','x1','x2','x4','x5']);
 assert.deepEqual(['I1','I2','I3','I4','I5'].map(ref=>{const cell=w.getValue(ref);return cell.error?.code??cell.value;}),['x2','x4','x1','x3','x5']);
});
test('v10 RANDARRAY accepts Excel five-argument whole-number flag',()=>{
 const result=functions.RANDARRAY.run([4,3,7,7,true]);
 assert.ok(Array.isArray(result));
 assert.equal(result.length,4);
 result.forEach(row=>assert.deepEqual(row,[7,7,7]));
});

