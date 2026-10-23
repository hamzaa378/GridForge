import test from 'node:test';
import assert from 'node:assert/strict';
import {Workbook} from '../dist/engine/workbook.js';
import {SpreadsheetDocument} from '../dist/engine/document.js';
import {functions,functionNames,v5FunctionNames} from '../dist/engine/functions.js';
import {categorizedFunctions,categoryNames} from '../dist/engine/categories.js';
import {validateMacro,runActionMacro} from '../dist/engine/macros.js';
import {exportXlsx,importXlsx} from '../dist/engine/xlsx.js';

function evalFormula(formula){
 const book=new Workbook(400,52);
 const edits={A1:'1',A2:'2',A3:'3',A4:'4',A5:'5',B1:'3',B2:'5',B3:'7',B4:'9',B5:'11',
 F1:'-1000',F2:'300',F3:'400',F4:'500',F5:'200',
 G1:'=DATE(2020,1,1)',G2:'=DATE(2021,1,1)',G3:'=DATE(2022,1,1)',G4:'=DATE(2023,1,1)',G5:'=DATE(2024,1,1)',Z20:formula};
 book.applyEdits(Object.entries(edits).map(([ref,raw])=>({ref,raw})),false);
 return book.getValue('Z20');
}
const almost=(expected)=>(v)=>typeof v==='number'&&Math.abs(v-expected)<=Math.max(1,Math.abs(expected))*1e-5;
const between=(lo,hi)=>(v)=>typeof v==='number'&&v>lo&&v<hi;
const cases=[
 // Math (18)
 ['ACOSH','=ACOSH(1)',0],['ASINH','=ASINH(0)',0],['ATANH','=ATANH(0)',0],
 ['ACOT','=ACOT(1)',Math.PI/4],['ACOTH','=ACOTH(2)',Math.atanh(.5)],
 ['COT','=COT(PI()/4)',1],['COTH','=COTH(1)',1/Math.tanh(1)],
 ['CSC','=CSC(PI()/2)',1],['CSCH','=CSCH(1)',1/Math.sinh(1)],
 ['SEC','=SEC(0)',1],['SECH','=SECH(0)',1],['MROUND','=MROUND(14,5)',15],
 ['FACTDOUBLE','=FACTDOUBLE(7)',105],['MULTINOMIAL','=MULTINOMIAL(2,1)',3],
 ['ARABIC','=ARABIC("XIV")',14],['ROMAN','=ROMAN(14)','XIV'],
 ['BASE','=BASE(255,16)','FF'],['DECIMAL','=DECIMAL("FF",16)',255],
 // Stats (12)
 ['SKEW','=SKEW(A1:A5)',0],['KURT','=KURT(A1:A5)',-1.2],
 ['COVARIANCEP','=COVARIANCEP(A1:A5,B1:B5)',4],
 ['COVARIANCES','=COVARIANCES(A1:A5,B1:B5)',5],
 ['PEARSON','=PEARSON(A1:A5,B1:B5)',1],['STEYX','=STEYX(B1:B5,A1:A5)',0],
 ['STANDARDIZE','=STANDARDIZE(10,8,2)',1],['TRIMMEAN','=TRIMMEAN(A1:A5,0.4)',3],
 ['NORMDIST','=NORMDIST(0,0,1,TRUE)',.5],['NORMSDIST','=NORMSDIST(0)',.5],
 ['NORMINV','=NORMINV(.5,0,1)',0],['NORMSINV','=NORMSINV(.5)',0],
 // Text (14)
 ['T','=T("Hello")','Hello'],['N','=N(TRUE)',1],
 ['DOLLAR','=DOLLAR(1234.5,2)','$1,234.50'],['FIXED','=FIXED(1234.5,1,TRUE)','1234.5'],
 ['TEXT','=TEXT(0.125,"0.0%")','12.5%'],['LENB','=LENB("abc")',3],
 ['LEFTB','=LEFTB("abc",2)','ab'],['RIGHTB','=RIGHTB("abc",2)','bc'],
 ['MIDB','=MIDB("abc",2,2)','bc'],['FINDB','=FINDB("b","abc")',2],
 ['SEARCHB','=SEARCHB("B","abc")',2],['REPLACEB','=REPLACEB("abc",2,1,"Z")','aZc'],
 ['REGEXTEST','=REGEXTEST("ab12","[0-9]+")',true],
 ['REGEXREPLACE','=REGEXREPLACE("ab12","[0-9]+","ZZ")','abZZ'],
 // Date (8)
 ['DATEDIF','=DATEDIF(DATE(2020,1,1),DATE(2022,4,1),"Y")',2],
 ['YEARFRAC','=YEARFRAC(DATE(2020,1,1),DATE(2021,1,1),3)',366/365],
 ['DAYS360','=DAYS360(DATE(2026,1,1),DATE(2026,4,1))',90],
 ['WEEKNUM','=WEEKNUM(DATE(2026,1,1),1)',1],
 ['ISOWEEKNUM','=ISOWEEKNUM(DATE(2021,1,4))',1],
 ['DATEVALUE','=DATEVALUE("2026-10-20")',almost((Date.UTC(2026,9,20)-Date.UTC(1899,11,31))/86400000+1)],
 ['TIMEVALUE','=TIMEVALUE("12:30:00")',12.5/24],
 ['NETWORKDAYSINTL','=NETWORKDAYSINTL(DATE(2026,10,19),DATE(2026,10,23),1)',5],
 // Finance (15)
 ['RATE','=RATE(10,-100,1000)',almost(0)],
 ['NPER','=NPER(0,-100,1000)',10],
 ['IPMT','=IPMT(0.1,1,2,100)',-10],
 ['PPMT','=PPMT(0.1,1,2,100)',-47.619047619],
 ['EFFECT','=EFFECT(0.12,12)',almost((1+.12/12)**12-1)],
 ['NOMINAL','=NOMINAL(EFFECT(0.12,12),12)',almost(.12)],
 ['SLN','=SLN(1000,0,5)',200],
 ['SYD','=SYD(1000,0,5,1)',1000*5*2/(5*6)],
 ['DDB','=DDB(1000,0,5,1)',400],
 ['MIRR','=MIRR(F1:F5,0.1,0.1)',between(0,.5)],
 ['IRR','=IRR(F1:F5)',between(0,.5)],
 ['XNPV','=XNPV(0,F1:F5,G1:G5)',400],
 ['XIRR','=XIRR(F1:F5,G1:G5)',between(0,.5)],
 ['FVSCHEDULE','=FVSCHEDULE(100,A1:A2)',600],
 ['RRI','=RRI(2,100,121)',.1],
 // Information (8)
 ['ISERR','=ISERR(1/0)',true],['ISERROR','=ISERROR(1/0)',true],
 ['ISNA','=ISNA(MATCH(99,A1:A5,0))',true],['ISEVEN','=ISEVEN(4)',true],
 ['ISODD','=ISODD(3)',true],['ISLOGICAL','=ISLOGICAL(TRUE)',true],
 ['ISNONTEXT','=ISNONTEXT(123)',true],['TYPE','=TYPE("hello")',2],
 // Lookup (5)
 ['LOOKUP','=LOOKUP(3,A1:A5,B1:B5)',7],
 ['CHOOSE','=CHOOSE(2,1/0,42)',42],
 ['ROWS','=ROWS(A1:C5)',5],['COLUMNS','=COLUMNS(A1:C5)',3],
 ['ADDRESS','=ADDRESS(4,2,1)','$B$4'],
];

test('v5: exactly 80 unique new functions and 216 overall, categorized by Excel-style families',()=>{
 assert.equal(v5FunctionNames.length,80);
 assert.equal(new Set(v5FunctionNames).size,80);
 assert.equal(functionNames.length,506);
 assert.deepEqual(cases.map(c=>c[0]).sort(),v5FunctionNames);
 assert.ok(categoryNames.includes('Financial'));
 const groups=Object.values(categorizedFunctions).flat();
 assert.equal(groups.length,506,'All functions appear in one and only one group.');
 assert.equal(new Set(groups).size,506);
 for(const name of v5FunctionNames)assert.ok(functions[name],`Missing ${name}`);
});
for(const [name,formula,expected] of cases){
 test(`v5 ${name}: ${formula}`,()=>{
  const {value,error}=evalFormula(formula);
  assert.equal(error,undefined,`${name}: ${error?.code??''} ${error?.message??''}`);
  if(typeof expected==='function')assert.ok(expected(value),`${name}: Unexpected value ${value}`);
  else if(typeof expected==='number')assert.ok(typeof value==='number'&&Math.abs(value-expected)<1e-6*Math.max(1,Math.abs(expected)),`${name}: Expected ${expected}, received ${value}`);
  else assert.deepEqual(value,expected);
 });
}
test('v5: information tests catch errors without evaluating other branches',()=>{
 for(const [formula,expected] of [['=ISERROR(1/0)',true],['=ISERR(MATCH(7,A1:A5,0))',false],['=ISNA(MATCH(7,A1:A5,0))',true],['=CHOOSE(1,5,1/0)',5]]){
  const {value,error}=evalFormula(formula);assert.equal(error,undefined);assert.equal(value,expected);
 }
});
test('v5: financial, date and engineering boundary validation emits spreadsheet errors',()=>{
 const bad=[['=DECIMAL("FG",16)','#NUM!'],['=FACTDOUBLE(-2)','#NUM!'],['=NORMINV(1,0,1)','#NUM!'],['=DATEVALUE("2026-02-30")','#VALUE!'],['=TIMEVALUE("28:00")','#VALUE!'],['=NETWORKDAYSINTL(1,2,"1210000")','#VALUE!'],['=RATE(0,0,0)','#NUM!']];
 for(const [formula,expected]of bad){const {error}=evalFormula(formula);assert.equal(error?.code,expected,formula);}
});
test('v5: action macros apply validated transactions, undo and reject executable instructions',()=>{
 const doc=new SpreadsheetDocument();const sheet=doc.active;
 const macro=validateMacro({version:1,name:'Set total',sheetName:sheet.name,steps:[
  {kind:'cells',edits:[{ref:'A1',raw:'15'},{ref:'B1',raw:'=A1*2'}]},
  {kind:'styles',edits:[{ref:'A1',style:{bold:true,background:'#ff0000'}}]},
 ]});
 const result=runActionMacro(doc,macro);assert.deepEqual(result,{values:2,styles:1});
 assert.equal(sheet.book.getValue('B1').value,30);assert.equal(sheet.book.getStyle('A1').bold,true);
 assert.equal(sheet.book.undo(),true);assert.equal(sheet.book.getRaw('A1'),'');
 for(const malicious of [
  {version:1,name:'X',sheetName:sheet.name,steps:[{kind:'javascript',source:'alert(1)'}]},
  {version:1,name:'X',sheetName:sheet.name,steps:[{kind:'cells',edits:[{ref:'A1',raw:42}]}]},
  {version:1,name:'X',sheetName:sheet.name,steps:Array.from({length:201},()=>({kind:'cells',edits:[]}))},
 ])assert.throws(()=>validateMacro(malicious));
});
test('v5: edited XLSX preserves supported new formulas without introducing executable macros',async()=>{
 const doc=new SpreadsheetDocument();doc.active.book.setCell('A1','=NETWORKDAYSINTL(DATE(2026,10,19),DATE(2026,10,23),1)');
 const bytes=exportXlsx(doc);const payload=new TextDecoder().decode(bytes);
 assert.ok(payload.includes('NETWORKDAYS.INTL'));
 assert.ok(!payload.includes('vbaProject.bin'));
 if(typeof DOMParser!=='undefined'){
   const restored=await importXlsx(bytes);assert.equal(restored.document.active.book.getValue('A1').value,5);
 }
});
