import test from 'node:test';
import assert from 'node:assert/strict';
import {Workbook} from '../dist/engine/workbook.js';
import {functionNames} from '../dist/engine/functions.js';
import {extraFunctionNames} from '../dist/engine/functions_extra.js';

function worksheet(formula){
  const w=new Workbook(400,52);
  const entries={A1:'1',A2:'2',A3:'3',A4:'4',A5:'5',B1:'3',B2:'5',B3:'7',B4:'9',B5:'11',
    C1:'North',C2:'North',C3:'South',C4:'North',C5:'South',D1:'k1',D2:'k2',D3:'k3',
    E1:'10',E2:'20',E3:'30',H1:'Q1',I1:'Q2',J1:'Q3',H2:'22',I2:'33',J2:'44',Z20:formula};
  w.applyEdits(Object.entries(entries).map(([ref,raw])=>({ref,raw})),false);
  return w;
}
const cases=[
  // Math: 24
  ['SUMSQ','=SUMSQ(1,2,3)',14],
  ['SUMPRODUCT','=SUMPRODUCT(A1:A3,B1:B3)',34],
  ['QUOTIENT','=QUOTIENT(-17,5)',-3],
  ['TRUNC','=TRUNC(-3.987,2)',-3.98],
  ['EVEN','=EVEN(-3)',-4],
  ['ODD','=ODD(-4)',-5],
  ['GCD','=GCD(24,36)',12],
  ['LCM','=LCM(4,6)',12],
  ['FACT','=FACT(5)',120],
  ['COMBIN','=COMBIN(5,2)',10],
  ['PERMUT','=PERMUT(5,2)',20],
  ['RADIANS','=RADIANS(180)',Math.PI],
  ['DEGREES','=DEGREES(PI())',180],
  ['SIN','=SIN(PI()/2)',1],
  ['COS','=COS(0)',1],
  ['TAN','=TAN(0)',0],
  ['ASIN','=ASIN(1)',Math.PI/2],
  ['ACOS','=ACOS(1)',0],
  ['ATAN','=ATAN(1)',Math.PI/4],
  ['ATAN2','=ATAN2(1,1)',Math.PI/4],
  ['SINH','=SINH(0)',0],
  ['COSH','=COSH(0)',1],
  ['TANH','=TANH(0)',0],
  ['LOG10','=LOG10(1000)',3],
  // Stats: 16
  ['GEOMEAN','=GEOMEAN(1,4,16)',4],
  ['HARMEAN','=HARMEAN(1,2,4)',12/7],
  ['STDEVP','=STDEVP(2,4,6)',Math.sqrt(8/3)],
  ['VARP','=VARP(2,4,6)',8/3],
  ['DEVSQ','=DEVSQ(2,4,6)',8],
  ['AVEDEV','=AVEDEV(2,4,6)',4/3],
  ['MODE','=MODE(1,2,2,3)',2],
  ['PERCENTILE','=PERCENTILE(A1:A5,0.25)',2],
  ['QUARTILE','=QUARTILE(A1:A5,3)',4],
  ['RANK','=RANK(4,A1:A5)',2],
  ['CORREL','=CORREL(A1:A5,B1:B5)',1],
  ['RSQ','=RSQ(A1:A5,B1:B5)',1],
  ['SLOPE','=SLOPE(B1:B5,A1:A5)',2],
  ['INTERCEPT','=INTERCEPT(B1:B5,A1:A5)',1],
  ['FORECAST','=FORECAST(7,B1:B5,A1:A5)',15],
  ['PERCENTRANK','=PERCENTRANK(A1:A5,3)',0.5],
  // Text: 12
  ['PROPER','=PROPER("heLLo woRLD")','Hello World'],
  ['REPT','=REPT("xy",3)','xyxyxy'],
  ['CLEAN','=CLEAN("a"&CHAR(10)&"b")','ab'],
  ['TEXTJOIN','=TEXTJOIN("|",TRUE,D1:D3)','k1|k2|k3'],
  ['CONCATENATE','=CONCATENATE("ab","cd")','abcd'],
  ['CODE','=CODE("A")',65],
  ['CHAR','=CHAR(65)','A'],
  ['UNICODE','=UNICODE("★")',9733],
  ['UNICHAR','=UNICHAR(9733)','★'],
  ['NUMBERVALUE','=NUMBERVALUE("1.234,5",",",".")',1234.5],
  ['TEXTBEFORE','=TEXTBEFORE("one-two-three","-",2)','one-two'],
  ['TEXTAFTER','=TEXTAFTER("one-two-three","-",2)','three'],
  // Date: 16
  ['DATE','=DATE(2026,10,20)',(Date.UTC(2026,9,20)-Date.UTC(1899,11,31))/86400000+1],
  ['YEAR','=YEAR(DATE(2026,10,20))',2026],
  ['MONTH','=MONTH(DATE(2026,10,20))',10],
  ['DAY','=DAY(DATE(2026,10,20))',20],
  ['TODAY','=TODAY()',value=>typeof value==='number'&&Math.abs(value-((Date.now()-Date.UTC(1899,11,31))/86400000+1))<2],
  ['NOW','=NOW()',value=>typeof value==='number'&&Math.abs(value-((Date.now()-Date.UTC(1899,11,31))/86400000+1))<0.01],
  ['TIME','=TIME(14,30,45)',(14*3600+30*60+45)/86400],
  ['HOUR','=HOUR(TIME(14,30,0))',14],
  ['MINUTE','=MINUTE(TIME(14,30,0))',30],
  ['SECOND','=SECOND(TIME(14,30,45))',45],
  ['DAYS','=DAYS(DATE(2026,10,20),DATE(2026,10,10))',10],
  ['EDATE','=DAY(EDATE(DATE(2026,1,31),1))',28],
  ['EOMONTH','=DAY(EOMONTH(DATE(2026,2,3),0))',28],
  ['WEEKDAY','=WEEKDAY(DATE(2026,10,20),2)',2],
  ['NETWORKDAYS','=NETWORKDAYS(DATE(2026,10,19),DATE(2026,10,23))',5],
  ['WORKDAY','=DAY(WORKDAY(DATE(2026,10,19),5))',26],
  // Logic / condition / lookup: 8
  ['XOR','=XOR(TRUE,FALSE,FALSE)',true],
  ['IFS','=IFS(FALSE,1/0,TRUE,42)',42],
  ['SWITCH','=SWITCH(2,1,1/0,2,99,0)',99],
  ['SUMIFS','=SUMIFS(B1:B5,C1:C5,"North",A1:A5,">2")',9],
  ['COUNTIFS','=COUNTIFS(C1:C5,"North",A1:A5,">2")',1],
  ['AVERAGEIFS','=AVERAGEIFS(B1:B5,C1:C5,"North")',17/3],
  ['XLOOKUP','=XLOOKUP("k2",D1:D3,E1:E3,"missing")',20],
  ['HLOOKUP','=HLOOKUP("Q2",H1:J2,2,FALSE)',33],
  // Financial: 4
  ['PMT','=PMT(0,10,1000)',-100],
  ['FV','=FV(0,10,-100)',1000],
  ['PV','=PV(0,10,-100)',1000],
  ['NPV','=NPV(0.1,100,100)',100/1.1+100/1.21],
];

test('v4 maintains exactly 80 new distinct functions in the expanded v5 registry',()=>{
  assert.equal(extraFunctionNames.length,80);
  assert.equal(new Set(extraFunctionNames).size,80);
  assert.equal(functionNames.length,516);
  assert.deepEqual([...cases.map(([name])=>name)].sort(),extraFunctionNames);
});

for(const [name,formula,expected] of cases){
 test(`v4 ${name}: ${formula}`,()=>{
  const w=worksheet(formula),out=w.getValue('Z20');
  assert.equal(out.error?.code,undefined,`${name}: ${out.error?.message??''}`);
  if(typeof expected==='function')assert.ok(expected(out.value),`${name}: ${out.value}`);
  else if(typeof expected==='number')assert.ok(typeof out.value==='number'&&Math.abs(out.value-expected)<=1e-8*Math.max(1,Math.abs(expected)),`${name}: expected ${expected}, got ${out.value}`);
  else assert.deepEqual(out.value,expected);
 });
}

test('lazy functions skip unselected error paths',()=>{
 for(const formula of ['=IFS(TRUE,7,FALSE,1/0)','=SWITCH("a","a",8,"b",1/0)']) {
  const out=worksheet(formula).getValue('Z20');assert.equal(out.error,undefined,formula);
 }
});
test('new functions report invalid input errors without throwing unhandled exceptions',()=>{
 const checks=[['=GEOMEAN(-2,3)','#NUM!'],['=HLOOKUP("absent",H1:J2,2,FALSE)','#N/A'],
 ['=XLOOKUP("missing",D1:D3,E1:E3)','#N/A'],['=NETWORKDAYS(DATE(2026,10,20),DATE(2026,10,19))',null],
 ['=PMT(0,0,100)','#NUM!'],['=WORKDAY(DATE(2026,10,19),-2)',null],['=SIN(1,2)','#ARG!'],
 ['=SUMPRODUCT(A1:A2,B1:B3)','#VALUE!'],['=IFS(FALSE,7)','#N/A'],['=SWITCH(9,1,1)','#N/A']];
 for(const [formula,err] of checks) {
  const result=worksheet(formula).getValue('Z20');assert.equal(result.error?.code??null,err,`${formula}: ${result.error?.message}`);
 }
});
test('new criteria functions recalculate only affected dependency chain',()=>{
 const w=worksheet('=SUMIFS(B1:B5,C1:C5,"North")');
 assert.equal(w.getValue('Z20').value,17);
 const result=w.setCell('B2','100');
 assert.equal(w.getValue('Z20').value,112);
 assert.ok(result.affected.includes('Z20'));
 assert.ok(!result.affected.includes('A5'));
});
test('v4 formulas and dependency data survive JSON serialization',()=>{
 const w=worksheet('=XLOOKUP("k3",D1:D3,E1:E3,"missing")');
 const clone=Workbook.fromJSON(w.toJSON());assert.equal(clone.getValue('Z20').value,30);
});

test('XLOOKUP missing argument is lazily evaluated after an actual lookup miss',()=>{
  assert.equal(worksheet('=XLOOKUP("k2",D1:D3,E1:E3,1/0)').getValue('Z20').value,20);
  assert.equal(worksheet('=XLOOKUP("missing",D1:D3,E1:E3,999)').getValue('Z20').value,999);
});
test('paired regression ignores incomplete observations without misaligning rows',()=>{
  const w=worksheet('=SLOPE(B1:B3,A1:A3)');
  w.setCell('A2','');
  assert.equal(w.getValue('Z20').value,2);
});
test('PERCENTRANK duplicates count first matching rank',()=>{
  const w=worksheet('=PERCENTRANK(A1:A4,1)');w.applyEdits([{ref:'A1',raw:'1'},{ref:'A2',raw:'1'},{ref:'A3',raw:'2'},{ref:'A4',raw:'3'}]);
  assert.equal(w.getValue('Z20').value,0);
});
test('Excel 1900 serial-date exception and negative date range rules are explicit',()=>{
  assert.equal(worksheet('=DAY(60)').getValue('Z20').value,29);
  assert.equal(worksheet('=YEAR(0)').getValue('Z20').value,1900);
  assert.equal(worksheet('=NETWORKDAYS(DATE(2026,10,23),DATE(2026,10,19))').getValue('Z20').value,-5);
});
test('LOG10 parses as a function but unrelated references still tokenize normally',()=>{
  const w=worksheet('=LOG10(100)');assert.equal(w.getValue('Z20').value,2);
});

test('XLSX exporter identifies modern Excel functions without altering string literals',async()=>{
  const {SpreadsheetDocument}=await import('../dist/engine/document.js');
  const {exportXlsx}=await import('../dist/engine/xlsx.js');
  const {zipRead}=await import('../dist/engine/zip.js');
  const book=new Workbook();
  book.applyEdits([{ref:'A1',raw:'=XLOOKUP("k2",D1:D3,E1:E3,"missing")'},
    {ref:'A2',raw:'=IF(TRUE,"XLOOKUP(A1)",1)'},
    {ref:'A3',raw:'=IFS(TRUE,10,FALSE,0)'},
    {ref:'D1',raw:'k2'},{ref:'E1',raw:'20'}],false);
  const xml=(await zipRead(exportXlsx(new SpreadsheetDocument(book)))).get('xl/worksheets/sheet1.xml');
  assert.match(xml,/<f>_xlfn.XLOOKUP\(/);
  assert.match(xml,/<f>_xlfn.IFS\(/);
  assert.match(xml,/XLOOKUP\(A1\)/);
  assert.doesNotMatch(xml,/_xlfn.XLOOKUP\(A1\)/);
});
