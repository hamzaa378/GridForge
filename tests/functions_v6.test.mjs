import test from 'node:test';
import assert from 'node:assert/strict';
import {Workbook} from '../dist/engine/workbook.js';
import {functionNames,v6FunctionNames,v6Categories} from '../dist/engine/functions.js';
import {functionCategory,categorizedFunctions} from '../dist/engine/categories.js';
const inputs=[...Array.from({length:5},(_,i)=>({ref:`A${i+1}`,raw:String(i+1)})),...Array.from({length:5},(_,i)=>({ref:`B${i+1}`,raw:String(i+2)})),...['East','West','East','West','East'].map((raw,i)=>({ref:`D${i+1}`,raw})),{ref:'E2',raw:'=SUM(A1:A2)'}];
const run=formula=>{const b=new Workbook(400,52);b.applyEdits([...inputs,{ref:'Z20',raw:formula}],false);return b.getValue('Z20');};
const close=(a,b)=>Math.abs(a-b)<=2e-5*Math.max(1,Math.abs(b));
const cases=[
['SQRTPI','=SQRTPI(4)',Math.sqrt(4*Math.PI)],['SUMX2MY2','=SUMX2MY2(A1:A2,B1:B2)',-8],['SUMX2PY2','=SUMX2PY2(A1:A2,B1:B2)',18],['SUMXMY2','=SUMXMY2(A1:A2,B1:B2)',2],
['COMBINA','=COMBINA(3,2)',6],['PERMUTATIONA','=PERMUTATIONA(3,2)',9],['BITAND','=BITAND(6,3)',2],['BITOR','=BITOR(6,3)',7],['BITXOR','=BITXOR(6,3)',5],['BITLSHIFT','=BITLSHIFT(3,2)',12],['BITRSHIFT','=BITRSHIFT(16,2)',4],['DELTA','=DELTA(4,4)',1],['GESTEP','=GESTEP(7,5)',1],['ERF','=ERF(0)',0],['ERFC','=ERFC(0)',1],['CEILING.MATH','=CEILING.MATH(7.2,2)',8],['FLOOR.MATH','=FLOOR.MATH(7.2,2)',6],['CEILING.PRECISE','=CEILING.PRECISE(-4.1,2)',-4],['FLOOR.PRECISE','=FLOOR.PRECISE(-4.1,2)',-6],['ISO.CEILING','=ISO.CEILING(5.2,2)',6],['GAMMA','=GAMMA(5)',24],['GAMMALN','=GAMMALN(5)',Math.log(24)],['GAMMALN.PRECISE','=GAMMALN.PRECISE(5)',Math.log(24)],['SERIESSUM','=SERIESSUM(2,1,2,A1:A3)',114],
['MAXIFS','=MAXIFS(A1:A5,D1:D5,"East")',5],['MINIFS','=MINIFS(A1:A5,D1:D5,"East")',1],['AVERAGEA','=AVERAGEA(TRUE,2)',1.5],['MAXA','=MAXA(-5,-1)',-1],['MINA','=MINA(1,3)',1],['STDEV.S','=STDEV.S(A1:A5)',Math.sqrt(2.5)],['STDEV.P','=STDEV.P(A1:A5)',Math.sqrt(2)],['VAR.S','=VAR.S(A1:A5)',2.5],['VAR.P','=VAR.P(A1:A5)',2],['MODE.SNGL','=MODE.SNGL(1,2,2,3)',2],['PERCENTILE.INC','=PERCENTILE.INC(A1:A5,0.75)',4],['PERCENTILE.EXC','=PERCENTILE.EXC(A1:A5,0.5)',3],['PERCENTRANK.INC','=PERCENTRANK.INC(A1:A5,3)',0.5],['PERCENTRANK.EXC','=PERCENTRANK.EXC(A1:A5,3)',0.5],['QUARTILE.INC','=QUARTILE.INC(A1:A5,1)',2],['QUARTILE.EXC','=QUARTILE.EXC(A1:A5,1)',1.5],['RANK.EQ','=RANK.EQ(4,A1:A5)',2],['RANK.AVG','=RANK.AVG(2,A1:A5)',4],['FORECAST.LINEAR','=FORECAST.LINEAR(6,B1:B5,A1:A5)',7],['SKEW.P','=SKEW.P(A1:A3)',0],['FISHER','=FISHER(0)',0],['FISHERINV','=FISHERINV(0)',0],['BINOM.DIST','=BINOM.DIST(2,4,0.5,FALSE)',.375],['POISSON.DIST','=POISSON.DIST(0,2,FALSE)',Math.exp(-2)],['EXPON.DIST','=EXPON.DIST(2,0.5,TRUE)',1-Math.exp(-1)],['WEIBULL.DIST','=WEIBULL.DIST(2,1,2,TRUE)',1-Math.exp(-1)],['NORM.DIST','=NORM.DIST(0,0,1,TRUE)',.5],['NORM.INV','=NORM.INV(0.5,0,1)',0],['NORM.S.DIST','=NORM.S.DIST(0,FALSE)',1/Math.sqrt(2*Math.PI)],['NORM.S.INV','=NORM.S.INV(0.5)',0],['LOGNORM.DIST','=LOGNORM.DIST(1,0,1,TRUE)',.5],['LOGNORM.INV','=LOGNORM.INV(0.5,0,1)',1],['CONFIDENCE.NORM','=CONFIDENCE.NORM(0.05,2,100)',.3919927969],['PHI','=PHI(0)',1/Math.sqrt(2*Math.PI)],['GAUSS','=GAUSS(0)',0],['Z.TEST','=Z.TEST(A1:A5,3)',.5],
['ENCODEURL','=ENCODEURL("a b")','a%20b'],['ARRAYTOTEXT','=ARRAYTOTEXT(A1:A2)','1, 2'],['VALUETOTEXT','=VALUETOTEXT("cat",1)','"cat"'],['REGEXEXTRACT','=REGEXEXTRACT("AB123","[0-9]+")','123'],
['IFNA','=IFNA(NA(),7)',7],['XMATCH','=XMATCH("East",D1:D5)',1],['ROW','=ROW(A4)',4],['COLUMN','=COLUMN(B3)',2],['FORMULATEXT','=FORMULATEXT(E2)','=SUM(A1:A2)'],['ISFORMULA','=ISFORMULA(E2)',true],['INDIRECT','=INDIRECT("A3")',3],['OFFSET','=OFFSET(A1,2,1)',4],['HYPERLINK','=HYPERLINK("https://example.com","Example")','Example'],['AREAS','=AREAS(A1:A5)',1],['NA','=NA()','#N/A'],['ERROR.TYPE','=ERROR.TYPE(1/0)',2],['ISREF','=ISREF(A1)',true],['CELL','=CELL("address",B3)','$B$3'],
['CUMIPMT','=CUMIPMT(0.01,12,1000,1,1,0)',-10],['CUMPRINC','=CUMPRINC(0.01,12,1000,1,1,0)',-78.8487886783416],['ISPMT','=ISPMT(0.05,2,12,1000)',-41.6666666666667],['DB','=DB(10000,1000,5,1)',3690],['VDB','=VDB(10000,1000,5,0,1)',4000],['PDURATION','=PDURATION(0.1,1000,2000)',Math.log(2)/Math.log(1.1)],['DOLLARDE','=DOLLARDE(1.02,16)',1.125],['DOLLARFR','=DOLLARFR(1.125,16)',1.02],['PRICEDISC','=PRICEDISC(45000,45100,0.05,100)',100*(1-0.05*98/360)],['DISC','=DISC(45000,45100,98,100)',.02*360/98],
['WORKDAY.INTL','=WORKDAY.INTL(DATE(2026,10,19),5,1)',null],['NETWORKDAYS.INTL','=NETWORKDAYS.INTL(DATE(2026,10,19),DATE(2026,10,23),1)',5],
];
// Use the engine DATE implementation as the independent expected serial for relative workday.
cases[cases.findIndex(c=>c[0]==='WORKDAY.INTL')][2]=run('=DATE(2026,10,26)').value;
test('v6 adds exactly 90 unique functions in consistent Excel-style categories',()=>{
 assert.equal(v6FunctionNames.length,90);assert.equal(new Set(v6FunctionNames).size,90);
 assert.equal(functionNames.length,516);assert.equal(cases.length,90);
 assert.deepEqual(cases.map(c=>c[0]).sort(),v6FunctionNames);
 assert.equal(Object.values(categorizedFunctions).flat().length,516);
 for(const name of v6FunctionNames)assert.ok(functionCategory(name));
});
for(const [name,formula,expected] of cases)test(`v6 ${name} : ${formula}`,()=>{
 const value=run(formula);
 if(expected==='#N/A'){assert.equal(value.error?.code,expected);return;}
 assert.equal(value.error,undefined,`${name}: ${value.error?.message}`);
 if(typeof expected==='number')assert.ok(typeof value.value==='number'&&close(value.value,expected),`${formula}: expected ${expected}, got ${value.value}`);
 else assert.equal(value.value,expected);
});
test('v6 dynamic INDIRECT dependency invalidates when target changes',()=>{
 const book=new Workbook();book.applyEdits([{ref:'A1',raw:'3'},{ref:'B1',raw:'=INDIRECT("A1")+2'},{ref:'C1',raw:'=B1*2'}]);
 assert.equal(book.getValue('C1').value,10);const stats=book.setCell('A1','8');
 assert.equal(book.getValue('C1').value,20);assert.ok(stats.affected.includes('B1'));
});
test('v6 IFNA only catches N/A, and reference arguments are not prematurely evaluated',()=>{
 assert.equal(run('=IFNA(NA(),7)').value,7);
 assert.equal(run('=IFNA(1/0,7)').error?.code,'#DIV/0!');
 assert.equal(run('=FORMULATEXT(E2)').value,'=SUM(A1:A2)');
});
test('v6 formulas reject invalid probability and unsupported array spill',()=>{
 assert.equal(run('=NORM.INV(0,0,1)').error?.code,'#NUM!');
 assert.equal(run('=OFFSET(A1,1,1,3,1)').error?.code,'#VALUE!');
});
