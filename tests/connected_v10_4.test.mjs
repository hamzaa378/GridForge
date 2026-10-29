import test from 'node:test';
import assert from 'node:assert/strict';
import {functions,functionNames,connectedFunctionNames,connectedFunctionRequirements,SheetError} from '../dist/engine/functions.js';
import {functionCategory,categorizedFunctions} from '../dist/engine/categories.js';
const expected=['CUBEKPIMEMBER','CUBEMEMBER','CUBEMEMBERPROPERTY','CUBERANKEDMEMBER','CUBESET','CUBEVALUE','RTD','STOCKHISTORY','WEBSERVICE'].sort();
test('Connected function list excludes workbook-aware GETPIVOTDATA',()=>{
 assert.equal(functionNames.length,516);
 assert.equal(connectedFunctionNames.length,9);
 assert.deepEqual(connectedFunctionNames,expected);
 assert.equal(Object.keys(connectedFunctionRequirements).length,9);
 assert.deepEqual(categorizedFunctions['External & Connected'],expected);
});
for(const name of expected){
 test(`${name}: recognized, unconnected behavior fails safely with specific explanatory error`,()=>{
  assert.equal(functionCategory(name),'External & Connected');
  const def=functions[name];assert.ok(def?.example.startsWith('='+name+'('));
  const fake=[null,null,null,null,null];
  // Too many? all have at least one allowed arity; adapt function-specific minima
  const required=['CUBEKPIMEMBER','CUBEMEMBERPROPERTY','CUBERANKEDMEMBER','CUBESET','RTD'].includes(name)?3:name==='WEBSERVICE'?1:2;
  const args=fake.slice(0,required);
  assert.throws(()=>def.run(args),err=>err instanceof SheetError && ['#NAME?','#REF!','#CONNECT!','#N/A'].includes(err.code)&&err.message.includes(name));
 });
}
test('GETPIVOTDATA is classified as a local workbook-aware lookup',()=>{
 assert.equal(functionCategory('GETPIVOTDATA'),'Lookup & Reference');
 assert.ok(functions.GETPIVOTDATA?.example.startsWith('=GETPIVOTDATA('));
 assert.ok(!connectedFunctionNames.includes('GETPIVOTDATA'));
});
test('External functions do not open network connections, generate live data, or fabricate PivotTable values',()=>{
 const web=functions.WEBSERVICE,stock=functions.STOCKHISTORY,pivot=functions.GETPIVOTDATA;
 assert.throws(()=>web.run(['https://example.org/']),e=>e.code==='#CONNECT!');
 assert.throws(()=>stock.run(['MSFT',45000]),e=>e.code==='#CONNECT!');
 assert.throws(()=>pivot.run(['Sales',12]),e=>e.code==='#REF!');
});
test('Connected syntax validation and missing arguments produce useful errors',()=>{
 assert.throws(()=>functions.CUBEMEMBER.run(['Sales']),e=>e.code==='#VALUE!');
 assert.throws(()=>functions.WEBSERVICE.run([]),e=>e.code==='#VALUE!');
});
