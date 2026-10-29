#!/usr/bin/env node
/** Cross-platform fail-fast regression runner with clear commands and exit status.
 * Usage: npm run regression (all) | npm run regression:core (no browser)
 */
import {spawnSync} from 'node:child_process';
import {performance} from 'node:perf_hooks';
const coreOnly=process.argv.includes('--core-only');
const started=performance.now();
const commands=[
 ['TypeScript type checking','npm',['run','check']],
 ['Engine and document unit tests','npm',['test']],
 ['10,000-formula dependency benchmark','npm',['run','benchmark']],
 ['10,000-cell Data Studio stress and undo benchmark','npm',['run','benchmark:tools']],
];
if(!coreOnly){
 let python=process.platform==='win32'?'py':'python3';
 let pythonArgs=process.platform==='win32'?['-3','scripts/regression_browser.py']:['scripts/regression_browser.py'];
 if(process.platform==='win32' && spawnSync('py',['-3','--version'],{stdio:'ignore',shell:true}).status!==0){
   python='python';pythonArgs=['scripts/regression_browser.py'];
 }
 commands.push(['Isolated Chromium UI regression',python,pythonArgs]);
 commands.push(['v11.2 Excel-style ribbon, chart, data-bar, resizing, and XLSX browser workflows',python,process.platform==='win32'?(python==='py'?['-3','scripts/smoke_v111_browser.py']:['scripts/smoke_v111_browser.py']):['scripts/smoke_v111_browser.py']]);
 commands.push(['v11.5 Formula ribbon, function insertion, formula audit and display workflows',python,process.platform==='win32'?(python==='py'?['-3','scripts/smoke_formula_ribbon_v115.py']:['scripts/smoke_formula_ribbon_v115.py']):['scripts/smoke_formula_ribbon_v115.py']]);
 commands.push(['v11.6 worksheet names, Name Manager, create-from-selection and XLSX round-trip',python,process.platform==='win32'?(python==='py'?['-3','scripts/smoke_defined_names_v116.py']:['scripts/smoke_defined_names_v116.py']):['scripts/smoke_defined_names_v116.py']]);
 commands.push(['XLSM binary preservation and edit round-trip',python,process.platform==='win32'?(python==='py'?['-3','scripts/test_xlsm_browser.py']:['scripts/test_xlsm_browser.py']):['scripts/test_xlsm_browser.py']]);
 commands.push(['Local Excel VBA bridge protocol security',python,process.platform==='win32'?(python==='py'?['-3','scripts/test_vba_bridge_protocol.py']:['scripts/test_vba_bridge_protocol.py']):['scripts/test_vba_bridge_protocol.py']]);
}
let passed=0;
for(const [label,cmd,args] of commands){
 console.log(`\n${'='.repeat(72)}\n${label}\n$ ${cmd} ${args.join(' ')}\n${'='.repeat(72)}`);
 const result=spawnSync(cmd,args,{stdio:'inherit',shell:process.platform==='win32',env:process.env});
 if(result.error||result.status!==0){
  console.error(`\nFAIL: ${label} (${result.error?.message??`exit ${result.status}`}).`);
  if(label.includes('Chromium')){
   console.error('First-time setup: py -3 -m pip install playwright');
   console.error('                  py -3 -m playwright install chromium');
   console.error('On Linux/macOS use python3 in place of py -3.');
  }
  console.error(`Passed ${passed}/${commands.length} stages. Review the first failure above.`);
  process.exit(1);
 }
 console.log(`PASS: ${label}`);passed++;
}
console.log(`\nREGRESSION PASS: ${passed}/${commands.length} stages in ${((performance.now()-started)/1000).toFixed(1)}s.`);
