import {writeFileSync} from 'node:fs';
import {DATA_TOOLS,TOOL_CATEGORIES} from '../dist/engine/tools_v8.js';
import {Workbook} from '../dist/engine/workbook.js';
import {SpreadsheetDocument} from '../dist/engine/document.js';
const lines=[
'# GridForge v8 — 100 Data Studio Actions','',
'These are **100 individually implemented Excel-style data operations**, not 100 additional Excel formula functions or full feature parity with Microsoft Excel. GridForge retains 306 custom formula functions.','',
'## How to use','','Select a range in the worksheet, click **Data Studio** (or use the right-click menu), choose an operation, click **Preview changes**, review the count and any warnings, then **Apply selection**. Every apply is undoable with Ctrl+Z.','',
'## Safety and compatibility','','- Text and numeric cleanup skip formulas instead of silently replacing their source. Fill and range transformations intentionally move formulas with relative references adjusted.','- Each operation is limited to 12,000 selected cells. Transpose and rotations require square ranges.','- Quick aggregates write formulas below the selection; if those output cells are populated, GridForge prompts before overwriting.','- For imported XLSM workbooks, the original VBA binary is preserved when exporting supported cell-content changes. Cell style updates, structural changes and shared formula edits have limited round-trip support.','- Embedded VBA is never executed during import. A separate Windows desktop Excel bridge can run trusted macros only after explicit approval.','',
'## Complete 100-action catalog',''];
for(const cat of TOOL_CATEGORIES){
 lines.push(`### ${cat}`,'','| # | Action | Behavior |','|---:|---|---|');
 for(const tool of DATA_TOOLS.filter(t=>t.category===cat))lines.push(`| ${DATA_TOOLS.indexOf(tool)+1} | **${tool.label}** (\`${tool.id}\`) | ${tool.description} |`);
 lines.push('');
}
lines.push('## Testing','','`npm run regression` runs TypeScript checks, unit tests, benchmarks, Chromium UI checks, XLSM round-trip tests, and VBA bridge security checks. `npm run regression:core` runs without browser dependencies. See `docs/V8_CODE_AUDIT.md`.');
writeFileSync(new URL('../docs/V8_100_ACTION_REFERENCE.md',import.meta.url),lines.join('\n')+'\n');
const book=new Workbook(400,52);
const examples=[['  aLiCe  ','120','High'],['BOB','84','Low'],['carol smith','158','High'],['DÁVID','90','Medium'],['Alice','120','High']];
const edits=[{ref:'A1',raw:'DATA STUDIO · v8'},{ref:'A2',raw:'Select cells, choose an action, Preview and Apply'},
{ref:'A4',raw:'Name'},{ref:'B4',raw:'Revenue'},{ref:'C4',raw:'Priority'},{ref:'D4',raw:'Calculated'},
{ref:'F4',raw:'EXAMPLE OPERATIONS'},{ref:'F5',raw:'A5:A9: Title Case / Trim / Deduplicate cells'},
{ref:'F6',raw:'B5:B9: Increase 10% / Round / Quick SUM'},
{ref:'F7',raw:'A5:D9: Fill and rearrange / flip'},
{ref:'F8',raw:'Always preview and use Ctrl+Z to undo'}];
for(let i=0;i<examples.length;i++){const [name,amount,priority]=examples[i],r=i+5;edits.push({ref:`A${r}`,raw:name},{ref:`B${r}`,raw:amount},{ref:`C${r}`,raw:priority},{ref:`D${r}`,raw:`=B${r}*1.1`});}
book.applyEdits(edits,false);
book.applyStyles([{ref:'A1',style:{bold:true,color:'#53d3bb'}},...['A4','B4','C4','D4'].map(ref=>({ref,style:{bold:true}}))],false);
const doc=new SpreadsheetDocument(book);
writeFileSync(new URL('../examples/v8_data_studio_demo.json',import.meta.url),JSON.stringify(doc.toJSON(),null,2));
console.log(`Generated ${DATA_TOOLS.length} reference rows and one sample JSON workbook`);
