import { Workbook, display, type Edit } from './engine/workbook.js';
import { parseCSV,encodeCSV } from './engine/csv.js';
import { colToName, normalizeRef,parseRef,shiftFormula } from './engine/parser.js';
import { functions,functionNames } from './engine/functions.js';
import {SpreadsheetDocument} from './engine/document.js';
import {styledDisplay, type CellStyle, type NumberFormat, type CellAlign} from './engine/format.js';
import {analyzeSelection,findCells,replaceCells,sortRangeTransaction} from './engine/operations.js';
import {snapshotCells,pasteSnapshot,type ClipboardSnapshot} from './engine/clipboard.js';
import {exportXlsx,importXlsx,xlsxFileName} from './engine/xlsx.js';
import {filterVisibleRows,ruleColor} from './engine/sheettools.js';
const ROW_H=30, COL_W=126, HEAD_W=54, HEAD_H=34;
const STORAGE='gridforge:document:v2';
const LEGACY_STORAGE='gridforge:greeting:workbook:v1';
const TITLE_STORAGE='gridforge:workbook-title:v1';
const el = <T extends HTMLElement>(id:string):T => {const node=document.getElementById(id);if(!node)throw Error(`Missing UI element ${id}`);return node as T;};
const escapeHtml=(s:string):string=>s.replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]!));
const asRef=(r:number,c:number):string=>`${colToName(c)}${r+1}`;
const bounds=(a:{r:number;c:number},b:{r:number;c:number})=>({r1:Math.min(a.r,b.r),r2:Math.max(a.r,b.r),c1:Math.min(a.c,b.c),c2:Math.max(a.c,b.c)});
const sample=():Workbook=>{
  const w=new Workbook(400,52);const edits:Edit[]=[
    {ref:'A1',raw:'NORTHSTAR / Q4 SALES'}, {ref:'A2',raw:'A transparent spreadsheet engine · edit units to see recalculation'},
    {ref:'A3',raw:'SKU'}, {ref:'B3',raw:'Units'}, {ref:'C3',raw:'Price ($)'}, {ref:'D3',raw:'Revenue ($)'}, {ref:'E3',raw:'Margin'}, {ref:'F3',raw:'Profit ($)'}, {ref:'G3',raw:'Status'},
    {ref:'J3',raw:'LIVE LOOKUP'}, {ref:'J4',raw:'SKU'}, {ref:'K4',raw:'Price'}, {ref:'L4',raw:'Revenue'},
    {ref:'J7',raw:'QUICK STATS'}, {ref:'J8',raw:'Products'}, {ref:'J9',raw:'Total Revenue'}, {ref:'J10',raw:'Avg. Price'}, {ref:'J11',raw:'Best seller'},
  ];
  const items=[['P-100','125','32','0.24'],['P-101','88','48','0.32'],['P-102','166','19','0.21'],['P-103','51','120','0.41'],['P-104','210','14','0.18'],['P-105','65','75','0.36'],['P-106','94','55','0.29'],['P-107','152','28','0.26']];
  items.forEach(([sku,units,price,margin],i)=>{const r=i+4;edits.push({ref:`A${r}`,raw:sku!},{ref:`B${r}`,raw:units!},{ref:`C${r}`,raw:price!},{ref:`D${r}`,raw:`=B${r}*C${r}`},{ref:`E${r}`,raw:margin!},{ref:`F${r}`,raw:`=ROUND(D${r}*E${r},2)`},{ref:`G${r}`,raw:`=IF(F${r}>1500,"On track","Review")`});});
  edits.push({ref:'A14',raw:'TOTAL'},{ref:'B14',raw:'=SUM(B4:B11)'},{ref:'D14',raw:'=SUM(D4:D11)'},{ref:'F14',raw:'=SUM(F4:F11)'},
    {ref:'J5',raw:'P-103'},{ref:'K5',raw:'=VLOOKUP(J5,A4:C11,3,FALSE)'},{ref:'L5',raw:'=VLOOKUP(J5,A4:D11,4,FALSE)'},
    {ref:'K8',raw:'=COUNTA(A4:A11)'},{ref:'K9',raw:'=D14'},{ref:'K10',raw:'=ROUND(AVERAGE(C4:C11),2)'},{ref:'K11',raw:'=INDEX(A4:A11,MATCH(MAX(B4:B11),B4:B11,0))'},
    {ref:'A17',raw:'TRY THESE'},{ref:'A18',raw:'Change B4, inspect D4, or simulate a what-if scenario →'},{ref:'A19',raw:'Try =SUM(A1:A5), =IF(2>1,"yes","no"), or =VLOOKUP(...)'});
  w.applyEdits(edits,false);return w;
};
let book:Workbook;
let doc:SpreadsheetDocument;
let workbookTitle='Northstar / Q4 Sales';
try {workbookTitle=localStorage.getItem(TITLE_STORAGE)??workbookTitle;}catch { /* storage is optional */ }
try {const previous=localStorage.getItem(STORAGE);const old=localStorage.getItem(LEGACY_STORAGE);doc=previous?SpreadsheetDocument.fromJSON(JSON.parse(previous)):new SpreadsheetDocument(old?Workbook.fromJSON(JSON.parse(old)):sample());}
catch {doc=new SpreadsheetDocument(sample());}
book=doc.active.book;
let active={r:3,c:3},anchor={...active},end={...active};
let editing:HTMLInputElement|null=null, editingRef='';
let mode:'inspector'|'impact'|'scenario'|'functions'|'analyze'='inspector';
let clipboardCache:ClipboardSnapshot|null=null;
let fillDrag=false, rangeDrag=false, dragTarget={...active};
let lastPointerDown={ref:'',at:0};
let storeTimeout=0,scrollFrame=0;
let functionFilter='';
const root=el<HTMLDivElement>('app');
root.innerHTML=`
  <header class="topbar">
    <div class="logo"><div class="logo-mark"><span></span><span></span><span></span><span></span></div><span>grid<span class="logo-light">forge</span></span></div>
    <div class="title-group"><div class="title-line"><span id="book-title">${escapeHtml(workbookTitle)}</span> <span class="tiny-caret">⌄</span></div><div class="title-sub">Personal workspace <span class="dot">•</span> <span id="save-state">Saved locally</span></div></div>
    <div class="topbar-right"><span class="online"><span class="green-dot"></span> Offline ready</span><button id="btn-shortcuts" class="ghost-btn">⌨ <span>Shortcuts</span></button><span class="avatar">G</span></div>
  </header>
  <nav class="toolbar" aria-label="Workbook toolbar">
    <div class="tool-group"><button id="btn-new" class="tool-btn" title="New workbook">＋ <span>New</span></button><button id="btn-sample" class="tool-btn" title="Load demo workbook">▧ <span>Demo</span></button></div>
    <span class="separator"></span><div class="tool-group"><button id="btn-undo" class="tool-btn" title="Undo (Ctrl+Z)">↶ <span>Undo</span></button><button id="btn-redo" class="tool-btn" title="Redo (Ctrl+Y)">↷ <span>Redo</span></button></div>
    <span class="separator"></span><div class="tool-group"><button id="btn-import" class="tool-btn" title="Import XLSX, CSV or JSON">↧ <span>Import</span></button><button id="btn-xlsx" class="tool-btn" title="Download Excel .xlsx workbook">▦ <span>XLSX</span></button><button id="btn-csv" class="tool-btn" title="Export calculated values as CSV">↥ <span>CSV</span></button><button id="btn-json" class="tool-btn" title="Export formulas as JSON">◧ <span>JSON</span></button></div>
    <span class="separator"></span><div class="tool-group"><button id="btn-sort-asc" class="tool-btn" title="Sort selected rows ascending by first selected column">↑ <span>Sort A–Z</span></button><button id="btn-sort-desc" class="tool-btn" title="Sort selected rows descending by first selected column">↓ <span>Sort Z–A</span></button><button id="btn-sort-custom" class="tool-btn" title="Choose sorting column and header handling">▤ <span>Custom sort</span></button><button id="btn-find" class="tool-btn" title="Find and replace (Ctrl+F)">⌕ <span>Find</span></button><button id="btn-filter" class="tool-btn" title="Filter selected table rows">⛃ <span>Filter</span></button><button id="btn-clear-filter" class="tool-btn" title="Show all filtered rows">◯ <span>Show all</span></button></div><span class="separator"></span><span class="toolbar-tip">Tip: Format, sort, or analyze your selection</span><div class="toolbar-push"></div><span id="grid-size" class="cell-counter"></span>
  </nav>
  <div class="format-toolbar" aria-label="Cell formatting toolbar">
    <span class="format-caption">FORMAT</span><button id="fmt-bold" class="format-btn" title="Bold (Ctrl+B)"><b>B</b></button><button id="fmt-italic" class="format-btn" title="Italic (Ctrl+I)"><i>I</i></button><button id="fmt-underline" class="format-btn" title="Underline (Ctrl+U)"><u>U</u></button>
    <label class="fmt-selector-label">Number <select id="fmt-number" aria-label="Number format"><option value="general">General</option><option value="number">Number (0.00)</option><option value="integer">Integer</option><option value="currency">USD ($)</option><option value="percent">Percent (%)</option></select></label>
    <label class="fmt-selector-label">Align <select id="fmt-align" aria-label="Cell text alignment"><option value="left">Left</option><option value="center">Center</option><option value="right">Right</option></select></label>
    <label class="color-field">Text <input id="fmt-color" type="color" value="#d4dbef" title="Font color" aria-label="Font color"/></label>
    <label class="color-field">Fill <input id="fmt-fill" type="color" value="#131b2d" title="Cell background" aria-label="Background color"/></label>
    <button id="fmt-clear" class="format-btn reset-format" title="Remove formatting from selected cells">Clear style</button>
    <button id="btn-autosum" class="format-btn reset-format" title="Insert a SUM formula for the cells above the selection (Alt+=)">∑ AutoSum</button><button id="btn-freeze-row" class="format-btn reset-format" title="Freeze or unfreeze top row">Freeze row</button><button id="btn-freeze-col" class="format-btn reset-format" title="Freeze or unfreeze first column">Freeze col</button><button id="btn-conditional" class="format-btn reset-format" title="Highlight cells by numeric rule">Conditional</button><button id="btn-paste-values" class="format-btn reset-format" title="Paste values from the most recent copied grid range">Paste values</button>
  </div>
  <div class="workspace">
    <section class="sheet-area" aria-label="Spreadsheet">
      <div class="formula-row"><label class="name-box"><span class="sr-only">Cell address</span><input id="name-box" value="D4" spellcheck="false" aria-label="Go to cell"/></label><span class="fx-icon">ƒ<span>x</span></span><input id="formula" placeholder="Select a cell, or type a formula beginning with =" spellcheck="false" aria-label="Formula bar"/><button id="btn-commit" class="apply-formula" title="Apply formula">✓</button></div>
      <div id="grid-scroll" class="grid-scroll" tabindex="0" role="grid" aria-label="Spreadsheet grid" aria-rowcount="400" aria-colcount="52"><div id="grid-surface" class="grid-surface"></div></div>
      <div class="sheet-tabs"><div id="sheet-tab-list" class="sheet-tab-list"></div><button id="btn-add-sheet" class="sheet-action" title="Add worksheet">＋</button><span class="tab-divider"></span><div class="sheet-commands"><button id="btn-rename-sheet" class="sheet-action" title="Rename selected sheet">Rename</button><button id="btn-duplicate-sheet" class="sheet-action" title="Duplicate selected sheet">Duplicate</button><button id="btn-delete-sheet" class="sheet-action" title="Delete selected sheet">Delete</button></div><span class="grow"></span><span id="selection-status">Ready</span></div>
      <div class="metrics"><div class="metric"><span class="metric-icon">◉</span><div><span class="metric-label">Cells recalculated</span><strong id="metric-count">0</strong></div></div><div class="metric"><span class="metric-icon">◷</span><div><span class="metric-label">Recalculation time</span><strong id="metric-time">0 ms</strong></div></div><div class="metric"><span class="metric-icon">⌘</span><div><span class="metric-label">Populated cells</span><strong id="metric-populated">0</strong></div></div><div class="metric"><span class="metric-icon">▤</span><div><span class="metric-label">Rendered cells</span><strong id="metric-rendered">0</strong></div></div></div>
    </section>
    <aside class="sidebar" id="sidebar"><div class="sidebar-top"><div class="sidebar-title"><span class="sidebar-emoji">✧</span><div><strong>Formula intelligence</strong><small>Understand every calculation</small></div></div><button id="sidebar-toggle" class="sidebar-hide" title="Show or hide inspector">☰</button></div>
      <div class="side-tabs" role="tablist"><button data-mode="inspector" class="side-tab chosen">Inspector</button><button data-mode="impact" class="side-tab">Impact</button><button data-mode="scenario" class="side-tab">What-if</button><button data-mode="analyze" class="side-tab">Analyze</button><button data-mode="functions" class="side-tab">Functions</button></div>
      <div class="sidebar-body" id="sidebar-body"></div><div class="sidebar-footer"><span class="mini-spark">✦</span> Built from scratch in TypeScript <span class="side-version">v3.0</span></div>
    </aside>
  </div>
  <footer class="statusbar"><span class="status-left"><span class="green-dot"></span> Auto-saved in this browser</span><span id="status-text">Custom formula engine · No external spreadsheet library</span><span id="selection-aggregate"></span><span class="status-right">Up to 12 sheets <span class="dot">•</span> ${functionNames.length} functions</span></footer>
  <input type="file" id="file-picker" accept=".csv,.json,.xlsx,text/csv,application/json,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" hidden/>
  <div id="toast" class="toast" role="status" aria-live="polite"></div>
  <div id="cell-context-menu" class="cell-context-menu" role="menu" aria-label="Cell actions" hidden>
    <div class="context-heading" id="context-heading">CELL ACTIONS</div>
    <button role="menuitem" data-context="edit"><span class="menu-glyph">✎</span> Edit cell <kbd>F2</kbd></button>
    <div class="context-separator"></div>
    <button role="menuitem" data-context="cut"><span class="menu-glyph">✂</span> Cut <kbd>Ctrl+X</kbd></button>
    <button role="menuitem" data-context="copy"><span class="menu-glyph">▢</span> Copy <kbd>Ctrl+C</kbd></button>
    <button role="menuitem" data-context="paste"><span class="menu-glyph">▣</span> Paste <kbd>Ctrl+V</kbd></button>
    <button role="menuitem" data-context="paste-values"><span class="menu-glyph">123</span> Paste values</button>
    <div class="context-separator"></div>
    <button role="menuitem" data-context="clear"><span class="menu-glyph">⌫</span> Clear contents <kbd>Del</kbd></button>
    <button role="menuitem" data-context="clear-format"><span class="menu-glyph">◌</span> Clear formatting</button>
    <button role="menuitem" data-context="bold"><span class="menu-glyph"><b>B</b></span> Toggle bold <kbd>Ctrl+B</kbd></button>
    <div class="context-separator"></div>
    <button role="menuitem" data-context="sort-asc"><span class="menu-glyph">↑</span> Sort A–Z</button>
    <button role="menuitem" data-context="sort-desc"><span class="menu-glyph">↓</span> Sort Z–A</button>
    <button role="menuitem" data-context="autosum"><span class="menu-glyph">∑</span> AutoSum</button><button role="menuitem" data-context="filter"><span class="menu-glyph">⛃</span> Filter selection…</button><button role="menuitem" data-context="conditional"><span class="menu-glyph">◈</span> Conditional formatting…</button>
    <button role="menuitem" data-context="sort-custom"><span class="menu-glyph">▤</span> Custom sort…</button>
    <button role="menuitem" data-context="find"><span class="menu-glyph">⌕</span> Find and replace <kbd>Ctrl+F</kbd></button>
    <div class="context-separator"></div>
    <button role="menuitem" data-context="inspect"><span class="menu-glyph">ƒx</span> Formula Inspector</button>
    <button role="menuitem" data-context="analyze"><span class="menu-glyph">▥</span> Analyze selection</button>
  </div>
  <dialog id="sort-dialog" class="help-dialog sort-dialog"><div class="dialog-top"><h2>Custom sort</h2><button id="close-sort" class="tool-btn" aria-label="Close custom sort">✕</button></div>
    <p>Sort complete rows inside the selected rectangle. Columns outside the selection are not moved.</p>
    <label for="sort-column" class="field-label">Sort by</label><select id="sort-column" class="context-select" aria-label="Column to sort"></select>
    <label for="sort-direction" class="field-label">Order</label><select id="sort-direction" class="context-select"><option value="ascending">A–Z / Smallest to largest</option><option value="descending">Z–A / Largest to smallest</option></select>
    <label class="sort-checkbox"><input type="checkbox" id="sort-header"/> My selection has a header row</label>
    <p id="sort-preview" class="sort-preview"></p><button id="btn-sort-apply" class="primary-button">Sort selection</button>
  </dialog>
  <dialog id="find-dialog" class="help-dialog find-dialog"><div class="dialog-top"><h2>Find and replace</h2><button id="close-find" class="tool-btn" aria-label="Close">✕</button></div><label class="field-label" for="find-query">Find in the active worksheet</label><input id="find-query" class="scenario-input" placeholder="Search cell contents or formulas"/>
    <label class="field-label" for="replace-text">Replace with</label><input id="replace-text" class="scenario-input" placeholder="Replacement text"/>
    <div class="find-options"><label><input type="checkbox" id="find-case"/> Match case</label><label><input type="checkbox" id="find-formulas"/> Include formulas in Replace All</label></div>
    <div class="find-actions"><button id="btn-find-all" class="secondary-button">Find all</button><button id="btn-replace-all" class="primary-button">Replace all</button></div><div id="find-results" class="find-results" aria-live="polite"></div></dialog>
  <dialog id="import-notes-dialog" class="help-dialog"><div class="dialog-top"><h2>XLSX import report</h2><button class="tool-btn" id="close-import-notes">✕</button></div><p>GridForge imported the supported worksheet content. Review any formula or feature differences before editing or exporting.</p><div id="import-notes-body" class="import-notes-body"></div><button id="confirm-import-notes" class="primary-button">Continue to workbook</button></dialog>
  <dialog id="filter-dialog" class="help-dialog"><div class="dialog-top"><h2>Filter table rows</h2><button class="tool-btn" id="close-filter">✕</button></div><p>Select a header and data rectangle, choose a column, and show rows containing a value. The data is never deleted.</p><label class="field-label" for="filter-column">Filter by column</label><select id="filter-column" class="context-select"></select><label class="field-label" for="filter-query">Contains (case-insensitive)</label><input id="filter-query" class="scenario-input" placeholder="Example: Review"/><button id="apply-filter" class="primary-button">Apply filter</button></dialog>
  <dialog id="conditional-dialog" class="help-dialog"><div class="dialog-top"><h2>Conditional formatting</h2><button class="tool-btn" id="close-conditional">✕</button></div><p>Color selected numeric cells based on their actual numeric values. Rules recalculate automatically after edits.</p><label class="field-label">Condition</label><select class="context-select" id="rule-op"><option value="gt">Greater than</option><option value="lt">Less than</option><option value="eq">Equal to</option></select><label class="field-label">Numeric threshold</label><input id="rule-threshold" type="number" step="any" value="100" class="scenario-input"/><label class="field-label">Highlight color</label><input type="color" id="rule-color" value="#276a53"/><button class="primary-button" id="apply-conditional">Apply to selection</button><button class="secondary-button" id="clear-conditional">Clear rules on this sheet</button></dialog>
  <dialog id="help-dialog" class="help-dialog"><div class="dialog-top"><h2>Keyboard shortcuts</h2><button id="close-help" class="tool-btn">✕</button></div><p>Spreadsheet editing without taking your hands off the keyboard.</p><div class="shortcut-list"><span>Navigate cells</span><kbd>Arrow keys</kbd><span>Edit current cell</span><kbd>Enter / F2 / Type</kbd><span>Commit edit</span><kbd>Enter / Tab</kbd><span>Cancel edit</span><kbd>Esc</kbd><span>Select a range</span><kbd>Shift + Arrow</kbd><span>Cell context menu</span><kbd>Right-click / Shift+F10</kbd><span>AutoSum</span><kbd>Alt + =</kbd><span>Copy range</span><kbd>Ctrl + C</kbd><span>Paste range</span><kbd>Ctrl + V</kbd><span>Fill downward</span><kbd>Ctrl + D</kbd><span>Undo / Redo</span><kbd>Ctrl + Z / Y</kbd><span>Clear selection</span><kbd>Delete</kbd><span>Search cell address</span><kbd>Ctrl + L</kbd></div></dialog>
`;
const grid=el<HTMLDivElement>('grid-scroll');const surface=el<HTMLDivElement>('grid-surface');
const formulaInput=el<HTMLInputElement>('formula'),nameBox=el<HTMLInputElement>('name-box');
const side=el<HTMLDivElement>('sidebar-body');
const currentRef=()=>asRef(active.r,active.c);
const sel=()=>bounds(anchor,end);
const inBounds=(r:number,c:number)=>r>=0&&r<book.rows&&c>=0&&c<book.cols;
function toast(msg:string){const n=el<HTMLDivElement>('toast');n.textContent=msg;n.classList.add('visible');setTimeout(()=>n.classList.remove('visible'),2800);}
function save(){el<HTMLElement>('save-state').textContent='Saving…';clearTimeout(storeTimeout);storeTimeout=window.setTimeout(()=>{try{localStorage.setItem(STORAGE,JSON.stringify(doc.toJSON()));localStorage.setItem(TITLE_STORAGE,workbookTitle);el<HTMLElement>('save-state').textContent='Saved locally';}catch{el<HTMLElement>('save-state').textContent='Storage full — export JSON';toast('Local storage is full. Export your workbook to JSON.');}},250);}
function stats(){const s=book.lastStats;el<HTMLElement>('metric-count').textContent=String(s.recalculated);el<HTMLElement>('metric-time').textContent=`${s.durationMs.toFixed(3)} ms`;el<HTMLElement>('metric-populated').textContent=s.totalPopulated.toLocaleString();el<HTMLElement>('grid-size').textContent=`${(book.rows*book.cols).toLocaleString()}-cell canvas`;el<HTMLElement>('btn-undo').toggleAttribute('disabled',!book.canUndo);el<HTMLElement>('btn-redo').toggleAttribute('disabled',!book.canRedo);}
function visibleRows():number[]{
  const f=doc.active.tools.filter,show=filterVisibleRows(book,f),rows:number[]=[];
  for(let r=0;r<book.rows;r++)if(!f||r<=f.startRow||r>f.endRow||show.has(r))rows.push(r);
  return rows;
}
function ensureVisible(){const displayed=visibleRows();const idx=displayed.indexOf(active.r);const y=HEAD_H+(idx<0?active.r:idx)*ROW_H,x=HEAD_W+active.c*COL_W;if(y<grid.scrollTop+HEAD_H)grid.scrollTop=Math.max(0,y-HEAD_H);else if(y+ROW_H>grid.scrollTop+grid.clientHeight)grid.scrollTop=y+ROW_H-grid.clientHeight;if(x<grid.scrollLeft+HEAD_W)grid.scrollLeft=Math.max(0,x-HEAD_W);else if(x+COL_W>grid.scrollLeft+grid.clientWidth)grid.scrollLeft=x+COL_W-grid.clientWidth;}
function select(r:number,c:number,extend=false,focus=true){if(!inBounds(r,c))return;
  if(editing)commitEdit();active={r,c};end={r,c};if(!extend)anchor={r,c};
  ensureVisible();refresh();if(focus)grid.focus({preventScroll:true});
}
function refresh(){renderGrid();nameBox.value=currentRef();if(document.activeElement!==formulaInput)formulaInput.value=book.getRaw(currentRef());stats();renderTabs();renderFormatControls();renderSidebar();const b=sel();const selection=analyzeSelection(book,b);el<HTMLElement>('selection-aggregate').textContent=selection.numbers?`SUM ${Math.round(selection.sum*100)/100} · AVG ${Math.round((selection.average??0)*100)/100}`:'';el<HTMLElement>('selection-status').textContent=b.r1===b.r2&&b.c1===b.c2?`${currentRef()} selected`:`${asRef(b.r1,b.c1)}:${asRef(b.r2,b.c2)} selected`;}
function renderTabs(){
  el<HTMLDivElement>('sheet-tab-list').innerHTML=doc.sheets.map(x=>`<button type="button" class="sheet-tab ${doc.activeId===x.id?'active-tab':''}" data-sheet-id="${x.id}" title="Switch to ${escapeHtml(x.name)}">▦ <span>${escapeHtml(x.name)}</span></button>`).join('');
}
function renderFormatControls(){
  const style=book.getStyle(currentRef());
  for(const key of ['bold','italic','underline'] as const)el<HTMLButtonElement>(`fmt-${key}`).classList.toggle('format-on',!!style[key]);
  el<HTMLSelectElement>('fmt-number').value=style.numberFormat??'general';
  el<HTMLSelectElement>('fmt-align').value=style.align??'left';
  el<HTMLInputElement>('fmt-color').value=style.color??'#d4dbef';
  el<HTMLInputElement>('fmt-fill').value=style.background??'#131b2d';
}
function selectionRefs(){const b=sel(),refs:string[]=[];for(let r=b.r1;r<=b.r2;r++)for(let c=b.c1;c<=b.c2;c++)refs.push(asRef(r,c));return refs;}
function applyStyle(patch:Partial<CellStyle>){try{book.patchStyles(selectionRefs(),patch);save();refresh();}catch(e){toast(e instanceof Error?e.message:'Invalid format');}}
function renderGrid(){
  const visibleRowsMap=visibleRows(),tools=doc.active.tools;
  surface.style.width=`${HEAD_W+book.cols*COL_W}px`;surface.style.height=`${HEAD_H+visibleRowsMap.length*ROW_H}px`;
  grid.setAttribute('aria-rowcount',String(book.rows));grid.setAttribute('aria-colcount',String(book.cols));
  const startR=Math.max(0,Math.floor(Math.max(0,grid.scrollTop-HEAD_H)/ROW_H)-4);
  const endR=Math.min(visibleRowsMap.length,Math.ceil((grid.scrollTop+grid.clientHeight-HEAD_H)/ROW_H)+4);
  const startC=Math.max(0,Math.floor(Math.max(0,grid.scrollLeft-HEAD_W)/COL_W)-2);
  const endC=Math.min(book.cols,Math.ceil((grid.scrollLeft+grid.clientWidth-HEAD_W)/COL_W)+2);
  const rowViews=new Set<number>();for(let r=startR;r<endR;r++)rowViews.add(r);
  for(let r=0;r<Math.min(tools.freezeRows,visibleRowsMap.length);r++)rowViews.add(r);
  const colViews=new Set<number>();for(let c=startC;c<endC;c++)colViews.add(c);
  for(let c=0;c<tools.freezeCols;c++)colViews.add(c);
  const frag=document.createDocumentFragment();let rendered=0;
  const add=(klass:string,left:number,top:number,width:number,height:number,content:string,dataset?:Record<string,string>):HTMLDivElement=>{
    const div=document.createElement('div');div.className=klass;div.style.cssText=`left:${left}px;top:${top}px;width:${width}px;height:${height}px;`;div.textContent=content;
    if(dataset)for(const [k,v]of Object.entries(dataset))div.dataset[k]=v;
    frag.appendChild(div);return div;
  };
  const selected=sel(),demoTheme=book.getRaw('A1')==='NORTHSTAR / Q4 SALES';
  for(const viewR of [...rowViews].sort((a,b)=>b-a)){
    const r=visibleRowsMap[viewR];if(r===undefined)continue;
    const fixedRow=viewR<tools.freezeRows,top=HEAD_H+viewR*ROW_H+(fixedRow?grid.scrollTop:0);
    const header=add('row-heading',grid.scrollLeft,top,HEAD_W,ROW_H,String(r+1));if(fixedRow)header.style.zIndex='8';
    for(const c of [...colViews].sort((a,b)=>b-a)){
      const ref=asRef(r,c),res=book.getValue(ref),style=book.getStyle(ref),value=styledDisplay(res,style);
      const inSelection=r>=selected.r1&&r<=selected.r2&&c>=selected.c1&&c<=selected.c2;
      const primary=r===active.r&&c===active.c,fixedCol=c<tools.freezeCols;
      const classes=`cell${inSelection?' selected':''}${primary?' primary':''}${res.error?' has-error':''}${book.getRaw(ref).startsWith('=')?' calculated':''}${demoTheme&&r===2?' header-data':''}${demoTheme&&r===13?' total-data':''}`;
      const node=add(classes,HEAD_W+c*COL_W+(fixedCol?grid.scrollLeft:0),top,COL_W,ROW_H,value,{ref});
      node.setAttribute('role','gridcell');node.setAttribute('aria-selected',String(inSelection));node.title=res.error?res.error.message:`${ref} ${book.getRaw(ref)}`;
      if(fixedRow||fixedCol){node.style.zIndex=fixedRow&&fixedCol?'7':fixedRow?'6':'5';node.classList.add('frozen-cell');}
      if(style.bold)node.style.fontWeight='800';if(style.italic)node.style.fontStyle='italic';if(style.underline)node.style.textDecoration='underline';
      if(style.color)node.style.color=style.color;if(style.background)node.style.backgroundColor=style.background;
      const conditional=ruleColor(ref,res.value,tools.conditional);if(conditional)node.style.backgroundColor=conditional;
      if(style.align)node.style.justifyContent=style.align==='right'?'flex-end':style.align==='center'?'center':'flex-start';
      if(primary){const handle=document.createElement('span');handle.className='fill-handle';handle.title='Drag to fill';handle.dataset.handle='true';node.appendChild(handle);}
      rendered++;
    }
  }
  for(const c of [...colViews].sort((a,b)=>b-a)){
    const fixed=c<tools.freezeCols,heading=add('col-heading',HEAD_W+c*COL_W+(fixed?grid.scrollLeft:0),grid.scrollTop,COL_W,HEAD_H,colToName(c));
    if(fixed)heading.style.zIndex='10';
  }
  add('corner',grid.scrollLeft,grid.scrollTop,HEAD_W,HEAD_H,'').style.zIndex='12';
  surface.replaceChildren(frag);el<HTMLElement>('metric-rendered').textContent=String(rendered);
}
function gridPointerCell(clientX:number,clientY:number):{r:number;c:number}|null {
  const bounds=grid.getBoundingClientRect(),tools=doc.active.tools;
  const x=clientX-bounds.left,y=clientY-bounds.top;
  const viewC=Math.floor((x-HEAD_W+(x<HEAD_W+tools.freezeCols*COL_W?0:grid.scrollLeft))/COL_W);
  const viewR=Math.floor((y-HEAD_H+(y<HEAD_H+tools.freezeRows*ROW_H?0:grid.scrollTop))/ROW_H);
  const r=visibleRows()[viewR];return r===undefined||!inBounds(r,viewC)?null:{r,c:viewC};
}

function apply(edits:Edit[]){try{book.applyEdits(edits);save();refresh();}catch(e){toast(e instanceof Error?e.message:'Unable to modify cells.');}}
function commitEdit(move:{r:number;c:number}|null=null){if(!editing)return;const text=editing.value,ref=editingRef;editing=null;editingRef='';apply([{ref,raw:text}]);if(move)select(Math.min(book.rows-1,Math.max(0,move.r)),Math.min(book.cols-1,Math.max(0,move.c)));}
function beginEdit(replace?:string){if(editing)return;const ref=currentRef();const node=surface.querySelector<HTMLElement>(`.cell[data-ref="${ref}"]`);if(!node)return;
  const inp=document.createElement('input');inp.className='cell-editor';inp.value=replace??book.getRaw(ref);inp.spellcheck=false;inp.setAttribute('aria-label',`Edit ${ref}`);node.replaceChildren(inp);node.classList.add('editing');editing=inp;editingRef=ref;inp.focus();if(replace===undefined)inp.select();
  inp.addEventListener('keydown',ev=>{ev.stopPropagation();if(ev.key==='Enter'){ev.preventDefault();commitEdit({r:active.r+1,c:active.c});}else if(ev.key==='Tab'){ev.preventDefault();commitEdit({r:active.r,c:active.c+(ev.shiftKey?-1:1)});}else if(ev.key==='Escape'){ev.preventDefault();editing=null;editingRef='';refresh();grid.focus({preventScroll:true});}});
  inp.addEventListener('blur',()=>{if(editing===inp)commitEdit();});
}
function actionFor(rect:{r1:number;r2:number;c1:number;c2:number}, raw=''):Edit[]{const edits:Edit[]=[];for(let r=rect.r1;r<=rect.r2;r++)for(let c=rect.c1;c<=rect.c2;c++)edits.push({ref:asRef(r,c),raw});return edits;}
function fillTo(target:{r:number;c:number}){
  const r0=active.r,c0=active.c;const vertical=Math.abs(target.r-r0)>=Math.abs(target.c-c0);
  const cells:Edit[]=[];const source=book.getRaw(currentRef());
  if(vertical)for(let r=Math.min(r0,target.r);r<=Math.max(r0,target.r);r++){if(r===r0)continue;cells.push({ref:asRef(r,c0),raw:shiftFormula(source,r-r0,0)});}
  else for(let c=Math.min(c0,target.c);c<=Math.max(c0,target.c);c++){if(c===c0)continue;cells.push({ref:asRef(r0,c),raw:shiftFormula(source,0,c-c0)});}
  if(cells.length){apply(cells);toast(`Filled ${cells.length} cells with relative references`);}
}
function setMode(next:typeof mode){mode=next;document.querySelectorAll<HTMLButtonElement>('[data-mode]').forEach(x=>x.classList.toggle('chosen',x.dataset.mode===next));renderSidebar();}
const tagLink=(ref:string,kind='dep')=>`<button class="ref-pill ${kind}" data-jump="${escapeHtml(ref)}">${escapeHtml(ref)} <span>↗</span></button>`;
function renderSidebar(){const ref=currentRef(), info=book.inspect(ref), raw=info.raw, result=display(info.result), err=info.result.error;
  if(mode==='inspector'){
    side.innerHTML=`<div class="panel-spotlight"><div class="overline">SELECTED CELL <span class="panel-ref">${ref}</span></div><div class="panel-value ${err?'error-txt':''}">${escapeHtml(result||'Empty')}</div><div class="formula-small">${escapeHtml(raw||'No formula entered')}</div></div>
      <div class="panel-block"><h3>How it works <span class="hint">AST</span></h3><p>Every formula is parsed into an expression tree, then evaluated using live references.</p><pre class="ast-tree">${escapeHtml(info.expressionTree)}</pre></div>
      <div class="panel-block"><h3>Direct references <span class="hint">${info.dependencies.length}</span></h3>${info.dependencies.length?`<div class="ref-wrap">${info.dependencies.slice(0,22).map(r=>tagLink(r)).join('')}</div>`:'<p class="empty-sub">No cell references in this formula.</p>'}</div>
      ${err?`<div class="diagnostic"><div class="diagnostic-title">⚠ ${escapeHtml(err.code)} · ${escapeHtml(err.message)}</div><pre>${escapeHtml(raw+'\n'+' '.repeat(Math.max(0,Math.min(err.position+1,raw.length)))+'^')}</pre><p><strong>Suggestion</strong> ${escapeHtml(err.suggestion)}</p></div>`:''}
      ${info.referencedValues.length?`<div class="panel-block"><h3>Input values</h3>${info.referencedValues.slice(0,8).map(x=>`<div class="key-value"><button data-jump="${x.ref}">${x.ref}</button><strong>${escapeHtml(x.display||'—')}</strong></div>`).join('')}</div>`:''}`;
  }else if(mode==='impact'){
    const reach=book.transitiveDependents(ref);const visible=reach.slice(1,20);
    side.innerHTML=`<div class="panel-spotlight impact-spot"><div class="overline">DEPENDENCY EXPLORER</div><div class="panel-value">${ref}</div><p>Discover what feeds this cell and what this cell influences.</p></div>
      <div class="panel-block"><h3>Upstream inputs <span class="hint">${info.dependencies.length}</span></h3><div class="dependency-diagram"><div class="ref-wrap">${info.dependencies.length?info.dependencies.slice(0,30).map(r=>tagLink(r)).join(''):'<span class="empty-sub">No direct inputs</span>'}</div><div class="graph-line">↓</div><div class="main-node">${ref} <small>${escapeHtml(result||'empty')}</small></div><div class="graph-line">↓</div><div class="ref-wrap">${visible.length?visible.map(r=>tagLink(r,'out')).join(''):'<span class="empty-sub">No downstream dependencies</span>'}</div></div></div>
      <div class="panel-block"><h3>Change impact <span class="hint">${reach.length-1} affected</span></h3><p>Editing ${ref} invalidates exactly the reachable dependent cells rather than the whole sheet.</p><div class="insight-line">${(reach.length-1).toLocaleString()} downstream cells reachable</div></div>`;
  }else if(mode==='scenario'){
    side.innerHTML=`<div class="panel-spotlight scenario-spot"><div class="overline">WHAT-IF SANDBOX</div><div class="panel-value">Explore a change</div><p>Preview an edit on an isolated workbook without changing your real data.</p></div><div class="panel-block"><h3>1. Choose a hypothetical value</h3><label class="field-label" for="scenario-input">New content for ${ref}</label><input id="scenario-input" class="scenario-input" placeholder="Example: 500 or =B4*3" value="${escapeHtml(raw)}"/><button id="btn-preview" class="primary-button">Preview impact →</button><div id="preview-results" class="preview-results"></div></div><div class="panel-block"><h3>Why this matters</h3><p>Risk-free planning: see the consequences of changing an input without modifying the saved workbook. Apply only after reviewing the differences.</p></div>`;
  }else if(mode==='analyze'){
    const data=analyzeSelection(book,sel());const max=Math.max(1,...data.chart.map(x=>Math.abs(x.value)));
    side.innerHTML=`<div class="panel-spotlight"><div class="overline">DATA QUALITY LENS <span class="panel-ref">${escapeHtml(asRef(sel().r1,sel().c1))}:${escapeHtml(asRef(sel().r2,sel().c2))}</span></div><div class="panel-value">${data.count.toLocaleString()} cells</div><p>Live statistics, missing data, duplicate entries and a lightweight chart of the selected range.</p></div>
      <div class="panel-block"><h3>Selection statistics</h3><div class="analysis-stats">${[['Numeric values',data.numbers],['Sum',data.sum],['Average',data.average??'—'],['Minimum',data.min??'—'],['Maximum',data.max??'—'],['Blank cells',data.blanks],['Formula errors',data.errors],['Duplicate values',data.duplicates]].map(([k,v])=>`<div><span>${k}</span><b>${typeof v==='number'?Math.round(v*1e4)/1e4:escapeHtml(String(v))}</b></div>`).join('')}</div></div>
      <div class="panel-block"><h3>Quick chart <span class="hint">First two columns</span></h3>${data.chart.length?`<div class="quick-chart">${data.chart.slice(0,16).map(x=>`<div class="chart-line" title="${escapeHtml(x.label)}: ${x.value}"><span>${escapeHtml(x.label.slice(0,11))}</span><div class="chart-track"><div class="chart-bar ${x.value<0?'negative':''}" style="width:${Math.max(2,100*Math.abs(x.value)/max)}%"></div></div><b>${Math.round(x.value*100)/100}</b></div>`).join('')}</div>`:'<p>Select a rectangle with numeric cells to visualize it. The first column supplies labels and the second supplies values.</p>'}</div><div class="panel-block"><h3>How to use</h3><p>Select a data table, use Sort A–Z or Z–A, then inspect differences here. Analytics never changes your spreadsheet.</p></div>`;
  }else {
    const filtered=functionNames.filter(name=>name.includes(functionFilter.toUpperCase())||functions[name]!.description.toLowerCase().includes(functionFilter.toLowerCase()));
    side.innerHTML=`<div class="panel-spotlight"><div class="overline">FUNCTION LIBRARY</div><div class="panel-value">${functionNames.length} built-ins</div><p>Every function is implemented in this repository. Select one to insert a template.</p></div><div class="panel-block"><input class="function-search" id="function-search" aria-label="Filter functions" placeholder="Search formulas…" value="${escapeHtml(functionFilter)}"/><div class="function-list">${filtered.map(f=>`<button class="function-card" data-function="${f}"><div><strong>${f}</strong><span>＋</span></div><small>${escapeHtml(functions[f]!.description)}</small><code>${escapeHtml(functions[f]!.example)}</code></button>`).join('')}</div></div>`;
  }
}
function jump(ref:string){const p=parseRef(ref);if(p&&inBounds(p.row,p.col))select(p.row,p.col);}
side.addEventListener('click',ev=>{
  const target=ev.target as HTMLElement;const jumpEl=target.closest<HTMLElement>('[data-jump]');if(jumpEl?.dataset.jump){jump(jumpEl.dataset.jump);return;}
  const funEl=target.closest<HTMLElement>('[data-function]');if(funEl?.dataset.function){const f=funEl.dataset.function;formulaInput.value=`=${f}()`;formulaInput.focus();formulaInput.setSelectionRange(formulaInput.value.length-1,formulaInput.value.length-1);return;}
  if(target.closest('#btn-preview')){
    const proposed=el<HTMLInputElement>('scenario-input').value;try{
      const results=book.preview(currentRef(),proposed);const area=el<HTMLDivElement>('preview-results');
      area.innerHTML=`<h4>${results.changes.length} cell${results.changes.length===1?'':'s'} would change</h4><p>Hypothetical recalculation: ${results.stats.recalculated} populated cells in ${results.stats.durationMs.toFixed(3)} ms.</p><div class="scenario-diffs">${results.changes.slice(0,25).map(x=>`<div class="diff-line"><b>${x.ref}</b><span>${escapeHtml(x.before||'∅')}</span><span class="diff-arrow">→</span><strong>${escapeHtml(x.after||'∅')}</strong></div>`).join('')}</div><button id="btn-apply-preview" class="secondary-button">Apply this change</button>`;
    }catch(e){toast(e instanceof Error?e.message:'Unable to simulate');}
  }
  if(target.closest('#btn-apply-preview')){const proposed=el<HTMLInputElement>('scenario-input').value;apply([{ref:currentRef(),raw:proposed}]);toast('Scenario applied to workbook');}
});
side.addEventListener('input',e=>{if((e.target as HTMLElement).id==='function-search'){const input=e.target as HTMLInputElement;const caret=input.selectionStart;functionFilter=input.value;renderSidebar();const restored=el<HTMLInputElement>('function-search');restored.focus();restored.setSelectionRange(caret,caret);}});
document.querySelectorAll<HTMLButtonElement>('[data-mode]').forEach(node=>node.addEventListener('click',()=>setMode(node.dataset.mode as typeof mode)));
grid.addEventListener('scroll',()=>{if(editing)commitEdit();if(!scrollFrame)scrollFrame=requestAnimationFrame(()=>{scrollFrame=0;renderGrid();});});
surface.addEventListener('mousedown',ev=>{
  const target=ev.target as HTMLElement;const cell=target.closest<HTMLElement>('[data-ref]');if(!cell?.dataset.ref)return;
  if(target.dataset.handle){ev.preventDefault();fillDrag=true;dragTarget={...active};return;}
  if(ev.button!==0)return;
  if(editing)return;const p=parseRef(cell.dataset.ref)!;ev.preventDefault();
  const second=ev.detail>=2||(lastPointerDown.ref===cell.dataset.ref&&Date.now()-lastPointerDown.at<420);
  lastPointerDown={ref:cell.dataset.ref,at:Date.now()};
  if(second){rangeDrag=false;active={r:p.row,c:p.col};anchor={...active};end={...active};refresh();beginEdit();return;}
  rangeDrag=true;
  if(ev.shiftKey){end={r:p.row,c:p.col};active={r:p.row,c:p.col};}else{active={r:p.row,c:p.col};anchor={...active};end={...active};}
  refresh();grid.focus({preventScroll:true});
});
grid.addEventListener('dblclick',ev=>{const where=gridPointerCell(ev.clientX,ev.clientY);if(where){select(where.r,where.c);beginEdit();}});
surface.addEventListener('mouseover',ev=>{
  const cell=(ev.target as HTMLElement).closest<HTMLElement>('[data-ref]');if(!cell?.dataset.ref)return;const p=parseRef(cell.dataset.ref)!;
  if(fillDrag)dragTarget={r:p.row,c:p.col};
  if(rangeDrag){active={r:p.row,c:p.col};end={...active};renderGrid();}
});
window.addEventListener('mouseup',()=>{if(fillDrag){fillDrag=false;fillTo(dragTarget);}if(rangeDrag){rangeDrag=false;refresh();}});
grid.addEventListener('keydown',ev=>{
  if(editing)return;
  if(ev.key==='ContextMenu'||(ev.shiftKey&&ev.key==='F10')){
    ev.preventDefault();const selectedNode=surface.querySelector<HTMLElement>(`.cell[data-ref="${currentRef()}"]`);
    const at=selectedNode?.getBoundingClientRect()??grid.getBoundingClientRect();openContextMenu(at.left+30,at.top+20);return;
  }
  const ctrl=ev.ctrlKey||ev.metaKey;
  if(ctrl&&ev.key.toLowerCase()==='x'){ev.preventDefault();copySelection(true);return;}
  if(ctrl&&ev.key.toLowerCase()==='z'){ev.preventDefault();if(ev.shiftKey)book.redo();else book.undo();save();refresh();return;}
  if(ctrl&&ev.key.toLowerCase()==='y'){ev.preventDefault();book.redo();save();refresh();return;}
  if(ctrl&&ev.key.toLowerCase()==='l'){ev.preventDefault();nameBox.focus();nameBox.select();return;}
  if(ctrl&&ev.key.toLowerCase()==='f'){ev.preventDefault();openFind();return;}
  if(ctrl&&['b','i','u'].includes(ev.key.toLowerCase())){ev.preventDefault();const k=({b:'bold',i:'italic',u:'underline'} as const)[ev.key.toLowerCase() as 'b'|'i'|'u'];applyStyle({[k]:!book.getStyle(currentRef())[k]});return;}

  if(ctrl&&ev.key.toLowerCase()==='d'){ev.preventDefault();const b=sel();if(b.r2>b.r1){const edits:Edit[]=[];for(let c=b.c1;c<=b.c2;c++)for(let r=b.r1+1;r<=b.r2;r++)edits.push({ref:asRef(r,c),raw:shiftFormula(book.getRaw(asRef(b.r1,c)),r-b.r1,0)});apply(edits);}return;}
  if(ctrl&&ev.key.toLowerCase()==='a'){ev.preventDefault();anchor={r:0,c:0};end={r:book.rows-1,c:book.cols-1};refresh();return;}
  if(ctrl&&['c','v'].includes(ev.key.toLowerCase()))return;
  let dr=0,dc=0;switch(ev.key){case 'ArrowUp':dr=-1;break;case 'ArrowDown':dr=1;break;case 'ArrowLeft':dc=-1;break;case 'ArrowRight':dc=1;break;case 'Tab':dc=ev.shiftKey?-1:1;break;case 'Enter':beginEdit();ev.preventDefault();return;case 'F2':beginEdit();ev.preventDefault();return;case 'Delete':case 'Backspace':ev.preventDefault();apply(actionFor(sel()));return;default:if(ev.key.length===1&&!ctrl&&!ev.altKey){ev.preventDefault();beginEdit(ev.key);}return;}
  ev.preventDefault();select(Math.min(book.rows-1,Math.max(0,active.r+dr)),Math.min(book.cols-1,Math.max(0,active.c+dc)),ev.shiftKey&&ev.key!=='Tab');
});
function copySelection(cut=false){
  clipboardCache=snapshotCells(book,sel(),doc.activeId,cut);
  // Menu and Ctrl+X copying is always usable within the app even when browser clipboard permission is denied.
  void navigator.clipboard?.writeText(clipboardCache.plain).catch(()=>{});
  toast(`${cut?'Cut':'Copied'} ${clipboardCache.rows.length} row(s) · ${cut?'paste to move':'paste to duplicate'}`);
}
function pasteCached(valuesOnly=false){
  if(!clipboardCache){toast('Copy a range first (Ctrl+C or right-click → Copy).');return;}
  try{
    const plan=pasteSnapshot(book,clipboardCache,{row:active.r,col:active.c},doc.activeId,valuesOnly);
    book.applyTransaction(plan.edits,plan.styles);
    if(plan.moved)clipboardCache=null;
    save();refresh();toast(`${plan.moved?'Moved':'Pasted'} ${plan.edits.length} cells${valuesOnly?' as values':''}`);
  }catch(e){toast(e instanceof Error?e.message:'Paste failed');}
}
document.addEventListener('copy',ev=>{
  if(document.activeElement!==grid)return;
  clipboardCache=snapshotCells(book,sel(),doc.activeId,false);
  ev.clipboardData?.setData('text/plain',clipboardCache.plain);
  // Preserve the existing same-app MIME metadata for browsers that allow it.
  ev.clipboardData?.setData('application/x-gridforge-cells',JSON.stringify(clipboardCache));
  ev.preventDefault();toast('Copied cells, formulas and formatting');
});
document.addEventListener('paste',ev=>{
  if(document.activeElement!==grid)return;
  const data=ev.clipboardData,txt=data?.getData('text/plain');if(txt===undefined||!data)return;
  ev.preventDefault();
  if(clipboardCache&&txt===clipboardCache.plain){pasteCached();return;}
  const lines=txt.replace(/\r\n/g,'\n').split('\n');if(lines[lines.length-1]==='')lines.pop();const edits:Edit[]=[];
  for(let r=0;r<lines.length;r++){const cols=lines[r]!.split('\t');for(let c=0;c<cols.length;c++)if(inBounds(active.r+r,active.c+c))edits.push({ref:asRef(active.r+r,active.c+c),raw:cols[c]!});}
  apply(edits);toast(`Pasted ${edits.length} cells from text`);
});
formulaInput.addEventListener('keydown',ev=>{if(ev.key==='Enter'){ev.preventDefault();apply([{ref:currentRef(),raw:formulaInput.value}]);grid.focus({preventScroll:true});}else if(ev.key==='Escape'){formulaInput.value=book.getRaw(currentRef());grid.focus({preventScroll:true});}});
el<HTMLButtonElement>('btn-commit').addEventListener('click',()=>{apply([{ref:currentRef(),raw:formulaInput.value}]);grid.focus({preventScroll:true});});
nameBox.addEventListener('keydown',ev=>{if(ev.key==='Enter'){const p=parseRef(nameBox.value.trim());if(p&&inBounds(p.row,p.col))select(p.row,p.col);else toast('Enter a valid address, e.g. AA120');}if(ev.key==='Escape')nameBox.value=currentRef();});
function download(name:string,contents:string|Uint8Array,type:string){const url=URL.createObjectURL(new Blob([typeof contents==='string'?contents:new Uint8Array(contents)],{type}));const a=document.createElement('a');a.href=url;a.download=name;a.click();setTimeout(()=>URL.revokeObjectURL(url),1500);}
el<HTMLButtonElement>('btn-xlsx').addEventListener('click',()=>{try{download(xlsxFileName(workbookTitle),exportXlsx(doc),'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');toast('Exported multi-sheet XLSX.');}catch(e){toast(e instanceof Error?e.message:'XLSX export failed');}});
el<HTMLButtonElement>('btn-json').addEventListener('click',()=>download('gridforge-workbook-v3.json',JSON.stringify(doc.toJSON(),null,2),'application/json'));
el<HTMLButtonElement>('btn-csv').addEventListener('click',()=>{
  const entries=book.allEntries();let maxR=0,maxC=0;for(const x of entries){const p=parseRef(x.ref)!;maxR=Math.max(maxR,p.row);maxC=Math.max(maxC,p.col);}
  const rows:string[][]=[];for(let r=0;r<=maxR;r++){const line:string[]=[];for(let c=0;c<=maxC;c++){const result=book.getValue(asRef(r,c)),shown=display(result);
      line.push(typeof result.value==='string' && /^[\t\r\n ]*[=+\-@]/.test(shown)?`'${shown}`:shown);
    }rows.push(line);}
  download('gridforge-values.csv',encodeCSV(rows),'text/csv');toast('Exported calculated values to CSV');
});
el<HTMLButtonElement>('btn-undo').addEventListener('click',()=>{if(book.undo()){save();refresh();}});
el<HTMLButtonElement>('btn-redo').addEventListener('click',()=>{if(book.redo()){save();refresh();}});
el<HTMLButtonElement>('btn-new').addEventListener('click',()=>{if(!confirm('Create a blank workbook? Export your current workbook first if you need a copy.'))return;doc=new SpreadsheetDocument();book=doc.active.book;workbookTitle='Untitled Workbook';el<HTMLElement>('book-title').textContent=workbookTitle;save();select(0,0);toast('Blank workbook created');});
el<HTMLButtonElement>('btn-sample').addEventListener('click',()=>{if(!confirm('Replace current workbook with the demo data?'))return;doc=new SpreadsheetDocument(sample());book=doc.active.book;workbookTitle='Northstar / Q4 Sales';el<HTMLElement>('book-title').textContent=workbookTitle;save();select(3,3);toast('Demo workbook loaded');});
const picker=el<HTMLInputElement>('file-picker');el<HTMLButtonElement>('btn-import').addEventListener('click',()=>picker.click());
picker.addEventListener('change',async()=>{const file=picker.files?.[0];if(!file)return;try{
  const content=file.name.toLowerCase().endsWith('.xlsx')?'':await file.text();if(file.name.toLowerCase().endsWith('.xlsx')){
    const response=await importXlsx(new Uint8Array(await file.arrayBuffer()));doc=response.document;book=doc.active.book;
    if(response.warnings.length){console.warn('XLSX import notes:',response.warnings);
      el<HTMLElement>('import-notes-body').innerHTML=`<strong>${doc.sheets.length} worksheet(s) loaded</strong><ul>${response.warnings.slice(0,20).map(w=>`<li>${escapeHtml(w)}</li>`).join('')}</ul>`;
      el<HTMLDialogElement>('import-notes-dialog').showModal();
    }
  }else if(file.name.toLowerCase().endsWith('.json')){doc=SpreadsheetDocument.fromJSON(JSON.parse(content));book=doc.active.book;}
  else {const rows=parseCSV(content);const fresh=new Workbook(Math.max(400,Math.min(1000,rows.length)),Math.max(52,Math.min(104,rows.reduce((longest,row)=>Math.max(longest,row.length),0))));const edits:Edit[]=[];
    rows.forEach((line,r)=>line.forEach((text,c)=>{if(r<fresh.rows&&c<fresh.cols&&text!=='')edits.push({ref:asRef(r,c),raw:text});}));fresh.applyEdits(edits,false);doc=new SpreadsheetDocument(fresh);book=doc.active.book;}
  workbookTitle=file.name.replace(/\.(csv|json|xlsx)$/i,'').slice(0,45)||'Imported Workbook';el<HTMLElement>('book-title').textContent=workbookTitle;save();select(0,0);toast(`Imported ${file.name}; verify formulas and styles`);
}catch(e){toast(e instanceof Error?e.message:'Import failed');}finally{picker.value='';}});
for(const id of ['close-import-notes','confirm-import-notes'])el<HTMLButtonElement>(id).addEventListener('click',()=>el<HTMLDialogElement>('import-notes-dialog').close());
const dialog=el<HTMLDialogElement>('help-dialog');el<HTMLButtonElement>('btn-shortcuts').addEventListener('click',()=>dialog.showModal());el<HTMLButtonElement>('close-help').addEventListener('click',()=>dialog.close());
el<HTMLButtonElement>('sidebar-toggle').addEventListener('click',()=>{el<HTMLElement>('sidebar').classList.toggle('collapsed');});
window.addEventListener('resize',()=>renderGrid());
// Excel-style selection formatting, with versioned JSON persistence and undo.
for(const key of ['bold','italic','underline'] as const)el<HTMLButtonElement>(`fmt-${key}`).addEventListener('click',()=>applyStyle({[key]:!book.getStyle(currentRef())[key]}));
el<HTMLSelectElement>('fmt-number').addEventListener('change',ev=>applyStyle({numberFormat:(ev.target as HTMLSelectElement).value as NumberFormat}));
el<HTMLSelectElement>('fmt-align').addEventListener('change',ev=>applyStyle({align:(ev.target as HTMLSelectElement).value as CellAlign}));
el<HTMLInputElement>('fmt-color').addEventListener('change',ev=>applyStyle({color:(ev.target as HTMLInputElement).value}));
el<HTMLInputElement>('fmt-fill').addEventListener('change',ev=>applyStyle({background:(ev.target as HTMLInputElement).value}));
el<HTMLButtonElement>('fmt-clear').addEventListener('click',()=>{book.clearStyles(selectionRefs());save();refresh();});
el<HTMLButtonElement>('btn-paste-values').addEventListener('click',()=>pasteCached(true));
// Styling follows the sorted records. Sorting plus style moves is one undoable transaction.
function sortCurrent(byColumn:number,descending:boolean,hasHeader=false){
  const selected=sel();if(selected.r1===selected.r2||(hasHeader&&selected.r2-selected.r1<2)){
    toast('Select at least two data rows to sort.');return false;
  }
  if((selected.r2-selected.r1+1)*(selected.c2-selected.c1+1)>15000){toast('Sort up to 15,000 cells at a time.');return false;}
  try{
    const tx=sortRangeTransaction(book,selected,byColumn,descending,hasHeader);
    book.applyTransaction(tx.edits,tx.styles);save();refresh();
    toast(`Sorted ${selected.r2-selected.r1+1-(hasHeader?1:0)} rows by ${colToName(byColumn)} (${descending?'Z–A':'A–Z'})`);
    return true;
  }catch(e){toast(e instanceof Error?e.message:'Sort failed');return false;}
}
for(const [id,descending]of [['btn-sort-asc',false],['btn-sort-desc',true]] as const)
  el<HTMLButtonElement>(id).addEventListener('click',()=>sortCurrent(sel().c1,descending));
const sortDialog=el<HTMLDialogElement>('sort-dialog');
function openCustomSort(){
  const selected=sel();if(selected.r1===selected.r2){toast('Select at least two rows first.');return;}
  const choices=[];
  for(let c=selected.c1;c<=selected.c2;c++){
    const name=colToName(c),label=book.getRaw(asRef(selected.r1,c)).slice(0,40);
    choices.push(`<option value="${c}">${name}${label?' — '+escapeHtml(label):''}</option>`);
  }
  el<HTMLSelectElement>('sort-column').innerHTML=choices.join('');
  el<HTMLSelectElement>('sort-column').value=String(selected.c1);
  el<HTMLInputElement>('sort-header').checked=false;
  el<HTMLElement>('sort-preview').textContent=`Selected ${asRef(selected.r1,selected.c1)}:${asRef(selected.r2,selected.c2)} · ${selected.r2-selected.r1+1} rows`;
  sortDialog.showModal();
}
el<HTMLButtonElement>('btn-sort-custom').addEventListener('click',openCustomSort);
el<HTMLButtonElement>('close-sort').addEventListener('click',()=>sortDialog.close());
el<HTMLButtonElement>('btn-sort-apply').addEventListener('click',()=>{
  if(sortCurrent(Number(el<HTMLSelectElement>('sort-column').value),el<HTMLSelectElement>('sort-direction').value==='descending',el<HTMLInputElement>('sort-header').checked))sortDialog.close();
});
const contextMenu=el<HTMLDivElement>('cell-context-menu');
function closeContextMenu(){contextMenu.hidden=true;}
function openContextMenu(clientX:number,clientY:number){
  if(editing)commitEdit();
  el<HTMLElement>('context-heading').textContent=`${doc.active.name} · ${asRef(sel().r1,sel().c1)}${sel().r1!==sel().r2||sel().c1!==sel().c2?':'+asRef(sel().r2,sel().c2):''}`;
  contextMenu.hidden=false;
  // Make the complete menu visible near all four viewport boundaries.
  contextMenu.style.left=`${Math.max(6,Math.min(clientX,window.innerWidth-contextMenu.offsetWidth-8))}px`;
  contextMenu.style.top=`${Math.max(6,Math.min(clientY,window.innerHeight-contextMenu.offsetHeight-8))}px`;
  contextMenu.querySelector<HTMLButtonElement>('[data-context="edit"]')?.focus({preventScroll:true});
}
grid.addEventListener('contextmenu',event=>{
  const where=gridPointerCell(event.clientX,event.clientY);if(!where)return;
  const {r,c}=where;
  event.preventDefault();
  const b=sel();
  if(r<b.r1||r>b.r2||c<b.c1||c>b.c2){active={r,c};anchor={r,c};end={r,c};refresh();}
  openContextMenu(event.clientX,event.clientY);
});
contextMenu.addEventListener('click',event=>{
  const btn=(event.target as HTMLElement).closest<HTMLButtonElement>('[data-context]');
  if(!btn)return;const action=btn.dataset.context;closeContextMenu();
  switch(action){
    case 'edit':beginEdit();break;
    case 'copy':copySelection();break;
    case 'cut':copySelection(true);break;
    case 'paste':pasteCached();break;
    case 'paste-values':pasteCached(true);break;
    case 'clear':apply(actionFor(sel()));break;
    case 'clear-format':book.clearStyles(selectionRefs());save();refresh();toast('Cleared selection formatting');break;
    case 'bold':applyStyle({bold:!book.getStyle(currentRef()).bold});break;
    case 'sort-asc':sortCurrent(sel().c1,false);break;
    case 'sort-desc':sortCurrent(sel().c1,true);break;
    case 'sort-custom':openCustomSort();break;
    case 'autosum':autoSum();break;
    case 'filter':el<HTMLButtonElement>('btn-filter').click();break;
    case 'conditional':el<HTMLButtonElement>('btn-conditional').click();break;
    case 'find':openFind();break;
    case 'inspect':setMode('inspector');break;
    case 'analyze':setMode('analyze');break;
  }
});
contextMenu.addEventListener('keydown',event=>{
  const buttons=Array.from(contextMenu.querySelectorAll<HTMLButtonElement>('[data-context]'));
  const index=buttons.indexOf(document.activeElement as HTMLButtonElement);
  if(event.key==='ArrowDown'||event.key==='ArrowUp'){
    event.preventDefault();buttons[(index+(event.key==='ArrowDown'?1:buttons.length-1))%buttons.length]?.focus();
  }
  if(event.key==='Escape'){event.preventDefault();closeContextMenu();grid.focus({preventScroll:true});}
});
document.addEventListener('pointerdown',event=>{if(!contextMenu.contains(event.target as Node))closeContextMenu();});
window.addEventListener('resize',closeContextMenu);
grid.addEventListener('wheel',closeContextMenu,{passive:true});
window.addEventListener('blur',closeContextMenu);
const findDialog=el<HTMLDialogElement>('find-dialog');
function openFind(){findDialog.showModal();el<HTMLInputElement>('find-query').focus();}
el<HTMLButtonElement>('btn-find').addEventListener('click',openFind);
el<HTMLButtonElement>('close-find').addEventListener('click',()=>findDialog.close());
el<HTMLButtonElement>('btn-find-all').addEventListener('click',()=>{
  const query=el<HTMLInputElement>('find-query').value,caseSensitive=el<HTMLInputElement>('find-case').checked;
  const matches=findCells(book,query,caseSensitive);
  el<HTMLDivElement>('find-results').innerHTML=`<p>${matches.length} matches in ${escapeHtml(doc.active.name)}</p>${matches.slice(0,100).map(({ref,raw})=>`<button data-search-jump="${ref}" class="search-result"><b>${ref}</b><span>${escapeHtml(raw.slice(0,95))}</span></button>`).join('')}${matches.length>100?'<p>Showing first 100 matches</p>':''}`;
});
el<HTMLDivElement>('find-results').addEventListener('click',event=>{
  const button=(event.target as HTMLElement).closest<HTMLElement>('[data-search-jump]');if(button?.dataset.searchJump){findDialog.close();jump(button.dataset.searchJump);}
});
el<HTMLButtonElement>('btn-replace-all').addEventListener('click',()=>{
  const query=el<HTMLInputElement>('find-query').value,replacement=el<HTMLInputElement>('replace-text').value;
  const caseSensitive=el<HTMLInputElement>('find-case').checked,includeFormulas=el<HTMLInputElement>('find-formulas').checked;
  try{const edits=replaceCells(book,query,replacement,caseSensitive,includeFormulas);
    if(!edits.length){toast('No matching cells to replace.');return;}
    if(!confirm(`Replace text in ${edits.length} cells? ${includeFormulas?'Formula source will also be changed.':'Formulas will not be changed.'}`))return;
    apply(edits);toast(`Updated ${edits.length} cells. Undo with Ctrl+Z.`);
  }catch(e){toast(e instanceof Error?e.message:'Replace failed');}
});
// Multiple independent sheets with a stable document container. Sheet-local formulas by design.
function changeSheet(id:string){
  if(editing)commitEdit();doc.selectSheet(id);book=doc.active.book;active={r:0,c:0};anchor={...active};end={...active};grid.scrollLeft=0;grid.scrollTop=0;save();refresh();grid.focus({preventScroll:true});
}
el<HTMLDivElement>('sheet-tab-list').addEventListener('click',event=>{const target=(event.target as HTMLElement).closest<HTMLElement>('[data-sheet-id]');if(target?.dataset.sheetId)changeSheet(target.dataset.sheetId);});
el<HTMLButtonElement>('btn-add-sheet').addEventListener('click',()=>{try{const created=doc.addSheet();changeSheet(created.id);toast(`Created ${created.name}`);}catch(e){toast((e as Error).message);}});
el<HTMLButtonElement>('btn-rename-sheet').addEventListener('click',()=>{
  const name=prompt('Rename worksheet:',doc.active.name);if(name===null)return;
  try{doc.renameSheet(doc.activeId,name);save();renderTabs();}catch(e){toast((e as Error).message);}
});
el<HTMLButtonElement>('btn-duplicate-sheet').addEventListener('click',()=>{try{const copy=doc.duplicateSheet();changeSheet(copy.id);toast(`Duplicated as ${copy.name}`);}catch(e){toast((e as Error).message);}});
el<HTMLButtonElement>('btn-delete-sheet').addEventListener('click',()=>{
  if(!confirm(`Delete worksheet ${doc.active.name}? This cannot be undone.`))return;
  try{doc.deleteSheet();changeSheet(doc.activeId);toast('Worksheet deleted');}catch(e){toast((e as Error).message);}
});
el<HTMLElement>('book-title').addEventListener('click',()=>{
  const next=prompt('Workbook title:',workbookTitle);if(next===null)return;workbookTitle=next.trim().slice(0,60)||'Untitled Workbook';
  el<HTMLElement>('book-title').textContent=workbookTitle;save();
});

// GridForge v3: Excel-style productivity actions.
function autoSum(){const selection=sel(),ref=currentRef();
  const col=colToName(active.c),last=selection.r1===selection.r2?active.r-1:selection.r2;
  let first=selection.r1;
  if(selection.r1===selection.r2){first=last;while(first>0){const val=book.getValue(`${col}${first+1}`);if(typeof val.value!=='number'||val.error)break;first--;}
    if(first!==last)first++;}
  if(last<first){toast('Select a range above the destination cell.');return;}
  const destination=selection.r1===selection.r2?ref:`${col}${Math.min(book.rows,selection.r2+2)}`;
  const p=parseRef(destination);if(!p||!inBounds(p.row,p.col)){toast('No space for AutoSum below this selection');return;}
  apply([{ref:destination,raw:`=SUM(${col}${first+1}:${col}${last+1})`}]);select(p.row,p.col);toast(`Inserted AutoSum in ${destination}`);
}
el<HTMLButtonElement>('btn-autosum').addEventListener('click',autoSum);
window.addEventListener('keydown',e=>{if(e.altKey&&e.key==='='){e.preventDefault();autoSum();}});
function freeze(axis:'row'|'col'){
  const tool=doc.active.tools,key=axis==='row'?'freezeRows':'freezeCols';tool[key]=tool[key]?0:1;
  save();renderGrid();toast(`${axis==='row'?'Top row':'First column'} ${tool[key]?'frozen':'unfrozen'}`);
}
el<HTMLButtonElement>('btn-freeze-row').addEventListener('click',()=>freeze('row'));
el<HTMLButtonElement>('btn-freeze-col').addEventListener('click',()=>freeze('col'));
const filterDialog=el<HTMLDialogElement>('filter-dialog');
el<HTMLButtonElement>('btn-filter').addEventListener('click',()=>{
  const s=sel();if(s.r2===s.r1){toast('Select a header row and at least one data row to filter');return;}
  el<HTMLSelectElement>('filter-column').innerHTML=Array.from({length:s.c2-s.c1+1},(_,i)=>{
    const col=s.c1+i,heading=book.getRaw(asRef(s.r1,col));return `<option value="${col}">${colToName(col)} — ${escapeHtml(heading)}</option>`;
  }).join('');el<HTMLInputElement>('filter-query').value=doc.active.tools.filter?.query??'';filterDialog.showModal();
});
el<HTMLButtonElement>('close-filter').addEventListener('click',()=>filterDialog.close());
el<HTMLButtonElement>('apply-filter').addEventListener('click',()=>{
  const b=sel();doc.active.tools.filter={startRow:b.r1,endRow:b.r2,startCol:b.c1,endCol:b.c2,column:Number(el<HTMLSelectElement>('filter-column').value),query:el<HTMLInputElement>('filter-query').value.slice(0,100)};
  save();refresh();filterDialog.close();toast('Filter active: rows are hidden, never deleted');
});
el<HTMLButtonElement>('btn-clear-filter').addEventListener('click',()=>{doc.active.tools.filter=null;save();refresh();toast('All table rows visible');});
const conditionalDialog=el<HTMLDialogElement>('conditional-dialog');
el<HTMLButtonElement>('btn-conditional').addEventListener('click',()=>conditionalDialog.showModal());
el<HTMLButtonElement>('close-conditional').addEventListener('click',()=>conditionalDialog.close());
el<HTMLButtonElement>('apply-conditional').addEventListener('click',()=>{
  const threshold=Number(el<HTMLInputElement>('rule-threshold').value),b=sel();
  if(!Number.isFinite(threshold)){toast('Enter a valid numeric threshold');return;}
  if(doc.active.tools.conditional.length>=30){toast('Maximum 30 rules per sheet');return;}
  doc.active.tools.conditional.push({start:asRef(b.r1,b.c1),end:asRef(b.r2,b.c2),operator:el<HTMLSelectElement>('rule-op').value as 'gt'|'lt'|'eq',threshold,color:el<HTMLInputElement>('rule-color').value});
  save();refresh();conditionalDialog.close();toast('Live conditional formatting rule added');
});
el<HTMLButtonElement>('clear-conditional').addEventListener('click',()=>{
  doc.active.tools.conditional=[];save();refresh();conditionalDialog.close();toast('Conditional formatting rules cleared');
});

refresh();grid.scrollTop=0;grid.scrollLeft=0;grid.focus({preventScroll:true});
