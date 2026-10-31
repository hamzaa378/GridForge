import { zipReadBytes, zipWriteBytes } from './zip.js';
import { colToName, parseRef, tokenize } from './parser.js';
const MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const DECODER = new TextDecoder(), ENCODER = new TextEncoder();
const nodeList = (n, tag) => Array.from(n.getElementsByTagNameNS('*', tag));
const getAttribute = (n, key) => n.getAttribute(key) ?? '';
function readXml(input) {
    if (/<!DOCTYPE|<!ENTITY/i.test(input))
        throw Error('Unsafe OOXML XML entities.');
    const result = new DOMParser().parseFromString(input, 'application/xml');
    if (result.getElementsByTagName('parsererror').length)
        throw Error('Invalid OOXML XML.');
    return result;
}
function writeXml(doc) { const text = new XMLSerializer().serializeToString(doc); return ENCODER.encode(text.trimStart().startsWith('<?xml') ? text : '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' + text); }
function cellChildren(c, tag) { return Array.from(c.children).find(x => x.localName === tag); }
function nextInside(parent, tag, index) {
    return Array.from(parent.children).find(x => x.localName === tag && Number(x.getAttribute('r')) > index);
}
function createCell(sheet, reference) {
    const p = parseRef(reference);
    if (!p)
        throw Error(`Invalid cell ${reference}`);
    const table = nodeList(sheet, 'sheetData')[0];
    if (!table)
        throw Error('Worksheet has no sheetData.');
    let row = Array.from(table.children).find(x => x.localName === 'row' && Number(x.getAttribute('r')) === p.row + 1);
    if (!row) {
        row = sheet.createElementNS(MAIN, 'row');
        row.setAttribute('r', String(p.row + 1));
        table.insertBefore(row, nextInside(table, 'row', p.row + 1) ?? null);
    }
    const c = sheet.createElementNS(MAIN, 'c');
    c.setAttribute('r', reference);
    const after = Array.from(row.children).find(x => { if (x.localName !== 'c')
        return false; const other = parseRef(x.getAttribute('r') ?? ''); return !!other && other.col > p.col; });
    row.insertBefore(c, after ?? null);
    return c;
}
function formulaForExcel(formula) {
    // Preserve supported future-function names during edited XLSM formula writes.
    // No spreadsheet engine is involved; this is only a format convention.
    const future = new Set(['XLOOKUP', 'IFS', 'SWITCH', 'TEXTJOIN', 'TEXTBEFORE', 'TEXTAFTER', 'CONCAT']);
    try {
        let out = formula;
        const matches = tokenize(formula).filter(x => x.kind === 'name' && future.has(x.text));
        for (const t of matches.reverse()) {
            if (out.slice(t.end).trimStart().startsWith('('))
                out = out.slice(0, t.start) + '_xlfn.' + out.slice(t.start);
        }
        return out;
    }
    catch {
        return formula;
    }
}
function updateCell(doc, c, raw) {
    if (['shared', 'array', 'dataTable'].includes(cellChildren(c, 'f')?.getAttribute('t') ?? ''))
        throw Error(`Excel shared/array/data-table formula in ${getAttribute(c, 'r')} cannot be modified safely in macro-preserving mode. Use desktop Excel.`);
    for (const child of Array.from(c.children))
        if (['f', 'v', 'is'].includes(child.localName))
            c.removeChild(child);
    c.removeAttribute('t');
    if (raw === '')
        return;
    if (raw.startsWith('=')) {
        const f = doc.createElementNS(MAIN, 'f');
        f.textContent = formulaForExcel(raw.slice(1));
        c.appendChild(f);
    }
    else if (/^'/.test(raw)) {
        c.setAttribute('t', 'inlineStr');
        const is = doc.createElementNS(MAIN, 'is'), t = doc.createElementNS(MAIN, 't');
        t.setAttribute('xml:space', 'preserve');
        t.textContent = raw.slice(1);
        is.appendChild(t);
        c.appendChild(is);
    }
    else if (/^(TRUE|FALSE)$/i.test(raw)) {
        c.setAttribute('t', 'b');
        const v = doc.createElementNS(MAIN, 'v');
        v.textContent = /^TRUE$/i.test(raw) ? '1' : '0';
        c.appendChild(v);
    }
    else if (/^[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?$/.test(raw) && Number.isFinite(Number(raw))) {
        const v = doc.createElementNS(MAIN, 'v');
        v.textContent = String(Number(raw));
        c.appendChild(v);
    }
    else {
        c.setAttribute('t', 'inlineStr');
        const is = doc.createElementNS(MAIN, 'is'), t = doc.createElementNS(MAIN, 't');
        t.setAttribute('xml:space', 'preserve');
        t.textContent = raw;
        is.appendChild(t);
        c.appendChild(is);
    }
}
function originalSheetParts(parts) {
    const workbook = parts.get('xl/workbook.xml'), rels = parts.get('xl/_rels/workbook.xml.rels');
    if (!workbook || !rels)
        throw Error('Missing required XLSM workbook relationships.');
    const wb = readXml(DECODER.decode(workbook)), rs = readXml(DECODER.decode(rels));
    const targets = new Map(nodeList(rs, 'Relationship').map(x => [getAttribute(x, 'Id'), getAttribute(x, 'Target')]));
    const files = new Map();
    for (const s of nodeList(wb, 'sheet')) {
        const name = getAttribute(s, 'name'), id = s.getAttributeNS(REL, 'id') ?? getAttribute(s, 'r:id');
        const target = targets.get(id);
        if (!target)
            throw Error(`Missing sheet relationship for ${name}.`);
        const location = target.startsWith('/') ? target.substring(1) : (target.startsWith('xl/') ? target : `xl/${target}`);
        if (!/^xl\/worksheets\/[a-z0-9_.-]+\.xml$/i.test(location) || !parts.has(location))
            throw Error(`Unsafe or missing worksheet path: ${location}`);
        files.set(name, location);
    }
    return files;
}
export async function exportEditedXlsm(original, before, after) {
    if (original.length > 25 * 1024 * 1024)
        throw Error('Source XLSM exceeds the 25 MB import limit.');
    const parts = await zipReadBytes(original);
    const types = parts.get('[Content_Types].xml');
    if (!types || !DECODER.decode(types).includes('macroEnabled.main+xml'))
        throw Error('Original file is not an Excel macro-enabled workbook.');
    if (after.sheets.length !== before.sheets.length || after.sheets.some((sheet, i) => sheet.name !== before.sheets[i]?.name)) {
        throw Error('Adding, deleting, reordering or renaming sheets cannot be merged into the original XLSM safely. Restore the original sheets or export XLSX.');
    }
    const sheetPaths = originalSheetParts(parts);
    const warnings = [];
    let editedCells = 0, editedNames = 0, styleEdits = 0;
    const workbook = readXml(DECODER.decode(parts.get('xl/workbook.xml')));
    for (const [sheetIndex, sheet] of after.sheets.entries()) {
        const source = before.sheets.find(x => x.name === sheet.name);
        const path = sheetPaths.get(sheet.name);
        if (!path)
            throw Error(`Missing original worksheet '${sheet.name}'.`);
        const originalEntries = new Map(source.book.allEntries().map(x => [x.ref, x.raw]));
        const currentEntries = new Map(sheet.book.allEntries().map(x => [x.ref, x.raw]));
        const changed = [...new Set([...originalEntries.keys(), ...currentEntries.keys()])].filter(ref => (originalEntries.get(ref) ?? '') !== (currentEntries.get(ref) ?? ''));
        if (editedCells + changed.length > 5000)
            throw Error('Macro-preserving export limited to 5,000 changed cells.');
        const beforeStyles = new Map(source.book.allStyles().map(x => [x.ref, JSON.stringify(x.style)]));
        const afterStyles = new Map(sheet.book.allStyles().map(x => [x.ref, JSON.stringify(x.style)]));
        styleEdits += [...new Set([...beforeStyles.keys(), ...afterStyles.keys()])].filter(ref => (beforeStyles.get(ref) ?? '{}') !== (afterStyles.get(ref) ?? '{}')).length;
        if (changed.length) {
            const sheetXml = readXml(DECODER.decode(parts.get(path)));
            const existing = new Map(nodeList(sheetXml, 'c').map(x => [getAttribute(x, 'r').toUpperCase(), x]));
            for (const ref of changed) {
                const c = existing.get(ref) ?? createCell(sheetXml, ref);
                updateCell(sheetXml, c, currentEntries.get(ref) ?? '');
                editedCells++;
            }
            parts.set(path, writeXml(sheetXml));
        }
        const oldNames = source.book.getDefinedNames(), newNames = sheet.book.getDefinedNames();
        if (JSON.stringify(oldNames) !== JSON.stringify(newNames)) {
            const oldByKey = new Map(oldNames.map(n => [n.name.toLowerCase(), n])), newByKey = new Map(newNames.map(n => [n.name.toLowerCase(), n]));
            const changedKeys = new Set([...oldByKey.keys(), ...newByKey.keys()].filter(key => JSON.stringify(oldByKey.get(key)) !== JSON.stringify(newByKey.get(key))));
            let container = nodeList(workbook, 'definedNames')[0];
            if (!container) {
                container = workbook.createElementNS(MAIN, 'definedNames');
                const calc = nodeList(workbook, 'calcPr')[0];
                workbook.documentElement.insertBefore(container, calc ?? null);
            }
            for (const node of nodeList(container, 'definedName'))
                if (getAttribute(node, 'localSheetId') !== '' && Number(getAttribute(node, 'localSheetId')) === sheetIndex && changedKeys.has(getAttribute(node, 'name').toLowerCase()))
                    container.removeChild(node);
            for (const item of newNames) {
                if (!changedKeys.has(item.name.toLowerCase()))
                    continue;
                const node = workbook.createElementNS(MAIN, 'definedName');
                node.setAttribute('name', item.name);
                node.setAttribute('localSheetId', String(sheetIndex));
                if (item.comment)
                    node.setAttribute('comment', item.comment);
                if (item.hidden)
                    node.setAttribute('hidden', '1');
                const range = item.ref.startsWith('=') ? item.ref.slice(1) : `'${sheet.name.replace(/'/g, "''")}'!${item.ref.split(':').map(ref => { const p = parseRef(ref); return `$${colToName(p.col)}$${p.row + 1}`; }).join(':')}`;
                node.textContent = range;
                container.appendChild(node);
                editedNames++;
            }
            editedNames += oldNames.filter(item => changedKeys.has(item.name.toLowerCase()) && !newByKey.has(item.name.toLowerCase())).length;
        }
    }
    const oldWorkbookNames = before.getWorkbookDefinedNames(), newWorkbookNames = after.getWorkbookDefinedNames();
    if (JSON.stringify(oldWorkbookNames) !== JSON.stringify(newWorkbookNames)) {
        const oldByKey = new Map(oldWorkbookNames.map(item => [item.name.toLowerCase(), item])), newByKey = new Map(newWorkbookNames.map(item => [item.name.toLowerCase(), item]));
        const changedKeys = new Set([...new Set([...oldByKey.keys(), ...newByKey.keys()])].filter(key => JSON.stringify(oldByKey.get(key)) !== JSON.stringify(newByKey.get(key))));
        let container = nodeList(workbook, 'definedNames')[0];
        if (!container) {
            container = workbook.createElementNS(MAIN, 'definedNames');
            const calc = nodeList(workbook, 'calcPr')[0];
            workbook.documentElement.insertBefore(container, calc ?? null);
        }
        for (const node of nodeList(container, 'definedName'))
            if (getAttribute(node, 'localSheetId') === '' && changedKeys.has(getAttribute(node, 'name').toLowerCase()))
                container.removeChild(node);
        for (const item of newWorkbookNames) {
            if (!changedKeys.has(item.name.toLowerCase()))
                continue;
            const node = workbook.createElementNS(MAIN, 'definedName');
            node.setAttribute('name', item.name);
            if (item.comment)
                node.setAttribute('comment', item.comment);
            if (item.hidden)
                node.setAttribute('hidden', '1');
            node.textContent = formulaForExcel(item.ref.replace(/^=/, ''));
            container.appendChild(node);
            editedNames++;
        }
        editedNames += oldWorkbookNames.filter(item => changedKeys.has(item.name.toLowerCase()) && !newByKey.has(item.name.toLowerCase())).length;
    }
    // Recalculate formula caches when Microsoft Excel opens the file again.
    if (editedCells || editedNames) {
        let calc = nodeList(workbook, 'calcPr')[0];
        if (!calc) {
            calc = workbook.createElementNS(MAIN, 'calcPr');
            workbook.documentElement.appendChild(calc);
        }
        calc.setAttribute('fullCalcOnLoad', '1');
        calc.setAttribute('forceFullCalc', '1');
        parts.set('xl/workbook.xml', writeXml(workbook));
    }
    if (styleEdits)
        warnings.push(`${styleEdits} GridForge formatting changes were not merged. Original Excel formatting is preserved; edit styles in Excel for this XLSM.`);
    warnings.push('VBA project and unrelated Office parts are preserved byte-for-byte as ZIP entry payloads, not executed.');
    warnings.push('Digital document-level signatures may become invalid when worksheets change. Back up and verify this file in Microsoft Excel.');
    return { bytes: zipWriteBytes(Object.fromEntries(parts)), editedCells, editedNames, preservedParts: parts.size, warnings };
}
