/** XLSX interoperability: Office Open XML reader/writer, NOT a formula engine.
 * Parsed formulas are always evaluated by GridForge's existing handwritten parser.
 * Deliberately bounded importer: never run macros, external links or uploaded code.
 */
import { SpreadsheetDocument } from './document.js';
import { Workbook } from './workbook.js';
import { parseRef, colToName, shiftFormula, tokenize } from './parser.js';
import { cleanStyle } from './format.js';
import { zipWrite, zipRead } from './zip.js';
import { freshTools, validateTools } from './sheettools.js';
const ns = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const relns = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const pkgns = 'http://schemas.openxmlformats.org/package/2006/relationships';
// Modern Excel functions often use _xlfn. in Office Open XML. The evaluator
// only sees native GridForge identifiers, while XLSX output keeps Excel's prefix.
const futureExcelNames = new Set(['XLOOKUP', 'IFS', 'SWITCH', 'TEXTJOIN', 'TEXTBEFORE', 'TEXTAFTER', 'CONCAT', 'REGEXTEST', 'REGEXREPLACE', 'ISOWEEKNUM', 'FILTER', 'SORT', 'SORTBY', 'UNIQUE', 'SEQUENCE', 'RANDARRAY', 'TAKE', 'DROP', 'CHOOSECOLS', 'CHOOSEROWS', 'HSTACK', 'VSTACK', 'TEXTSPLIT', 'TOCOL', 'TOROW', 'WRAPROWS', 'WRAPCOLS', 'BYROW', 'BYCOL', 'MAP', 'REDUCE', 'SCAN', 'LAMBDA', 'LET', 'MAKEARRAY', 'EXPAND', 'TRIMRANGE']);
const excelAliasNames = { NETWORKDAYSINTL: 'NETWORKDAYS.INTL', NORMDIST: 'NORM.DIST', NORMSDIST: 'NORM.S.DIST', NORMINV: 'NORM.INV', NORMSINV: 'NORM.S.INV', COVARIANCEP: 'COVARIANCE.P', COVARIANCES: 'COVARIANCE.S' };
function xlsxFormula(formula) {
    try {
        const tokens = tokenize(formula), replacements = [];
        for (let i = 0; i < tokens.length - 1; i++) {
            const t = tokens[i];
            if (t.kind === 'name' && tokens[i + 1].text === '(' && (futureExcelNames.has(t.text) || excelAliasNames[t.text]))
                replacements.push({ start: t.start, end: t.end, text: `${futureExcelNames.has(t.text) ? '_xlfn.' : ''}${excelAliasNames[t.text] ?? t.text}` });
        }
        let pos = 0, out = '';
        for (const r of replacements) {
            out += formula.slice(pos, r.start) + r.text;
            pos = r.end;
        }
        return out + formula.slice(pos);
    }
    catch {
        return formula;
    }
}
function fromXlsxFormula(formula) {
    // Do not remove extension prefixes from quoted literal text.
    let out = '', inside = false;
    for (let i = 0; i < formula.length;) {
        if (formula[i] === '"') {
            if (inside && formula[i + 1] === '"') {
                out += '""';
                i += 2;
                continue;
            }
            inside = !inside;
            out += '"';
            i++;
            continue;
        }
        if (!inside) {
            const m = /^_xlfn\.(?:_xlws\.)?/i.exec(formula.slice(i));
            if (m) {
                i += m[0].length;
                continue;
            }
        }
        out += formula[i++];
    }
    // Normalize supported dotted Excel function names outside string literals.
    let normalized = '', quoted = false;
    for (let i = 0; i < out.length;) {
        if (out[i] === '"') {
            if (quoted && out[i + 1] === '"') {
                normalized += '""';
                i += 2;
                continue;
            }
            quoted = !quoted;
            normalized += out[i++];
            continue;
        }
        if (!quoted) {
            const match = /^(NETWORKDAYS\.INTL|NORM\.S\.DIST|NORM\.S\.INV|NORM\.DIST|NORM\.INV|COVARIANCE\.P|COVARIANCE\.S)(?=\s*\()/i.exec(out.slice(i));
            if (match) {
                const target = Object.entries(excelAliasNames).find(([, alias]) => alias === match[1].toUpperCase());
                normalized += target?.[0] ?? match[0];
                i += match[0].length;
                continue;
            }
        }
        normalized += out[i++];
    }
    return normalized;
}
const esc = (x) => x.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;').replace(/'/g, '&apos;').replace(/[\u0000-\u0008\u000b\u000c\u000e-\u001f]/g, '');
const xml = (body) => `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>${body}`;
function nodes(node, local) { return Array.from(node.getElementsByTagNameNS('*', local)); }
function direct(node, local) { return Array.from(node.children).find(n => n.localName === local); }
function parseXml(text) {
    if (/<!DOCTYPE|<!ENTITY/i.test(text))
        throw Error('XML external entities are not supported.');
    if (typeof DOMParser === 'undefined')
        throw Error('XLSX parsing requires a modern browser DOM parser.');
    const doc = new DOMParser().parseFromString(text, 'application/xml');
    if (doc.getElementsByTagName('parsererror').length)
        throw Error('Invalid XML inside XLSX.');
    return doc;
}
const attr = (n, name, def = '') => n.getAttribute(name) ?? def;
const int = (x, def = 0) => x && Number.isFinite(Number(x)) ? Number(x) : def;
const color = (n) => {
    const v = n?.getAttribute('rgb');
    return v && /^[0-9a-f]{8}$/i.test(v) ? '#' + v.slice(-6).toLowerCase() : undefined;
};
function fileName(name) { return name.replace(/[\\/:*?"<>|]/g, '_').slice(0, 72) || 'GridForge'; }
function styleCatalog(doc) {
    const map = new Map([['{}', 0]]), styles = [{}];
    for (const sheet of doc.sheets)
        for (const { style } of sheet.book.allStyles()) {
            const normalized = cleanStyle(style), key = JSON.stringify(normalized);
            if (!map.has(key)) {
                map.set(key, styles.length);
                styles.push(normalized);
            }
        }
    const fonts = [{ bold: false, italic: false, underline: false, color: '' }], fills = [''], fontMap = new Map(), fillMap = new Map();
    function fontId(s) { const key = JSON.stringify([!!s.bold, !!s.italic, !!s.underline, s.color ?? '']); if (!fontMap.has(key)) {
        if (!s.bold && !s.italic && !s.underline && !s.color) {
            fontMap.set(key, 0);
        }
        else {
            fontMap.set(key, fonts.length);
            fonts.push({ bold: !!s.bold, italic: !!s.italic, underline: !!s.underline, color: s.color ?? '' });
        }
    } return fontMap.get(key); }
    function fillId(s) { if (!s.background)
        return 0; if (!fillMap.has(s.background)) {
        fillMap.set(s.background, fills.length + 1);
        fills.push(s.background);
    } return fillMap.get(s.background); }
    const xfs = styles.map(s => { const font = fontId(s), fill = fillId(s); return `<xf numFmtId="${({ general: 0, number: 2, integer: 1, currency: 164, percent: 10 })[s.numberFormat ?? 'general']}" fontId="${font}" fillId="${fill}" borderId="0" xfId="0" applyFont="1" applyFill="1" applyNumberFormat="1" applyAlignment="1">${s.align ? `<alignment horizontal="${s.align}"/>` : ''}</xf>`; });
    const fontXML = fonts.map(f => `<font><sz val="11"/><name val="Calibri"/>${f.bold ? '<b/>' : ''}${f.italic ? '<i/>' : ''}${f.underline ? '<u/>' : ''}${f.color ? `<color rgb="FF${f.color.slice(1).toUpperCase()}"/>` : '<color theme="1"/>'}</font>`).join('');
    const fillXML = ['<fill><patternFill patternType="none"/></fill>', '<fill><patternFill patternType="gray125"/></fill>', ...fills.slice(1).map(f => `<fill><patternFill patternType="solid"><fgColor rgb="FF${f.slice(1).toUpperCase()}"/><bgColor indexed="64"/></patternFill></fill>`)].join('');
    return { xml: xml(`<styleSheet xmlns="${ns}"><numFmts count="1"><numFmt numFmtId="164" formatCode="&quot;$&quot;#,##0.00"/></numFmts><fonts count="${fonts.length}">${fontXML}</fonts><fills count="${fills.length + 1}">${fillXML}</fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="${xfs.length}">${xfs.join('')}</cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>`), index: map };
}
function styleFromXlsx(parts) {
    const file = parts.get('xl/styles.xml');
    if (!file)
        return [{}];
    const root = parseXml(file).documentElement;
    const fonts = nodes(direct(root, 'fonts') ?? root, 'font').map(n => ({ bold: !!direct(n, 'b'), italic: !!direct(n, 'i'), underline: !!direct(n, 'u'), color: color(direct(n, 'color')) }));
    const fills = nodes(direct(root, 'fills') ?? root, 'fill').map(n => { const pf = direct(n, 'patternFill'); return attr(pf ?? n, 'patternType') === 'solid' ? color(direct(pf, 'fgColor')) : undefined; });
    const formats = new Map(nodes(direct(root, 'numFmts') ?? root, 'numFmt').map(x => [int(attr(x, 'numFmtId')), attr(x, 'formatCode')]));
    const cellxfs = direct(root, 'cellXfs');
    if (!cellxfs)
        return [{}];
    return nodes(cellxfs, 'xf').map(x => {
        const font = fonts[int(attr(x, 'fontId'))] ?? fonts[0], fill = fills[int(attr(x, 'fillId'))];
        const s = {};
        if (font?.bold)
            s.bold = true;
        if (font?.italic)
            s.italic = true;
        if (font?.underline)
            s.underline = true;
        if (font?.color)
            s.color = font.color;
        if (fill)
            s.background = fill;
        const alignment = direct(x, 'alignment');
        if (alignment && ['center', 'right'].includes(attr(alignment, 'horizontal')))
            s.align = attr(alignment, 'horizontal');
        const fmtId = int(attr(x, 'numFmtId')), fmt = (formats.get(fmtId) ?? '').toLowerCase();
        if (fmtId === 10 || fmtId === 9 || fmt.includes('%'))
            s.numberFormat = 'percent';
        else if ([5, 6, 7, 8, 44].includes(fmtId) || /[$€£]/.test(fmt))
            s.numberFormat = 'currency';
        else if (fmtId === 2 || fmtId === 4 || /0\.0+/.test(fmt))
            s.numberFormat = 'number';
        else if (fmtId === 1 || fmtId === 3)
            s.numberFormat = 'integer';
        return cleanStyle(s);
    });
}
function sheetXML(book, tool, styles) {
    const content = new Map();
    for (const item of book.allEntries()) {
        const p = parseRef(item.ref);
        if (!content.has(p.row))
            content.set(p.row, []);
        content.get(p.row).push(item);
    }
    const styleMap = new Map(book.allStyles().map(({ ref, style }) => [ref, styles.get(JSON.stringify(cleanStyle(style))) ?? 0]));
    for (const ref of styleMap.keys()) {
        const p = parseRef(ref);
        if (!content.has(p.row))
            content.set(p.row, []);
        if (!content.get(p.row).some(c => c.ref === ref))
            content.get(p.row).push({ ref, raw: '' });
    }
    const rows = [...content].sort((a, b) => a[0] - b[0]).map(([r, values]) => `<row r="${r + 1}">${values.sort((a, b) => parseRef(a.ref).col - parseRef(b.ref).col).map(({ ref, raw }) => {
        const st = styleMap.get(ref) ?? 0, attrS = st ? ` s="${st}"` : '';
        if (raw === '')
            return `<c r="${ref}"${attrS}/>`;
        if (raw.startsWith('=')) {
            const result = book.getValue(ref);
            const cached = typeof result.value === 'number' && Number.isFinite(result.value) ? `<v>${result.value}</v>` : '';
            return `<c r="${ref}"${attrS}><f>${esc(xlsxFormula(raw.slice(1)))}</f>${cached}</c>`;
        }
        if (/^'/.test(raw))
            return `<c r="${ref}"${attrS} t="inlineStr"><is><t xml:space="preserve">${esc(raw.slice(1))}</t></is></c>`;
        if (/^(true|false)$/i.test(raw))
            return `<c r="${ref}"${attrS} t="b"><v>${raw.toLowerCase() === 'true' ? 1 : 0}</v></c>`;
        if (raw.trim() !== '' && /^[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:e[+-]?\d+)?$/i.test(raw))
            return `<c r="${ref}"${attrS}><v>${Number(raw)}</v></c>`;
        return `<c r="${ref}"${attrS} t="inlineStr"><is><t xml:space="preserve">${esc(raw)}</t></is></c>`;
    }).join('')}</row>`).join('');
    const pane = tool.freezeRows || tool.freezeCols ? `<pane xSplit="${tool.freezeCols}" ySplit="${tool.freezeRows}" topLeftCell="${colToName(tool.freezeCols)}${tool.freezeRows + 1}" state="frozen"/>` : '';
    const filter = tool.filter ? `<autoFilter ref="${colToName(tool.filter.startCol)}${tool.filter.startRow + 1}:${colToName(tool.filter.endCol)}${tool.filter.endRow + 1}">${tool.filter.query ? `<filterColumn colId="${tool.filter.column - tool.filter.startCol}"><customFilters><customFilter operator="equal" val="*${esc(tool.filter.query)}*"/></customFilters></filterColumn>` : ''}</autoFilter>` : '';
    const cf = tool.conditional.map((r, i) => `<conditionalFormatting sqref="${esc(r.start + ':' + r.end)}"><cfRule type="cellIs" priority="${i + 1}" operator="${({ gt: 'greaterThan', lt: 'lessThan', eq: 'equal' })[r.operator]}" dxfId="${i}"><formula>${r.threshold}</formula></cfRule></conditionalFormatting>`).join('');
    return xml(`<worksheet xmlns="${ns}"><sheetViews><sheetView workbookViewId="0">${pane}</sheetView></sheetViews><sheetFormatPr defaultRowHeight="15"/><sheetData>${rows}</sheetData>${filter}${cf}</worksheet>`);
}
export function exportXlsx(doc) {
    const parts = {}, sty = styleCatalog(doc);
    // DXFs are intentionally not persisted as named rules by Excel unless a dxf is written.
    const dxfs = doc.sheets.flatMap(s => s.tools.conditional);
    // Native conditional rules need a per-sheet index, so the exporter writes a shared differential-format table.
    const dxfNodes = dxfs.map(x => `<dxf><fill><patternFill patternType="solid"><fgColor rgb="FF${x.color.slice(1).toUpperCase()}"/><bgColor indexed="64"/></patternFill></fill></dxf>`).join('');
    // Each sheet's rules get their own global dxf offsets.
    const offsets = [];
    let accum = 0;
    for (const sheet of doc.sheets) {
        offsets.push(accum);
        accum += sheet.tools.conditional.length;
    }
    parts['xl/styles.xml'] = sty.xml.replace('</styleSheet>', `<dxfs count="${dxfs.length}">${dxfNodes}</dxfs></styleSheet>`);
    const typeEntries = [];
    doc.sheets.forEach((sheet, i) => {
        let worksheet = sheetXML(sheet.book, sheet.tools, sty.index);
        const offset = offsets[i];
        let next = 0;
        worksheet = worksheet.replace(/dxfId="\d+"/g, () => `dxfId="${offset + next++}"`);
        parts[`xl/worksheets/sheet${i + 1}.xml`] = worksheet;
        typeEntries.push(`<Override PartName="/xl/worksheets/sheet${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`);
    });
    parts['[Content_Types].xml'] = xml(`<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>${typeEntries.join('')}</Types>`);
    parts['_rels/.rels'] = xml(`<Relationships xmlns="${pkgns}"><Relationship Id="rId1" Type="${relns}/officeDocument" Target="xl/workbook.xml"/></Relationships>`);
    parts['xl/_rels/workbook.xml.rels'] = xml(`<Relationships xmlns="${pkgns}">${doc.sheets.map((_, i) => `<Relationship Id="rId${i + 1}" Type="${relns}/worksheet" Target="worksheets/sheet${i + 1}.xml"/>`).join('')}<Relationship Id="rId${doc.sheets.length + 1}" Type="${relns}/styles" Target="styles.xml"/></Relationships>`);
    parts['xl/workbook.xml'] = xml(`<workbook xmlns="${ns}" xmlns:r="${relns}"><bookViews><workbookView activeTab="${Math.max(0, doc.sheets.findIndex(s => s.id === doc.activeId))}"/></bookViews><sheets>${doc.sheets.map((s, i) => `<sheet name="${esc(s.name)}" sheetId="${i + 1}" r:id="rId${i + 1}"/>`).join('')}</sheets><calcPr calcId="191029" fullCalcOnLoad="1"/></workbook>`);
    return zipWrite(parts);
}
export async function importXlsx(bytes) {
    if (bytes.length > 25 * 1024 * 1024)
        throw Error('XLSX exceeds the 25 MB import limit.');
    const parts = await zipRead(bytes), warnings = [];
    const workbookData = parts.get('xl/workbook.xml');
    if (!workbookData)
        throw Error('Not an XLSX workbook: missing xl/workbook.xml');
    const workbook = parseXml(workbookData), rels = parseXml(parts.get('xl/_rels/workbook.xml.rels') ?? '<Relationships/>');
    const relation = new Map(nodes(rels, 'Relationship').map(n => [attr(n, 'Id'), attr(n, 'Target')]));
    const sheets = nodes(workbook, 'sheet');
    if (!sheets.length || sheets.length > 12)
        throw Error('XLSX must contain between 1 and 12 worksheets.');
    const shared = parts.get('xl/sharedStrings.xml') ? nodes(parseXml(parts.get('xl/sharedStrings.xml')), 'si').map(si => nodes(si, 't').map(t => t.textContent ?? '').join('')) : [];
    const catalog = styleFromXlsx(parts);
    const dxfXml = parts.get('xl/styles.xml'), dxfs = dxfXml ? nodes(direct(parseXml(dxfXml).documentElement, 'dxfs') ?? parseXml(dxfXml).documentElement, 'dxf').map(node => color(nodes(node, 'fgColor')[0]) ?? '#ffe0a2') : [];
    const result = new SpreadsheetDocument(); // Replace the default sheet with the first imported sheet below.
    let first = true;
    for (const item of sheets) {
        const name = attr(item, 'name'), relId = item.getAttributeNS(relns, 'id') ?? attr(item, 'r:id');
        const target = relation.get(relId);
        if (!target)
            throw Error(`Worksheet ${name} has no relationship target.`);
        // External targets and path traversal are not permitted.
        const relative = target.replace(/^\//, '');
        const path = relative.startsWith('xl/') ? relative : `xl/${relative}`;
        if (path.includes('..') || !path.startsWith('xl/worksheets/'))
            throw Error(`Unsupported worksheet target ${target}`);
        const content = parts.get(path);
        if (!content)
            throw Error(`Missing worksheet ${path}`);
        const sheet = parseXml(content), cells = nodes(sheet, 'c'), edits = [], styleEdits = [];
        let maxR = 400, maxC = 52;
        const sharedMasters = new Map();
        for (const source of cells) {
            const f = direct(source, 'f');
            if (f && attr(f, 't') === 'shared' && (f.textContent ?? '').trim()) {
                sharedMasters.set(attr(f, 'si'), { ref: attr(source, 'r'), formula: '=' + fromXlsxFormula((f.textContent ?? '').trim()) });
            }
        }
        if (cells.length > 50000)
            throw Error(`${name} has more than 50,000 cells. Split this workbook before importing.`);
        for (const c of cells) {
            const ref = attr(c, 'r'), p = parseRef(ref);
            if (!p)
                continue;
            if (p.row >= 1000 || p.col >= 104)
                throw Error(`${name}!${ref} is outside GridForge's 1000-row × 104-column import limit.`);
            maxR = Math.max(maxR, p.row + 1);
            maxC = Math.max(maxC, p.col + 1);
            const f = direct(c, 'f'), type = attr(c, 't'), value = direct(c, 'v')?.textContent ?? '', inline = direct(c, 'is');
            let raw = '';
            if (f) {
                const formula = (f.textContent ?? '').trim();
                if (formula)
                    raw = '=' + fromXlsxFormula(formula);
                else if (attr(f, 't') === 'shared' && sharedMasters.has(attr(f, 'si'))) {
                    const master = sharedMasters.get(attr(f, 'si')), origin = parseRef(master.ref);
                    try {
                        raw = shiftFormula(master.formula, p.row - origin.row, p.col - origin.col);
                    }
                    catch {
                        warnings.push(`${name}!${ref}: could not shift shared formula.`);
                        raw = value;
                    }
                }
                else {
                    warnings.push(`${name}!${ref}: unsupported shared/array formula (formula text unavailable).`);
                    raw = value;
                }
            }
            else if (type === 's') {
                raw = shared[int(value)] ?? '';
                if (raw.startsWith('='))
                    raw = "'" + raw;
            }
            else if (type === 'inlineStr') {
                raw = inline ? nodes(inline, 't').map(n => n.textContent ?? '').join('') : '';
                if (raw.startsWith('='))
                    raw = "'" + raw;
            }
            else if (type === 'b')
                raw = value === '1' ? 'TRUE' : 'FALSE';
            else if (type === 'e') {
                warnings.push(`${name}!${ref}: cached Excel error imported as text.`);
                raw = "'" + value;
            }
            else if (type === 'str') {
                raw = value;
                if (raw.startsWith('='))
                    raw = "'" + raw;
            }
            else
                raw = value;
            if (raw)
                edits.push({ ref, raw });
            const st = catalog[int(attr(c, 's'))] ?? {};
            if (Object.keys(st).length)
                styleEdits.push({ ref, style: st });
        }
        const book = new Workbook(maxR, maxC);
        book.applyEdits(edits, false);
        book.applyStyles(styleEdits, false);
        let unsupported = 0;
        for (const item of edits)
            if (item.raw.startsWith('=')) {
                const fault = book.getValue(item.ref).error;
                if (fault && ['#PARSE!', '#NAME?', '#REF!'].includes(fault.code)) {
                    unsupported++;
                    if (unsupported <= 5)
                        warnings.push(`${name}!${item.ref}: ${fault.code} — ${fault.message}`);
                }
            }
        if (unsupported > 5)
            warnings.push(`${name}: ${unsupported - 5} additional formulas could not be evaluated by the custom engine.`);
        let added;
        if (first) {
            const old = result.activeId;
            result.renameSheet(old, name);
            added = result.active; // Cannot resize initial sheet; swap through JSON once all sheets are loaded.
            first = false;
            added.book = book;
        }
        else
            added = result.addSheet(name, book);
        const tool = freshTools(), pane = nodes(sheet, 'pane')[0];
        if (pane) {
            tool.freezeRows = Math.min(maxR - 1, Math.max(0, int(attr(pane, 'ySplit'))));
            tool.freezeCols = Math.min(maxC - 1, Math.max(0, int(attr(pane, 'xSplit'))));
        }
        const auto = nodes(sheet, 'autoFilter')[0];
        if (auto) {
            const m = /^(\$?[A-Z]+\$?\d+):(\$?[A-Z]+\$?\d+)$/i.exec(attr(auto, 'ref'));
            if (m) {
                const a = parseRef(m[1]), b = parseRef(m[2]);
                if (a && b && b.row > a.row) {
                    const column = nodes(auto, 'filterColumn')[0], custom = column ? nodes(column, 'customFilter')[0] : undefined;
                    tool.filter = { startRow: a.row, endRow: Math.min(b.row, maxR - 1), startCol: a.col, endCol: Math.min(b.col, maxC - 1), column: Math.min(maxC - 1, a.col + int(column ? attr(column, 'colId') : '0')), query: custom ? attr(custom, 'val').replace(/^\*|\*$/g, '') : '' };
                }
            }
        }
        for (const node of nodes(sheet, 'conditionalFormatting')) {
            const range = attr(node, 'sqref').split(/\s+/)[0] ?? '', m = /^([A-Z]+\d+):([A-Z]+\d+)$/.exec(range), rule = nodes(node, 'cfRule')[0];
            if (!m || !rule || attr(rule, 'type') !== 'cellIs')
                continue;
            const op = { greaterThan: 'gt', lessThan: 'lt', equal: 'eq' }[attr(rule, 'operator')], threshold = Number(nodes(rule, 'formula')[0]?.textContent ?? 'NaN');
            if (!op || !Number.isFinite(threshold))
                continue;
            tool.conditional.push({ start: m[1], end: m[2], operator: op, threshold, color: dxfs[int(attr(rule, 'dxfId'))] ?? '#ffe0a2' });
        }
        added.tools = validateTools(tool, maxR, maxC);
    }
    const activeTab = int(attr(nodes(workbook, 'workbookView')[0] ?? workbook.documentElement, 'activeTab'));
    if (result.sheets[activeTab])
        result.selectSheet(result.sheets[activeTab].id);
    warnings.push('GridForge previews only a supported subset of Excel. Rebuilt XLSX exports omit VBA and some unsupported objects; when an XLSM is imported, Save XLSM can preserve unrelated original OOXML parts while merging supported cell edits. VBA never runs during import.');
    return { document: result, warnings };
}
export function xlsxFileName(title) { return `${fileName(title)}.xlsx`; }
