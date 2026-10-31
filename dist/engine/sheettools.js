import { display } from './workbook.js';
import { parseRef, colToName, nameToCol } from './parser.js';
import { validateChart } from './charts.js';
import { validateTable, tableRange, intersectTables } from './tables.js';
import { validatePivotSummary } from './pivot.js';
import { excelDateSerial } from './date_serial.js';
import { validateQuerySpec } from './power_query.js';
export const defaultPageSetup = () => ({ orientation: 'portrait', paper: 'letter', pageOrder: 'downThenOver', centerHorizontally: false, centerVertically: false, fitWidth: 1, fitHeight: 0, printGridlines: false, printHeadings: false, marginTop: 0.75, marginBottom: 0.75, marginLeft: 0.7, marginRight: 0.7 });
export const freshTools = () => ({ freezeRows: 0, freezeCols: 0, filter: null, conditional: [], dataValidations: [], hyperlinks: [], charts: [], drawings: [], tables: [], pivotTables: [], pivotSlicers: [], queries: [], pageSetup: defaultPageSetup(), columnWidths: {}, rowHeights: {}, hiddenRows: [], hiddenColumns: [] });
export function validateTools(data, rows, cols) {
    const t = freshTools();
    if (!data || typeof data !== 'object')
        return t;
    const obj = data;
    if (obj.pageSetup !== undefined) {
        const p = obj.pageSetup;
        if (!p || typeof p !== 'object' || !['portrait', 'landscape'].includes(p.orientation) || !['letter', 'legal', 'a3', 'a4', 'tabloid'].includes(p.paper) || (p.pageOrder !== undefined && !['downThenOver', 'overThenDown'].includes(p.pageOrder)) || ![p.fitWidth, p.fitHeight].every(v => Number.isInteger(v) && v >= 0 && v <= 4) || ![p.printGridlines, p.printHeadings].every(v => typeof v === 'boolean') || [p.centerHorizontally, p.centerVertically].some(v => v !== undefined && typeof v !== 'boolean') || ![p.marginTop, p.marginBottom, p.marginLeft, p.marginRight].every(v => Number.isFinite(v) && v >= 0 && v <= 2))
            throw Error('Invalid worksheet page setup.');
        for (const key of ['headerLeft', 'headerCenter', 'headerRight', 'footerLeft', 'footerCenter', 'footerRight'])
            if (p[key] !== undefined && (typeof p[key] !== 'string' || p[key].length > 255))
                throw Error(`${key} must be text with no more than 255 characters.`);
        const printArea = p.printArea?.replace(/\$/g, '').toUpperCase();
        if (p.printArea !== undefined) {
            const parts = printArea.split(':'), a = parseRef(parts[0] ?? ''), b = parseRef(parts[1] ?? parts[0] ?? '');
            if (parts.length > 2 || !a || !b || a.row > b.row || a.col > b.col || b.row >= rows || b.col >= cols)
                throw Error('Print area must be a valid in-worksheet range.');
        }
        let repeatRows, repeatColumns;
        if (p.repeatRows !== undefined) {
            const value = p.repeatRows.replace(/\$/g, '');
            const match = /^([1-9]\d*):([1-9]\d*)$/.exec(value);
            if (!match || Number(match[1]) > Number(match[2]) || Number(match[2]) > 1_048_576)
                throw Error('Print title rows must be a whole-row range such as 1:3.');
            repeatRows = `${Number(match[1])}:${Number(match[2])}`;
        }
        if (p.repeatColumns !== undefined) {
            const value = p.repeatColumns.replace(/\$/g, '').toUpperCase();
            const match = /^([A-Z]{1,3}):([A-Z]{1,3})$/.exec(value), start = match ? nameToCol(match[1]) : -1, end = match ? nameToCol(match[2]) : -1;
            if (!match || start < 0 || start > end || end > 16_383)
                throw Error('Print title columns must be a whole-column range such as A:B.');
            repeatColumns = `${colToName(start)}:${colToName(end)}`;
        }
        t.pageSetup = { ...p, pageOrder: p.pageOrder ?? 'downThenOver', centerHorizontally: p.centerHorizontally ?? false, centerVertically: p.centerVertically ?? false, ...(printArea ? { printArea: `${colToName(parseRef(printArea.split(':')[0]).col)}${parseRef(printArea.split(':')[0]).row + 1}:${colToName(parseRef(printArea.split(':')[1] ?? printArea.split(':')[0]).col)}${parseRef(printArea.split(':')[1] ?? printArea.split(':')[0]).row + 1}` } : {}), ...(repeatRows ? { repeatRows } : {}), ...(repeatColumns ? { repeatColumns } : {}) };
    }
    for (const [key, max, limit] of [['columnWidths', 8, 600], ['rowHeights', 12, 400]]) {
        const values = obj[key];
        if (values !== undefined) {
            if (!values || typeof values !== 'object' || Array.isArray(values))
                throw Error(`Invalid ${key}`);
            const clean = {};
            for (const [index, value] of Object.entries(values)) {
                const i = Number(index);
                if (!Number.isInteger(i) || i < 0 || i >= (key === 'columnWidths' ? cols : rows) || typeof value !== 'number' || !Number.isFinite(value) || value < max || value > limit)
                    throw Error(`Invalid ${key}`);
                clean[i] = value;
            }
            t[key] = clean;
        }
    }
    for (const key of ['freezeRows', 'freezeCols']) {
        const v = obj[key];
        if (v !== undefined) {
            if (!Number.isInteger(v) || v < 0 || v > (key === 'freezeRows' ? rows - 1 : cols - 1))
                throw Error(`Invalid ${key}`);
            t[key] = v;
        }
    }
    for (const [key, limit] of [['hiddenRows', rows], ['hiddenColumns', cols]]) {
        const values = obj[key];
        if (values !== undefined) {
            if (!Array.isArray(values) || values.length >= limit || values.length > limit || values.some(v => !Number.isInteger(v) || v < 0 || v >= limit) || new Set(values).size !== values.length)
                throw Error(`Invalid ${key}; at least one ${key === 'hiddenRows' ? 'row' : 'column'} must remain visible.`);
            t[key] = [...values].sort((a, b) => a - b);
        }
    }
    if (obj.filter) {
        const f = obj.filter;
        if (![f.startRow, f.endRow, f.startCol, f.endCol, f.column].every(Number.isInteger) || f.startRow < 0 || f.endRow >= rows || f.endRow <= f.startRow || f.startCol < 0 || f.endCol >= cols || f.column < f.startCol || f.column > f.endCol || typeof f.query !== 'string' || f.query.length > 100)
            throw Error('Invalid table filter');
        if (f.criteria !== undefined) {
            if (!Array.isArray(f.criteria) || f.criteria.length > 16 || new Set(f.criteria.map(c => c.column)).size !== f.criteria.length)
                throw Error('Invalid column-filter rules');
            for (const criterion of f.criteria) {
                if (!Number.isInteger(criterion.column) || criterion.column < f.startCol || criterion.column > f.endCol || !Array.isArray(criterion.values) || criterion.values.length > 150 || criterion.values.some(v => typeof v !== 'string' || v.length > 200))
                    throw Error('Invalid filter values');
            }
        }
        t.filter = { ...f, criteria: f.criteria?.map(c => ({ column: c.column, values: [...c.values] })) };
    }
    if (obj.conditional) {
        if (!Array.isArray(obj.conditional) || obj.conditional.length > 30)
            throw Error('Too many conditional rules');
        for (const rule of obj.conditional) {
            const a = parseRef(rule.start), b = parseRef(rule.end);
            if (!a || !b || a.row >= rows || b.row >= rows || a.col >= cols || b.col >= cols || !['gt', 'gte', 'lt', 'lte', 'eq', 'ne', 'between', 'notBetween', 'scale', 'dataBar'].includes(rule.operator) || !Number.isFinite(rule.threshold) || !/^#[0-9a-f]{6}$/i.test(rule.color) || (rule.operator === 'between' || rule.operator === 'notBetween') && (!Number.isFinite(rule.threshold2) || rule.threshold > rule.threshold2) || (rule.operator === 'scale' && ((rule.startColor !== undefined && !/^#[0-9a-f]{6}$/i.test(rule.startColor)) || (rule.midColor !== undefined && !/^#[0-9a-f]{6}$/i.test(rule.midColor)))) || (rule.operator === 'dataBar' && ((rule.barStyle !== undefined && !['solid', 'gradient'].includes(rule.barStyle)) || (rule.negativeColor !== undefined && !/^#[0-9a-f]{6}$/i.test(rule.negativeColor)) || (rule.minValue !== undefined && !Number.isFinite(rule.minValue)) || (rule.maxValue !== undefined && !Number.isFinite(rule.maxValue)) || (rule.minValue !== undefined && rule.maxValue !== undefined && rule.minValue >= rule.maxValue) || (rule.showValue !== undefined && typeof rule.showValue !== 'boolean'))))
                throw Error('Invalid conditional formatting rule');
            t.conditional.push({ ...rule });
        }
    }
    if (obj.dataValidations !== undefined) {
        if (!Array.isArray(obj.dataValidations) || obj.dataValidations.length > 100)
            throw Error('Too many data validation rules.');
        for (const raw of obj.dataValidations) {
            if (!raw || typeof raw !== 'object')
                throw Error('Invalid data validation rule.');
            const rule = raw, a = parseRef(rule.start), b = parseRef(rule.end);
            if (!a || !b || a.row >= rows || b.row >= rows || a.col >= cols || b.col >= cols || a.row > b.row || a.col > b.col || !['list', 'whole', 'decimal', 'date', 'textLength'].includes(rule.type) || typeof rule.allowBlank !== 'boolean')
                throw Error('Invalid data validation range or type.');
            if ((rule.errorStyle !== undefined && !['stop', 'warning', 'information'].includes(rule.errorStyle)) || [rule.showErrorMessage, rule.showInputMessage].some(value => value !== undefined && typeof value !== 'boolean') || [rule.promptTitle, rule.promptMessage, rule.errorTitle, rule.errorMessage].some(value => value !== undefined && (typeof value !== 'string' || value.length > 255)))
                throw Error('Invalid data validation prompt or alert settings.');
            const messages = { ...(rule.showErrorMessage !== undefined ? { showErrorMessage: rule.showErrorMessage } : {}), ...(rule.showInputMessage !== undefined ? { showInputMessage: rule.showInputMessage } : {}), ...(rule.errorStyle ? { errorStyle: rule.errorStyle } : {}), ...(rule.promptTitle ? { promptTitle: rule.promptTitle } : {}), ...(rule.promptMessage ? { promptMessage: rule.promptMessage } : {}), ...(rule.errorTitle ? { errorTitle: rule.errorTitle } : {}), ...(rule.errorMessage ? { errorMessage: rule.errorMessage } : {}) };
            if (rule.type === 'list') {
                if (rule.sourceRange !== undefined) {
                    const range = rule.sourceRange.replace(/\$/g, '').toUpperCase().split(':');
                    const a = parseRef(range[0] ?? ''), b = parseRef(range[1] ?? range[0] ?? '');
                    if (rule.values !== undefined || !a || !b || a.row > b.row || a.col > b.col || b.row >= rows || b.col >= cols)
                        throw Error('List source must be a valid range on this worksheet.');
                    t.dataValidations.push({ start: rule.start.toUpperCase(), end: rule.end.toUpperCase(), type: 'list', sourceRange: `${colToName(a.col)}${a.row + 1}:${colToName(b.col)}${b.row + 1}`, allowBlank: rule.allowBlank, ...messages });
                }
                else {
                    if (!Array.isArray(rule.values) || rule.values.length < 1 || rule.values.length > 50 || rule.values.some(v => typeof v !== 'string' || !v.trim() || v.length > 80) || rule.values.join(',').length > 255)
                        throw Error('List validation needs 1–50 values and at most 255 total characters.');
                    t.dataValidations.push({ start: rule.start.toUpperCase(), end: rule.end.toUpperCase(), type: 'list', values: [...rule.values], allowBlank: rule.allowBlank, ...messages });
                }
            }
            else {
                const operator = rule.operator ?? (rule.min !== undefined && rule.max !== undefined ? 'between' : rule.min !== undefined ? 'greaterThanOrEqual' : 'lessThanOrEqual'), two = ['between', 'notBetween'].includes(operator);
                if (rule.values !== undefined || rule.sourceRange !== undefined || !['between', 'notBetween', 'equalTo', 'notEqualTo', 'greaterThan', 'greaterThanOrEqual', 'lessThan', 'lessThanOrEqual'].includes(operator) || (two && (rule.min === undefined || rule.max === undefined)) || (!two && (rule.min === undefined || rule.max !== undefined)) || (rule.min !== undefined && !Number.isFinite(rule.min)) || (rule.max !== undefined && !Number.isFinite(rule.max)) || (rule.min !== undefined && rule.max !== undefined && rule.min > rule.max) || (rule.type === 'textLength' && [rule.min, rule.max].some(value => value !== undefined && !Number.isInteger(value))))
                    throw Error('Validation needs criteria that match its comparison operator. Text-length limits must be whole numbers.');
                t.dataValidations.push({ start: rule.start.toUpperCase(), end: rule.end.toUpperCase(), type: rule.type, operator, ...(rule.min === undefined ? {} : { min: rule.min }), ...(rule.max === undefined ? {} : { max: rule.max }), allowBlank: rule.allowBlank, ...messages });
            }
        }
    }
    if (obj.hyperlinks !== undefined) {
        if (!Array.isArray(obj.hyperlinks) || obj.hyperlinks.length > 10_000)
            throw Error('A worksheet can contain at most 10,000 hyperlinks.');
        const refs = new Set();
        for (const raw of obj.hyperlinks) {
            if (!raw || typeof raw !== 'object')
                throw Error('Invalid worksheet hyperlink.');
            const link = raw, point = typeof link.ref === 'string' ? parseRef(link.ref) : null, hasTarget = typeof link.target === 'string', hasLocation = typeof link.location === 'string';
            if (!point || point.row >= rows || point.col >= cols || !hasTarget && !hasLocation || refs.has(`${point.row}:${point.col}`))
                throw Error('Worksheet hyperlinks must use unique in-range cell references and one target type.');
            if (hasTarget) {
                const target = link.target;
                if (target.length > 2048 || /[\u0000-\u0020\\]/.test(target))
                    throw Error('Hyperlink target is malformed or too long.');
                let safe = false;
                try {
                    const url = new URL(target);
                    safe = (url.protocol === 'https:' || url.protocol === 'http:') && !!url.hostname || url.protocol === 'mailto:' && /^mailto:[^\s@]+@[^\s@]+$/i.test(target);
                }
                catch {
                    safe = false;
                }
                if (!safe)
                    throw Error('Hyperlink targets must use HTTP, HTTPS, or mailto.');
            }
            if (hasLocation && (link.location.length > 2048 || /[\u0000-\u001f]/.test(link.location)))
                throw Error('Internal hyperlink location is malformed or too long.');
            for (const key of ['tooltip', 'display'])
                if (link[key] !== undefined && (typeof link[key] !== 'string' || link[key].length > 255 || /[\u0000-\u001f]/.test(link[key])))
                    throw Error(`Hyperlink ${key} text is invalid.`);
            refs.add(`${point.row}:${point.col}`);
            t.hyperlinks.push({ ref: `${colToName(point.col)}${point.row + 1}`, ...(hasTarget ? { target: link.target } : {}), ...(hasLocation ? { location: link.location } : {}), ...(link.tooltip ? { tooltip: link.tooltip } : {}), ...(link.display ? { display: link.display } : {}) });
        }
    }
    if (obj.charts !== undefined) {
        if (!Array.isArray(obj.charts) || obj.charts.length > 24)
            throw Error('Too many charts.');
        const ids = new Set();
        for (const item of obj.charts) {
            const chart = validateChart(item, rows, cols);
            if (ids.has(chart.id))
                throw Error('Duplicate chart identifier.');
            ids.add(chart.id);
            t.charts.push(chart);
        }
    }
    if (obj.drawings !== undefined) {
        if (!Array.isArray(obj.drawings) || obj.drawings.length > 300)
            throw Error('A worksheet can contain at most 300 ink strokes.');
        const ids = new Set();
        let totalPoints = 0;
        for (const raw of obj.drawings) {
            if (!raw || typeof raw !== 'object')
                throw Error('Invalid ink stroke.');
            const stroke = raw;
            if (typeof stroke.id !== 'string' || !/^ink-[A-Za-z0-9_-]{1,64}$/.test(stroke.id) || ids.has(stroke.id) || !['pen', 'highlighter', 'line', 'arrow', 'rectangle', 'ellipse', 'text'].includes(stroke.tool) || typeof stroke.color !== 'string' || !/^#[0-9a-f]{6}$/i.test(stroke.color) || !Number.isFinite(stroke.width) || stroke.width < 1 || stroke.width > 40 || !Array.isArray(stroke.points) || stroke.points.length < (stroke.tool === 'text' ? 1 : 2) || stroke.points.length > (stroke.tool === 'pen' || stroke.tool === 'highlighter' ? 1200 : 2) || stroke.text !== undefined && (typeof stroke.text !== 'string' || stroke.text.length > 500) || stroke.tool === 'text' && typeof stroke.text !== 'string' || stroke.fill !== undefined && typeof stroke.fill !== 'boolean' || stroke.fill && !['rectangle', 'ellipse'].includes(stroke.tool))
                throw Error('Invalid drawing object settings.');
            ids.add(stroke.id);
            totalPoints += stroke.points.length;
            for (const point of stroke.points)
                if (!Array.isArray(point) || point.length !== 2 || !point.every(value => typeof value === 'number' && Number.isFinite(value) && value >= 0 && value <= 30000))
                    throw Error('Drawing coordinates must be finite worksheet coordinates.');
            t.drawings.push({ id: stroke.id, tool: stroke.tool, color: stroke.color.toLowerCase(), width: stroke.width, points: stroke.points.map(point => [point[0], point[1]]), ...(stroke.text !== undefined ? { text: stroke.text } : {}), ...(stroke.fill ? { fill: true } : {}) });
        }
        if (totalPoints > 75000)
            throw Error('The worksheet ink data exceeds the supported point limit.');
    }
    if (obj.tables !== undefined) {
        if (!Array.isArray(obj.tables) || obj.tables.length > 32)
            throw Error('Too many tables.');
        const names = new Set();
        for (const raw of obj.tables) {
            const table = validateTable(raw, rows, cols);
            const name = table.name.toLowerCase();
            if (names.has(name) || t.tables.some(other => intersectTables(tableRange(other.range, rows, cols), tableRange(table.range, rows, cols))))
                throw Error('Overlapping or duplicate spreadsheet table.');
            names.add(name);
            t.tables.push(table);
        }
    }
    if (obj.pivotTables !== undefined) {
        if (!Array.isArray(obj.pivotTables) || obj.pivotTables.length > 16)
            throw Error('A worksheet can contain at most 16 pivot summaries.');
        const ids = new Set();
        for (const raw of obj.pivotTables) {
            const pivot = validatePivotSummary(raw, rows, cols);
            if (ids.has(pivot.id))
                throw Error('Duplicate pivot summary identifier.');
            ids.add(pivot.id);
            t.pivotTables.push(pivot);
        }
    }
    if (obj.pivotSlicers !== undefined) {
        if (!Array.isArray(obj.pivotSlicers) || obj.pivotSlicers.length > 16)
            throw Error('A worksheet can contain at most 16 PivotTable slicers.');
        const ids = new Set(), bindings = new Set();
        for (const raw of obj.pivotSlicers) {
            if (!raw || typeof raw !== 'object')
                throw Error('Invalid PivotTable slicer.');
            const slicer = raw, pivot = t.pivotTables.find(item => item.id === slicer.pivotId), connected = [...new Set([slicer.pivotId, ...(slicer.pivotIds ?? [])])], eligible = (item) => { const fields = [...(item.rowFields ?? [item.rowField]), ...(item.columnFields ?? (item.columnField === undefined ? [] : [item.columnField]))], grouped = [...(item.dateGroups ?? []).map(group => group.field), ...(item.numberGroups ?? []).map(group => group.field), ...(item.manualGroups ?? []).map(group => group.field)]; return fields.includes(slicer.field) && !grouped.includes(slicer.field); };
            const priorFilters = slicer.priorFilters;
            if (priorFilters !== undefined && (!Array.isArray(priorFilters) || priorFilters.length > 16 || new Set(priorFilters.map(item => item?.pivotId)).size !== priorFilters.length || priorFilters.some(item => !item || !connected.includes(item.pivotId) || item.values !== null && (!Array.isArray(item.values) || item.values.length > 12000 || item.values.some(value => value !== null && (typeof value !== 'string' || value.length > 32767))))))
                throw Error('PivotTable slicer filter snapshots are invalid.');
            if (typeof slicer.id !== 'string' || !/^slicer-[A-Za-z0-9_-]{1,60}$/.test(slicer.id) || ids.has(slicer.id) || !pivot || !Number.isInteger(slicer.field) || !eligible(pivot) || connected.length > 16 || connected.some(id => !t.pivotTables.some(item => item.id === id && eligible(item) && (item.sourceSheet ?? '').toLocaleLowerCase() === (pivot.sourceSheet ?? '').toLocaleLowerCase() && item.sourceRange.replace(/\$/g, '').toUpperCase() === pivot.sourceRange.replace(/\$/g, '').toUpperCase())) || [slicer.x, slicer.y].some(value => !Number.isFinite(value) || value < 0 || value > 30000) || ![slicer.width, slicer.height].every(value => Number.isFinite(value) && value >= 180 && value <= 900) || typeof slicer.caption !== 'string' || slicer.caption.length > 100)
                throw Error('PivotTable slicers must connect compatible summaries that share a source and expose the same ungrouped field.');
            for (const pivotId of connected) {
                const binding = `${pivotId}:${slicer.field}`;
                if (bindings.has(binding))
                    throw Error('A PivotTable can have only one slicer for each field.');
                bindings.add(binding);
            }
            ids.add(slicer.id);
            t.pivotSlicers.push({ ...slicer, pivotIds: connected, ...(priorFilters ? { priorFilters: priorFilters.map(item => ({ pivotId: item.pivotId, values: item.values === null ? null : [...item.values] })) } : {}) });
        }
    }
    if (obj.queries !== undefined) {
        if (!Array.isArray(obj.queries) || obj.queries.length > 24)
            throw Error('A worksheet can contain at most 24 saved queries.');
        const ids = new Set();
        for (const raw of obj.queries) {
            const query = validateQuerySpec(raw, rows, cols);
            if (ids.has(query.id))
                throw Error('Duplicate query identifier.');
            ids.add(query.id);
            t.queries.push(query);
        }
    }
    return t;
}
export function dataValidationAt(ref, rules) {
    const p = parseRef(ref);
    if (!p)
        return undefined;
    return [...rules].reverse().find(rule => { const a = parseRef(rule.start), b = parseRef(rule.end); return p.row >= a.row && p.row <= b.row && p.col >= a.col && p.col <= b.col; });
}
export function dataValidationValues(rule, book) {
    if (rule.type !== 'list')
        return [];
    if (rule.values)
        return [...rule.values];
    if (!rule.sourceRange || !book)
        return [];
    const [from, to = ''] = rule.sourceRange.split(':'), a = parseRef(from ?? ''), b = parseRef(to || from || '');
    if (!a || !b)
        return [];
    const values = [], seen = new Set();
    for (let r = a.row; r <= b.row && values.length < 200; r++)
        for (let c = a.col; c <= b.col && values.length < 200; c++) {
            const value = display(book.getValue(`${colToName(c)}${r + 1}`));
            if (value && !seen.has(value)) {
                seen.add(value);
                values.push(value);
            }
        }
    return values;
}
export function dataValidationError(raw, rule, values = dataValidationValues(rule)) {
    if (rule.showErrorMessage === false)
        return null;
    if (raw.startsWith('='))
        return null;
    if (raw.trim() === '')
        return rule.allowBlank ? null : 'Blank values are not allowed for this cell.';
    if (rule.type === 'list') {
        const value = raw.startsWith("'") ? raw.slice(1) : raw;
        return values.includes(value) ? null : `Choose a value from the allowed list: ${values.join(', ')}.`;
    }
    if (rule.type === 'textLength') {
        const value = raw.startsWith("'") ? raw.slice(1) : raw, operator = rule.operator ?? (rule.min !== undefined && rule.max !== undefined ? 'between' : rule.min !== undefined ? 'greaterThanOrEqual' : 'lessThanOrEqual');
        return validationComparison(value.length, rule, operator, 'character count');
    }
    let value;
    if (rule.type === 'date') {
        const input = raw.startsWith("'") ? raw.slice(1) : raw, iso = /^\d{4}-\d{2}-\d{2}$/.test(input) ? input : null;
        if (iso) {
            const serial = excelDateSerial(iso);
            if (serial === null)
                return 'Enter a valid date.';
            value = serial;
        }
        else
            value = Number(input);
        if (!Number.isFinite(value))
            return 'Enter a valid date.';
    }
    else {
        value = Number(raw);
        if (raw.trim() === '' || !Number.isFinite(value))
            return 'Enter a numeric value allowed by this cell rule.';
    }
    if (rule.type === 'whole' && !Number.isInteger(value))
        return 'Enter a whole number.';
    const operator = rule.operator ?? (rule.min !== undefined && rule.max !== undefined ? 'between' : rule.min !== undefined ? 'greaterThanOrEqual' : 'lessThanOrEqual');
    return validationComparison(value, rule, operator, rule.type === 'date' ? 'date serial' : 'value');
}
function validationComparison(value, rule, operator, label) {
    const invalid = operator === 'between' ? value < rule.min || value > rule.max : operator === 'notBetween' ? value >= rule.min && value <= rule.max : operator === 'equalTo' ? value !== rule.min : operator === 'notEqualTo' ? value === rule.min : operator === 'greaterThan' ? value <= rule.min : operator === 'greaterThanOrEqual' ? value < rule.min : operator === 'lessThan' ? value >= rule.min : value > rule.min;
    if (invalid) {
        if (operator === 'between')
            return `Enter a ${label} from ${rule.min} to ${rule.max}.`;
        if (operator === 'notBetween')
            return `Enter a ${label} below ${rule.min} or above ${rule.max}.`;
        if (operator === 'equalTo')
            return `Enter exactly ${rule.min} for the ${label}.`;
        if (operator === 'notEqualTo')
            return `Enter a ${label} other than ${rule.min}.`;
        if (operator === 'greaterThan')
            return `Enter a ${label} greater than ${rule.min}.`;
        if (operator === 'greaterThanOrEqual')
            return `Enter a ${label} greater than or equal to ${rule.min}.`;
        if (operator === 'lessThan')
            return `Enter a ${label} less than ${rule.min}.`;
        return `Enter a ${label} less than or equal to ${rule.min}.`;
    }
    return null;
}
export function filterVisibleRows(book, spec) {
    const shown = new Set();
    if (!spec)
        return shown;
    for (let r = spec.startRow + 1; r <= spec.endRow; r++) {
        const ref = `${colToName(spec.column)}${r + 1}`;
        const value = display(book.getValue(ref));
        const legacy = spec.criteria?.length ? true : (!spec.query || value.toLowerCase().includes(spec.query.toLowerCase()));
        const combined = (spec.criteria ?? []).every(c => c.values.includes(display(book.getValue(`${colToName(c.column)}${r + 1}`))));
        if (legacy && combined)
            shown.add(r);
    }
    return shown;
}
/** Cell-is and Excel-style 2-color scale rules, computed against live numeric cell results.
 *  Cache min/max by rule during a render so virtualization stays responsive. */
export function ruleColor(ref, value, rules, book, cache = new Map()) {
    if (typeof value !== 'number')
        return null;
    const p = parseRef(ref);
    if (!p)
        return null;
    const rgb = (s) => [1, 3, 5].map(i => parseInt(s.slice(i, i + 2), 16));
    for (const rule of rules) {
        const a = parseRef(rule.start), b = parseRef(rule.end);
        if (p.row < Math.min(a.row, b.row) || p.row > Math.max(a.row, b.row) || p.col < Math.min(a.col, b.col) || p.col > Math.max(a.col, b.col))
            continue;
        if (rule.operator === 'scale' && book) {
            let extents = cache.get(rule);
            if (!extents) {
                let min = Infinity, max = -Infinity;
                for (let r = Math.min(a.row, b.row); r <= Math.max(a.row, b.row); r++)
                    for (let c = Math.min(a.col, b.col); c <= Math.max(a.col, b.col); c++) {
                        const res = book.getValue(`${colToName(c)}${r + 1}`);
                        if (!res.error && typeof res.value === 'number' && Number.isFinite(res.value)) {
                            min = Math.min(min, res.value);
                            max = Math.max(max, res.value);
                        }
                    }
                extents = [min, max];
                cache.set(rule, extents);
            }
            if (!Number.isFinite(extents[0]))
                return null;
            const ratio = extents[1] === extents[0] ? .5 : Math.max(0, Math.min(1, (value - extents[0]) / (extents[1] - extents[0])));
            const start = rgb(rule.startColor ?? '#ffffff'), end = rgb(rule.color), middle = rule.midColor ? rgb(rule.midColor) : null, localRatio = middle ? (ratio <= .5 ? ratio * 2 : (ratio - .5) * 2) : ratio, from = middle && ratio > .5 ? middle : start, to = middle && ratio <= .5 ? middle : end;
            return '#' + from.map((v, i) => Math.round(v + (to[i] - v) * localRatio).toString(16).padStart(2, '0')).join('');
        }
        if (rule.operator === 'gt' && value > rule.threshold || rule.operator === 'gte' && value >= rule.threshold || rule.operator === 'lt' && value < rule.threshold || rule.operator === 'lte' && value <= rule.threshold || rule.operator === 'eq' && value === rule.threshold || rule.operator === 'ne' && value !== rule.threshold || rule.operator === 'between' && value >= rule.threshold && value <= rule.threshold2 || rule.operator === 'notBetween' && (value < rule.threshold || value > rule.threshold2))
            return rule.color;
    }
    return null;
}
export function dataBarFill(ref, value, rules, book, cache = new Map()) {
    if (typeof value !== 'number' || !Number.isFinite(value))
        return null;
    const p = parseRef(ref);
    if (!p)
        return null;
    for (const rule of rules) {
        if (rule.operator !== 'dataBar')
            continue;
        const a = parseRef(rule.start), b = parseRef(rule.end);
        if (p.row < Math.min(a.row, b.row) || p.row > Math.max(a.row, b.row) || p.col < Math.min(a.col, b.col) || p.col > Math.max(a.col, b.col))
            continue;
        let extent = cache.get(rule);
        if (!extent) {
            let min = rule.minValue ?? Infinity, max = rule.maxValue ?? -Infinity;
            for (let r = Math.min(a.row, b.row); r <= Math.max(a.row, b.row); r++)
                for (let c = Math.min(a.col, b.col); c <= Math.max(a.col, b.col); c++) {
                    const x = book.getValue(`${colToName(c)}${r + 1}`);
                    if (!x.error && typeof x.value === 'number' && Number.isFinite(x.value)) {
                        if (rule.minValue === undefined)
                            min = Math.min(min, x.value);
                        if (rule.maxValue === undefined)
                            max = Math.max(max, x.value);
                    }
                }
            if (rule.minValue === undefined && min > 0)
                min = 0;
            if (rule.maxValue === undefined && max < 0)
                max = 0;
            extent = [min, max];
            cache.set(rule, extent);
        }
        const [min, max] = extent;
        if (!Number.isFinite(min) || !Number.isFinite(max) || min === max)
            return { background: '', hideValue: rule.showValue === false };
        const zero = max === min ? 0 : Math.max(0, Math.min(100, 100 * (0 - min) / (max - min))), scale = 100 / Math.max(1e-12, max - min), amount = Math.max(0, Math.min(100, value < 0 ? Math.abs(value) * scale : value * scale)), from = value < 0 ? Math.max(0, zero - amount) : zero, to = value < 0 ? zero : Math.min(100, zero + amount), color = value < 0 ? (rule.negativeColor ?? rule.color) : rule.color;
        const fill = rule.barStyle === 'gradient' ? (value < 0 ? `linear-gradient(to right, transparent ${from}%, ${color} ${zero}%, transparent ${zero}%)` : `linear-gradient(to right, transparent ${zero}%, ${color} ${to}%, transparent ${to}%)`) : `linear-gradient(to right, transparent ${from}%, ${color} ${from}%, ${color} ${to}%, transparent ${to}%)`;
        return { background: fill, hideValue: rule.showValue === false };
    }
    return null;
}
