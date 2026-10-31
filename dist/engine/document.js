import { Workbook, display } from './workbook.js';
import { filterVisibleRows, freshTools, validateTools } from './sheettools.js';
import { colToName, parseRef, rewriteSheetReferences } from './parser.js';
import { error, text } from './functions.js';
import { matchesPivotLabelFilter, pivotDateGroupLabel, pivotDateSerial, pivotNumberGroupLabel, pivotNumberGroupUsesIntegerLabels, pivotSourceFieldHeaders, resolvePivotNumberGroupRange } from './pivot.js';
const validTabColor = (value) => { if (value === undefined || value === '')
    return undefined; if (!/^#[0-9a-f]{6}$/i.test(value))
    throw Error('Worksheet tab color must be a six-digit hex color.'); return value.toLowerCase(); };
const validSheetView = (value) => { if (value === undefined)
    return undefined; if (!value || typeof value !== 'object')
    throw Error('Invalid worksheet view settings.'); const raw = value, clean = {}; if (raw.showGridLines !== undefined) {
    if (typeof raw.showGridLines !== 'boolean')
        throw Error('Invalid worksheet gridline setting.');
    clean.showGridLines = raw.showGridLines;
} for (const key of ['zoomScale', 'zoomScaleNormal'])
    if (raw[key] !== undefined) {
        const scale = raw[key];
        if (typeof scale !== 'number' || !Number.isInteger(scale) || scale < 10 || scale > 400)
            throw Error('Worksheet zoom must be a whole percentage from 10 to 400.');
        clean[key] = scale;
    } return clean; };
export class SpreadsheetDocument {
    items = [];
    filteredRowsCache = new Map();
    nextId = 1;
    changeQueue = [];
    notifyingChanges = false;
    pendingExternalInvalidations = new Map();
    workbookDefinedNames = [];
    activeId = '';
    constructor(initial, deferRecalculation = false) { this.addSheet('Sheet 1', initial ?? new Workbook(), deferRecalculation); }
    get sheets() { return this.items; }
    get active() { return this.items.find(x => x.id === this.activeId) ?? this.items[0]; }
    validateName(name, excludeId) {
        const n = name.trim();
        if (!n || n.length > 31 || /[\\/?*\[\]:]/.test(n))
            throw Error('Sheet name must be 1–31 characters and cannot contain \\ / ? * [ ] :.');
        if (this.items.some(s => s.id !== excludeId && s.name.toLowerCase() === n.toLowerCase()))
            throw Error('A sheet with that name already exists.');
        return n;
    }
    addSheet(name, book, deferRecalculation = false) {
        if (this.items.length >= 12)
            throw Error('This version supports a maximum of 12 sheets.');
        let title = name;
        if (!title) {
            let n = 1;
            while (this.items.some(x => x.name.toLowerCase() === `sheet ${n}`))
                n++;
            title = `Sheet ${n}`;
        }
        title = this.validateName(title);
        const item = { id: `s${this.nextId++}`, name: title, book: book ?? new Workbook(), tools: freshTools(), hidden: false };
        this.items.push(item);
        this.activeId = item.id;
        this.bindSheetContexts(deferRecalculation);
        return item;
    }
    bindSheetContexts(deferRecalculation = false) {
        for (const item of this.items)
            item.book.setSheetContext(item.name, (name, ref, visiting) => {
                const target = this.items.find(sheet => sheet.name.toLowerCase() === name.toLowerCase());
                return target ? target.book.getValue(ref, visiting) : { value: null, error: { code: '#REF!', message: `Worksheet '${name}' does not exist.`, position: 0, length: 1, suggestion: 'Check the worksheet name in the reference.' } };
            }, refs => this.notifyExternalChanges(item.name, refs), (name, ref, visiting) => {
                const target = this.items.find(sheet => sheet.name.toLowerCase() === name.toLowerCase());
                if (!target)
                    return error('#REF!', `Worksheet '${name}' does not exist.`);
                return target.book.getSpillValues(ref, visiting);
            }, this.items.map(sheet => sheet.name), (sheetName, includeManuallyHidden) => { const target = this.items.find(sheet => sheet.name.toLocaleLowerCase() === sheetName.toLocaleLowerCase()); if (!target)
                return []; const filtered = target.tools.filter, key = filtered ? `${target.id}:${target.book.valueRevision}:${JSON.stringify(filtered)}` : '', cached = key ? this.filteredRowsCache.get(key) : undefined; let filteredRows = cached ?? []; if (filtered && !cached) {
                const visible = filterVisibleRows(target.book, filtered), skipped = [];
                for (let row = filtered.startRow + 1; row <= filtered.endRow; row++)
                    if (!visible.has(row))
                        skipped.push(row);
                filteredRows = skipped;
                this.filteredRowsCache.set(key, filteredRows);
                if (this.filteredRowsCache.size > 24)
                    this.filteredRowsCache.delete(this.filteredRowsCache.keys().next().value);
            } return includeManuallyHidden ? [...new Set([...target.tools.hiddenRows, ...filteredRows])] : filteredRows; }, (sheetName, ref) => this.items.find(sheet => sheet.name.toLocaleLowerCase() === sheetName.toLocaleLowerCase())?.book.hasNestedAggregateFormula(ref) ?? false, (sheetName, ref) => this.items.find(sheet => sheet.name.toLocaleLowerCase() === sheetName.toLocaleLowerCase())?.book.getRaw(ref), deferRecalculation, sheetName => this.items.find(sheet => sheet.name.toLocaleLowerCase() === sheetName.toLocaleLowerCase())?.tools.hiddenColumns ?? []);
        for (const item of this.items)
            item.book.setPivotDataContext((pivotSheet, pivotRef, dataField, criteria, trackRange) => this.resolvePivotData(pivotSheet, pivotRef, dataField, criteria, trackRange), deferRecalculation);
        for (const item of this.items)
            item.book.setWorkbookDefinedNames(this.workbookDefinedNames, deferRecalculation);
    }
    getWorkbookDefinedNames() { return this.workbookDefinedNames.map(item => ({ ...item })); }
    setWorkbookDefinedNames(items, deferRecalculation = false) {
        const validator = new Workbook();
        validator.setWorkbookDefinedNames(items, true);
        const next = validator.getWorkbookDefinedNames();
        if (JSON.stringify(next) === JSON.stringify(this.workbookDefinedNames))
            return;
        this.workbookDefinedNames = next;
        for (const item of this.items)
            item.book.setWorkbookDefinedNames(next, deferRecalculation);
    }
    /** Return formula cells recalculated by cross-sheet edits so dependent tools can refresh. */
    consumeExternalInvalidations() {
        const changes = [...this.pendingExternalInvalidations.values()].map(change => ({ sheetName: change.sheetName, refs: [...change.refs] }));
        this.pendingExternalInvalidations.clear();
        return changes;
    }
    resolvePivotData(pivotSheet, pivotRef, dataField, criteria, trackRange) {
        const owner = this.items.find(sheet => sheet.name.toLocaleLowerCase() === pivotSheet.toLocaleLowerCase()), anchor = parseRef(pivotRef);
        if (!owner || !anchor)
            error('#REF!', `PivotTable reference '${pivotSheet}!${pivotRef}' is invalid.`);
        if (anchor.row >= owner.book.rows || anchor.col >= owner.book.cols)
            error('#REF!', `PivotTable reference '${pivotSheet}!${pivotRef}' is outside the worksheet.`);
        const spec = owner.tools.pivotTables.find(pivot => {
            if (!pivot.lastOutputRange)
                return false;
            const parts = pivot.lastOutputRange.replace(/\$/g, '').split(':'), first = parseRef(parts[0] ?? ''), last = parseRef(parts[1] ?? parts[0] ?? '');
            return !!first && !!last && anchor.row >= Math.min(first.row, last.row) && anchor.row <= Math.max(first.row, last.row) && anchor.col >= Math.min(first.col, last.col) && anchor.col <= Math.max(first.col, last.col);
        });
        if (!spec)
            error('#REF!', `Cell '${pivotSheet}!${pivotRef}' is not inside a generated PivotTable.`);
        // Watch the remaining grid from the PivotTable anchor so formulas also
        // refresh when a PivotTable grows after a newly introduced item appears.
        trackRange(owner.name, `${spec.destination}:${colToName(owner.book.cols - 1)}${owner.book.rows}`, spec.lastOutputRange);
        const sourceSheet = spec.sourceSheet ? this.items.find(sheet => sheet.name.toLocaleLowerCase() === spec.sourceSheet.toLocaleLowerCase()) : owner;
        if (!sourceSheet)
            error('#REF!', `PivotTable source worksheet '${spec.sourceSheet}' does not exist.`);
        const bounds = spec.sourceRange.split(':').map(ref => parseRef(ref.replace(/\$/g, '')));
        if (!bounds[0] || !bounds[1])
            error('#REF!', 'The PivotTable source range is invalid.');
        const first = bounds[0], last = bounds[1], width = last.col - first.col + 1, sourceHeaders = pivotSourceFieldHeaders(sourceSheet.book, spec.sourceRange), headerAt = (field) => sourceHeaders[field] ?? `Column ${field + 1}`;
        const rowFields = spec.rowFields ?? (spec.rowField === -1 ? [] : [spec.rowField]), columnFields = spec.columnFields ?? (spec.columnField === undefined ? [] : [spec.columnField]), measures = spec.measures ?? (spec.valueFields ?? [spec.valueField]).map(field => ({ field, aggregation: spec.aggregation }));
        const filters = spec.filters ?? (spec.filterField === undefined ? [] : [{ field: spec.filterField, value: spec.filterValue }]), criteriaByField = new Map();
        for (let index = 0; index < criteria.length; index++) {
            const name = text(criteria[index].field), matches = Array.from({ length: width }, (_, fieldIndex) => fieldIndex).filter(fieldIndex => {
                const dateGroup = spec.dateGroups?.find(group => group.field === fieldIndex), numberGroup = spec.numberGroups?.find(group => group.field === fieldIndex), manualGroup = spec.manualGroups?.find(group => group.field === fieldIndex), groupCaption = dateGroup ? `${{ years: 'Years', quarters: 'Quarters', months: 'Months', weeks: 'ISO Weeks', days: 'Days' }[dateGroup.unit]} (${headerAt(fieldIndex)})` : numberGroup || manualGroup ? `Groups (${headerAt(fieldIndex)})` : '';
                return headerAt(fieldIndex).toLocaleLowerCase() === name.toLocaleLowerCase() || groupCaption !== '' && groupCaption.toLocaleLowerCase() === name.toLocaleLowerCase();
            }), field = matches.length === 1 ? matches[0] : undefined;
            if (field === undefined || criteriaByField.has(field))
                return error('#REF!', `'${name}' is not a unique PivotTable field.`);
            if (!rowFields.includes(field) && !columnFields.includes(field) && !filters.some(filter => filter.field === field))
                error('#REF!', `'${name}' is not a field displayed or filtered by this PivotTable.`);
            criteriaByField.set(field, criteria[index].item);
        }
        const aggLabels = { countNumbers: 'Count Numbers', stdDev: 'StdDev', stdDevp: 'StdDevp', var: 'Var', varp: 'VarP' }, showLabels = { percentGrandTotal: '% of Grand Total', percentRowTotal: '% of Row Total', percentColumnTotal: '% of Column Total', percentParentRow: '% of Parent Row Total', percentParentColumn: '% of Parent Column Total', percentOf: '% Of', runningTotal: 'Running Total In', percentRunningTotal: '% Running Total In', differenceFrom: 'Difference From', percentDifferenceFrom: '% Difference From', rankAscending: 'Rank Smallest to Largest', rankDescending: 'Rank Largest to Smallest' };
        const measureCaption = (measure) => measure.caption?.trim() || `${aggLabels[measure.aggregation] ?? `${measure.aggregation[0].toUpperCase()}${measure.aggregation.slice(1)}`} of ${headerAt(measure.field)}${showLabels[measure.showAs ?? ''] ? ` (${showLabels[measure.showAs]})` : ''}`;
        const captionCounts = new Map(), measureCaptions = measures.map(measure => { const base = measureCaption(measure), count = (captionCounts.get(base.toLocaleLowerCase()) ?? 0) + 1; captionCounts.set(base.toLocaleLowerCase(), count); return count === 1 ? base : `${base} (${count})`; });
        const wanted = text(dataField), matches = measures.map((measure, index) => ({ measure, index, caption: measureCaptions[index] })).filter(item => item.caption.toLocaleLowerCase() === wanted.toLocaleLowerCase() || measures.filter(other => other.field === item.measure.field).length === 1 && headerAt(item.measure.field).toLocaleLowerCase() === wanted.toLocaleLowerCase());
        if (matches.length !== 1)
            error('#REF!', `'${wanted}' does not identify exactly one value field in this PivotTable.`);
        const selected = matches[0], measure = selected.measure;
        const dateGroups = new Map((spec.dateGroups ?? []).map(group => [group.field, group.unit])), numberGroups = new Map((spec.numberGroups ?? []).map(group => [group.field, group])), manualGroups = new Map((spec.manualGroups ?? []).map(group => [group.field, group.groups])), manualValueLabels = new Map(), groupRanges = new Map(), groupIntegerLabels = new Map();
        for (const [field, groups] of manualGroups) {
            const labels = new Map();
            for (const group of groups)
                for (const item of group.values)
                    labels.set((item ?? '').toLocaleLowerCase(), group.label);
            manualValueLabels.set(field, labels);
        }
        for (const [field, group] of numberGroups) {
            const values = [];
            for (let row = first.row + 1; row <= last.row; row++) {
                const value = sourceSheet.book.getValue(`${colToName(first.col + field)}${row + 1}`).value;
                if (typeof value === 'number' && Number.isFinite(value))
                    values.push(value);
            }
            const range = resolvePivotNumberGroupRange(values, group);
            groupRanges.set(field, range);
            groupIntegerLabels.set(field, pivotNumberGroupUsesIntegerLabels(values, group, range));
        }
        const fieldValue = (field, row) => {
            const ref = `${colToName(first.col + field)}${row + 1}`, cell = sourceSheet.book.getValue(ref), value = cell.value, label = display(cell), manualGroupLabel = manualValueLabels.get(field)?.get(label.toLocaleLowerCase());
            if (manualGroupLabel !== undefined)
                return { value, label: manualGroupLabel };
            if (!cell.error && (value === null || value === ''))
                return { value, label: '' };
            const dateUnit = dateGroups.get(field);
            if (dateUnit) {
                const label = pivotDateGroupLabel(value, dateUnit);
                if (label === null)
                    error('#REF!', `PivotTable grouping for '${headerAt(field)}' contains an invalid date.`);
                return { value, label: label };
            }
            const group = numberGroups.get(field);
            if (group) {
                const label = pivotNumberGroupLabel(value, group, groupRanges.get(field), groupIntegerLabels.get(field));
                if (label === null)
                    error('#REF!', `PivotTable grouping for '${headerAt(field)}' contains a nonnumeric value.`);
                return { value, label: label };
            }
            return { value, label: text(value) };
        };
        const equal = (left, right) => left.toLocaleLowerCase() === right.toLocaleLowerCase(), criteriaMatches = (field, value, label) => {
            if (!criteriaByField.has(field))
                return true;
            const wantedValue = criteriaByField.get(field), wantedText = text(wantedValue), blank = wantedValue === null || wantedText === '' || wantedText === '(blank)';
            return blank ? label === '' : equal(label, wantedText);
        };
        const allowed = (field, value, label, values) => values.length === 0 || values.some(item => item === null ? label === '' : equal(label, item));
        const dateSerial = pivotDateSerial;
        const activeFilters = filters, manualFilters = spec.itemFilters ?? [], captionFilters = spec.labelFilters ?? [], dateFilters = spec.dateFilters ?? [];
        const includedByReport = (row) => {
            let include = true;
            for (const filter of activeFilters) {
                const item = fieldValue(filter.field, row), choices = filter.values ?? (filter.value === undefined ? [] : [filter.value]);
                if (!allowed(filter.field, item.value, item.label, choices)) {
                    include = false;
                    break;
                }
            }
            if (!include)
                return false;
            for (const filter of manualFilters) {
                const item = fieldValue(filter.field, row);
                if (!allowed(filter.field, item.value, item.label, filter.values))
                    return false;
            }
            for (const filter of dateFilters) {
                const value = sourceSheet.book.getValue(`${colToName(first.col + filter.field)}${row + 1}`).value, serial = dateSerial(value);
                if (serial === null || Math.floor(serial) < filter.start || Math.floor(serial) > filter.end)
                    return false;
            }
            for (const filter of captionFilters) {
                const item = fieldValue(filter.field, row);
                if (!matchesPivotLabelFilter(item.label, filter.operator, filter.value, filter.value2))
                    return false;
            }
            return true;
        };
        const reportRows = [];
        for (let row = first.row + 1; row <= last.row; row++)
            if (includedByReport(row))
                reportRows.push(row);
        const valueFilter = spec.valueFilter, visibleFilteredItems = new Set();
        if (valueFilter) {
            const axis = rowFields.includes(valueFilter.field) ? rowFields : columnFields, depth = axis.indexOf(valueFilter.field), filterMeasureIndex = valueFilter.measureIndex ?? measures.findIndex(item => item.field === valueFilter.measureField), groups = new Map();
            if (depth < 0 || filterMeasureIndex < 0)
                error('#REF!', 'PivotTable value-filter metadata is invalid.');
            for (const row of reportRows) {
                const path = axis.map(field => fieldValue(field, row).label), label = path[depth] ?? '', parent = path.slice(0, depth), id = JSON.stringify([parent.map(value => value.toLocaleLowerCase()), label.toLocaleLowerCase()]);
                let group = groups.get(id);
                if (!group) {
                    group = { id, label, parent, values: [], raws: [] };
                    groups.set(id, group);
                }
                const sourceRef = `${colToName(first.col + measures[filterMeasureIndex].field)}${row + 1}`, cell = sourceSheet.book.getValue(sourceRef);
                group.values.push(cell.value);
                group.raws.push(sourceSheet.book.getRaw(sourceRef));
            }
            const measureValue = (items, rawValues, aggregation) => { let count = 0, countA = 0, sum = 0, product = 1, mean = 0, m2 = 0, min = Infinity, max = -Infinity; items.forEach((value, index) => { const numeric = typeof value === 'number' && Number.isFinite(value) ? value : null; if (value !== null || rawValues[index] !== '')
                countA++; if (numeric !== null) {
                count++;
                sum += numeric;
                const delta = numeric - mean;
                mean += delta / count;
                m2 += delta * (numeric - mean);
                product *= numeric;
                min = Math.min(min, numeric);
                max = Math.max(max, numeric);
            } }); switch (aggregation) {
                case 'count': return countA;
                case 'countNumbers': return count;
                case 'sum': return sum;
                case 'average': return count ? mean : 0;
                case 'min': return count ? min : 0;
                case 'max': return count ? max : 0;
                case 'product': return count ? product : 0;
                case 'stdDev': return count > 1 ? Math.sqrt(m2 / (count - 1)) : 0;
                case 'stdDevp': return count ? Math.sqrt(m2 / count) : 0;
                case 'var': return count > 1 ? m2 / (count - 1) : 0;
                default: return count ? m2 / count : 0;
            } };
            const buckets = new Map();
            for (const group of groups.values()) {
                const parent = JSON.stringify(group.parent.map(value => value.toLocaleLowerCase())), bucket = buckets.get(parent) ?? [];
                bucket.push({ id: group.id, label: group.label, value: measureValue(group.values, group.raws, measures[filterMeasureIndex].aggregation) });
                buckets.set(parent, bucket);
            }
            for (const bucket of buckets.values()) {
                bucket.sort((left, right) => { const compared = left.value === right.value ? left.label.localeCompare(right.label, undefined, { sensitivity: 'base', numeric: true }) : left.value < right.value ? -1 : 1; return valueFilter.direction === 'top' ? -compared : compared; });
                for (const entry of bucket.slice(0, valueFilter.count))
                    visibleFilteredItems.add(entry.id);
            }
        }
        const effectiveRows = reportRows.filter(row => { if (!valueFilter)
            return true; const axis = rowFields.includes(valueFilter.field) ? rowFields : columnFields, depth = axis.indexOf(valueFilter.field), path = axis.map(field => fieldValue(field, row).label), id = JSON.stringify([path.slice(0, depth).map(value => value.toLocaleLowerCase()), (path[depth] ?? '').toLocaleLowerCase()]); return visibleFilteredItems.has(id); });
        const filteredRows = (omitFields) => effectiveRows.filter(row => { for (const [field, item] of criteriaByField) {
            if (omitFields.has(field))
                continue;
            const current = fieldValue(field, row);
            if (!criteriaMatches(field, current.value, current.label))
                return false;
        } return true; });
        const aggregateRows = (rows, target) => { let count = 0, countA = 0, sum = 0, product = 1, mean = 0, m2 = 0, min = Infinity, max = -Infinity; for (const row of rows) {
            const sourceRef = `${colToName(first.col + target.field)}${row + 1}`, cell = sourceSheet.book.getValue(sourceRef), raw = sourceSheet.book.getRaw(sourceRef), numeric = typeof cell.value === 'number' && Number.isFinite(cell.value) ? cell.value : null;
            if (cell.error || raw !== '')
                countA++;
            if (numeric !== null) {
                count++;
                sum += numeric;
                const delta = numeric - mean;
                mean += delta / count;
                m2 += delta * (numeric - mean);
                product *= numeric;
                min = Math.min(min, numeric);
                max = Math.max(max, numeric);
            }
        } switch (target.aggregation) {
            case 'count': return countA;
            case 'countNumbers': return count;
            case 'sum': return sum;
            case 'average': return count ? mean : 0;
            case 'min': return count ? min : 0;
            case 'max': return count ? max : 0;
            case 'product': return count ? product : 0;
            case 'stdDev': return count > 1 ? Math.sqrt(m2 / (count - 1)) : 0;
            case 'stdDevp': return count ? Math.sqrt(m2 / count) : 0;
            case 'var': return count > 1 ? m2 / (count - 1) : 0;
            case 'varp': return count ? m2 / count : 0;
            default: return error('#VALUE!', 'Unsupported PivotTable aggregation.');
        } };
        const selectedValue = aggregateRows(filteredRows(new Set()), measure), showAs = measure.showAs ?? 'normal';
        if (showAs === 'normal')
            return selectedValue;
        const rowFieldSet = new Set(rowFields), columnFieldSet = new Set(columnFields), rowTotal = () => aggregateRows(filteredRows(columnFieldSet), measure), columnTotal = () => aggregateRows(filteredRows(rowFieldSet), measure), grandTotal = () => aggregateRows(filteredRows(new Set([...rowFields, ...columnFields])), measure), ratio = (numerator, denominator) => denominator ? numerator / denominator : 0;
        if (showAs === 'percentGrandTotal')
            return ratio(selectedValue, grandTotal());
        if (showAs === 'percentRowTotal')
            return ratio(selectedValue, rowTotal());
        if (showAs === 'percentColumnTotal')
            return ratio(selectedValue, columnTotal());
        if (showAs === 'percentParentRow' || showAs === 'percentParentColumn') {
            const axis = showAs === 'percentParentRow' ? rowFields : columnFields, deepest = [...axis].reverse().find(field => criteriaByField.has(field));
            if (deepest === undefined)
                return selectedValue;
            return ratio(selectedValue, aggregateRows(filteredRows(new Set([deepest])), measure));
        }
        if (showAs === 'index')
            return ratio(selectedValue * grandTotal(), rowTotal() * columnTotal());
        if (['percentOf', 'runningTotal', 'percentRunningTotal', 'differenceFrom', 'percentDifferenceFrom', 'rankAscending', 'rankDescending'].includes(showAs)) {
            const baseField = measure.baseField, axis = rowFields.includes(baseField ?? -1) ? rowFields : columnFields, depth = axis.indexOf(baseField ?? -1);
            if (baseField === undefined || depth < 0)
                return error('#REF!', `GETPIVOTDATA needs a valid Base field for '${showAs}'.`);
            // The Base field's siblings form the current subtotal scope. Build one
            // aggregate per visible item after applying the pivot's saved filters.
            const scopeRows = filteredRows(new Set([baseField])), items = new Map();
            for (const row of scopeRows) {
                const item = fieldValue(baseField, row), key = item.label.toLocaleLowerCase(), current = items.get(key);
                if (current)
                    current.rows.push(row);
                else
                    items.set(key, { label: item.label, rows: [row], sourceValue: item.value, order: items.size });
            }
            const collator = new Intl.Collator(undefined, { numeric: true, sensitivity: 'base' }), sortOrder = spec.fieldSortOrders?.[String(baseField)] ?? spec.sortOrder ?? 'ascending', valueSort = (spec.valueSorts ?? (spec.valueSort ? [spec.valueSort] : [])).find(sort => sort.field === baseField), sortMeasure = valueSort ? measures[valueSort.measureIndex ?? measures.findIndex(item => item.field === valueSort.measureField)] : undefined;
            const entries = [...items.values()].map(item => ({ ...item, value: aggregateRows(item.rows, measure), sortValue: sortMeasure ? aggregateRows(item.rows, sortMeasure) : 0 }));
            entries.sort((left, right) => {
                if (sortOrder === 'manual')
                    return left.order - right.order;
                let compared = 0;
                if (valueSort) {
                    compared = left.sortValue === right.sortValue ? collator.compare(left.label, right.label) : left.sortValue < right.sortValue ? -1 : 1;
                    return valueSort.order === 'descending' ? -compared : compared;
                }
                if (sortOrder === 'valueAscending' || sortOrder === 'valueDescending') {
                    compared = left.sortValue === right.sortValue ? collator.compare(left.label, right.label) : left.sortValue < right.sortValue ? -1 : 1;
                    return sortOrder === 'valueDescending' ? -compared : compared;
                }
                const a = left.sourceValue, b = right.sourceValue;
                if (a === b)
                    compared = 0;
                else if (a === null || a === '')
                    compared = -1;
                else if (b === null || b === '')
                    compared = 1;
                else if (typeof a === 'number' && typeof b === 'number')
                    compared = a < b ? -1 : 1;
                else if (typeof a === 'boolean' && typeof b === 'boolean')
                    compared = a ? 1 : -1;
                else
                    compared = collator.compare(String(a), String(b));
                if (numberGroups.has(baseField)) {
                    const numberCompare = (label) => label.startsWith('<') ? -Infinity : label.startsWith('>') ? Infinity : Number(/^(-?(?:\d+\.?\d*|\.\d+))(?:-|–<)/.exec(label)?.[1] ?? NaN);
                    compared = numberCompare(left.label) < numberCompare(right.label) ? -1 : numberCompare(left.label) > numberCompare(right.label) ? 1 : 0;
                }
                if (dateGroups.get(baseField) === 'years')
                    compared = Number(left.label) - Number(right.label);
                else if (dateGroups.get(baseField) === 'quarters')
                    compared = Number(left.label.slice(1)) - Number(right.label.slice(1));
                else if (dateGroups.get(baseField) === 'days') {
                    const aDate = Date.parse(`${left.label}T00:00:00Z`), bDate = Date.parse(`${right.label}T00:00:00Z`);
                    compared = aDate < bDate ? -1 : aDate > bDate ? 1 : 0;
                }
                else if (dateGroups.get(baseField) === 'months')
                    compared = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'].indexOf(left.label) - ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'].indexOf(right.label);
                return sortOrder === 'descending' ? -compared : compared;
            });
            const requested = criteriaByField.get(baseField), requestedLabel = requested === undefined ? undefined : requested === null || text(requested) === '' || text(requested) === '(blank)' ? '' : text(requested), currentIndex = requestedLabel === undefined ? -1 : entries.findIndex(entry => entry.label.toLocaleLowerCase() === requestedLabel.toLocaleLowerCase());
            if (currentIndex < 0)
                return selectedValue;
            const current = entries[currentIndex], customBaseItem = measure.baseItem && typeof measure.baseItem === 'object' ? measure.baseItem.label : undefined, baseIndex = measure.baseItem === 'next' ? currentIndex + 1 : measure.baseItem === 'previous' || measure.baseItem === undefined ? currentIndex - 1 : entries.findIndex(entry => entry.label.toLocaleLowerCase() === customBaseItem?.toLocaleLowerCase()), base = entries[baseIndex];
            if (showAs === 'runningTotal' || showAs === 'percentRunningTotal') {
                const cumulative = entries.slice(0, currentIndex + 1).reduce((total, entry) => total + entry.value, 0);
                return showAs === 'runningTotal' ? cumulative : ratio(cumulative, entries.reduce((total, entry) => total + entry.value, 0));
            }
            if (showAs === 'rankAscending' || showAs === 'rankDescending') {
                const ordered = [...entries].sort((left, right) => showAs === 'rankAscending' ? left.value - right.value : right.value - left.value);
                let rank = 1;
                for (let index = 0; index < ordered.length; index++) {
                    if (index > 0 && ordered[index].value !== ordered[index - 1].value)
                        rank = index + 1;
                    if (ordered[index].label.toLocaleLowerCase() === current.label.toLocaleLowerCase())
                        return rank;
                }
                return error('#REF!', 'Unable to rank the selected PivotTable item.');
            }
            if (!base)
                return '';
            const difference = current.value - base.value;
            if (showAs === 'differenceFrom')
                return difference;
            if (showAs === 'percentDifferenceFrom')
                return ratio(difference, base.value);
            if (showAs === 'percentOf')
                return ratio(current.value, base.value);
        }
        return error('#REF!', `GETPIVOTDATA does not yet calculate the '${showAs}' display mode.`);
    }
    notifyExternalChanges(sheet, refs) {
        this.changeQueue.push({ sheet, refs: [...new Set(refs)] });
        if (this.notifyingChanges)
            return;
        this.notifyingChanges = true;
        const seen = new Set();
        try {
            while (this.changeQueue.length) {
                const change = this.changeQueue.shift(), signature = `${change.sheet.toLowerCase()}|${change.refs.slice().sort().join(',')}`;
                if (seen.has(signature))
                    continue;
                seen.add(signature);
                for (const target of this.items) {
                    if (target.name.toLowerCase() === change.sheet.toLowerCase())
                        continue;
                    const affected = target.book.invalidateExternal(change.sheet, change.refs);
                    if (affected.length) {
                        const pending = this.pendingExternalInvalidations.get(target.id) ?? { sheetName: target.name, refs: new Set() };
                        for (const ref of affected)
                            pending.refs.add(ref);
                        this.pendingExternalInvalidations.set(target.id, pending);
                        this.changeQueue.push({ sheet: target.name, refs: affected });
                    }
                }
            }
        }
        finally {
            this.changeQueue = [];
            this.notifyingChanges = false;
        }
    }
    selectSheet(id) { const sheet = this.items.find(x => x.id === id); if (!sheet)
        throw Error('Sheet not found.'); if (sheet.hidden)
        throw Error('Unhide this worksheet before selecting it.'); this.activeId = id; }
    hideSheet(id = this.activeId) { const sheet = this.items.find(x => x.id === id); if (!sheet)
        throw Error('Sheet not found.'); if (sheet.hidden)
        return; if (this.items.filter(x => !x.hidden).length <= 1)
        throw Error('A workbook must keep at least one worksheet visible.'); sheet.hidden = true; if (this.activeId === id) {
        const index = this.items.indexOf(sheet);
        this.activeId = this.items.slice(index + 1).find(x => !x.hidden)?.id ?? this.items.slice(0, index).reverse().find(x => !x.hidden).id;
    } }
    unhideSheet(id) { const sheet = this.items.find(x => x.id === id); if (!sheet)
        throw Error('Sheet not found.'); sheet.hidden = false; }
    renameSheet(id, name) { const found = this.items.find(x => x.id === id); if (!found)
        throw Error('Sheet not found.'); const previous = found.name, next = this.validateName(name, id); if (previous === next)
        return; found.name = next; for (const item of this.items)
        for (const pivot of item.tools.pivotTables)
            if (pivot.sourceSheet?.toLocaleLowerCase() === previous.toLocaleLowerCase())
                pivot.sourceSheet = next; this.bindSheetContexts(); this.setWorkbookDefinedNames(this.workbookDefinedNames.map(item => ({ ...item, ref: rewriteSheetReferences(item.ref, previous, next) }))); for (const item of this.items) {
        const names = item.book.getDefinedNames();
        item.book.setDefinedNames(names.map(named => ({ ...named, ref: rewriteSheetReferences(named.ref, previous, next) })), false);
        const entries = item.book.allEntries(), edits = entries.map(cell => ({ ref: cell.ref, raw: rewriteSheetReferences(cell.raw, previous, next) })).filter((cell, index) => cell.raw !== entries[index]?.raw);
        if (edits.length)
            item.book.applyEdits(edits);
    } }
    setSheetTabColor(id, color) { const sheet = this.items.find(x => x.id === id); if (!sheet)
        throw Error('Sheet not found.'); sheet.tabColor = validTabColor(color); }
    reorderSheet(id, index) { const from = this.items.findIndex(x => x.id === id); if (from < 0)
        throw Error('Sheet not found.'); if (!Number.isInteger(index) || index < 0 || index >= this.items.length)
        throw Error('Worksheet position is out of range.'); const [sheet] = this.items.splice(from, 1); this.items.splice(index, 0, sheet); if (from !== index)
        this.bindSheetContexts(); }
    duplicateSheet(id = this.activeId) {
        const original = this.items.find(x => x.id === id);
        if (!original)
            throw Error('Sheet not found.');
        let name = `${original.name} copy`, suffix = 2;
        while (this.items.some(x => x.name.toLowerCase() === name.toLowerCase()))
            name = `${original.name.slice(0, 23)} ${suffix++}`;
        const duplicate = this.addSheet(name, Workbook.fromJSON(original.book.toJSON()));
        duplicate.tools = validateTools(JSON.parse(JSON.stringify(original.tools)), duplicate.book.rows, duplicate.book.cols);
        duplicate.book.setTableContext(duplicate.tools.tables);
        duplicate.tabColor = original.tabColor;
        duplicate.viewOptions = original.viewOptions ? { ...original.viewOptions } : undefined;
        this.bindSheetContexts();
        duplicate.book.recalculateAggregatesForRowVisibilityChange();
        return duplicate;
    }
    deleteSheet(id = this.activeId) { const idx = this.items.findIndex(x => x.id === id); if (idx === -1)
        throw Error('Sheet not found.'); if (this.items.length === 1)
        throw Error('At least one sheet must remain.'); if (!this.items[idx].hidden && this.items.filter(x => !x.hidden).length <= 1)
        throw Error('A workbook must keep at least one worksheet visible.'); const removed = this.items[idx].name; this.items.splice(idx, 1); this.bindSheetContexts(); this.setWorkbookDefinedNames(this.workbookDefinedNames.map(item => ({ ...item, ref: rewriteSheetReferences(item.ref, removed, null) }))); for (const item of this.items) {
        const names = item.book.getDefinedNames();
        item.book.setDefinedNames(names.map(named => ({ ...named, ref: rewriteSheetReferences(named.ref, removed, null) })), false);
        const removedPivotIds = new Set(item.tools.pivotTables.filter(pivot => pivot.sourceSheet?.toLocaleLowerCase() === removed.toLocaleLowerCase()).map(pivot => pivot.id));
        if (removedPivotIds.size) {
            item.tools.pivotTables = item.tools.pivotTables.filter(pivot => !removedPivotIds.has(pivot.id));
            item.tools.pivotSlicers = item.tools.pivotSlicers.flatMap(slicer => { const connected = [slicer.pivotId, ...(slicer.pivotIds ?? [])], remaining = connected.filter(pivotId => !removedPivotIds.has(pivotId)); if (remaining.length === connected.length)
                return [slicer]; if (!remaining.length)
                return []; return [{ ...slicer, pivotId: remaining[0], pivotIds: remaining.slice(1), ...(slicer.priorFilters ? { priorFilters: slicer.priorFilters.filter(item => remaining.includes(item.pivotId)) } : {}) }]; });
            for (const chart of item.tools.charts)
                if (chart.pivotId && removedPivotIds.has(chart.pivotId))
                    chart.pivotId = undefined;
        }
        const entries = item.book.allEntries(), edits = entries.map(cell => ({ ref: cell.ref, raw: rewriteSheetReferences(cell.raw, removed, null) })).filter((cell, index) => cell.raw !== entries[index]?.raw);
        if (edits.length)
            item.book.applyEdits(edits);
    } if (this.activeId === id)
        this.activeId = this.items.slice(Math.min(idx, this.items.length)).find(x => !x.hidden)?.id ?? this.items.find(x => !x.hidden).id; }
    toJSON() { return { format: 'gridforge', version: 5, activeId: this.activeId, workbookDefinedNames: this.getWorkbookDefinedNames(), sheets: this.items.map(x => ({ id: x.id, name: x.name, workbook: x.book.toJSON(), tools: x.tools, hidden: x.hidden, ...(x.tabColor ? { tabColor: x.tabColor } : {}), ...(x.viewOptions ? { viewOptions: x.viewOptions } : {}) })) }; }
    static fromJSON(raw, options = {}) {
        if (!raw || typeof raw !== 'object')
            throw Error('Invalid document JSON.');
        const data = raw;
        if (!Array.isArray(data.sheets)) {
            return new SpreadsheetDocument(Workbook.fromJSON(raw, options), options.deferCalculation === true);
        }
        if (data.format !== 'gridforge' || ![2, 3, 4, 5].includes(data.version) || !data.sheets.length || data.sheets.length > 12)
            throw Error('Invalid GridForge document.');
        const doc = new SpreadsheetDocument();
        doc.items = [];
        doc.nextId = 1;
        for (const s of data.sheets) {
            if (!s || typeof s !== 'object')
                throw Error('Invalid sheet.');
            const x = s;
            if (typeof x.name !== 'string' || (x.hidden !== undefined && typeof x.hidden !== 'boolean') || (x.tabColor !== undefined && typeof x.tabColor !== 'string'))
                throw Error('Invalid sheet name, visibility or tab color.');
            const name = doc.validateName(x.name);
            const id = `s${doc.nextId++}`;
            const book = Workbook.fromJSON(x.workbook, options);
            const tools = validateTools(x.tools, book.rows, book.cols);
            if (tools.tables.length)
                book.setTableContext(tools.tables);
            const viewOptions = validSheetView(x.viewOptions);
            doc.items.push({ id, name, book, tools, hidden: x.hidden === true, ...(x.tabColor ? { tabColor: validTabColor(x.tabColor) } : {}), ...(viewOptions ? { viewOptions } : {}) });
        }
        if (doc.items.every(sheet => sheet.hidden))
            throw Error('A workbook must contain at least one visible worksheet.');
        doc.bindSheetContexts(options.deferCalculation === true);
        if (data.workbookDefinedNames !== undefined)
            doc.setWorkbookDefinedNames(data.workbookDefinedNames, options.deferCalculation === true);
        if (!options.deferCalculation)
            for (const sheet of doc.items)
                sheet.book.recalculateAggregatesForRowVisibilityChange();
        if (typeof data.activeId === 'string') {
            const i = data.sheets.findIndex((x) => !!x && typeof x === 'object' && x.id === data.activeId);
            const requested = doc.items[Math.max(0, i)];
            doc.activeId = requested.hidden ? doc.items.find(sheet => !sheet.hidden).id : requested.id;
        }
        else
            doc.activeId = doc.items.find(sheet => !sheet.hidden).id;
        return doc;
    }
}
