/** Native editable OOXML PivotTable and PivotChart export for GridForge's bounded model. */
import { display } from './workbook.js';
import { colToName, parseRef } from './parser.js';
import { buildPivotSummary, matchesPivotLabelFilter, pivotDateGroupIsoBoundary, pivotDateGroupLabel, pivotDateSerial, pivotNumberGroupLabel, pivotNumberGroupLabels, pivotNumberGroupSortValue, pivotNumberGroupUsesIntegerLabels, pivotSourceFieldHeaders, resolvePivotNumberGroupRange } from './pivot.js';
const ns = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const relNs = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const pkgNs = 'http://schemas.openxmlformats.org/package/2006/relationships';
const esc = (s) => s.replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&apos;' }[c]));
const normalized = (label) => label.toLocaleLowerCase();
function cellValue(book, col, row) {
    const cell = book.getValue(`${colToName(col)}${row + 1}`), label = display(cell);
    if (cell.error)
        return { key: normalized(label), label, kind: 'e', value: cell.error.code };
    if (cell.value === null || cell.value === '')
        return { key: '', label: '', kind: 'm', value: '' };
    if (typeof cell.value === 'number')
        return { key: normalized(label), label, kind: 'n', value: String(cell.value) };
    if (typeof cell.value === 'boolean')
        return { key: normalized(label), label, kind: 'b', value: cell.value ? '1' : '0' };
    return { key: normalized(label), label, kind: 's', value: String(cell.value) };
}
const valueXml = (value) => value.kind === 'm' ? '<m/>' : `<${value.kind} v="${esc(value.value)}"/>`;
const sharedItemXml = (value) => valueXml(value);
const aggregate = (value) => value === 'countNumbers' ? 'countNums' : value === 'stdDevp' ? 'stdDevp' : value;
/** Generates the coordinated cache, record, definition and relationship parts for one local-source PivotTable. */
export function nativePivotParts(book, sheetName, spec, cacheId, tableName, outputRange, destinationBook = book, pivotChartIndexes = [], expansionState = {}) {
    const [first, last] = spec.sourceRange.split(':'), a = parseRef(first), b = parseRef(last), width = b.col - a.col + 1, sourceRows = Array.from({ length: b.row - a.row }, (_, i) => Array.from({ length: width }, (_, j) => cellValue(book, a.col + j, a.row + i + 1)));
    const uniqueHeaders = pivotSourceFieldHeaders(book, spec.sourceRange);
    const rowFields = spec.rowFields ?? [spec.rowField], labelFilters = spec.labelFilters ?? [], valueFields = spec.measures?.map(measure => measure.field) ?? spec.valueFields ?? [spec.valueField], columnFields = spec.columnFields ?? (spec.columnField === undefined ? [] : [spec.columnField]), filters = spec.filters ?? (spec.filterField === undefined ? [] : [{ field: spec.filterField, value: spec.filterValue }]), itemFilters = spec.itemFilters ?? [], dateGroups = spec.dateGroups ?? [], numberGroups = spec.numberGroups ?? [], manualGroups = spec.manualGroups ?? [], allGroups = [...dateGroups, ...numberGroups, ...manualGroups], groupIndexByField = new Map(allGroups.map((group, index) => [group.field, width + index])), rowAxisFields = rowFields.map(field => groupIndexByField.get(field) ?? field), columnAxisFields = columnFields.map(field => groupIndexByField.get(field) ?? field), cacheWidth = width + allGroups.length, groupHeaders = [...dateGroups.map(group => `${{ years: 'Years', quarters: 'Quarters', months: 'Months', weeks: 'ISO Weeks', days: 'Days' }[group.unit]} (${uniqueHeaders[group.field]})`), ...numberGroups.map(group => `Groups (${uniqueHeaders[group.field]})`), ...manualGroups.map(group => `Groups (${uniqueHeaders[group.field]})`)], numberGroupRanges = new Map(numberGroups.map(group => { const source = sourceRows.map(row => row[group.field]), values = source.flatMap(value => value.kind === 'n' ? [Number(value.value)] : []); if (source.some(value => value.kind !== 'm' && value.kind !== 'n'))
        throw Error(`Native numeric grouping requires numeric values in “${uniqueHeaders[group.field]}”.`); return [group.field, resolvePivotNumberGroupRange(values, group)]; })), numberGroupIntegerLabels = new Map(numberGroups.map(group => [group.field, pivotNumberGroupUsesIntegerLabels(sourceRows.map(row => row[group.field]).flatMap(value => value.kind === 'n' ? [Number(value.value)] : []), group, numberGroupRanges.get(group.field))])), manualGroupMaps = new Map(manualGroups.map(group => [group.field, new Map(group.groups.flatMap(item => item.values.map(value => [(value ?? '').toLocaleLowerCase(), item.label])))])), recordRows = sourceRows.map(source => [...source, ...dateGroups.map(group => { const value = source[group.field]; if (value.kind === 'm')
            return { key: '', label: '', kind: 'm', value: '' }; const raw = value.kind === 'n' ? Number(value.value) : value.kind === 's' ? value.value : null, label = pivotDateGroupLabel(raw, group.unit); if (label === null)
            throw Error(`Native date grouping requires valid date values in “${uniqueHeaders[group.field]}”.`); return { key: normalized(label), label, kind: 's', value: label }; }), ...numberGroups.map(group => { const value = source[group.field]; if (value.kind === 'm')
            return { key: '', label: '', kind: 'm', value: '' }; const range = numberGroupRanges.get(group.field), label = pivotNumberGroupLabel(value.kind === 'n' ? Number(value.value) : null, group, range, numberGroupIntegerLabels.get(group.field)); if (label === null)
            throw Error(`Native numeric grouping requires numeric values in “${uniqueHeaders[group.field]}”.`); return { key: normalized(label), label, kind: 's', value: label }; }), ...manualGroups.map(group => { const value = source[group.field], label = manualGroupMaps.get(group.field)?.get(value.label.toLocaleLowerCase()) ?? value.label; if (value.kind === 'm' && label === '')
            return { key: '', label: '', kind: 'm', value: '' }; return { key: normalized(label), label, kind: 's', value: label }; })]), dimensions = [...new Set([...rowAxisFields, ...columnAxisFields, ...allGroups.map(group => group.field), ...filters.map(filter => filter.field)])], shared = new Map();
    const exportMeasures = spec.measures ?? valueFields.map(field => ({ field, aggregation: spec.aggregation })), valueSorts = spec.valueSorts ?? (spec.valueSort ? [spec.valueSort] : []), valueSortByField = new Map(valueSorts.map(sort => [sort.field, sort])), valueSortChildOrders = new Map(), valueSortSummary = valueSorts.length ? buildPivotSummary(destinationBook, spec, book) : undefined;
    for (const valueSort of valueSorts) {
        const isRow = rowFields.includes(valueSort.field), fields = isRow ? rowFields : columnFields, depth = fields.indexOf(valueSort.field), paths = valueSortSummary?.[isRow ? 'rowPaths' : 'columnPaths'] ?? [], fieldOrders = new Map();
        if (depth < 0)
            continue;
        for (const path of paths) {
            if (path.length <= depth)
                continue;
            const parentKey = JSON.stringify(path.slice(0, depth).map(label => normalized(label === '\u0000' ? '' : label))), itemKey = normalized(path[depth] === '\u0000' ? '' : path[depth]), order = fieldOrders.get(parentKey) ?? new Map();
            if (!order.has(itemKey))
                order.set(itemKey, order.size);
            fieldOrders.set(parentKey, order);
        }
        valueSortChildOrders.set(valueSort.field, fieldOrders);
    }
    for (const field of dimensions) {
        const values = [], indices = new Set();
        for (const row of recordRows) {
            const value = row[field];
            if (!indices.has(value.key)) {
                indices.add(value.key);
                values.push(value);
            }
        }
        shared.set(field, values);
    }
    for (let index = 0; index < numberGroups.length; index++) {
        const group = numberGroups[index], cacheIndex = width + dateGroups.length + index, values = shared.get(cacheIndex) ?? [], indices = new Set(values.map(value => value.key)), range = numberGroupRanges.get(group.field);
        for (const label of pivotNumberGroupLabels(group, range, numberGroupIntegerLabels.get(group.field))) {
            const key = normalized(label);
            if (indices.has(key))
                continue;
            indices.add(key);
            values.push({ key, label, kind: 's', value: label });
        }
        shared.set(cacheIndex, values);
    }
    const sharedIndexByKey = new Map(), fieldItemOrders = new Map(), fieldItemOrdinals = new Map(), itemCollator = new Intl.Collator(undefined, { numeric: true, sensitivity: 'base' }), dateGroupByIndex = new Map(dateGroups.map((group, index) => [width + index, group])), numberGroupByIndex = new Map(numberGroups.map((group, index) => [width + dateGroups.length + index, group])), manualGroupByIndex = new Map(manualGroups.map((group, index) => [width + dateGroups.length + numberGroups.length + index, group]));
    const sourceAxisField = (field) => dateGroupByIndex.get(field)?.field ?? numberGroupByIndex.get(field)?.field ?? manualGroupByIndex.get(field)?.field ?? field, fieldSortOrder = (field) => { const order = spec.fieldSortOrders?.[String(sourceAxisField(field))] ?? spec.sortOrder ?? 'ascending'; return order === 'valueAscending' ? 'ascending' : order === 'valueDescending' ? 'descending' : order; }, manualItemRanks = new Map((spec.manualItemOrders ?? []).map(order => [order.field, new Map(order.items.map((item, index) => [normalized(item ?? ''), index]))])), manualRank = (field, label) => manualItemRanks.get(sourceAxisField(field))?.get(normalized(label)), compareDateGroup = (field, left, right) => { const unit = dateGroupByIndex.get(field)?.unit; if (unit === 'months')
        return ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'].indexOf(left) - ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'].indexOf(right); if (unit === 'quarters')
        return Number(left.slice(1)) - Number(right.slice(1)); if (numberGroupByIndex.has(field)) {
        const rank = (label) => pivotNumberGroupSortValue(label);
        return rank(left) - rank(right);
    } return itemCollator.compare(left, right); };
    for (const field of dimensions) {
        const values = shared.get(field), sharedIndices = new Map(values.map((value, index) => [value.key, index]));
        sharedIndexByKey.set(field, sharedIndices);
        const order = [...values.keys()], sortOrder = fieldSortOrder(field);
        if ((rowAxisFields.includes(field) || columnAxisFields.includes(field)) && sortOrder === 'manual' && manualItemRanks.has(sourceAxisField(field)))
            order.sort((left, right) => { const leftRank = manualRank(field, values[left].label), rightRank = manualRank(field, values[right].label); return leftRank === undefined ? rightRank === undefined ? 0 : 1 : rightRank === undefined ? -1 : leftRank - rightRank; });
        else if ((rowAxisFields.includes(field) || columnAxisFields.includes(field)) && sortOrder !== 'manual')
            order.sort((left, right) => { const compared = compareDateGroup(field, values[left].label || '(blank)', values[right].label || '(blank)'); return sortOrder === 'descending' ? -compared : compared; });
        fieldItemOrders.set(field, order);
        fieldItemOrdinals.set(field, new Map(order.map((sharedIndex, index) => [values[sharedIndex].key, index])));
    }
    const sharedIndex = (field, value) => dateGroupByIndex.has(field) || numberGroupByIndex.has(field) || manualGroupByIndex.has(field) ? fieldItemOrdinals.get(field)?.get(value.key) ?? -1 : sharedIndexByKey.get(field)?.get(value.key) ?? -1;
    const nativeBaseItemIndex = (measure) => { if (!['percentOf', 'differenceFrom', 'percentDifferenceFrom'].includes(measure.showAs ?? ''))
        return 0; const item = measure.baseItem ?? 'previous'; if (item === 'previous')
        return 1048828; if (item === 'next')
        return 1048829; const field = groupIndexByField.get(measure.baseField) ?? measure.baseField, index = fieldItemOrdinals.get(field)?.get(normalized(item.label)); if (index === undefined)
        throw Error('Selected Base Item is not present in the PivotTable field cache.'); return index; };
    const dateGroupCacheFields = dateGroups.map((group, index) => { const cacheIndex = width + index, items = shared.get(cacheIndex) ?? [], groupOrder = fieldItemOrders.get(cacheIndex) ?? items.map((_, itemIndex) => itemIndex), groupItems = groupOrder.map(item => sharedItemXml(items[item])).join(''), dateValues = sourceRows.flatMap(row => { const value = row[group.field], boundary = pivotDateGroupIsoBoundary(value.kind === 'n' ? Number(value.value) : value.kind === 's' ? value.value : null); return boundary === null ? [] : [boundary]; }).sort(), earliest = dateValues[0], latest = dateValues.at(-1), startDate = earliest ? (() => { if (earliest.startsWith('1900-02-29'))
        return undefined; if (group.unit !== 'weeks')
        return earliest; const date = new Date(`${earliest.slice(0, 10)}T00:00:00Z`); date.setUTCDate(date.getUTCDate() - ((date.getUTCDay() + 6) % 7)); return date.toISOString().replace(/\.\d{3}Z$/, ''); })() : undefined, endDate = latest?.startsWith('1900-02-29') ? undefined : latest, range = `<rangePr groupBy="${group.unit === 'weeks' ? 'days' : group.unit}"${group.unit === 'weeks' ? ' groupInterval="7"' : ''} autoStart="1" autoEnd="1"${startDate ? ` startDate="${esc(startDate)}"` : ''}${endDate ? ` endDate="${esc(endDate)}"` : ''}/>`; return `<cacheField name="${esc(groupHeaders[index])}" databaseField="0"><sharedItems/> <fieldGroup base="${group.field}">${range}<groupItems count="${groupOrder.length}">${groupItems}</groupItems></fieldGroup></cacheField>`; }).join(''), numberGroupCacheFields = numberGroups.map((group, index) => { const cacheIndex = width + dateGroups.length + index, items = shared.get(cacheIndex) ?? [], groupOrder = fieldItemOrders.get(cacheIndex) ?? items.map((_, itemIndex) => itemIndex), groupItems = groupOrder.map(item => sharedItemXml(items[item])).join(''), range = numberGroupRanges.get(group.field); return `<cacheField name="${esc(groupHeaders[dateGroups.length + index])}" databaseField="0"><sharedItems/><fieldGroup base="${group.field}"><rangePr groupBy="range" autoStart="0" autoEnd="0" startNum="${range.start}" endNum="${range.end}" groupInterval="${group.interval}"/><groupItems count="${groupOrder.length}">${groupItems}</groupItems></fieldGroup></cacheField>`; }).join(''), manualGroupCacheFields = manualGroups.map((group, index) => { const cacheIndex = width + dateGroups.length + numberGroups.length + index, items = shared.get(cacheIndex) ?? [], groupOrder = fieldItemOrders.get(cacheIndex) ?? items.map((_, itemIndex) => itemIndex), groupItems = groupOrder.map(item => sharedItemXml(items[item])).join(''), ordinal = new Map(groupOrder.map((sharedIndex, position) => [items[sharedIndex].key, position])), baseItems = shared.get(group.field) ?? [], mapping = baseItems.map(item => { const label = manualGroupMaps.get(group.field)?.get(item.label.toLocaleLowerCase()) ?? item.label, result = ordinal.get(normalized(label)); if (result === undefined)
        throw Error(`Manual group mapping for “${uniqueHeaders[group.field]}” is incomplete.`); return `<x v="${result}"/>`; }).join(''); return `<cacheField name="${esc(groupHeaders[dateGroups.length + numberGroups.length + index])}" databaseField="0"><sharedItems/><fieldGroup base="${group.field}"><discretePr count="${baseItems.length}">${mapping}</discretePr><groupItems count="${groupOrder.length}">${groupItems}</groupItems></fieldGroup></cacheField>`; }).join(''), cacheFields = uniqueHeaders.map((name, index) => `<cacheField name="${esc(name)}">${shared.has(index) ? `<sharedItems count="${shared.get(index).length}">${shared.get(index).map(sharedItemXml).join('')}</sharedItems>` : '<sharedItems/>'}</cacheField>`).join('') + dateGroupCacheFields + numberGroupCacheFields + manualGroupCacheFields;
    const records = recordRows.map(row => `<r>${row.map((value, index) => shared.has(index) ? `<x v="${sharedIndex(index, value)}"/>` : valueXml(value)).join('')}</r>`).join('');
    const worksheetSource = spec.sourceTable ? `<worksheetSource name="${esc(spec.sourceTable)}"/>` : `<worksheetSource ref="${esc(spec.sourceRange)}" sheet="${esc(sheetName)}"/>`;
    const cacheDefinition = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><pivotCacheDefinition xmlns="${ns}" xmlns:r="${relNs}" r:id="rId1" saveData="1" refreshOnLoad="0" recordCount="${recordRows.length}" createdVersion="6" refreshedVersion="6" minRefreshableVersion="3"><cacheSource type="worksheet">${worksheetSource}</cacheSource><cacheFields count="${cacheWidth}">${cacheFields}</cacheFields></pivotCacheDefinition>`;
    const cacheRecords = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><pivotCacheRecords xmlns="${ns}" count="${recordRows.length}">${records}</pivotCacheRecords>`;
    const cacheRelationships = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="${pkgNs}"><Relationship Id="rId1" Type="${relNs}/pivotCacheRecords" Target="pivotCacheRecords${cacheId}.xml"/></Relationships>`;
    const reportFilterKeys = filters.map(filter => ({ field: filter.field, values: filter.values ?? (filter.value === undefined ? [] : [filter.value]) })).filter(filter => filter.values.length).map(filter => ({ field: filter.field, keys: new Set(filter.values.map(value => normalized(value ?? ''))) })), itemFilterKeys = itemFilters.map(filter => ({ field: groupIndexByField.get(filter.field) ?? filter.field, keys: new Set(filter.values.map(value => normalized(value ?? ''))) })), visibleRows = recordRows.filter(row => !(spec.dateFilters ?? []).some(filter => { const value = row[filter.field], serial = value.kind === "n" ? pivotDateSerial(Number(value.value)) : value.kind === "s" ? pivotDateSerial(value.value) : null; return serial === null || Math.floor(serial) < filter.start || Math.floor(serial) > filter.end; }) && !reportFilterKeys.some(filter => !filter.keys.has(row[filter.field].key)) && !itemFilterKeys.some(filter => !filter.keys.has(row[filter.field].key)) && !labelFilters.some(filter => { const index = groupIndexByField.get(filter.field) ?? filter.field, value = row[index]; return !value || !matchesPivotLabelFilter(value.label, filter.operator, filter.value, filter.value2); }));
    const collapsedItemKeys = new Set();
    for (const [paths, fields] of [[expansionState.collapsedRows ?? [], rowFields], [expansionState.collapsedColumns ?? [], columnFields]])
        for (const path of paths) {
            const depth = path.length - 1, sourceField = fields[depth];
            if (sourceField === undefined)
                continue;
            const cacheField = groupIndexByField.get(sourceField) ?? sourceField;
            collapsedItemKeys.add(`${cacheField}\u0000${normalized(path[depth] === '\u0000' ? '' : path[depth] ?? '')}`);
        }
    const rowKey = (row) => JSON.stringify(rowAxisFields.map(field => row[field].key)), actualRows = [], seenRows = new Set();
    for (const row of visibleRows) {
        const key = rowKey(row);
        if (!seenRows.has(key)) {
            seenRows.add(key);
            actualRows.push(row);
        }
    }
    const actualColumns = [], seenColumns = new Set();
    for (const row of visibleRows) {
        const key = JSON.stringify(columnAxisFields.map(field => row[field].key));
        if (columnFields.length && !seenColumns.has(key)) {
            seenColumns.add(key);
            actualColumns.push(columnAxisFields.map(field => row[field]));
        }
    }
    // Keep source order here; the hierarchy traversal below applies each column field's own sort.
    const pivotFields = Array.from({ length: cacheWidth }, (_, index) => {
        const sourceField = sourceAxisField(index), axis = rowAxisFields.includes(index) ? 'axisRow' : columnAxisFields.includes(index) ? 'axisCol' : filters.some(filter => filter.field === index) ? 'axisPage' : undefined;
        if (valueFields.includes(index) && !axis)
            return '<pivotField dataField="1" showAll="0"/>';
        if (axis) {
            const items = shared.get(index), order = fieldItemOrders.get(index) ?? items.map((_, itemIndex) => itemIndex), filter = filters.find(entry => entry.field === sourceField), allowed = filter?.values ?? (filter?.value === undefined ? [] : [filter.value]), allowedKeys = new Set(allowed.map(value => normalized(value ?? ''))), itemFilter = itemFilters.find(entry => entry.field === sourceField), selected = itemFilter?.values ?? [], selectedKeys = new Set(selected.map(value => normalized(value ?? ''))), dateFilter = spec.dateFilters?.find(entry => entry.field === sourceField), multi = allowed.length > 1 || selected.length > 1 || !!itemFilter || !!dateFilter, hiddenDefault = order.map(itemIndex => { const value = items[itemIndex], itemOrdinal = dateGroupByIndex.has(index) || numberGroupByIndex.has(index) || manualGroupByIndex.has(index) ? fieldItemOrdinals.get(index)?.get(value.key) ?? itemIndex : itemIndex, dateFilter = spec.dateFilters?.find(entry => entry.field === sourceField), serial = value.kind === "n" ? pivotDateSerial(Number(value.value)) : value.kind === "s" ? pivotDateSerial(value.value) : null, hidden = allowedKeys.size > 0 && !allowedKeys.has(value.key) || !!itemFilter && !selectedKeys.has(value.key) || labelFilters.some(filter => filter.field === sourceField && !matchesPivotLabelFilter(value.label, filter.operator, filter.value, filter.value2)) || !!dateFilter && (serial === null || Math.floor(serial) < dateFilter.start || Math.floor(serial) > dateFilter.end); const collapsed = collapsedItemKeys.has(`${index}\u0000${value.key}`); return `<item x="${itemOrdinal}"${hidden ? ' h="1"' : ''}${collapsed ? ' sd="1"' : ''}/>`; }).join('') + '<item t="default"/>', subtotalFields = spec.subtotalFields ?? rowFields, valueSort = valueSortByField.get(sourceField), sortMeasureIndex = valueSort ? (valueSort.measureIndex ?? exportMeasures.findIndex(measure => measure.field === valueSort.measureField)) : -1, autoSortScope = valueSort && sortMeasureIndex >= 0 ? `<autoSortScope><pivotArea dataOnly="0" outline="0" fieldPosition="0"><references count="1"><reference field="4294967294" count="1" selected="0"><x v="${sortMeasureIndex}"/></reference></references></pivotArea></autoSortScope>` : '', valueFilter = (spec.valueFilters ?? (spec.valueFilter ? [spec.valueFilter] : [])).find(entry => entry.field === sourceField);
            return `<pivotField${valueFields.includes(index) ? ' dataField="1"' : ''} axis="${axis}" sortType="${fieldSortOrder(index)}" showAll="${spec.showEmptyItemsFields?.includes(sourceField) || spec.showEmptyItems ? 1 : 0}"${valueFilter ? ` autoShow="1" topAutoShow="${valueFilter.direction === 'top' ? 1 : 0}" itemPageCount="${valueFilter.count}" rankBy="${valueFilter.measureIndex ?? exportMeasures.findIndex(measure => measure.field === valueFilter.measureField)}"` : ''}${axis === 'axisRow' && spec.blankRows ? ` insertBlankRow="1"` : ''}${axis === 'axisRow' && (spec.subtotals === false || !subtotalFields.includes(sourceField)) || axis === 'axisCol' && (spec.columnSubtotals === false || spec.columnSubtotalFields !== undefined && !spec.columnSubtotalFields.includes(sourceField)) ? ` defaultSubtotal="0"` : ''}${filter ? ` multipleItemSelectionAllowed="${multi ? 1 : 0}"` : ''}${itemFilter ? ` multipleItemSelectionAllowed="${multi ? 1 : 0}"` : ''}><items count="${items.length + 1}">${hiddenDefault}</items>${autoSortScope}</pivotField>`;
        }
        return '<pivotField showAll="0"/>';
    }).join('');
    const rowFieldsXml = '<rowFields count="' + rowAxisFields.length + '">' + rowAxisFields.map(field => '<field x="' + field + '"/>').join('') + '</rowFields>';
    const axisItemIndex = (field, value) => fieldItemOrdinals.get(field)?.get(value.key) ?? -1;
    const root = { label: { key: '', label: '', kind: 'm', value: '' }, children: new Map() };
    for (const row of actualRows) {
        let node = root;
        for (const field of rowAxisFields) {
            const value = row[field], id = value.key;
            if (!node.children.has(id))
                node.children.set(id, { label: value, children: new Map() });
            node = node.children.get(id);
        }
    }
    const rowLines = [], collator = new Intl.Collator(undefined, { numeric: true, sensitivity: 'base' }), dateLabelCompare = (sourceField, left, right) => { const unit = dateGroups.find(group => group.field === sourceField)?.unit; if (unit === 'months') {
        const names = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
        return names.indexOf(left) - names.indexOf(right);
    } if (unit === 'quarters')
        return Number(left.slice(1)) - Number(right.slice(1)); return collator.compare(left, right); }, subtotalFields = spec.subtotalFields ?? rowFields, visit = (node, path, depth) => { const children = Array.from(node.children.values()), field = rowFields[depth], sortOrder = fieldSortOrder(field); if (sortOrder !== 'manual' || manualItemRanks.has(field))
        children.sort((left, right) => { if (sortOrder === 'manual') {
            const leftRank = manualRank(field, left.label.label), rightRank = manualRank(field, right.label.label);
            return leftRank === undefined ? rightRank === undefined ? 0 : 1 : rightRank === undefined ? -1 : leftRank - rightRank;
        } if (valueSortByField.has(field)) {
            const order = valueSortChildOrders.get(field)?.get(JSON.stringify(path.map(value => value.key))), leftRank = order?.get(left.label.key), rightRank = order?.get(right.label.key);
            if (leftRank !== undefined && rightRank !== undefined && leftRank !== rightRank)
                return leftRank - rightRank;
        } const compared = dateLabelCompare(field, left.label.label || '(blank)', right.label.label || '(blank)'); return sortOrder === 'descending' ? -compared : compared; }); for (const child of children) {
        const next = [...path, child.label], hasSubtotal = spec.subtotals !== false && subtotalFields.includes(field) && child.children.size > 0;
        if (hasSubtotal && spec.subtotalPosition === 'top')
            rowLines.push({ values: next, subtotal: true });
        if (child.children.size)
            visit(child, next, depth + 1);
        else
            rowLines.push({ values: next, subtotal: false });
        if (hasSubtotal && spec.subtotalPosition !== 'top')
            rowLines.push({ values: next, subtotal: true });
        if (spec.blankRows && child.children.size)
            rowLines.push({ values: [], subtotal: false, blank: true });
    } };
    visit(root, [], 0);
    const rowItemNodes = [];
    let previous = [];
    for (const line of rowLines) {
        if (line.blank) {
            rowItemNodes.push('<i t="blank"/>');
            previous = [];
            continue;
        }
        const indices = line.values.map((value, level) => axisItemIndex(rowAxisFields[level], value)), common = line.subtotal ? 0 : indices.reduce((count, value, index) => index < previous.length && previous[index] === value ? count + 1 : count, 0), encoded = indices.slice(common).map(value => `<x v="${value}"/>`).join('');
        rowItemNodes.push(`<i${line.subtotal ? ' t="default"' : ''}${common ? ` r="${common}"` : ''}>${encoded}</i>`);
        previous = indices;
    }
    if (spec.rowGrandTotals !== false)
        rowItemNodes.push(`<i t="grand">${rowAxisFields.map(() => '<x/>').join('')}</i>`);
    const rowItemsXml = '<rowItems count="' + rowItemNodes.length + '">' + rowItemNodes.join('') + '</rowItems>';
    const colFieldIndices = [...columnAxisFields, ...(valueFields.length > 1 ? [-2] : [])], colFields = colFieldIndices.length ? `<colFields count="${colFieldIndices.length}">${colFieldIndices.map(field => `<field x="${field}"/>`).join('')}</colFields>` : '';
    const colLines = [], columnRoot = { label: { key: '', label: '', kind: 'm', value: '' }, children: new Map() };
    if (columnFields.length) {
        for (const row of actualColumns) {
            let node = columnRoot;
            for (let level = 0; level < columnAxisFields.length; level++) {
                const value = row[level], id = value.key;
                if (!node.children.has(id))
                    node.children.set(id, { label: value, children: new Map() });
                node = node.children.get(id);
            }
        }
        const visitColumn = (node, path, depth) => { const children = Array.from(node.children.values()), field = columnFields[depth], sortOrder = fieldSortOrder(field); if (sortOrder !== 'manual' || manualItemRanks.has(field))
            children.sort((left, right) => { if (sortOrder === 'manual') {
                const leftRank = manualRank(field, left.label.label), rightRank = manualRank(field, right.label.label);
                return leftRank === undefined ? rightRank === undefined ? 0 : 1 : rightRank === undefined ? -1 : leftRank - rightRank;
            } if (valueSortByField.has(field)) {
                const order = valueSortChildOrders.get(field)?.get(JSON.stringify(path.map(value => value.key))), leftRank = order?.get(left.label.key), rightRank = order?.get(right.label.key);
                if (leftRank !== undefined && rightRank !== undefined && leftRank !== rightRank)
                    return leftRank - rightRank;
            } const compared = dateLabelCompare(field, left.label.label || '(blank)', right.label.label || '(blank)'); return sortOrder === 'descending' ? -compared : compared; }); for (const child of children) {
            const next = [...path, child.label];
            if (child.children.size)
                visitColumn(child, next, depth + 1);
            else
                colLines.push({ values: next, subtotal: false });
            if (spec.columnSubtotals && (spec.columnSubtotalFields === undefined || spec.columnSubtotalFields.includes(field)) && child.children.size)
                colLines.push({ values: next, subtotal: true });
        } };
        visitColumn(columnRoot, [], 0);
    }
    else
        colLines.push({ values: [], subtotal: false });
    const measureIndices = valueFields.length > 1 ? valueFields.map((_, index) => index) : [undefined], colCombinations = colLines.flatMap(line => measureIndices.map(measure => ({ line, measure }))), colItemsXml = colFieldIndices.length ? `<colItems count="${colCombinations.length}">${colCombinations.map(({ line, measure }) => `<i${line.subtotal ? ' t="default"' : ''}>${line.values.map((value, index) => `<x v="${axisItemIndex(columnAxisFields[index], value)}"/>`).join('')}${measure === undefined ? '' : `<x v="${measure}"/>`}</i>`).join('')}</colItems>` : '';
    let pageFields = '';
    if (filters.length) {
        const pageItems = filters.map(filter => { const allowed = filter.values ?? (filter.value === undefined ? [] : [filter.value]); if (allowed.length !== 1)
            return `<pageField fld="${filter.field}"/>`; const item = allowed[0] === null ? shared.get(filter.field).findIndex(x => x.label === '') : shared.get(filter.field).findIndex(x => normalized(x.label) === normalized(allowed[0])); if (item < 0)
            throw Error('Pivot filter item is not present in its source cache.'); return `<pageField fld="${filter.field}" item="${item}"/>`; });
        pageFields = `<pageFields count="${pageItems.length}">${pageItems.join('')}</pageFields>`;
    }
    const showDataAs = (measure) => measure.showAs === 'percentGrandTotal' ? 'percentOfTotal' : measure.showAs === 'percentRowTotal' ? 'percentOfRow' : measure.showAs === 'percentColumnTotal' ? 'percentOfCol' : measure.showAs === 'percentOf' ? 'percent' : measure.showAs === 'index' ? 'index' : measure.showAs === 'runningTotal' ? 'runTotal' : measure.showAs === 'differenceFrom' ? 'difference' : measure.showAs === 'percentDifferenceFrom' ? 'percentDiff' : 'normal', pivotShowAs = (measure) => measure.showAs === 'percentParentRow' ? 'percentOfParentRow' : measure.showAs === 'percentParentColumn' ? 'percentOfParentCol' : measure.showAs === 'percentRunningTotal' ? 'percentOfRunningTotal' : measure.showAs === 'rankAscending' ? 'rankAscending' : measure.showAs === 'rankDescending' ? 'rankDescending' : undefined;
    const rowCaption = rowFields.map(field => groupIndexByField.has(field) ? groupHeaders[groupIndexByField.get(field) - width] : uniqueHeaders[field]).join(' · '), columnCaption = columnFields.map(field => groupIndexByField.has(field) ? groupHeaders[groupIndexByField.get(field) - width] : uniqueHeaders[field]).join(' · '), measures = spec.measures ?? valueFields.map(field => ({ field, aggregation: spec.aggregation })), captionCounts = new Map(), valueCaptions = measures.map(measure => { const aggregation = { countNumbers: 'Count Numbers', stdDev: 'StdDev', stdDevp: 'StdDevp', var: 'Var', varp: 'VarP' }[measure.aggregation] ?? `${measure.aggregation[0].toUpperCase() + measure.aggregation.slice(1)}`, show = { percentGrandTotal: ' (% of Grand Total)', percentRowTotal: ' (% of Row Total)', percentColumnTotal: ' (% of Column Total)', percentOf: ' (% Of)', runningTotal: ' (Running Total In)', percentRunningTotal: ' (% Running Total In)', differenceFrom: ' (Difference From)', percentDifferenceFrom: ' (% Difference From)', rankAscending: ' (Rank Smallest to Largest)', rankDescending: ' (Rank Largest to Smallest)', index: ' (Index)' }[measure.showAs ?? 'normal'] ?? '', base = measure.caption?.trim() || `${aggregation} of ${uniqueHeaders[measure.field]}${show}`, count = (captionCounts.get(base) ?? 0) + 1; captionCounts.set(base, count); return count === 1 ? base : `${base} (${count})`; });
    const layout = spec.layout ?? 'tabular', compact = layout === 'compact', outline = layout === 'outline', chartFormats = pivotChartIndexes.length ? `<chartFormats count="${pivotChartIndexes.length}">${pivotChartIndexes.map(index => `<chartFormat chart="${index}" format="0" series="0"><pivotArea type="data" outline="0"/></chartFormat>`).join('')}</chartFormats>` : '', tableDefinition = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><pivotTableDefinition xmlns="${ns}" xmlns:r="${relNs}" name="${esc(tableName)}" cacheId="${cacheId}" dataCaption="${esc(spec.valuesCaption ?? 'Values')}" grandTotalCaption="${esc(spec.grandTotalCaption ?? 'Grand Total')}" rowHeaderCaption="${esc(spec.rowHeaderCaption ?? rowCaption)}"${columnCaption ? ` colHeaderCaption="${esc(columnCaption)}"` : ''} rowGrandTotals="${spec.rowGrandTotals === false ? 0 : 1}" colGrandTotals="${spec.columnGrandTotals === false ? 0 : 1}" subtotalTop="${spec.subtotalPosition === 'top' ? 1 : 0}" multipleFieldFilters="1" useAutoFormatting="1" applyNumberFormats="1" applyBorderFormats="1" applyFontFormats="1" applyPatternFormats="1" applyAlignmentFormats="1" applyWidthHeightFormats="1" compact="${compact ? 1 : 0}" compactData="${compact ? 1 : 0}" outline="${outline ? 1 : 0}" outlineData="${outline ? 1 : 0}" showHeaders="1" showDrill="1"${spec.missingCaption !== undefined ? ` showMissing="1" missingCaption="${esc(spec.missingCaption)}"` : ""} createdVersion="6" updatedVersion="6" minRefreshableVersion="3"${pivotChartIndexes.length ? ` chartFormat="1"` : ''}><location ref="${esc(outputRange)}" firstHeaderRow="1" firstDataRow="1" firstDataCol="1"/><pivotFields count="${cacheWidth}">${pivotFields}</pivotFields>${rowFieldsXml}${rowItemsXml}${colFields}${colItemsXml}${pageFields}<dataFields count="${valueFields.length}">${measures.map((measure, index) => { const extension = pivotShowAs(measure); return `<dataField name="${esc(valueCaptions[index])}" fld="${measure.field}" subtotal="${aggregate(measure.aggregation)}"${extension ? '' : ` showDataAs="${showDataAs(measure)}"`} baseField="${measure.baseField === undefined ? 0 : groupIndexByField.get(measure.baseField) ?? measure.baseField}" baseItem="${measure.showAs === 'percentOf' || measure.showAs === 'differenceFrom' || measure.showAs === 'percentDifferenceFrom' ? nativeBaseItemIndex(measure) : 0}"${extension ? `><extLst><ext uri="{E15A36E0-9728-4e99-A89B-3F7291B0FE68}"><x14:dataField xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" pivotShowAs="${extension}"/></ext></extLst></dataField>` : '/>'}`; }).join('')}</dataFields>${chartFormats}<pivotTableStyleInfo name="PivotStyleMedium2" showRowHeaders="1" showColHeaders="1" showRowStripes="1" showColStripes="0"/>${labelFilters.length ? `<filters count="${labelFilters.length}">${labelFilters.map((filter, index) => `<filter fld="${groupIndexByField.get(filter.field) ?? filter.field}" type="${{ equals: 'captionEqual', notEqual: 'captionNotEqual', beginsWith: 'captionBeginsWith', notBeginsWith: 'captionNotBeginsWith', endsWith: 'captionEndsWith', notEndsWith: 'captionNotEndsWith', contains: 'captionContains', notContains: 'captionNotContains', greaterThan: 'captionGreaterThan', greaterOrEqual: 'captionGreaterThanOrEqual', lessThan: 'captionLessThan', lessOrEqual: 'captionLessThanOrEqual', between: 'captionBetween', notBetween: 'captionNotBetween' }[filter.operator]}" evalOrder="-1" id="${index}" stringValue1="${esc(filter.value)}"${filter.value2 !== undefined ? ` stringValue2="${esc(filter.value2)}"` : ''}/>`).join('')}</filters>` : ''}</pivotTableDefinition>`;
    const [outputStartRef, outputEndRef] = outputRange.split(':'), outputStart = parseRef(outputStartRef), outputEnd = parseRef(outputEndRef ?? outputStartRef), pageRows = filters.length, locationRange = `${colToName(outputStart.col)}${outputStart.row + pageRows + 1}:${colToName(outputEnd.col)}${outputEnd.row + 1}`, locationAttrs = `ref="${esc(locationRange)}" firstHeaderRow="1" firstDataRow="1" firstDataCol="${Math.max(1, compact ? 1 : rowFields.length)}"${pageRows ? ` rowPageCount="${pageRows}"` : ''}`, nativeTableDefinition = tableDefinition.replace(`<location ref="${esc(outputRange)}" firstHeaderRow="1" firstDataRow="1" firstDataCol="1"/>`, `<location ${locationAttrs}/>`), columnGrandItems = columnFields.length && spec.columnGrandTotals !== false ? measureIndices.map(measure => `<i t="grand">${columnFields.map(() => '<x/>').join('')}${measure === undefined ? '' : `<x v="${measure}"/>`}</i>`).join('') : '';
    const withColumnItems = colFieldIndices.length ? nativeTableDefinition.replace(`<colItems count="${colCombinations.length}">`, `<colItems count="${colCombinations.length + (columnFields.length && spec.columnGrandTotals !== false ? measureIndices.length : 0)}">`).replace('</colItems>', `${columnGrandItems}</colItems>`) : nativeTableDefinition.replace(pageFields || '<dataFields', `<colItems count="1"><i t="grand"><x v="0"/></i></colItems>${pageFields || '<dataFields'}`);
    const tableRelationships = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="${pkgNs}"><Relationship Id="rId1" Type="${relNs}/pivotCacheDefinition" Target="../pivotCache/pivotCacheDefinition${cacheId}.xml"/></Relationships>`;
    // SlicerCache tabular item indexes address the cacheField sharedItems sequence,
    // not the PivotField's independently sorted item sequence.
    const fieldItems = [...new Set([...rowFields, ...columnFields])].map(field => ({ field, items: (shared.get(field) ?? []).map(value => value.label) }));
    return { cacheDefinition, cacheRecords, cacheRelationships, tableDefinition: withColumnItems, tableRelationships, outputRange, fieldItems };
}
