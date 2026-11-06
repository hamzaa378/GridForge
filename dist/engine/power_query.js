/** Offline, repeatable query pipelines over a bounded worksheet rectangle. */
import { Workbook } from './workbook.js';
import { colToName, parseRef } from './parser.js';
import { DATA_TOOLS, previewDataTool } from './tools_v8.js';
import { parseCSVAsync } from './csv.js';
import { executePowerQueryM, validatePowerQueryM } from './power_query_m.js';
const queryable = new Set(DATA_TOOLS.filter(tool => ['text', 'number', 'cleanup'].includes(tool.kind)).map(tool => tool.id));
const range = (value, rows, cols) => { const parts = value.replace(/\$/g, '').toUpperCase().split(':'), a = parseRef(parts[0] ?? ''), b = parseRef(parts[1] ?? parts[0] ?? ''); if (parts.length > 2 || !a || !b || a.row > b.row || a.col > b.col || b.row >= rows || b.col >= cols || (b.row - a.row + 1) * (b.col - a.col + 1) > 12000)
    throw Error('Query source must be an in-bounds rectangle of at most 12,000 cells.'); return { a, b, ref: `${colToName(a.col)}${a.row + 1}:${colToName(b.col)}${b.row + 1}` }; };
export function validateQuerySpec(value, rows, cols) {
    if (!value || typeof value !== 'object' || Array.isArray(value))
        throw Error('Invalid query definition.');
    const q = value;
    const sourceKind = (q.sourceKind ?? 'worksheet');
    if (typeof q.id !== 'string' || !/^query-[A-Za-z0-9_-]{1,60}$/.test(q.id) || typeof q.name !== 'string' || !q.name.trim() || q.name.length > 80 || !['worksheet', 'csv-url', 'json-url'].includes(sourceKind) || typeof q.destination !== 'string' || !Array.isArray(q.steps) || q.steps.length > 16 || q.steps.some(step => typeof step !== 'string' || !queryable.has(step)))
        throw Error('Query needs a name, source, destination and up to 16 supported transformation steps.');
    const dest = parseRef(q.destination.replace(/\$/g, '').toUpperCase());
    if (!dest || dest.row >= rows || dest.col >= cols)
        throw Error('Query destination must be a valid cell address.');
    let sourceRange, sourceBounds, appendRanges, appendBounds = [];
    if (sourceKind === 'worksheet') {
        if (typeof q.sourceRange !== 'string')
            throw Error('Worksheet queries need a source range.');
        sourceBounds = range(q.sourceRange, rows, cols);
        sourceRange = sourceBounds.ref;
        if (q.appendRanges !== undefined) {
            if (!Array.isArray(q.appendRanges) || q.appendRanges.length > 7 || q.appendRanges.some(item => typeof item !== 'string'))
                throw Error('Append Queries accepts up to seven additional worksheet ranges.');
            appendBounds = q.appendRanges.map(item => range(item, rows, cols));
            if (appendBounds.some(item => item.a.col !== sourceBounds.a.col || item.b.col !== sourceBounds.b.col))
                throw Error('Appended ranges must have the same number of columns as the primary source.');
            const sources = [sourceBounds, ...appendBounds];
            for (let i = 0; i < sources.length; i++)
                for (let j = 0; j < i; j++) {
                    const left = sources[i], right = sources[j];
                    if (!(left.b.row < right.a.row || left.a.row > right.b.row || left.b.col < right.a.col || left.a.col > right.b.col))
                        throw Error('Append Query source ranges must not overlap.');
                }
            const totalCells = sources.reduce((sum, item) => sum + (item.b.row - item.a.row + 1) * (item.b.col - item.a.col + 1), 0);
            if (totalCells > 12000)
                throw Error('Combined Append Query sources are limited to 12,000 cells.');
            appendRanges = appendBounds.map(item => item.ref);
        }
        const sources = [sourceBounds, ...appendBounds], height = sources.reduce((sum, item, index) => sum + item.b.row - item.a.row + 1 - (index ? 1 : 0), 0), width = sourceBounds.b.col - sourceBounds.a.col + 1;
        if (dest.row + height > rows || dest.col + width > cols)
            throw Error('Query output would extend beyond the worksheet.');
        for (const source of sources)
            if (!(dest.row + height - 1 < source.a.row || dest.row > source.b.row || dest.col + width - 1 < source.a.col || dest.col > source.b.col))
                throw Error('Query output must not overlap any source range.');
    }
    let sourceUrl, jsonPointer;
    if (sourceKind !== 'worksheet') {
        if (typeof q.sourceUrl !== 'string')
            throw Error('External queries need a CSV or JSON URL.');
        let parsed;
        try {
            parsed = new URL(q.sourceUrl.trim());
        }
        catch {
            throw Error('Enter a complete HTTPS URL for the data source.');
        }
        if (parsed.protocol !== 'https:' || parsed.username || parsed.password || parsed.href.length > 2048)
            throw Error('External queries require an HTTPS URL without embedded credentials.');
        sourceUrl = parsed.href;
    }
    if (q.jsonPointer !== undefined) {
        if (sourceKind !== 'json-url' || typeof q.jsonPointer !== 'string' || q.jsonPointer.length > 300 || q.jsonPointer !== '' && !q.jsonPointer.startsWith('/') || /~(?![01])/.test(q.jsonPointer) || q.jsonPointer.split('/').length > 33)
            throw Error('JSON Pointer must be empty or use a valid slash path with at most 32 levels.');
        jsonPointer = q.jsonPointer;
    }
    let mExpression;
    if (q.mExpression !== undefined) {
        if (typeof q.mExpression !== 'string' || q.mExpression.length > 20_000)
            throw Error('Power Query M subset expressions are limited to 20,000 characters.');
        validatePowerQueryM(q.mExpression);
        mExpression = q.mExpression;
    }
    let cachedData;
    if (q.cachedData !== undefined) {
        if (!Array.isArray(q.cachedData) || !q.cachedData.length || q.cachedData.length > 1000 || q.cachedData.some(row => !Array.isArray(row) || row.length > 104 || row.some(cell => cell !== null && typeof cell !== 'string' && typeof cell !== 'number' && typeof cell !== 'boolean' || typeof cell === 'number' && !Number.isFinite(cell))))
            throw Error('Cached query data must fit within 1,000 rows, 104 columns, and contain only plain values.');
        const width = Math.max(0, ...q.cachedData.map(row => row.length));
        if (!width || q.cachedData.length * width > 12000 || dest.row + q.cachedData.length > rows || dest.col + width > cols)
            throw Error('Cached query output exceeds the worksheet or 12,000-cell limit.');
        cachedData = q.cachedData.map(row => [...row]);
    }
    let lastOutputRange;
    if (q.lastOutputRange !== undefined) {
        if (typeof q.lastOutputRange !== 'string')
            throw Error('Invalid previous query output range.');
        const old = range(q.lastOutputRange, rows, cols);
        if (sourceBounds && !(old.b.row < sourceBounds.a.row || old.a.row > sourceBounds.b.row || old.b.col < sourceBounds.a.col || old.a.col > sourceBounds.b.col) || appendBounds.some(source => !(old.b.row < source.a.row || old.a.row > source.b.row || old.b.col < source.a.col || old.a.col > source.b.col)))
            throw Error('Previous query output overlaps a source range.');
        lastOutputRange = old.ref;
    }
    return { id: q.id, name: q.name.trim(), sourceKind, ...(sourceRange ? { sourceRange } : {}), ...(appendRanges?.length ? { appendRanges } : {}), ...(sourceUrl ? { sourceUrl } : {}), ...(jsonPointer ? { jsonPointer } : {}), ...(mExpression ? { mExpression } : {}), destination: `${colToName(dest.col)}${dest.row + 1}`, steps: [...q.steps], ...(cachedData ? { cachedData } : {}), ...(lastOutputRange ? { lastOutputRange } : {}) };
}
/** Replays a saved pipeline on a private workbook copy and returns one atomic load-to-sheet transaction. */
export function buildQueryRefresh(book, value) {
    const q = validateQuerySpec(value, book.rows, book.cols), sources = [q.sourceRange, ...(q.appendRanges ?? [])].map(source => range(source, book.rows, book.cols)), dest = parseRef(q.destination), width = sources[0].b.col - sources[0].a.col + 1, headerKey = (row, bounds) => Array.from({ length: width }, (_, index) => { const result = book.getValue(`${colToName(bounds.a.col + index)}${row + 1}`); return result.error ? '' : String(result.value ?? '').trim(); }), headers = headerKey(sources[0].a.row, sources[0]);
    if (q.appendRanges?.length) {
        if (headers.some(header => !header) || new Set(headers.map(header => header.toLocaleLowerCase())).size !== headers.length)
            throw Error('Append Queries needs nonblank, unique column headers in the primary source.');
        for (const bounds of sources.slice(1)) {
            const other = headerKey(bounds.a.row, bounds);
            if (other.length !== headers.length || other.some((header, index) => header.toLocaleLowerCase() !== headers[index].toLocaleLowerCase()))
                throw Error('Append Query ranges must have the same column headers in the same order.');
        }
    }
    const data = [];
    for (const [sourceIndex, bounds] of sources.entries())
        for (let row = bounds.a.row + (sourceIndex ? 1 : 0); row <= bounds.b.row; row++)
            data.push(Array.from({ length: width }, (_, column) => { const result = book.getValue(`${colToName(bounds.a.col + column)}${row + 1}`); return result.error ? result.error.code : result.value; }));
    const snapshot = new Workbook(Math.max(1, data.length), Math.max(1, width)), seed = [];
    data.forEach((row, r) => row.forEach((cell, c) => seed.push({ ref: `${colToName(c)}${r + 1}`, raw: rawValue(cell) })));
    snapshot.applyEdits(seed, false);
    for (const step of q.steps) {
        const outcome = previewDataTool(snapshot, { r1: 0, r2: data.length - 1, c1: 0, c2: width - 1 }, step);
        snapshot.applyEdits(outcome.edits, false);
    }
    let rows = data.map((row, r) => row.map((_, c) => { const result = snapshot.getValue(`${colToName(c)}${r + 1}`); return result.error ? result.error.code : result.value; }));
    if (q.mExpression)
        rows = executePowerQueryM(rows, q.mExpression);
    const materialized = materializeRows(book, q, rows), edits = [];
    if (q.lastOutputRange) {
        const old = range(q.lastOutputRange, book.rows, book.cols);
        for (let r = old.a.row; r <= old.b.row; r++)
            for (let c = old.a.col; c <= old.b.col; c++)
                edits.push({ ref: `${colToName(c)}${r + 1}`, raw: '' });
    }
    return { ...materialized, edits: [...edits, ...materialized.edits] };
}
async function rowsFromExternal(source, kind, jsonPointer, onProgress) {
    let rows;
    if (kind === 'csv-url')
        rows = await parseCSVAsync(source, onProgress);
    else {
        let data;
        try {
            data = JSON.parse(source);
        }
        catch {
            throw Error('The JSON source is not valid JSON.');
        }
        if (jsonPointer) {
            for (const rawPart of jsonPointer.slice(1).split('/')) {
                const part = rawPart.replace(/~1/g, '/').replace(/~0/g, '~');
                if (Array.isArray(data)) {
                    if (!/^(?:0|[1-9]\d*)$/.test(part) || Number(part) >= data.length)
                        throw Error(`JSON Pointer “${jsonPointer}” does not exist in the response.`);
                    data = data[Number(part)];
                }
                else if (data && typeof data === 'object' && Object.hasOwn(data, part))
                    data = data[part];
                else
                    throw Error(`JSON Pointer “${jsonPointer}” does not exist in the response.`);
            }
        }
        if (!Array.isArray(data))
            throw Error('JSON queries expect an array of rows or an array of objects.');
        if (data.every(Array.isArray))
            rows = data.map(row => row.map(cell => externalCell(cell)));
        else if (data.every(item => item && typeof item === 'object' && !Array.isArray(item))) {
            const objects = data.map(flattenJsonRecord), headers = [...new Set(objects.flatMap(item => Object.keys(item)))];
            if (headers.length > 104)
                throw Error('Flattened JSON data is limited to 104 columns.');
            rows = [headers, ...objects.map(item => headers.map(header => externalCell(item[header])))];
        }
        else
            throw Error('JSON rows must all be arrays or all be objects.');
    }
    const width = Math.max(0, ...rows.map(row => row.length));
    if (!rows.length || !width || !rows.some(row => row.some(cell => cell !== null && cell !== '')))
        throw Error('The external source contains no tabular data.');
    if (rows.length > 1000 || width > 104 || rows.length * width > 12000)
        throw Error('External data is limited to 1,000 rows, 104 columns, and 12,000 cells.');
    return rows.map(row => Array.from({ length: width }, (_, i) => row[i] ?? null));
}
function externalCell(value) {
    if (value === null || value === undefined)
        return null;
    if (typeof value === 'string' || typeof value === 'boolean')
        return value;
    if (typeof value === 'number' && Number.isFinite(value))
        return value;
    if (typeof value === 'object')
        throw Error('JSON arrays must be selected or expanded before loading; nested arrays cannot fit into a single worksheet cell.');
    throw Error('External data contains an unsupported value.');
}
/** Flattens JSON records into deterministic worksheet columns while rejecting ambiguous paths. */
function flattenJsonRecord(record) {
    const result = Object.create(null), seen = new Set();
    const visit = (value, path, depth) => {
        if (depth > 32)
            throw Error('Nested JSON records are limited to 32 levels.');
        if (value !== null && typeof value === 'object' && !Array.isArray(value)) {
            const entries = Object.entries(value);
            if (!entries.length) {
                if (!path)
                    throw Error('A JSON row cannot be an empty object.');
                if (seen.has(path))
                    throw Error(`Nested JSON fields produce an ambiguous column name: ${path}`);
                seen.add(path);
                result[path] = null;
                return;
            }
            for (const [key, child] of entries) {
                const childPath = path ? `${path}.${key}` : key;
                if (!childPath || childPath.length > 512)
                    throw Error('Flattened JSON column names are limited to 512 characters.');
                visit(child, childPath, depth + 1);
            }
            return;
        }
        if (!path)
            throw Error('A JSON row must contain named fields.');
        if (seen.has(path))
            throw Error(`Nested JSON fields produce an ambiguous column name: ${path}`);
        seen.add(path);
        result[path] = value;
    };
    for (const [key, value] of Object.entries(record))
        visit(value, key, 0);
    return result;
}
function rawValue(value) { const raw = value === null ? '' : typeof value === 'boolean' ? value ? 'TRUE' : 'FALSE' : String(value); return raw.startsWith('=') ? `'${raw}` : raw; }
async function readBoundedText(response, limit) {
    const reader = response.body?.getReader();
    if (!reader) {
        const text = await response.text();
        if (new TextEncoder().encode(text).byteLength > limit)
            throw Error('External data is limited to 3 MB.');
        return text;
    }
    const chunks = [];
    let size = 0;
    for (;;) {
        const result = await reader.read();
        if (result.done)
            break;
        size += result.value.byteLength;
        if (size > limit) {
            await reader.cancel();
            throw Error('External data is limited to 3 MB.');
        }
        chunks.push(result.value);
    }
    const bytes = new Uint8Array(size);
    let offset = 0;
    for (const chunk of chunks) {
        bytes.set(chunk, offset);
        offset += chunk.byteLength;
    }
    try {
        return new TextDecoder('utf-8', { fatal: true }).decode(bytes);
    }
    catch {
        throw Error('External data must be valid UTF-8 text.');
    }
}
function materializeRows(book, q, rows) {
    const destination = parseRef(q.destination), width = Math.max(0, ...rows.map(row => row.length));
    if (!rows.length || !width || rows.length * width > 12000)
        throw Error('Query output is empty or exceeds 12,000 cells.');
    if (destination.row + rows.length > book.rows || destination.col + width > book.cols)
        throw Error('Query output would extend beyond the worksheet.');
    if (q.sourceKind === 'worksheet') {
        for (const sourceText of [q.sourceRange, ...(q.appendRanges ?? [])]) {
            if (!sourceText)
                continue;
            const source = range(sourceText, book.rows, book.cols);
            if (!(destination.row + rows.length - 1 < source.a.row || destination.row > source.b.row || destination.col + width - 1 < source.a.col || destination.col > source.b.col))
                throw Error('Query output must not overlap any worksheet source range after M transformations.');
        }
    }
    const edits = [];
    if (q.lastOutputRange) {
        const old = range(q.lastOutputRange, book.rows, book.cols);
        for (let row = old.a.row; row <= old.b.row; row++)
            for (let col = old.a.col; col <= old.b.col; col++)
                edits.push({ ref: `${colToName(col)}${row + 1}`, raw: '' });
    }
    rows.forEach((row, r) => row.forEach((value, c) => edits.push({ ref: `${colToName(destination.col + c)}${destination.row + r + 1}`, raw: rawValue(value) })));
    const end = `${colToName(destination.col + width - 1)}${destination.row + rows.length}`;
    return { edits, outputRange: `${q.destination}:${end}`, rows: rows.length, columns: width, cachedData: rows.map(row => [...row]) };
}
/** Fetches an explicitly configured public HTTPS CSV/JSON source without cookies, credentials, or redirects. */
export async function buildExternalQueryRefresh(book, value, onProgress) {
    const q = validateQuerySpec(value, book.rows, book.cols);
    if (q.sourceKind === 'worksheet' || !q.sourceUrl)
        throw Error('This is not an external query.');
    const kind = q.sourceKind;
    const controller = new AbortController(), timeout = setTimeout(() => controller.abort(), 15000);
    try {
        const response = await fetch(q.sourceUrl, { method: 'GET', mode: 'cors', credentials: 'omit', redirect: 'error', cache: 'no-store', signal: controller.signal, headers: { Accept: kind === 'csv-url' ? 'text/csv, text/plain;q=0.9' : 'application/json' } });
        if (!response.ok)
            throw Error(`The data source returned HTTP ${response.status}.`);
        const length = Number(response.headers.get('content-length'));
        if (Number.isFinite(length) && length > 3 * 1024 * 1024)
            throw Error('External data is limited to 3 MB.');
        const content = await readBoundedText(response, 3 * 1024 * 1024);
        let rows = await rowsFromExternal(content, kind, q.jsonPointer, onProgress);
        if (q.steps.length) {
            const width = rows[0].length, temp = new Workbook(rows.length, width), edits = [];
            rows.forEach((row, r) => row.forEach((cell, c) => edits.push({ ref: `${colToName(c)}${r + 1}`, raw: rawValue(cell) })));
            temp.applyEdits(edits, false);
            for (const step of q.steps) {
                const result = previewDataTool(temp, { r1: 0, r2: rows.length - 1, c1: 0, c2: width - 1 }, step);
                temp.applyEdits(result.edits, false);
            }
            rows = rows.map((row, r) => row.map((_, c) => { const result = temp.getValue(`${colToName(c)}${r + 1}`); return result.error ? result.error.code : result.value; }));
        }
        if (q.mExpression)
            rows = executePowerQueryM(rows, q.mExpression);
        return materializeRows(book, q, rows);
    }
    catch (error) {
        if (error instanceof Error && error.name === 'AbortError')
            throw Error('The data source request timed out after 15 seconds.');
        if (error instanceof TypeError)
            throw Error('The data source could not be reached. Check the URL, CORS access, and network connection.');
        throw error;
    }
    finally {
        clearTimeout(timeout);
    }
}
/** Apply a one-time desktop-authenticated fetch through the same bounded import/M pipeline. */
export async function buildAuthenticatedQueryImport(book, source, content) {
    const query = validateQuerySpec({ id: 'query-authenticated-import', name: 'Authenticated import', sourceKind: source.kind, sourceUrl: source.url, destination: source.destination, steps: [], ...(source.jsonPointer ? { jsonPointer: source.jsonPointer } : {}), ...(source.mExpression ? { mExpression: source.mExpression } : {}) }, book.rows, book.cols);
    let rows = await rowsFromExternal(content, query.sourceKind, query.jsonPointer);
    if (query.mExpression)
        rows = executePowerQueryM(rows, query.mExpression);
    return materializeRows(book, query, rows);
}
/** Replays a saved external query's last fetched values for deterministic, offline XLSX export. */
export function buildCachedQueryRefresh(book, value) {
    const q = validateQuerySpec(value, book.rows, book.cols);
    if (q.sourceKind === 'worksheet' || !q.cachedData)
        throw Error('This external query has no cached result. Refresh it while online first.');
    return materializeRows(book, q, q.cachedData);
}
