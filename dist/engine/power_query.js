/** Offline, repeatable query pipelines over a bounded worksheet rectangle. */
import { Workbook } from './workbook.js';
import { colToName, parseRef } from './parser.js';
import { DATA_TOOLS, previewDataTool } from './tools_v8.js';
import { parseCSVAsync } from './csv.js';
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
    let sourceRange, sourceBounds;
    if (sourceKind === 'worksheet') {
        if (typeof q.sourceRange !== 'string')
            throw Error('Worksheet queries need a source range.');
        sourceBounds = range(q.sourceRange, rows, cols);
        sourceRange = sourceBounds.ref;
        const height = sourceBounds.b.row - sourceBounds.a.row + 1, width = sourceBounds.b.col - sourceBounds.a.col + 1;
        if (dest.row + height > rows || dest.col + width > cols)
            throw Error('Query output would extend beyond the worksheet.');
        if (!(dest.row + height - 1 < sourceBounds.a.row || dest.row > sourceBounds.b.row || dest.col + width - 1 < sourceBounds.a.col || dest.col > sourceBounds.b.col))
            throw Error('Query output must not overlap its source range.');
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
        if (sourceBounds && !(old.b.row < sourceBounds.a.row || old.a.row > sourceBounds.b.row || old.b.col < sourceBounds.a.col || old.a.col > sourceBounds.b.col))
            throw Error('Previous query output overlaps its source range.');
        lastOutputRange = old.ref;
    }
    return { id: q.id, name: q.name.trim(), sourceKind, ...(sourceRange ? { sourceRange } : {}), ...(sourceUrl ? { sourceUrl } : {}), ...(jsonPointer ? { jsonPointer } : {}), destination: `${colToName(dest.col)}${dest.row + 1}`, steps: [...q.steps], ...(cachedData ? { cachedData } : {}), ...(lastOutputRange ? { lastOutputRange } : {}) };
}
/** Replays a saved pipeline on a private workbook copy and returns one atomic load-to-sheet transaction. */
export function buildQueryRefresh(book, value) {
    const q = validateQuerySpec(value, book.rows, book.cols), { a, b } = range(q.sourceRange, book.rows, book.cols), dest = parseRef(q.destination);
    const snapshot = Workbook.fromJSON(book.toJSON());
    for (const step of q.steps) {
        const outcome = previewDataTool(snapshot, { r1: a.row, r2: b.row, c1: a.col, c2: b.col }, step);
        snapshot.applyEdits(outcome.edits, false);
    }
    const edits = [];
    if (q.lastOutputRange) {
        const old = range(q.lastOutputRange, book.rows, book.cols);
        for (let r = old.a.row; r <= old.b.row; r++)
            for (let c = old.a.col; c <= old.b.col; c++)
                edits.push({ ref: `${colToName(c)}${r + 1}`, raw: '' });
    }
    for (let r = a.row; r <= b.row; r++)
        for (let c = a.col; c <= b.col; c++) {
            const result = snapshot.getValue(`${colToName(c)}${r + 1}`), raw = result.error ? `'${result.error.code}` : typeof result.value === 'number' ? String(result.value) : typeof result.value === 'boolean' ? String(result.value).toUpperCase() : String(result.value ?? '');
            edits.push({ ref: `${colToName(dest.col + c - a.col)}${dest.row + r - a.row + 1}`, raw: raw.startsWith('=') ? `'${raw}` : raw });
        }
    const end = `${colToName(dest.col + b.col - a.col)}${dest.row + b.row - a.row + 1}`;
    return { edits, outputRange: `${q.destination}:${end}`, rows: b.row - a.row + 1, columns: b.col - a.col + 1 };
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
/** Replays a saved external query's last fetched values for deterministic, offline XLSX export. */
export function buildCachedQueryRefresh(book, value) {
    const q = validateQuerySpec(value, book.rows, book.cols);
    if (q.sourceKind === 'worksheet' || !q.cachedData)
        throw Error('This external query has no cached result. Refresh it while online first.');
    return materializeRows(book, q, q.cachedData);
}
