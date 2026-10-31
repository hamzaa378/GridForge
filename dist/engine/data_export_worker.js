import { SpreadsheetDocument } from './document.js';
import { parseRef, colToName } from './parser.js';
import { display } from './workbook.js';
const MAX_CSV_CELLS = 2_000_000;
const MAX_EXPORT_CHARACTERS = 64 * 1024 * 1024;
const workerScope = self;
workerScope.onmessage = event => {
    const { requestId, format, document } = event.data;
    try {
        let output;
        if (format === 'json') {
            output = JSON.stringify(document, null, 2);
            if (output.length > MAX_EXPORT_CHARACTERS)
                throw Error('This workbook JSON is larger than the 64 MB export limit. Remove unused data or split the workbook before exporting.');
        }
        else if (format === 'csv') {
            const workbook = SpreadsheetDocument.fromJSON(document).active.book, entries = workbook.allEntries();
            let maxRow = 0, maxColumn = 0;
            for (const entry of entries) {
                const point = parseRef(entry.ref);
                if (point) {
                    maxRow = Math.max(maxRow, point.row);
                    maxColumn = Math.max(maxColumn, point.col);
                }
            }
            const cellCount = (maxRow + 1) * (maxColumn + 1);
            if (!Number.isSafeInteger(cellCount) || cellCount > MAX_CSV_CELLS)
                throw Error(`This sheet's CSV area is ${cellCount.toLocaleString()} cells, above the 2,000,000-cell export limit. Move or remove distant data, or export a smaller sheet.`);
            const rows = [];
            let outputCharacters = 0;
            for (let row = 0; row <= maxRow; row++) {
                const values = [];
                for (let column = 0; column <= maxColumn; column++) {
                    const result = workbook.getValue(`${colToName(column)}${row + 1}`), shown = display(result);
                    const safe = typeof result.value === 'string' && /^[\t\r\n ]*[=+\-@]/.test(shown) ? `'${shown}` : shown;
                    values.push(/[",\r\n]/.test(safe) ? `"${safe.replace(/"/g, '""')}"` : safe);
                }
                const line = values.join(',');
                outputCharacters += line.length + (row ? 2 : 0);
                if (outputCharacters > MAX_EXPORT_CHARACTERS)
                    throw Error('This CSV is larger than the 64 MB export limit. Export a smaller range or split the sheet.');
                rows.push(line);
            }
            output = rows.join('\r\n');
        }
        else
            throw Error('Unsupported data export type.');
        const bytes = new TextEncoder().encode(output), buffer = bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength);
        workerScope.postMessage({ requestId, kind: 'result', bytes: buffer }, [buffer]);
    }
    catch (error) {
        workerScope.postMessage({ requestId, kind: 'error', message: error instanceof Error ? error.message : String(error) });
    }
};
