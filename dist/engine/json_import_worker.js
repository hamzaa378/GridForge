import { SpreadsheetDocument } from './document.js';
const scope = self;
const MAX_INPUT_BYTES = 64 * 1024 * 1024;
scope.onmessage = event => {
    const { requestId, bytes } = event.data;
    try {
        if (!(bytes instanceof ArrayBuffer))
            throw Error('JSON import received invalid file data.');
        if (bytes.byteLength > MAX_INPUT_BYTES)
            throw Error('GridForge JSON files are limited to 64 MB. Export a smaller workbook or remove unused data.');
        scope.postMessage({ requestId, kind: 'progress', message: 'Reading workbook JSON…' });
        const source = new TextDecoder('utf-8', { fatal: false }).decode(bytes), parsed = JSON.parse(source);
        scope.postMessage({ requestId, kind: 'progress', message: 'Restoring workbook and formulas…' });
        const document = SpreadsheetDocument.fromJSON(parsed);
        scope.postMessage({ requestId, kind: 'result', document: document.toJSON() });
    }
    catch (error) {
        scope.postMessage({ requestId, kind: 'error', message: error instanceof Error ? error.message : 'Unable to import this GridForge JSON workbook.' });
    }
};
