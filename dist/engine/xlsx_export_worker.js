import { SpreadsheetDocument } from './document.js';
import { exportXlsx } from './xlsx.js';
const workerScope = self;
workerScope.onmessage = event => {
    const { requestId, document } = event.data;
    try {
        const workbook = SpreadsheetDocument.fromJSON(document);
        const bytes = exportXlsx(workbook);
        const buffer = bytes.slice().buffer;
        workerScope.postMessage({ requestId, kind: 'result', bytes: buffer }, [buffer]);
    }
    catch (error) {
        workerScope.postMessage({ requestId, kind: 'error', message: error instanceof Error ? error.message : String(error) });
    }
};
