import { configureXlsxXmlParser, importXlsx } from './xlsx.js';
const workerScope = self;
workerScope.onmessage = async (event) => {
    const { requestId, bytes } = event.data;
    try {
        const moduleUrl = new URL('../vendor/xmldom.js', import.meta.url).href;
        const xmlModule = await import(moduleUrl);
        const XmlDomParser = xmlModule.DOMParser ?? xmlModule.default?.DOMParser;
        if (!XmlDomParser)
            throw Error('The background XML parser did not load.');
        configureXlsxXmlParser(text => { let parseError = ''; const parser = new XmlDomParser({ onError: (level, message) => { if (level === 'error' || level === 'fatalError')
                parseError = message; } }), document = parser.parseFromString(text, 'application/xml'); if (parseError)
            throw Error(`Invalid XML inside XLSX: ${parseError}`); return document; });
        const imported = await importXlsx(new Uint8Array(bytes), { onProgress: message => workerScope.postMessage({ requestId, kind: 'progress', message }) });
        workerScope.postMessage({ requestId, kind: 'result', document: imported.document.toJSON(), warnings: imported.warnings });
    }
    catch (error) {
        workerScope.postMessage({ requestId, kind: 'error', message: error instanceof Error ? error.message : String(error) });
    }
};
