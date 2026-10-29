import { SpreadsheetDocument } from './document.js';
import { buildPivotSummary, validatePivotSummary } from './pivot.js';
/** Rebuild several summaries from one workbook snapshot and one deserialization. */
export function runPivotJobs(request) {
    try {
        const document = SpreadsheetDocument.fromJSON(request.document), destination = document.sheets.find(sheet => sheet.id === request.destinationSheetId);
        if (!destination)
            throw Error('The PivotTable destination worksheet no longer exists.');
        const results = [];
        for (const rawSpec of request.specs) {
            try {
                const spec = validatePivotSummary(rawSpec, destination.book.rows, destination.book.cols), sourceName = spec.sourceSheet ?? destination.name, source = document.sheets.find(sheet => sheet.name.toLocaleLowerCase() === sourceName.toLocaleLowerCase());
                if (!source)
                    throw Error(`Pivot source worksheet “${sourceName}” is unavailable.`);
                const result = buildPivotSummary(destination.book, spec, source.book);
                results.push({ id: spec.id, result });
                destination.book.applyEdits(result.edits, false);
                destination.book.applyStyles(result.styleEdits, false);
                const index = destination.tools.pivotTables.findIndex(item => item.id === spec.id);
                if (index >= 0)
                    destination.tools.pivotTables[index] = { ...spec, lastOutputRange: result.outputRange, lastPercentCells: result.percentCells };
            }
            catch (error) {
                results.push({ id: rawSpec.id, error: error instanceof Error ? error.message : 'Unable to calculate this Pivot Summary.' });
            }
        }
        return { requestId: request.requestId, results };
    }
    catch (error) {
        return { requestId: request.requestId, results: [], error: error instanceof Error ? error.message : 'Unable to calculate these Pivot Summaries.' };
    }
}
/** Rebuild a Pivot Summary from a JSON workbook snapshot in a background worker. */
export function runPivotJob(request) {
    try {
        const document = SpreadsheetDocument.fromJSON(request.document), destination = document.sheets.find(sheet => sheet.id === request.destinationSheetId);
        if (!destination)
            throw Error('The PivotTable destination worksheet no longer exists.');
        const spec = validatePivotSummary(request.spec, destination.book.rows, destination.book.cols), sourceName = spec.sourceSheet ?? destination.name, source = document.sheets.find(sheet => sheet.name.toLocaleLowerCase() === sourceName.toLocaleLowerCase());
        if (!source)
            throw Error(`Pivot source worksheet “${sourceName}” is unavailable.`);
        return { requestId: request.requestId, result: buildPivotSummary(destination.book, spec, source.book, request.detailRef) };
    }
    catch (error) {
        return { requestId: request.requestId, error: error instanceof Error ? error.message : 'Unable to calculate this Pivot Summary.' };
    }
}
