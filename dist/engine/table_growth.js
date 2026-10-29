/** Expand a formatted table when contiguous data is entered directly below it.
 *  Preserve pre-existing table IDs, charts, and style flags. No rows are inserted or shifted.
 *  Reject expansion into another table or beyond spreadsheet/chart limits. */
import { colToName } from './parser.js';
import { tableRange, intersectTables } from './tables.js';
import { chartRange } from './charts.js';
export function expandTablesAfterEdit(book, tools, edits) {
    const changed = new Set(edits.filter(e => e.raw.trim().length > 0).map(e => e.ref.replace(/\$/g, '').toUpperCase()));
    const results = [];
    for (const table of tools.tables) {
        const bounds = tableRange(table.range, book.rows, book.cols);
        let next = bounds.r2;
        while (next + 1 < book.rows) {
            const r = next + 1;
            // Grow only when the edit touches the next physical row in this table's columns.
            if (!Array.from({ length: bounds.c2 - bounds.c1 + 1 }, (_, i) => `${colToName(bounds.c1 + i)}${r + 1}`).some(ref => changed.has(ref)))
                break;
            const newBounds = { ...bounds, r2: r };
            if (tools.tables.some(other => other.id !== table.id && intersectTables(newBounds, tableRange(other.range, book.rows, book.cols))))
                break;
            // Require non-empty entered content on the growth row; do not extend over empty rows.
            if (!Array.from({ length: bounds.c2 - bounds.c1 + 1 }, (_, i) => book.getRaw(`${colToName(bounds.c1 + i)}${r + 1}`)).some(x => x.trim() !== ''))
                break;
            next = r;
        }
        if (next === bounds.r2)
            continue;
        const old = table.range, newRange = `${colToName(bounds.c1)}${bounds.r1 + 1}:${colToName(bounds.c2)}${next + 1}`;
        table.range = newRange;
        let updatedCharts = 0;
        for (const chart of tools.charts) {
            if (chart.range.toUpperCase() !== old.toUpperCase())
                continue;
            try {
                chartRange(newRange, book.rows, book.cols);
                chart.range = newRange;
                updatedCharts++;
            }
            catch { /* chart's 250-row/12-col limit: table may grow without chart */ }
        }
        // Existing filters track the table's new end as well.
        if (tools.filter && tools.filter.startRow === bounds.r1 && tools.filter.endRow === bounds.r2 && tools.filter.startCol === bounds.c1 && tools.filter.endCol === bounds.c2)
            tools.filter.endRow = next;
        results.push({ table: table.name, before: old, after: newRange, updatedCharts });
    }
    return results;
}
