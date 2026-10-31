import { display } from './workbook.js';
import { colToName, shiftFormula } from './parser.js';
const addr = (r, c) => `${colToName(c)}${r + 1}`;
export const MAX_CLIPBOARD_CELLS = 50_000;
export function snapshotCells(book, area, sourceSheet, cut = false) {
    const cellCount = (area.r2 - area.r1 + 1) * (area.c2 - area.c1 + 1);
    if (!Number.isSafeInteger(cellCount) || cellCount < 1 || cellCount > MAX_CLIPBOARD_CELLS)
        throw Error(`Copy and cut are limited to ${MAX_CLIPBOARD_CELLS.toLocaleString()} cells per operation.`);
    const rows = [];
    for (let r = area.r1; r <= area.r2; r++) {
        const line = [];
        for (let c = area.c1; c <= area.c2; c++) {
            const ref = addr(r, c);
            line.push({ raw: book.getRaw(ref), value: display(book.getValue(ref)), style: book.getStyle(ref) });
        }
        rows.push(line);
    }
    return { origin: { row: area.r1, col: area.c1 }, rows, plain: rows.map(row => row.map(x => x.value).join('\t')).join('\n'), sourceSheet, cut };
}
/** Generate one atomic transaction, including a safe within-sheet cut/move operation.
 * Moved formulas keep their raw source text: external references are not rewritten.
 * Cuts are only cleared when pasted back into the same worksheet.
 */
export function pasteSnapshot(book, s, destination, sheetId, valuesOnly = false) {
    if (!Array.isArray(s.rows) || !s.rows.length || s.rows.some(row => !Array.isArray(row)))
        throw Error('Clipboard data is invalid. Copy the range again.');
    const cellCount = s.rows.reduce((count, row) => count + row.length, 0);
    if (!Number.isSafeInteger(cellCount) || cellCount > MAX_CLIPBOARD_CELLS)
        throw Error(`Paste is limited to ${MAX_CLIPBOARD_CELLS.toLocaleString()} cells per operation.`);
    if (destination.row < 0 || destination.col < 0 || destination.row + s.rows.length > book.rows || destination.col + Math.max(...s.rows.map(x => x.length)) > book.cols) {
        throw new Error('Paste destination is outside the worksheet.');
    }
    const move = s.cut && s.sourceSheet === sheetId && !valuesOnly;
    const edits = new Map(), styles = new Map();
    if (move) {
        for (let r = 0; r < s.rows.length; r++)
            for (let c = 0; c < s.rows[r].length; c++) {
                const ref = addr(s.origin.row + r, s.origin.col + c);
                edits.set(ref, '');
                styles.set(ref, {});
            }
    }
    for (let r = 0; r < s.rows.length; r++)
        for (let c = 0; c < s.rows[r].length; c++) {
            const cell = s.rows[r][c], ref = addr(destination.row + r, destination.col + c);
            let raw = valuesOnly ? cell.value : move ? cell.raw : shiftFormula(cell.raw, destination.row - (s.origin.row + r), destination.col - (s.origin.col + c));
            // Calculated text that resembles a formula must remain text in Paste Values.
            if (valuesOnly && /^[=+]/.test(raw) && Number.isNaN(Number(raw)))
                raw = "'" + raw;
            edits.set(ref, raw);
            if (!valuesOnly)
                styles.set(ref, cell.style);
        }
    return { edits: [...edits].map(([ref, raw]) => ({ ref, raw })), styles: [...styles].map(([ref, style]) => ({ ref, style })), moved: move };
}
