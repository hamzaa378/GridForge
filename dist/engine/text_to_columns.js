import { display } from './workbook.js';
import { colToName, parseRef } from './parser.js';
/** Split one field using Excel-like text qualifiers and optional delimiter merging. */
export function splitDelimitedText(input, delimiter, mergeConsecutive = false) {
    if (delimiter.length !== 1)
        throw Error('Choose a single-character delimiter.');
    const fields = [];
    let field = '', quoted = false, atFieldStart = true;
    for (let i = 0; i < input.length; i++) {
        const char = input[i];
        if (char === '"') {
            if (quoted && input[i + 1] === '"') {
                field += '"';
                i++;
                continue;
            }
            if (quoted) {
                quoted = false;
                continue;
            }
            if (atFieldStart) {
                quoted = true;
                continue;
            }
        }
        if (char === delimiter && !quoted) {
            fields.push(field);
            field = '';
            atFieldStart = true;
            if (mergeConsecutive)
                while (input[i + 1] === delimiter)
                    i++;
            continue;
        }
        field += char;
        atFieldStart = false;
    }
    if (quoted)
        throw Error('A text qualifier was opened but not closed.');
    fields.push(field);
    return fields;
}
export function textToColumnsEdits(book, sourceStart, sourceEnd, destination, delimiter, mergeConsecutive = false) {
    const from = parseRef(sourceStart), to = parseRef(sourceEnd), dest = parseRef(destination);
    if (!from || !to || !dest || from.col !== to.col || from.row > to.row || from.row < 0 || to.row >= book.rows || from.col >= book.cols || to.col >= book.cols)
        throw Error('Select a single-column source range and a valid destination cell.');
    const sourceRows = to.row - from.row + 1;
    if (dest.row + sourceRows > book.rows)
        throw Error('The split output would extend beyond the last worksheet row.');
    const values = [];
    let columns = 1;
    for (let row = from.row; row <= to.row; row++) {
        const raw = display(book.getValue(`${colToName(from.col)}${row + 1}`)), parts = splitDelimitedText(raw, delimiter, mergeConsecutive);
        values.push(parts);
        columns = Math.max(columns, parts.length);
    }
    if (dest.col + columns > book.cols)
        throw Error('The split output would extend beyond the last worksheet column.');
    const edits = [];
    for (let row = 0; row < sourceRows; row++)
        for (let col = 0; col < columns; col++)
            edits.push({ ref: `${colToName(dest.col + col)}${dest.row + row + 1}`, raw: values[row][col] ?? '' });
    return { edits, rows: sourceRows, columns, startRow: dest.row, startCol: dest.col };
}
