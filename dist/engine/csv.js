/** CSV import/export: handles quoted commas, doubled quotes, LF and CRLF. */
export function parseCSV(source) {
    const rows = [];
    let row = [], field = '', quoted = false;
    for (let i = 0; i < source.length; i++) {
        const ch = source[i];
        if (quoted) {
            if (ch === '"') {
                if (source[i + 1] === '"') {
                    field += '"';
                    i++;
                }
                else
                    quoted = false;
            }
            else
                field += ch;
            continue;
        }
        if (ch === '"') {
            if (field === '')
                quoted = true;
            else
                field += ch;
        }
        else if (ch === ',') {
            row.push(field);
            field = '';
        }
        else if (ch === '\r' || ch === '\n') {
            if (ch === '\r' && source[i + 1] === '\n')
                i++;
            row.push(field);
            rows.push(row);
            row = [];
            field = '';
        }
        else
            field += ch;
    }
    if (quoted)
        throw Error('CSV contains an unterminated quoted field.');
    if (field !== '' || row.length) {
        row.push(field);
        rows.push(row);
    }
    return rows;
}
export function encodeCSV(rows) {
    const quote = (v) => /[",\r\n]/.test(v) ? `"${v.replace(/"/g, '""')}"` : v;
    return rows.map(r => r.map(quote).join(',')).join('\r\n');
}
