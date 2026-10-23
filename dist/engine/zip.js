/** Small standards-based ZIP adapter for OOXML. No third-party formula or archive runtime. */
const te = new TextEncoder();
const td = new TextDecoder();
const LIMIT_FILE = 25 * 1024 * 1024, LIMIT_TOTAL = 80 * 1024 * 1024;
const crcTable = new Uint32Array(256);
for (let n = 0; n < 256; n++) {
    let x = n;
    for (let i = 0; i < 8; i++)
        x = (x & 1) ? 0xedb88320 ^ (x >>> 1) : x >>> 1;
    crcTable[n] = x >>> 0;
}
function crc(data) { let n = 0xffffffff; for (const b of data)
    n = crcTable[(n ^ b) & 255] ^ (n >>> 8); return (n ^ 0xffffffff) >>> 0; }
function u16(buf, offset) { return (buf[offset] ?? 0) | ((buf[offset + 1] ?? 0) << 8); }
function u32(buf, offset) { return (u16(buf, offset) + (u16(buf, offset + 2) * 65536)) >>> 0; }
function set16(b, o, v) { b[o] = v & 255; b[o + 1] = (v >>> 8) & 255; }
function set32(b, o, v) { set16(b, o, v); set16(b, o + 2, Math.floor(v / 65536)); }
function join(chunks) { let size = 0; for (const c of chunks)
    size += c.length; const out = new Uint8Array(size); let at = 0; for (const c of chunks) {
    out.set(c, at);
    at += c.length;
} return out; }
/** Stored-entry ZIP writer: larger than deflated ZIP, universally readable by Excel/LibreOffice. */
export function zipWrite(entries) {
    return zipWriteBytes(Object.fromEntries(Object.entries(entries).map(([name, value]) => [name, te.encode(value)])));
}
/** Repackage all OOXML parts, including binary VBA projects, signatures and media. */
export function zipWriteBytes(entries) {
    const chunks = [], directory = [];
    let offset = 0;
    for (const [name, content] of Object.entries(entries)) {
        if (!name || name.includes('..') || name.startsWith('/'))
            throw Error('Invalid ZIP part name');
        const filename = te.encode(name), bytes = content;
        if (bytes.length > LIMIT_FILE)
            throw Error('ZIP part too large');
        const hash = crc(bytes), local = new Uint8Array(30 + filename.length);
        set32(local, 0, 0x04034b50);
        set16(local, 4, 20);
        set16(local, 6, 0x0800);
        set16(local, 8, 0);
        set32(local, 14, hash);
        set32(local, 18, bytes.length);
        set32(local, 22, bytes.length);
        set16(local, 26, filename.length);
        local.set(filename, 30);
        chunks.push(local, bytes);
        const central = new Uint8Array(46 + filename.length);
        set32(central, 0, 0x02014b50);
        set16(central, 4, 20);
        set16(central, 6, 20);
        set16(central, 8, 0x0800);
        set16(central, 10, 0);
        set32(central, 16, hash);
        set32(central, 20, bytes.length);
        set32(central, 24, bytes.length);
        set16(central, 28, filename.length);
        set32(central, 42, offset);
        central.set(filename, 46);
        directory.push(central);
        offset += local.length + bytes.length;
    }
    const dirSize = directory.reduce((a, b) => a + b.length, 0), end = new Uint8Array(22);
    set32(end, 0, 0x06054b50);
    set16(end, 8, directory.length);
    set16(end, 10, directory.length);
    set32(end, 12, dirSize);
    set32(end, 16, offset);
    return join([...chunks, ...directory, end]);
}
export async function zipRead(data) {
    const entries = await zipReadBytes(data), out = new Map();
    for (const [name, bytes] of entries) {
        if (name.endsWith('.xml') || name.endsWith('.rels'))
            out.set(name, td.decode(bytes));
    }
    return out;
}
/** Read and verify binary and XML entries without discarding VBA and unrelated features. */
export async function zipReadBytes(data) {
    let end = -1;
    for (let p = data.length - 22; p >= Math.max(0, data.length - 65557); p--)
        if (u32(data, p) === 0x06054b50) {
            end = p;
            break;
        }
    if (end < 0)
        throw Error('Not a ZIP / XLSX file.');
    const count = u16(data, end + 10), directoryOffset = u32(data, end + 16);
    if (count > 3000 || directoryOffset >= data.length)
        throw Error('Invalid or oversized XLSX archive.');
    const out = new Map();
    let at = directoryOffset, total = 0;
    for (let i = 0; i < count; i++) {
        if (u32(data, at) !== 0x02014b50)
            throw Error('Corrupt XLSX ZIP directory.');
        const method = u16(data, at + 10), expectedCrc = u32(data, at + 16), compressed = u32(data, at + 20), uncompressed = u32(data, at + 24);
        const len = u16(data, at + 28), extra = u16(data, at + 30), comment = u16(data, at + 32), offset = u32(data, at + 42);
        const name = td.decode(data.subarray(at + 46, at + 46 + len));
        at += 46 + len + extra + comment;
        if (uncompressed > LIMIT_FILE || (total += uncompressed) > LIMIT_TOTAL || name.includes('..') || name.startsWith('/'))
            throw Error('Unsafe or oversized XLSX ZIP archive.');
        if (!name || name.endsWith('/') || name.includes('\\') || name.includes('..') || name.startsWith('/'))
            throw Error('Unsafe ZIP entry name.');
        if (u32(data, offset) !== 0x04034b50)
            throw Error('Invalid XLSX local ZIP header.');
        const pos = offset + 30 + u16(data, offset + 26) + u16(data, offset + 28);
        if (pos + compressed > data.length)
            throw Error('Truncated XLSX part.');
        const src = data.subarray(pos, pos + compressed);
        let bytes;
        if (method === 0)
            bytes = src;
        else if (method === 8) {
            if (typeof DecompressionStream === 'undefined')
                throw Error('This browser does not support XLSX decompression. Use a current Chrome/Edge/Firefox browser.');
            const stream = new Blob([new Uint8Array(src)]).stream().pipeThrough(new DecompressionStream('deflate-raw'));
            bytes = new Uint8Array(await new Response(stream).arrayBuffer());
        }
        else
            throw Error('Unsupported ZIP compression method ' + method);
        if (bytes.length !== uncompressed || crc(bytes) !== expectedCrc)
            throw Error(`Invalid XLSX ZIP checksum in ${name}`);
        if (out.has(name))
            throw Error('Duplicate OOXML ZIP entry.');
        out.set(name, new Uint8Array(bytes));
    }
    return out;
}
