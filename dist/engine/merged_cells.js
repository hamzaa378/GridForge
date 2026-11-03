import { colToName, parseRef } from './parser.js';
const MAX_MERGED_CELLS = 2_000_000;
export function normalizeMergedRange(value, rows, cols) {
    const parts = value.replace(/\$/g, '').toUpperCase().split(':');
    if (parts.length !== 2)
        throw Error('A merged range must include two cell references.');
    const first = parseRef(parts[0] ?? ''), last = parseRef(parts[1] ?? '');
    if (!first || !last)
        throw Error('A merged range contains an invalid cell reference.');
    const bounds = { r1: Math.min(first.row, last.row), r2: Math.max(first.row, last.row), c1: Math.min(first.col, last.col), c2: Math.max(first.col, last.col) };
    if (bounds.r1 === bounds.r2 && bounds.c1 === bounds.c2)
        throw Error('Select at least two cells to merge.');
    if (bounds.r2 >= rows || bounds.c2 >= cols)
        throw Error('The merged range exceeds the worksheet.');
    return `${colToName(bounds.c1)}${bounds.r1 + 1}:${colToName(bounds.c2)}${bounds.r2 + 1}`;
}
export function mergedRangeBounds(value) {
    const parts = value.split(':'), first = parseRef(parts[0] ?? ''), last = parseRef(parts[1] ?? '');
    return first && last ? { r1: Math.min(first.row, last.row), r2: Math.max(first.row, last.row), c1: Math.min(first.col, last.col), c2: Math.max(first.col, last.col) } : null;
}
export function mergeContaining(row, column, ranges) {
    for (const range of ranges) {
        const bounds = mergedRangeBounds(range);
        if (bounds && row >= bounds.r1 && row <= bounds.r2 && column >= bounds.c1 && column <= bounds.c2)
            return bounds;
    }
    return null;
}
export function validateMergedRanges(value, rows, cols) {
    if (value === undefined)
        return [];
    if (!Array.isArray(value) || value.length > 2000)
        throw Error('A worksheet can contain at most 2,000 merged regions.');
    const ranges = value.map(item => { if (typeof item !== 'string')
        throw Error('Invalid merged range.'); return normalizeMergedRange(item, rows, cols); });
    if (new Set(ranges).size !== ranges.length)
        throw Error('Duplicate merged regions are not allowed.');
    const bounds = ranges.map(range => mergedRangeBounds(range)).sort((a, b) => a.r1 - b.r1 || a.c1 - b.c1);
    let covered = 0;
    for (const range of bounds) {
        covered += (range.r2 - range.r1 + 1) * (range.c2 - range.c1 + 1);
        if (covered > MAX_MERGED_CELLS)
            throw Error('Merged regions may cover at most 2,000,000 cells per worksheet.');
    }
    for (let i = 0; i < bounds.length; i++)
        for (let j = i + 1; j < bounds.length && bounds[j].r1 <= bounds[i].r2; j++)
            if (bounds[j].c1 <= bounds[i].c2 && bounds[j].c2 >= bounds[i].c1)
                throw Error('Merged regions cannot overlap.');
    return ranges.sort((a, b) => { const x = mergedRangeBounds(a), y = mergedRangeBounds(b); return x.r1 - y.r1 || x.c1 - y.c1; });
}
export function addMergedRange(existing, value, rows, cols) {
    if (existing.length >= 2000)
        throw Error('A worksheet can contain at most 2,000 merged regions.');
    const candidate = normalizeMergedRange(value, rows, cols), bounds = mergedRangeBounds(candidate);
    for (const range of existing) {
        const prior = mergedRangeBounds(range);
        if (prior && bounds.r1 <= prior.r2 && bounds.r2 >= prior.r1 && bounds.c1 <= prior.c2 && bounds.c2 >= prior.c1)
            throw Error('Merged regions cannot overlap.');
    }
    const coverage = existing.reduce((total, range) => { const area = mergedRangeBounds(range); return total + (area.r2 - area.r1 + 1) * (area.c2 - area.c1 + 1); }, (bounds.r2 - bounds.r1 + 1) * (bounds.c2 - bounds.c1 + 1));
    if (coverage > MAX_MERGED_CELLS)
        throw Error('Merged regions may cover at most 2,000,000 cells per worksheet.');
    const result = [...existing, candidate];
    return result.sort((left, right) => { const a = mergedRangeBounds(left), b = mergedRangeBounds(right); return a.r1 - b.r1 || a.c1 - b.c1; });
}
