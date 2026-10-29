/** GridForge v4: 80 additional Excel-style scalar and range functions.
 * Implemented in TypeScript without any spreadsheet formula engine.
 * Deliberate limits: no dynamic-array spill; dates use Excel's 1900 date system
 * with UTC arithmetic; unsupported advanced Excel options are documented.
 */
import { SheetError, flatten, number, text, bool, matchCriterion } from './functions.js';
const fail = (code, message) => { throw new SheetError(code, message); };
const def = (description, example, run) => ({ description, example, run });
const req = (a, min, max = min) => { if (a.length < min || a.length > max)
    fail('#ARG!', `Expected ${min}${min !== max ? ` to ${max}` : ''} arguments, got ${a.length}.`); };
const val = (a, i) => { const v = a[i]; return Array.isArray(v) ? fail('#VALUE!', 'This argument must be a single value.') : v ?? null; };
const n = (a, i) => number(val(a, i));
const integer = (a, i) => Math.trunc(n(a, i));
const nums = (a) => flatten(a).filter((x) => typeof x === 'number' && Number.isFinite(x));
const finite = (v) => Number.isFinite(v) ? v : fail('#NUM!', 'The result is not a finite number.');
const sum = (xs) => xs.reduce((x, y) => x + y, 0);
const numericRange = (v) => nums([v]);
const holidaySet = (v) => { const dates = numericRange(v); for (const date of dates)
    if (date < 0 || date > 2_958_465)
        fail('#NUM!', 'Holiday date is outside Excel’s supported range.'); return new Set(dates.map(Math.trunc)); };
const nonBlank = (v) => v !== null && v !== '';
const matrix = (v) => Array.isArray(v) ? v.map(row => row.map(c => Array.isArray(c) ? fail('#VALUE!', 'Nested array not supported.') : c)) : [[v]];
const vector = (v) => flatten([v]);
const equal = (a, b) => matchCriterion(a, b);
const paired = (a, count) => {
    req(a, count, count);
    const source = a.map(vector);
    if (source.some(x => x.length !== source[0].length))
        fail('#VALUE!', 'Data ranges must have equal lengths.');
    const results = Array.from({ length: count }, () => []);
    for (let i = 0; i < source[0].length; i++) {
        if (source.every(row => typeof row[i] === 'number' && Number.isFinite(row[i])))
            source.forEach((row, j) => results[j].push(row[i]));
    }
    if (results[0].length === 0)
        fail('#DIV/0!', 'At least one numeric pair is required.');
    return results;
};
const moments = (a) => {
    const values = nums(a);
    if (!values.length)
        fail('#DIV/0!', 'At least one numeric observation is required.');
    const avg = sum(values) / values.length;
    return { values, avg, squares: sum(values.map(x => (x - avg) ** 2)) };
};
const regression = (a) => {
    const [ys, xs] = paired(a, 2);
    if (xs.length < 2)
        fail('#DIV/0!', 'At least two observations required.');
    const mx = sum(xs) / xs.length, my = sum(ys) / ys.length;
    const sxx = sum(xs.map(x => (x - mx) ** 2)), syy = sum(ys.map(y => (y - my) ** 2));
    const sxy = sum(xs.map((x, i) => (x - mx) * (ys[i] - my)));
    return { xs: xs, ys: ys, mx, my, sxx, sxy, syy };
};
const quantile = (values, fraction) => {
    if (fraction < 0 || fraction > 1)
        fail('#NUM!', 'Percentile must be from 0 to 1.');
    if (!values.length)
        fail('#NUM!', 'The input range has no numbers.');
    const a = [...values].sort((x, y) => x - y), rank = (a.length - 1) * fraction, lo = Math.floor(rank), hi = Math.ceil(rank);
    return a[lo] + (a[hi] - a[lo]) * (rank - lo);
};
const factorial = (v) => {
    if (v < 0)
        fail('#NUM!', 'Factorial requires a nonnegative number.');
    let ans = 1;
    for (let i = 2; i <= Math.trunc(v); i++)
        ans = finite(ans * i);
    return ans;
};
const gcd = (a, b) => { a = Math.abs(Math.trunc(a)); b = Math.abs(Math.trunc(b)); while (b) {
    const r = a % b;
    a = b;
    b = r;
} return a; };
const one = (name, example, fn) => def(name, example, a => { req(a, 1); return finite(fn(n(a, 0))); });
// Excel 1900 serial dates: serial 60 is Excel's fictitious 1900-02-29.
const dayMs = 86400000, epoch = Date.UTC(1899, 11, 31);
const excelDate = (v) => { const t = v.getTime(); if (!Number.isFinite(t))
    fail('#NUM!', 'Date is outside Excel’s supported range.'); const result = (t - epoch) / dayMs + (t >= Date.UTC(1900, 2, 1) ? 1 : 0); if (result < 0 || result > 2_958_465)
    fail('#NUM!', 'Date is outside Excel’s supported range.'); return result; };
const serialDate = (serial) => {
    if (!Number.isFinite(serial) || serial < 0 || serial > 2_958_465)
        fail('#NUM!', 'Date serial must be within Excel’s supported range.');
    return new Date(epoch + (Math.floor(serial) - (Math.floor(serial) >= 60 ? 1 : 0)) * dayMs);
};
const excelWeekday = (serial) => Math.floor(serial) === 60 ? 3 : serialDate(serial).getUTCDay();
const dateSerial = (a, i) => n(a, i);
const dayDate = (serial) => serialDate(serial);
const monthShift = (serial, offset, end = false) => {
    const d = dayDate(serial), origDay = Math.floor(serial) === 60 ? 29 : d.getUTCDate();
    const first = new Date(Date.UTC(d.getUTCFullYear(), d.getUTCMonth() + offset + (end ? 1 : 0), end ? 0 : 1));
    if (end)
        return first.getUTCFullYear() === 1900 && first.getUTCMonth() === 1 ? 60 : excelDate(first);
    if (Math.floor(serial) === 60 && offset === 0)
        return serial;
    if (first.getUTCFullYear() === 1900 && first.getUTCMonth() === 1 && origDay >= 29)
        return 60;
    const daysInTarget = new Date(Date.UTC(first.getUTCFullYear(), first.getUTCMonth() + 1, 0)).getUTCDate();
    first.setUTCDate(Math.min(origDay, daysInTarget));
    return excelDate(first);
};
const matchesAll = (ranges, criteria, index) => ranges.every((v, i) => matchCriterion(v[index] ?? null, criteria[i]));
const multipleCriteria = (a, hasValues) => {
    const offset = hasValues ? 1 : 0;
    if (a.length < offset + 2 || (a.length - offset) % 2 !== 0)
        fail('#ARG!', 'Expected criteria_range, criterion pairs.');
    const ranges = [], criteria = [];
    for (let i = offset; i < a.length; i += 2) {
        ranges.push(vector(a[i]));
        criteria.push(val(a, i + 1));
    }
    const length = ranges[0].length;
    if (ranges.some(r => r.length !== length))
        fail('#VALUE!', 'Criteria ranges must have the same number of cells.');
    const values = hasValues ? vector(a[0]) : ranges[0];
    if (values.length !== length)
        fail('#VALUE!', 'Result and criteria ranges must be the same size.');
    return { values, ranges, criteria };
};
function lookupMatch(key, lookup) { return vector(lookup).findIndex(v => equal(v, key)); }
const lookupOrder = (left, right) => {
    const rank = (x) => x === null ? 0 : typeof x === 'number' ? 1 : typeof x === 'string' ? 2 : 3;
    const lr = rank(left), rr = rank(right);
    if (lr !== rr)
        return lr - rr;
    if (typeof left === 'number' && typeof right === 'number')
        return left < right ? -1 : left > right ? 1 : 0;
    if (typeof left === 'boolean' && typeof right === 'boolean')
        return Number(left) - Number(right);
    const a = text(left).toLowerCase(), b = text(right).toLowerCase();
    return a < b ? -1 : a > b ? 1 : 0;
};
const lookupEqual = (left, right) => typeof left === 'number' && typeof right === 'number' ? left === right : typeof left === 'boolean' && typeof right === 'boolean' ? left === right : text(left).toLowerCase() === text(right).toLowerCase();
const lookupWildcard = (pattern) => {
    let source = '^';
    for (let i = 0; i < pattern.length; i++) {
        let ch = pattern[i];
        if (ch === '~' && i + 1 < pattern.length) {
            ch = pattern[++i];
            source += ch.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
        }
        else if (ch === '*')
            source += '.*';
        else if (ch === '?')
            source += '.';
        else
            source += ch.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    }
    return new RegExp(source + '$', 'i');
};
const textDelimiterPositions = (value, delimiter, ignoreCase, matchEnd) => {
    const positions = [];
    if (ignoreCase) {
        const needle = delimiter.toLocaleLowerCase();
        for (let index = 0; index <= value.length - delimiter.length;) {
            if (value.slice(index, index + delimiter.length).toLocaleLowerCase() === needle) {
                positions.push(index);
                index += delimiter.length;
            }
            else
                index++;
        }
    }
    else
        for (let index = value.indexOf(delimiter); index >= 0; index = value.indexOf(delimiter, index + delimiter.length))
            positions.push(index);
    const ends = value.slice(-delimiter.length);
    if (matchEnd && !(ignoreCase ? ends.toLocaleLowerCase() === delimiter.toLocaleLowerCase() : ends === delimiter))
        positions.push(value.length);
    return positions;
};
const pmt = (rate, nper, pv, fv = 0, type = 0) => {
    if (nper <= 0 || !Number.isFinite(nper))
        fail('#NUM!', 'Number of periods must be positive.');
    if (rate === 0)
        return -(pv + fv) / nper;
    const growth = (1 + rate) ** nper;
    return finite(-(pv * growth + fv) * rate / ((1 + rate * type) * (growth - 1)));
};
export const moreFunctions = {
    // 24 mathematical, trigonometric and combinatorial functions
    SUMSQ: def('Sum of squares of all numeric arguments.', '=SUMSQ(A1:A4)', a => finite(sum(nums(a).map(x => x * x)))),
    SUMPRODUCT: def('Sum of products at matching positions in equally shaped arrays.', '=SUMPRODUCT(A1:A4,B1:B4)', a => {
        req(a, 1, 30);
        const arrays = a.map(matrix), rows = arrays[0].length, cols = arrays[0][0].length;
        if (arrays.some(m => m.length !== rows || m.some(row => row.length !== cols)))
            fail('#VALUE!', 'SUMPRODUCT arrays must have identical dimensions.');
        let total = 0;
        for (let r = 0; r < rows; r++)
            for (let c = 0; c < cols; c++)
                total += arrays.reduce((p, m) => p * (typeof m[r][c] === 'number' ? m[r][c] : 0), 1);
        return finite(total);
    }),
    QUOTIENT: def('Integer portion of division truncated toward zero.', '=QUOTIENT(17,5)', a => { req(a, 2); const d = n(a, 1); if (d === 0)
        fail('#DIV/0!', 'Division by zero.'); return Math.trunc(n(a, 0) / d); }),
    TRUNC: def('Truncate number at requested decimal places.', '=TRUNC(3.987,2)', a => { req(a, 1, 2); const p = 10 ** (a.length === 2 ? integer(a, 1) : 0); return finite(Math.trunc(n(a, 0) * p) / p); }),
    EVEN: one('Round away from zero to the nearest even integer.', '=EVEN(3)', x => Math.sign(x) * Math.ceil(Math.abs(x) / 2) * 2),
    ODD: one('Round away from zero to the nearest odd integer.', '=ODD(4)', x => x === 0 ? 1 : Math.sign(x) * (Math.ceil((Math.abs(x) - 1) / 2) * 2 + 1)),
    GCD: def('Greatest common divisor of nonnegative integers.', '=GCD(24,36)', a => { const x = flatten(a).map(number); if (x.some(v => v < 0))
        fail('#NUM!', 'GCD requires nonnegative arguments.'); return x.reduce(gcd, 0); }),
    LCM: def('Least common multiple of nonnegative integers.', '=LCM(4,6)', a => { const x = flatten(a).map(number); if (x.some(v => v < 0))
        fail('#NUM!', 'LCM requires nonnegative arguments.'); return finite(x.reduce((acc, v) => acc === 0 || v === 0 ? 0 : Math.abs(acc * Math.trunc(v) / gcd(acc, v)), 1)); }),
    FACT: one('Factorial of a nonnegative integer.', '=FACT(5)', factorial),
    COMBIN: def('Number of combinations without repetition.', '=COMBIN(5,2)', a => { req(a, 2); const x = integer(a, 0), k = integer(a, 1); if (x < 0 || k < 0 || k > x)
        fail('#NUM!', 'COMBIN expects 0 ≤ k ≤ n.'); let result = 1; for (let i = 1; i <= k; i++)
        result = result * (x - k + i) / i; return Math.round(finite(result)); }),
    PERMUT: def('Permutations without repetition.', '=PERMUT(5,2)', a => { req(a, 2); const x = integer(a, 0), k = integer(a, 1); if (x < 0 || k < 0 || k > x)
        fail('#NUM!', 'PERMUT expects 0 ≤ k ≤ n.'); let result = 1; for (let i = 0; i < k; i++)
        result = finite(result * (x - i)); return result; }),
    RADIANS: one('Convert degrees to radians.', '=RADIANS(180)', x => x * Math.PI / 180),
    DEGREES: one('Convert radians to degrees.', '=DEGREES(PI())', x => x * 180 / Math.PI),
    SIN: one('Sine of an angle in radians.', '=SIN(PI()/2)', Math.sin),
    COS: one('Cosine of an angle in radians.', '=COS(0)', Math.cos),
    TAN: one('Tangent of an angle in radians.', '=TAN(0)', Math.tan),
    ASIN: one('Inverse sine in radians.', '=ASIN(1)', Math.asin),
    ACOS: one('Inverse cosine in radians.', '=ACOS(1)', Math.acos),
    ATAN: one('Inverse tangent in radians.', '=ATAN(1)', Math.atan),
    ATAN2: def('Angle from origin to point (x,y) in radians.', '=ATAN2(1,1)', a => { req(a, 2); if (n(a, 0) === 0 && n(a, 1) === 0)
        fail('#DIV/0!', 'ATAN2(0,0) is undefined.'); return Math.atan2(n(a, 1), n(a, 0)); }),
    SINH: one('Hyperbolic sine.', '=SINH(1)', Math.sinh),
    COSH: one('Hyperbolic cosine.', '=COSH(0)', Math.cosh),
    TANH: one('Hyperbolic tangent.', '=TANH(1)', Math.tanh),
    LOG10: one('Base 10 logarithm.', '=LOG10(100)', Math.log10),
    // 16 statistics and regression functions
    GEOMEAN: def('Geometric mean of positive numeric values.', '=GEOMEAN(A1:A5)', a => { const x = nums(a); if (!x.length || x.some(v => v <= 0))
        fail('#NUM!', 'GEOMEAN requires positive values.'); return Math.exp(sum(x.map(Math.log)) / x.length); }),
    HARMEAN: def('Harmonic mean of positive numeric values.', '=HARMEAN(A1:A5)', a => { const x = nums(a); if (!x.length || x.some(v => v <= 0))
        fail('#NUM!', 'HARMEAN requires positive values.'); return x.length / sum(x.map(v => 1 / v)); }),
    STDEVP: def('Population standard deviation.', '=STDEVP(A1:A8)', a => { const m = moments(a); return Math.sqrt(m.squares / m.values.length); }),
    VARP: def('Population variance.', '=VARP(A1:A8)', a => { const m = moments(a); return m.squares / m.values.length; }),
    DEVSQ: def('Sum of squared deviations from the mean.', '=DEVSQ(A1:A8)', a => moments(a).squares),
    AVEDEV: def('Average absolute deviation from mean.', '=AVEDEV(A1:A8)', a => { const m = moments(a); return sum(m.values.map(x => Math.abs(x - m.avg))) / m.values.length; }),
    MODE: def('Most frequent numeric value, lowest value breaks ties.', '=MODE(A1:A8)', a => { const x = nums(a).sort((p, q) => p - q); const counts = new Map(); for (const v of x)
        counts.set(v, (counts.get(v) ?? 0) + 1); const best = Math.max(0, ...counts.values()); if (best < 2)
        fail('#N/A', 'MODE found no repeated numeric value.'); return [...counts].find(([, count]) => count === best)[0]; }),
    PERCENTILE: def('Inclusive interpolated percentile, k between 0 and 1.', '=PERCENTILE(A1:A8,0.75)', a => { req(a, 2); return quantile(numericRange(a[0]), n(a, 1)); }),
    QUARTILE: def('Inclusive quartile 0 to 4.', '=QUARTILE(A1:A8,3)', a => { req(a, 2); const k = integer(a, 1); if (k < 0 || k > 4)
        fail('#NUM!', 'QUARTILE index must be 0–4.'); return quantile(numericRange(a[0]), k / 4); }),
    RANK: def('Position of number in numeric range, descending by default.', '=RANK(7,A1:A8,0)', a => { req(a, 2, 3); const x = n(a, 0), arr = numericRange(a[1]); if (!arr.includes(x))
        fail('#N/A', 'The number is not in the reference range.'); const ascending = a.length === 3 && n(a, 2) !== 0; return 1 + arr.filter(v => ascending ? v < x : v > x).length; }),
    CORREL: def('Pearson correlation coefficient for paired ranges.', '=CORREL(A1:A6,B1:B6)', a => { const r = regression(a); if (!r.sxx || !r.syy)
        fail('#DIV/0!', 'Correlation requires nonconstant ranges.'); return r.sxy / Math.sqrt(r.sxx * r.syy); }),
    RSQ: def('Square of Pearson correlation coefficient.', '=RSQ(A1:A6,B1:B6)', a => { const r = regression(a); if (!r.sxx || !r.syy)
        fail('#DIV/0!', 'RSQ requires nonconstant ranges.'); return (r.sxy * r.sxy) / (r.sxx * r.syy); }),
    SLOPE: def('Least squares regression slope.', '=SLOPE(B1:B6,A1:A6)', a => { const r = regression(a); if (r.sxx === 0)
        fail('#DIV/0!', 'SLOPE cannot use constant x values.'); return r.sxy / r.sxx; }),
    INTERCEPT: def('Least squares regression intercept.', '=INTERCEPT(B1:B6,A1:A6)', a => { const r = regression(a); if (r.sxx === 0)
        fail('#DIV/0!', 'INTERCEPT cannot use constant x values.'); return r.my - r.mx * r.sxy / r.sxx; }),
    FORECAST: def('Predict y using least squares at a specified x.', '=FORECAST(7,B1:B6,A1:A6)', a => { req(a, 3); const r = regression(a.slice(1)); if (r.sxx === 0)
        fail('#DIV/0!', 'FORECAST cannot use constant x values.'); return r.my + (n(a, 0) - r.mx) * r.sxy / r.sxx; }),
    PERCENTRANK: def('Inclusive percentile rank by interpolation.', '=PERCENTRANK(A1:A8,4)', a => { req(a, 2, 3); const x = n(a, 1), values = [...numericRange(a[0])].sort((p, q) => p - q); if (values.length < 2 || x < values[0] || x > values[values.length - 1])
        fail('#N/A', 'PERCENTRANK value outside range.'); let rank = values.indexOf(x) >= 0 ? values.indexOf(x) / (values.length - 1) : 0; for (let i = 0; i < values.length - 1 && values.indexOf(x) === -1; i++) {
        if (values[i] !== values[i + 1] && x >= values[i] && x <= values[i + 1]) {
            rank = (i + (x - values[i]) / (values[i + 1] - values[i])) / (values.length - 1);
            break;
        }
        if (x === values[values.length - 1])
            rank = 1;
    } const digits = a.length === 3 ? integer(a, 2) : 3; if (digits < 1 || digits > 15)
        fail('#NUM!', 'Digits must be between 1 and 15.'); const factor = 10 ** digits; return Math.floor(rank * factor) / factor; }),
    // 12 text and character functions
    PROPER: def('Capitalize the first letter of each word.', '=PROPER("hello world")', a => { req(a, 1); return text(val(a, 0)).toLowerCase().replace(/(^|[^\p{L}\p{N}])\p{L}/gu, m => m.toUpperCase()); }),
    REPT: def('Repeat text an integer number of times.', '=REPT("*",5)', a => { req(a, 2); const count = integer(a, 1); if (count < 0 || count > 32767)
        fail('#VALUE!', 'Repeat count must be 0–32767.'); const src = text(val(a, 0)); if (src.length * count > 32767)
        fail('#VALUE!', 'Result would exceed Excel text limit.'); return src.repeat(count); }),
    CLEAN: def('Remove ASCII control characters 0 through 31.', '=CLEAN(A1)', a => { req(a, 1); return text(val(a, 0)).replace(/[\x00-\x1F]/g, ''); }),
    TEXTJOIN: def('Join text or ranges with delimiter, optionally ignoring blanks.', '=TEXTJOIN(", ",TRUE,A1:A4)', a => { req(a, 3, 256); const separator = text(val(a, 0)), ignore = bool(val(a, 1)); const items = flatten(a.slice(2)); const out = items.filter(v => !ignore || nonBlank(v)).map(text).join(separator); if (out.length > 32767)
        fail('#VALUE!', 'Joined text exceeds Excel length limit.'); return out; }),
    CONCATENATE: def('Join multiple scalar strings.', '=CONCATENATE(A1," ",B1)', a => { req(a, 1, 255); return a.map(v => Array.isArray(v) ? fail('#VALUE!', 'CONCATENATE does not accept ranges.') : text(v)).join(''); }),
    CODE: def('ASCII/Unicode code point of the first character.', '=CODE("A")', a => { req(a, 1); const s = text(val(a, 0)); if (!s)
        fail('#VALUE!', 'CODE needs text.'); return s.charCodeAt(0); }),
    CHAR: def('Character from code 1 through 255.', '=CHAR(65)', a => { req(a, 1); const code = integer(a, 0); if (code < 1 || code > 255)
        fail('#VALUE!', 'CHAR requires a code from 1 to 255.'); return String.fromCharCode(code); }),
    UNICODE: def('Unicode code point of first character.', '=UNICODE("é")', a => { req(a, 1); const s = text(val(a, 0)); if (!s)
        fail('#VALUE!', 'UNICODE needs text.'); return s.codePointAt(0); }),
    UNICHAR: def('Character from Unicode code point.', '=UNICHAR(9733)', a => { req(a, 1); const cp = integer(a, 0); if (cp < 1 || cp > 0x10ffff || (cp >= 0xd800 && cp <= 0xdfff))
        fail('#VALUE!', 'Invalid Unicode code point.'); return String.fromCodePoint(cp); }),
    NUMBERVALUE: def('Parse a number with selectable decimal and group separators.', '=NUMBERVALUE("1.234,5",",",".")', a => {
        req(a, 1, 3);
        const raw = text(val(a, 0)), decimalArg = a.length >= 2 ? text(val(a, 1)) : '.', groupArg = a.length === 3 ? text(val(a, 2)) : ',';
        // Excel uses only the first character of each separator argument.
        const decimal = decimalArg[0] ?? '', group = groupArg[0] ?? '';
        if (!decimal || (group && decimal === group))
            fail('#VALUE!', 'Invalid decimal or grouping separator.');
        // Excel ignores ordinary spaces throughout the number and applies each trailing %.
        let input = raw.replace(/ /g, '');
        if (input === '')
            return 0;
        let percentCount = 0;
        while (input.endsWith('%')) {
            percentCount++;
            input = input.slice(0, -1);
        }
        if (input.includes('%'))
            fail('#VALUE!', 'Percent signs must be at the end of the text.');
        const exponentParts = input.split(/[eE]/);
        if (exponentParts.length > 2)
            fail('#VALUE!', 'NUMBERVALUE could not parse the text.');
        let mantissa = exponentParts[0];
        const exponent = exponentParts.length === 2 ? `e${exponentParts[1]}` : '';
        const decimalIndex = mantissa.indexOf(decimal);
        if (decimalIndex >= 0 && mantissa.indexOf(decimal, decimalIndex + decimal.length) >= 0)
            fail('#VALUE!', 'NUMBERVALUE could not parse the text.');
        // Group separators are ignored before the decimal point; after it Excel returns #VALUE!.
        const integerPart = decimalIndex < 0 ? mantissa : mantissa.slice(0, decimalIndex);
        const fractionPart = decimalIndex < 0 ? '' : mantissa.slice(decimalIndex + decimal.length);
        if (group && fractionPart.includes(group))
            fail('#VALUE!', 'Grouping separator cannot follow the decimal separator.');
        mantissa = integerPart.split(group).join('') + (decimalIndex < 0 ? '' : '.' + fractionPart);
        const cleaned = mantissa + exponent;
        if (!/^[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:e[+-]?\d+)?$/i.test(cleaned))
            fail('#VALUE!', 'NUMBERVALUE could not parse the text.');
        let result = Number(cleaned);
        for (let i = 0; i < percentCount; i++)
            result /= 100;
        return finite(result);
    }),
    TEXTBEFORE: def('Text before a delimiter with positive/negative instance, case mode, end matching and fallback.', '=TEXTBEFORE("a-b-c","-",-1)', a => { req(a, 2, 6); const str = text(val(a, 0)), sep = text(val(a, 1)), instance = a.length >= 3 ? integer(a, 2) : 1, matchMode = a.length >= 4 ? n(a, 3) : 0, matchEnd = a.length >= 5 ? n(a, 4) : 0; if (str.length > 32767 || sep.length > 32767)
        fail('#VALUE!', 'TEXTBEFORE text and delimiter are limited to 32,767 characters.'); if (instance === 0)
        fail('#VALUE!', 'Instance must be nonzero.'); if (matchMode !== 0 && matchMode !== 1)
        fail('#VALUE!', 'match_mode must be 0 or 1.'); if (matchEnd !== 0 && matchEnd !== 1)
        fail('#VALUE!', 'match_end must be 0 or 1.'); if (!str)
        return ''; if (!sep)
        return instance > 0 ? '' : str; const positions = textDelimiterPositions(str, sep, matchMode === 1, matchEnd === 1), index = instance > 0 ? instance - 1 : positions.length + instance; if (index < 0 || index >= positions.length)
        return a.length >= 6 ? val(a, 5) : fail('#N/A', 'TEXTBEFORE delimiter instance not found.'); return str.slice(0, positions[index]); }),
    TEXTAFTER: def('Text after a delimiter with positive/negative instance, case mode, end matching and fallback.', '=TEXTAFTER("a-b-c","-",-1)', a => { req(a, 2, 6); const str = text(val(a, 0)), sep = text(val(a, 1)), instance = a.length >= 3 ? integer(a, 2) : 1, matchMode = a.length >= 4 ? n(a, 3) : 0, matchEnd = a.length >= 5 ? n(a, 4) : 0; if (str.length > 32767 || sep.length > 32767)
        fail('#VALUE!', 'TEXTAFTER text and delimiter are limited to 32,767 characters.'); if (instance === 0)
        fail('#VALUE!', 'Instance must be nonzero.'); if (matchMode !== 0 && matchMode !== 1)
        fail('#VALUE!', 'match_mode must be 0 or 1.'); if (matchEnd !== 0 && matchEnd !== 1)
        fail('#VALUE!', 'match_end must be 0 or 1.'); if (!str)
        return ''; if (!sep)
        return instance > 0 ? str : ''; const positions = textDelimiterPositions(str, sep, matchMode === 1, matchEnd === 1), index = instance > 0 ? instance - 1 : positions.length + instance; if (index < 0 || index >= positions.length)
        return a.length >= 6 ? val(a, 5) : fail('#N/A', 'TEXTAFTER delimiter instance not found.'); return str.slice(positions[index] + sep.length); }),
    // 16 Excel 1900 date-system scalar functions (UTC, no locale-dependent parsing)
    DATE: def('Construct an Excel date serial from year, month and day.', '=DATE(2026,10,20)', a => { req(a, 3); let year = integer(a, 0); if (year >= 0 && year < 1900)
        year += 1900; const month = integer(a, 1), day = integer(a, 2); if (year === 1900 && month === 2 && day === 29)
        return 60; const d = new Date(Date.UTC(year, month - 1, day)); if (d.getUTCFullYear() < 1900 || d.getUTCFullYear() > 9999)
        fail('#NUM!', 'DATE year outside 1900–9999.'); return excelDate(d); }),
    YEAR: def('Year of a date serial.', '=YEAR(DATE(2026,10,20))', a => { req(a, 1); return Math.floor(dateSerial(a, 0)) === 0 ? 1900 : dayDate(dateSerial(a, 0)).getUTCFullYear(); }),
    MONTH: def('Month of a date serial.', '=MONTH(DATE(2026,10,20))', a => { req(a, 1); return Math.floor(dateSerial(a, 0)) === 60 ? 2 : dayDate(dateSerial(a, 0)).getUTCMonth() + 1; }),
    DAY: def('Day of month for date serial.', '=DAY(DATE(2026,10,20))', a => { req(a, 1); return Math.floor(dateSerial(a, 0)) === 60 ? 29 : dayDate(dateSerial(a, 0)).getUTCDate(); }),
    TODAY: def('Current UTC calendar day as Excel date serial (refreshes on recalculation).', '=TODAY()', a => { req(a, 0); const d = new Date(); return excelDate(new Date(Date.UTC(d.getUTCFullYear(), d.getUTCMonth(), d.getUTCDate()))); }),
    NOW: def('Current UTC date and time as Excel serial (refreshes on recalculation).', '=NOW()', a => { req(a, 0); return excelDate(new Date()); }),
    TIME: def('Time fraction of day from hours, minutes and seconds.', '=TIME(14,30,0)', a => { req(a, 3); if ([0, 1, 2].some(i => n(a, i) < 0))
        fail('#NUM!', 'TIME inputs must be nonnegative.'); return ((integer(a, 0) * 3600 + integer(a, 1) * 60 + integer(a, 2)) % 86400) / 86400; }),
    HOUR: def('Hour component of a serial time.', '=HOUR(TIME(14,30,0))', a => { req(a, 1); const x = n(a, 0); if (x < 0)
        fail('#NUM!', 'Time serial must be nonnegative.'); return Math.floor((x - Math.floor(x)) * 24 + 1e-9); }),
    MINUTE: def('Minute component of a serial time.', '=MINUTE(TIME(14,30,0))', a => { req(a, 1); const x = n(a, 0); if (x < 0)
        fail('#NUM!', 'Time serial must be nonnegative.'); return Math.floor(((x - Math.floor(x)) * 1440) % 60 + 1e-8); }),
    SECOND: def('Second component of a serial time.', '=SECOND(TIME(14,30,45))', a => { req(a, 1); const x = n(a, 0); if (x < 0)
        fail('#NUM!', 'Time serial must be nonnegative.'); return Math.floor(((x - Math.floor(x)) * 86400) % 60 + 1e-6); }),
    DAYS: def('Number of days between date serials.', '=DAYS(B1,A1)', a => { req(a, 2); return Math.trunc(n(a, 0)) - Math.trunc(n(a, 1)); }),
    EDATE: def('Shift date by whole months, keeping day when possible.', '=EDATE(A1,3)', a => { req(a, 2); return monthShift(n(a, 0), integer(a, 1)); }),
    EOMONTH: def('Last day of month shifted by whole months.', '=EOMONTH(A1,1)', a => { req(a, 2); return monthShift(n(a, 0), integer(a, 1), true); }),
    WEEKDAY: def('Day of week using Excel return type 1, 2 or 3.', '=WEEKDAY(A1,2)', a => { req(a, 1, 2); const serial = n(a, 0), w = excelWeekday(serial), type = a.length === 2 ? integer(a, 1) : 1; if (type === 1)
        return w + 1; if (type === 2)
        return (w + 6) % 7 + 1; if (type === 3)
        return (w + 6) % 7; if (type >= 11 && type <= 17)
        return (w - (type - 11 + 1) + 7) % 7 + 1; return fail('#NUM!', 'WEEKDAY supports return types 1–3 and 11–17.'); }),
    NETWORKDAYS: def('Count weekdays inclusively, excluding optional holidays.', '=NETWORKDAYS(A1,B1)', a => { req(a, 2, 3); let start = integer(a, 0), end = integer(a, 1), sign = 1; if (start < 0 || end < 0 || start > 2_958_465 || end > 2_958_465)
        fail('#NUM!', 'Date serial must be within Excel’s supported range.'); if (start > end) {
        [start, end] = [end, start];
        sign = -1;
    } if (end - start > 365250)
        fail('#NUM!', 'Date interval too large.'); const holidays = a.length === 3 ? holidaySet(a[2]) : new Set(); let total = 0; for (let d = start; d <= end; d++) {
        const day = excelWeekday(d);
        if (day !== 0 && day !== 6 && !holidays.has(d))
            total++;
    } return total * sign; }),
    WORKDAY: def('Move a number of weekdays, skipping optional holidays.', '=WORKDAY(A1,10)', a => { req(a, 2, 3); const start = integer(a, 0), steps = integer(a, 1), holidays = a.length === 3 ? holidaySet(a[2]) : new Set(); if (start < 0 || start > 2_958_465)
        fail('#NUM!', 'Date serial must be within Excel’s supported range.'); if (Math.abs(steps) > 100000)
        fail('#NUM!', 'Offset too large.'); let serial = start, left = Math.abs(steps), iterations = 0; while (left > 0) {
        serial += Math.sign(steps);
        if (serial < 0 || serial > 2_958_465)
            fail('#NUM!', 'Result date is outside Excel’s supported range.');
        const day = excelWeekday(serial);
        if (day !== 0 && day !== 6 && !holidays.has(serial))
            left--;
        if (++iterations > 200000)
            fail('#NUM!', 'WORKDAY too many iterations.');
    } return serial; }),
    // 8 logical / conditional aggregations / lookup functions
    XOR: def('TRUE for an odd number of truthy arguments.', '=XOR(TRUE,FALSE,TRUE)', a => flatten(a).reduce((acc, v) => acc + (bool(v) ? 1 : 0), 0) % 2 === 1),
    IFS: def('First true condition and corresponding value, lazily evaluated.', '=IFS(A1>90,"A",A1>80,"B",TRUE,"C")', () => fail('#INTERNAL!', 'IFS must be evaluated lazily.')),
    SWITCH: def('Match expression against cases, with optional default, lazily.', '=SWITCH(A1,1,"one",2,"two","other")', () => fail('#INTERNAL!', 'SWITCH must be evaluated lazily.')),
    SUMIFS: def('Sum values matching multiple criteria ranges.', '=SUMIFS(C1:C5,A1:A5,"North",B1:B5,">3")', a => { const x = multipleCriteria(a, true); return finite(sum(x.values.filter((v, i) => matchesAll(x.ranges, x.criteria, i) && typeof v === 'number'))); }),
    COUNTIFS: def('Count rows satisfying all conditions.', '=COUNTIFS(A1:A5,"North",B1:B5,">3")', a => { const x = multipleCriteria(a, false); return x.values.filter((_, i) => matchesAll(x.ranges, x.criteria, i)).length; }),
    AVERAGEIFS: def('Average numeric values matching all criteria ranges.', '=AVERAGEIFS(C1:C5,A1:A5,"North")', a => { const x = multipleCriteria(a, true); const values = x.values.filter((v, i) => matchesAll(x.ranges, x.criteria, i) && typeof v === 'number'); if (!values.length)
        fail('#DIV/0!', 'No matching numeric values.'); return sum(values) / values.length; }),
    XLOOKUP: def('Exact, wildcard or approximate one-dimensional lookup with row or column array returns.', '=XLOOKUP(E1,A1:A8,B1:C8,"missing",0,1)', a => {
        req(a, 3, 6);
        const lookupMatrix = matrix(a[1]), returnMatrix = matrix(a[2]), lookup = vector(a[1]);
        if (!lookupMatrix.length || !lookupMatrix[0].length || !returnMatrix.length || !returnMatrix[0].length)
            fail('#VALUE!', 'XLOOKUP arrays must not be empty.');
        const vertical = lookupMatrix.length > 1 && lookupMatrix[0].length === 1, horizontal = lookupMatrix.length === 1 && lookupMatrix[0].length > 1;
        if (lookupMatrix.some(row => row.length !== lookupMatrix[0].length) || lookupMatrix.length > 1 && lookupMatrix[0].length > 1)
            fail('#VALUE!', 'XLOOKUP lookup_array must be a single row or column.');
        if (vertical ? returnMatrix.length !== lookup.length : horizontal ? returnMatrix.some(row => row.length !== lookup.length) : returnMatrix.length !== 1 && returnMatrix[0].length !== 1)
            fail('#VALUE!', 'XLOOKUP return_array must align with lookup_array.');
        if (returnMatrix.some(row => row.length !== returnMatrix[0].length))
            fail('#VALUE!', 'XLOOKUP return_array must be rectangular.');
        const key = val(a, 0), matchMode = a.length >= 5 ? n(a, 4) : 0, searchMode = a.length === 6 ? n(a, 5) : 1;
        if (![-1, 0, 1, 2].includes(matchMode))
            fail('#VALUE!', 'XLOOKUP match_mode must be -1, 0, 1 or 2.');
        if (![-2, -1, 1, 2].includes(searchMode))
            fail('#VALUE!', 'XLOOKUP search_mode must be 1, -1, 2 or -2.');
        const descending = searchMode === -2, binary = searchMode === 2 || descending;
        if (binary && lookup.length > 1_048_576)
            fail('#VALUE!', 'Binary XLOOKUP is limited to one Excel column of values.');
        let index = -1;
        const foundOrFallback = () => index < 0 ? (a.length >= 4 ? val(a, 3) : fail('#N/A', 'XLOOKUP did not find the key.')) : vertical ? [returnMatrix[index].map(value => value ?? null)] : horizontal ? returnMatrix.map(row => [row[index] ?? null]) : returnMatrix[0][0] ?? null;
        if (binary) {
            if (!lookup.length)
                return foundOrFallback();
            for (let i = 1; i < lookup.length; i++) {
                const order = lookupOrder(lookup[i - 1], lookup[i]);
                if (descending ? order < 0 : order > 0)
                    fail('#N/A', `XLOOKUP binary search requires ${descending ? 'descending' : 'ascending'} sorted data.`);
            }
            if (matchMode === 2) {
                const rx = lookupWildcard(text(key));
                if (descending) {
                    for (let i = lookup.length - 1; i >= 0; i--)
                        if (rx.test(text(lookup[i]))) {
                            index = i;
                            break;
                        }
                }
                else
                    for (let i = 0; i < lookup.length; i++)
                        if (rx.test(text(lookup[i]))) {
                            index = i;
                            break;
                        }
            }
            else {
                // Find an exact match first so duplicate keys resolve to the first item
                // in binary traversal order, including for approximate match modes.
                let lo = 0, hi = lookup.length;
                while (lo < hi) {
                    const mid = lo + Math.floor((hi - lo) / 2), cmp = lookupOrder(lookup[mid], key);
                    if (descending ? cmp > 0 : cmp < 0)
                        lo = mid + 1;
                    else
                        hi = mid;
                }
                if (lo < lookup.length && lookupOrder(lookup[lo], key) === 0 && lookupEqual(lookup[lo], key))
                    index = lo;
                if (index < 0 && matchMode !== 0) {
                    lo = 0;
                    hi = lookup.length;
                    if (!descending) {
                        if (matchMode === -1) {
                            while (lo < hi) {
                                const mid = lo + Math.floor((hi - lo) / 2);
                                if (lookupOrder(lookup[mid], key) <= 0)
                                    lo = mid + 1;
                                else
                                    hi = mid;
                            }
                            index = lo - 1;
                        }
                        else {
                            while (lo < hi) {
                                const mid = lo + Math.floor((hi - lo) / 2);
                                if (lookupOrder(lookup[mid], key) < 0)
                                    lo = mid + 1;
                                else
                                    hi = mid;
                            }
                            index = lo < lookup.length ? lo : -1;
                        }
                    }
                    else {
                        if (matchMode === 1) {
                            while (lo < hi) {
                                const mid = lo + Math.floor((hi - lo) / 2);
                                if (lookupOrder(lookup[mid], key) >= 0)
                                    lo = mid + 1;
                                else
                                    hi = mid;
                            }
                            index = lo - 1;
                        }
                        else {
                            while (lo < hi) {
                                const mid = lo + Math.floor((hi - lo) / 2);
                                if (lookupOrder(lookup[mid], key) > 0)
                                    lo = mid + 1;
                                else
                                    hi = mid;
                            }
                            index = lo < lookup.length ? lo : -1;
                        }
                    }
                }
            }
        }
        else {
            const order = searchMode === -1 ? Array.from({ length: lookup.length }, (_, i) => lookup.length - 1 - i) : Array.from({ length: lookup.length }, (_, i) => i);
            const rx = matchMode === 2 ? lookupWildcard(text(key)) : null;
            for (const i of order)
                if (rx ? rx.test(text(lookup[i])) : lookupEqual(lookup[i], key)) {
                    index = i;
                    break;
                }
            if (index < 0 && matchMode !== 0 && matchMode !== 2) {
                let bestOrder = matchMode === 1 ? Infinity : -Infinity;
                for (const i of order) {
                    const cmp = lookupOrder(lookup[i], key);
                    if (matchMode === 1 && cmp > 0 && cmp < bestOrder) {
                        bestOrder = cmp;
                        index = i;
                    }
                    else if (matchMode === -1 && cmp < 0 && cmp > bestOrder) {
                        bestOrder = cmp;
                        index = i;
                    }
                }
            }
        }
        return foundOrFallback();
    }),
    HLOOKUP: def('Lookup in the top row; exact or next-smaller match from ascending data.', '=HLOOKUP("Q2",A1:D3,2,FALSE)', a => { req(a, 3, 4); const rows = matrix(a[1]), headers = rows[0] ?? [], row = integer(a, 2) - 1; if (row < 0 || row >= rows.length)
        fail('#REF!', 'HLOOKUP result row is outside the range.'); const key = val(a, 0), approx = a.length < 4 || bool(val(a, 3)); let col = -1; if (!approx) {
        const pattern = typeof key === 'string' && /[*?~]/.test(key) ? lookupWildcard(key) : null;
        col = headers.findIndex(value => pattern ? pattern.test(text(value)) : lookupEqual(value, key));
    }
    else {
        for (let i = 1; i < headers.length; i++)
            if (lookupOrder(headers[i - 1], headers[i]) > 0)
                fail('#N/A', 'HLOOKUP approximate lookup requires ascending sorted data.');
        for (let i = 0; i < headers.length; i++) {
            const order = lookupOrder(headers[i], key);
            if (order === 0 && lookupEqual(headers[i], key))
                col = i;
            else if (order < 0)
                col = i;
            else
                break;
        }
    } if (col < 0)
        fail('#N/A', `HLOOKUP could not find '${text(key)}'.`); return rows[row][col] ?? null; }),
    // 4 finance functions: standard end-of-period payments, optional type 0/1
    PMT: def('Payment per period for a fixed-rate loan.', '=PMT(0.05/12,60,10000)', a => { req(a, 3, 5); return pmt(n(a, 0), n(a, 1), n(a, 2), a.length >= 4 ? n(a, 3) : 0, a.length === 5 ? n(a, 4) : 0); }),
    FV: def('Future value of an annuity with constant payments.', '=FV(0.05/12,60,-200)', a => { req(a, 3, 5); const rate = n(a, 0), periods = n(a, 1), payment = n(a, 2), pv = a.length >= 4 ? n(a, 3) : 0, type = a.length === 5 ? n(a, 4) : 0; if (periods < 0)
        fail('#NUM!', 'Periods must be nonnegative.'); if (rate === 0)
        return -(pv + payment * periods); const growth = (1 + rate) ** periods; return finite(-pv * growth - payment * (1 + rate * type) * (growth - 1) / rate); }),
    PV: def('Present value of an annuity with constant payments.', '=PV(0.05/12,60,-200)', a => { req(a, 3, 5); const rate = n(a, 0), periods = n(a, 1), payment = n(a, 2), future = a.length >= 4 ? n(a, 3) : 0, type = a.length === 5 ? n(a, 4) : 0; if (periods < 0)
        fail('#NUM!', 'Periods must be nonnegative.'); if (rate === 0)
        return -(future + payment * periods); const growth = (1 + rate) ** periods; return finite(-(future + payment * (1 + rate * type) * (growth - 1) / rate) / growth); }),
    NPV: def('Net present value of periodic future cash flows.', '=NPV(0.1,B1:B5)', a => { req(a, 2, 255); const rate = n(a, 0); if (rate <= -1)
        fail('#NUM!', 'NPV rate must be greater than -1.'); const values = flatten(a.slice(1)); let total = 0; for (let i = 0; i < values.length; i++)
        total += number(values[i]) / (1 + rate) ** (i + 1); return finite(total); }),
};
export const extraFunctionNames = Object.keys(moreFunctions).sort();
