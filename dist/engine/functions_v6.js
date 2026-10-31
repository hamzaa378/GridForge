/** GridForge v6: 90 additional Excel-compatible scalar functions.
 * Deliberately no external formula engine. Dynamic array spill/VBA are out of scope.
 */
import { SheetError, number, text, flatten, matchCriterion } from './functions.js';
const def = (description, example, run) => ({ description, example, run });
const err = (code, message) => { throw new SheetError(code, message); };
const arg = (a, min, max = min) => { if (a.length < min || a.length > max)
    err('#ARG!', `Expected ${min}${min !== max ? `–${max}` : ''} arguments, received ${a.length}.`); };
const sc = (d) => Array.isArray(d) ? err('#VALUE!', 'Expected a single value.') : d ?? null;
const n = (a, i) => number(sc(a[i]));
const s = (a, i) => text(sc(a[i]));
const v = (x) => flatten([x]);
const ns = (x) => v(x).filter((x) => typeof x === 'number' && Number.isFinite(x));
const finite = (x) => Number.isFinite(x) ? x : err('#NUM!', 'Result is not a finite number.');
const sum = (a) => a.reduce((x, y) => x + y, 0);
const avg = (a) => a.length ? sum(a) / a.length : err('#DIV/0!', 'No numeric observations.');
const sorted = (a) => a.sort((x, y) => x - y);
const meanVar = (a, sample) => a.length > (sample ? 1 : 0) ? sum(a.map(x => (x - avg(a)) ** 2)) / (a.length - (sample ? 1 : 0)) : err('#DIV/0!', 'Insufficient observations.');
const percentile = (x, p, exclusive = false) => { const a = sorted([...x]); if (!a.length || p < 0 || p > 1)
    err('#NUM!', 'Invalid percentile or empty range.'); const r = exclusive ? p * (a.length + 1) : p * (a.length - 1) + 1; if (r < 1 || r > a.length)
    err('#NUM!', 'Exclusive percentile outside available rank.'); const k = Math.floor(r); return finite(a[k - 1] + (r - k) * ((a[Math.min(k, a.length - 1)] ?? a[k - 1]) - a[k - 1])); };
const rank = (value, x, ascending = false) => 1 + x.filter(k => ascending ? k < value : k > value).length;
const binomial = (k, p) => { if (k < 0 || !Number.isInteger(k))
    return 0; if (p < 0 || p > 1)
    err('#NUM!', 'Probability outside 0–1.'); return p === 0 ? (k === 0 ? 1 : 0) : p === 1 ? 0 : Math.exp(k * Math.log(p) - logGamma(k + 1)); };
// Lanczos approximation, reflection for negative noninteger arguments.
function logGamma(z) { const c = [.99999999999980993, 676.5203681218851, -1259.1392167224028, 771.32342877765313, -176.6150291621406, 12.507343278686905, -.13857109526572012, 9.9843695780195716e-6, 1.5056327351493116e-7]; if (z < .5)
    return Math.log(Math.PI) - Math.log(Math.abs(Math.sin(Math.PI * z))) - logGamma(1 - z); z--; let x = c[0]; for (let i = 1; i < c.length; i++)
    x += c[i] / (z + i); const t = z + 7.5; return .5 * Math.log(2 * Math.PI) + (z + .5) * Math.log(t) - t + Math.log(x); }
function gamma(z) { if (z <= 0 && Number.isInteger(z))
    err('#NUM!', 'Gamma pole at non-positive integer.'); if (z < .5)
    return Math.PI / (Math.sin(Math.PI * z) * Math.exp(logGamma(1 - z))); return finite(Math.exp(logGamma(z))); }
function erf(x) { if (x === 0)
    return 0; const sign = x < 0 ? -1 : 1, z = Math.abs(x); if (z >= 6)
    return sign; let term = z, total = z; for (let k = 1; k < 160; k++) {
    term *= -(z * z) / k;
    const add = term / (2 * k + 1);
    total += add;
    if (Math.abs(add) < Math.abs(total) * 2e-16)
        break;
} return sign * 2 / Math.sqrt(Math.PI) * total; }
const normal = (x) => .5 * (1 + erf(x / Math.SQRT2));
function invNormal(p) { if (p <= 0 || p >= 1)
    err('#NUM!', 'Probability must be strictly between 0 and 1.'); let lo = -11, hi = 11; for (let i = 0; i < 90; i++) {
    const mid = (lo + hi) / 2;
    if (normal(mid) < p)
        lo = mid;
    else
        hi = mid;
} return (lo + hi) / 2; }
const comb = (m, k) => { if (!Number.isInteger(m) || !Number.isInteger(k) || m < 0 || k < 0 || k > m)
    err('#NUM!', 'Invalid combination.'); return Math.round(Math.exp(logGamma(m + 1) - logGamma(k + 1) - logGamma(m - k + 1))); };
const criterionPairs = (a, offset) => { if (a.length < offset + 2 || (a.length - offset) % 2 !== 0)
    err('#ARG!', 'Criteria require range, criterion pairs.'); const rows = [], criteria = []; for (let i = offset; i < a.length; i += 2) {
    rows.push(v(a[i]));
    const criterion = sc(a[i + 1]);
    criteria.push(criterion === null ? 0 : criterion);
} const length = rows[0].length; if (rows.some(row => row.length !== length))
    err('#VALUE!', 'Criteria ranges differ in length.'); return { rows, criteria, length }; };
const matchAll = (c, i) => c.rows.every((row, j) => matchCriterion(row[i] ?? null, c.criteria[j]));
const toMatrix = (x) => Array.isArray(x) ? x.map(row => row.map(sc)) : [[sc(x)]];
const maxExcelDate = 2_958_465;
const parseRadix = (value, base) => { const alphabet = base === 2 ? '01' : base === 8 ? '01234567' : '0123456789ABCDEF', textValue = value.toUpperCase(); if (!textValue || textValue.length > 10 || [...textValue].some(char => !alphabet.includes(char)))
    err('#NUM!', 'Invalid base-conversion text or input exceeds 10 characters.'); const magnitude = parseInt(textValue, base), signBit = base ** 9 * (base === 2 ? 1 : base === 8 ? 4 : 8); return textValue.length === 10 && magnitude >= signBit ? magnitude - base ** 10 : magnitude; };
const radixResult = (value, base, places) => { const bits = base === 2 ? 10 : base === 8 ? 10 : 40, min = -(2 ** (bits - 1)), max = 2 ** (bits - 1) - 1; if (!Number.isInteger(value) || value < min || value > max)
    err('#NUM!', 'Decimal value is outside the supported signed range.'); if (places !== undefined && (!Number.isInteger(places) || places < 1 || places > 10))
    err('#NUM!', 'Places must be an integer from 1 to 10.'); if (value < 0)
    return (2 ** bits + value).toString(base).toUpperCase().padStart(10, '0'); const result = value.toString(base).toUpperCase(); if (places !== undefined && result.length > places)
    err('#NUM!', 'Result requires more characters than places allows.'); return result.padStart(places ?? 1, '0'); };
const validateDateSerial = (days, label = 'Date') => { if (!Number.isFinite(days) || days < 0 || days > maxExcelDate)
    err('#NUM!', `${label} is outside Excel’s supported date range.`); };
const holidaySerials = (x) => { const values = ns(x); for (const value of values)
    validateDateSerial(value, 'Holiday date'); return new Set(values.map(Math.trunc)); };
const dateFrom = (days) => { const whole = Math.trunc(days); return new Date(Date.UTC(1899, 11, 31) + (whole - (whole >= 60 ? 1 : 0)) * 86400000); };
const excelWeekday = (days) => Math.trunc(days) === 60 ? 3 : dateFrom(days).getUTCDay();
const serial = (d) => Math.floor((d.getTime() - Date.UTC(1899, 11, 31)) / 86400000) + (d.getTime() >= Date.UTC(1900, 2, 1) ? 1 : 0);
const workday = (start, days, weekend, holidays) => { validateDateSerial(start, 'Start date'); if (weekend === '1111111')
    err('#NUM!', 'The weekend definition must leave at least one working day.'); let current = Math.trunc(start), step = Math.sign(days), remaining = Math.abs(Math.trunc(days)); if (remaining > 366000)
    err('#NUM!', 'Workday interval too large.'); while (remaining) {
    current += step;
    validateDateSerial(current, 'Result date');
    const dow = (excelWeekday(current) + 6) % 7;
    if (weekend[dow] === '0' && !holidays.has(current))
        remaining--;
} return current; };
const weekendCode = (x) => { const codes = { '1': '0000011', '2': '1000001', '3': '1100000', '4': '0110000', '5': '0011000', '6': '0001100', '7': '0000110', '11': '0000001', '12': '1000000', '13': '0100000', '14': '0010000', '15': '0001000', '16': '0000100', '17': '0000010' }; const mask = typeof x === 'string' && /^[01]{7}$/.test(x) ? x : codes[String(x)]; if (!mask)
    err('#VALUE!', 'Invalid weekend code/mask.'); return mask; };
const usThirty360 = (a, b) => { const one = dateFrom(a), two = dateFrom(b); let d1 = one.getUTCDate(), d2 = two.getUTCDate(); const lastFeb = (d) => d.getUTCMonth() === 1 && d.getUTCDate() === new Date(Date.UTC(d.getUTCFullYear(), 2, 0)).getUTCDate(); if (d1 === 31 || lastFeb(one))
    d1 = 30; if (d2 === 31 && d1 === 30)
    d2 = 30; if (lastFeb(one) && lastFeb(two))
    d2 = 30; return 360 * (two.getUTCFullYear() - one.getUTCFullYear()) + 30 * (two.getUTCMonth() - one.getUTCMonth()) + d2 - d1; };
const pm = (r, t, pv, fv = 0, type = 0) => { if (t <= 0 || r <= -1 || ![0, 1].includes(type))
    err('#NUM!', 'Invalid loan arguments.'); if (r === 0)
    return -(pv + fv) / t; const growth = (1 + r) ** t; return finite(-(pv * growth + fv) * r / ((1 + r * type) * (growth - 1))); };
const testFunction = () => err('#INTERNAL!', 'This reference-aware function is evaluated by the GridForge workbook engine.');
export const v6Functions = {
    // Binary, octal, decimal and hexadecimal conversions (Excel's signed fixed-width format).
    'BIN2DEC': def('Convert a signed 10-bit binary value to decimal.', '=BIN2DEC("1111111111")', a => { arg(a, 1); return parseRadix(s(a, 0), 2); }),
    'BIN2HEX': def('Convert a signed 10-bit binary value to hexadecimal.', '=BIN2HEX("1110",4)', a => { arg(a, 1, 2); return radixResult(parseRadix(s(a, 0), 2), 16, a.length === 2 ? Math.trunc(n(a, 1)) : undefined); }),
    'BIN2OCT': def('Convert a signed 10-bit binary value to octal.', '=BIN2OCT("1001",3)', a => { arg(a, 1, 2); return radixResult(parseRadix(s(a, 0), 2), 8, a.length === 2 ? Math.trunc(n(a, 1)) : undefined); }),
    'DEC2BIN': def('Convert a signed decimal integer to 10-bit binary.', '=DEC2BIN(-100)', a => { arg(a, 1, 2); const value = Math.trunc(n(a, 0)), places = a.length === 2 ? Math.trunc(n(a, 1)) : undefined; if (value < 0)
        return radixResult(value, 2); return radixResult(value, 2, places); }),
    'DEC2HEX': def('Convert a signed decimal integer to hexadecimal.', '=DEC2HEX(-54)', a => { arg(a, 1, 2); return radixResult(Math.trunc(n(a, 0)), 16, a.length === 2 ? Math.trunc(n(a, 1)) : undefined); }),
    'DEC2OCT': def('Convert a signed decimal integer to octal.', '=DEC2OCT(58,3)', a => { arg(a, 1, 2); return radixResult(Math.trunc(n(a, 0)), 8, a.length === 2 ? Math.trunc(n(a, 1)) : undefined); }),
    'HEX2BIN': def('Convert a signed 40-bit hexadecimal value to binary.', '=HEX2BIN("FFFFFFFFFF")', a => { arg(a, 1, 2); const value = parseRadix(s(a, 0), 16); if (value < -512 || value > 511)
        err('#NUM!', 'Value is outside the signed 10-bit binary range.'); return radixResult(value, 2, a.length === 2 ? Math.trunc(n(a, 1)) : undefined); }),
    'HEX2DEC': def('Convert a signed 40-bit hexadecimal value to decimal.', '=HEX2DEC("A5")', a => { arg(a, 1); return parseRadix(s(a, 0), 16); }),
    'HEX2OCT': def('Convert a signed 40-bit hexadecimal value to octal.', '=HEX2OCT("A5")', a => { arg(a, 1, 2); return radixResult(parseRadix(s(a, 0), 16), 8, a.length === 2 ? Math.trunc(n(a, 1)) : undefined); }),
    'OCT2BIN': def('Convert a signed 30-bit octal value to binary.', '=OCT2BIN("7777777777")', a => { arg(a, 1, 2); const value = parseRadix(s(a, 0), 8); if (value < -512 || value > 511)
        err('#NUM!', 'Value is outside the signed 10-bit binary range.'); return radixResult(value, 2, a.length === 2 ? Math.trunc(n(a, 1)) : undefined); }),
    'OCT2DEC': def('Convert a signed 30-bit octal value to decimal.', '=OCT2DEC("7777777777")', a => { arg(a, 1); return parseRadix(s(a, 0), 8); }),
    'OCT2HEX': def('Convert a signed 30-bit octal value to hexadecimal.', '=OCT2HEX("7777777777")', a => { arg(a, 1, 2); return radixResult(parseRadix(s(a, 0), 8), 16, a.length === 2 ? Math.trunc(n(a, 1)) : undefined); }),
    // 24 mathematics and engineering functions
    'SQRTPI': def('Square root of pi times x.', '=SQRTPI(4)', a => { arg(a, 1); return finite(Math.sqrt(Math.PI * n(a, 0))); }),
    'SUMX2MY2': def('Sum of differences of squares.', '=SUMX2MY2(A1:A3,B1:B3)', a => { arg(a, 2); const x = v(a[0]), y = v(a[1]); if (x.length !== y.length)
        err('#VALUE!', 'Ranges must match.'); return finite(sum(x.map((t, i) => number(t) ** 2 - number(y[i]) ** 2))); }),
    'SUMX2PY2': def('Sum of sums of squares.', '=SUMX2PY2(A1:A3,B1:B3)', a => { arg(a, 2); const x = v(a[0]), y = v(a[1]); if (x.length !== y.length)
        err('#VALUE!', 'Ranges must match.'); return finite(sum(x.map((t, i) => number(t) ** 2 + number(y[i]) ** 2))); }),
    'SUMXMY2': def('Sum of squared differences.', '=SUMXMY2(A1:A3,B1:B3)', a => { arg(a, 2); const x = v(a[0]), y = v(a[1]); if (x.length !== y.length)
        err('#VALUE!', 'Ranges must match.'); return finite(sum(x.map((t, i) => (number(t) - number(y[i])) ** 2))); }),
    'COMBINA': def('Combinations with repetition.', '=COMBINA(3,2)', a => { arg(a, 2); const m = Math.trunc(n(a, 0)), k = Math.trunc(n(a, 1)); return m === 0 && k === 0 ? 1 : comb(m + k - 1, k); }),
    'PERMUTATIONA': def('Permutations with repetition.', '=PERMUTATIONA(3,2)', a => { arg(a, 2); const m = n(a, 0), k = n(a, 1); if (!Number.isInteger(m) || !Number.isInteger(k) || m < 0 || k < 0)
        err('#NUM!', 'Invalid permutation.'); return finite(m ** k); }),
    'BITAND': def('Bitwise AND for non-negative 48-bit integers.', '=BITAND(6,3)', a => { arg(a, 2); return bit(n(a, 0), n(a, 1), (x, y) => x & y); }),
    'BITOR': def('Bitwise OR for non-negative 48-bit integers.', '=BITOR(6,3)', a => { arg(a, 2); return bit(n(a, 0), n(a, 1), (x, y) => x | y); }),
    'BITXOR': def('Bitwise XOR for non-negative 48-bit integers.', '=BITXOR(6,3)', a => { arg(a, 2); return bit(n(a, 0), n(a, 1), (x, y) => x ^ y); }),
    'BITLSHIFT': def('Shift bits left within supported 48-bit range.', '=BITLSHIFT(3,2)', a => shift(n(a, 0), Math.trunc(n(a, 1)))),
    'BITRSHIFT': def('Shift bits right within supported 48-bit range.', '=BITRSHIFT(16,2)', a => shift(n(a, 0), -Math.trunc(n(a, 1)))),
    'DELTA': def('Returns 1 when two numbers are equal.', '=DELTA(4,4)', a => { arg(a, 1, 2); return n(a, 0) === (a.length === 2 ? n(a, 1) : 0) ? 1 : 0; }),
    'GESTEP': def('Returns 1 when a number reaches a step.', '=GESTEP(7,5)', a => { arg(a, 1, 2); return n(a, 0) >= (a.length === 2 ? n(a, 1) : 0) ? 1 : 0; }),
    'ERF': def('Approximate Gaussian error function.', '=ERF(1)', a => { arg(a, 1, 2); return a.length === 2 ? erf(n(a, 1)) - erf(n(a, 0)) : erf(n(a, 0)); }),
    'ERFC': def('Complementary Gaussian error function.', '=ERFC(1)', a => { arg(a, 1); return 1 - erf(n(a, 0)); }),
    'CEILING.MATH': def('Round up to multiple, with mode for negatives.', '=CEILING.MATH(7.2,2)', a => ceilMath(a, false)),
    'FLOOR.MATH': def('Round down to multiple, with mode for negatives.', '=FLOOR.MATH(7.2,2)', a => ceilMath(a, true)),
    'CEILING.PRECISE': def('Round toward positive infinity to significance.', '=CEILING.PRECISE(-4.1,2)', a => precise(a, true)),
    'FLOOR.PRECISE': def('Round toward negative infinity to significance.', '=FLOOR.PRECISE(-4.1,2)', a => precise(a, false)),
    'ISO.CEILING': def('ISO-compliant rounding to significance.', '=ISO.CEILING(5.2,2)', a => precise(a, true)),
    'GAMMA': def('Gamma function.', '=GAMMA(5)', a => { arg(a, 1); return gamma(n(a, 0)); }),
    'GAMMALN': def('Natural log of gamma.', '=GAMMALN(5)', a => { arg(a, 1); if (n(a, 0) <= 0)
        err('#NUM!', 'Positive input required.'); return logGamma(n(a, 0)); }),
    'GAMMALN.PRECISE': def('High-precision log gamma.', '=GAMMALN.PRECISE(5)', a => { arg(a, 1); if (n(a, 0) <= 0)
        err('#NUM!', 'Positive input required.'); return logGamma(n(a, 0)); }),
    'SERIESSUM': def('Power series polynomial weighted by coefficients.', '=SERIESSUM(2,1,2,A1:A3)', a => { arg(a, 4); const x = n(a, 0), degree = n(a, 1), step = n(a, 2); return finite(sum(ns(a[3]).map((c, i) => c * x ** (degree + i * step)))); }),
    // 36 statistical and probability functions
    'MAXIFS': def('Maximum value matching all criteria.', '=MAXIFS(C2:C6,A2:A6,"East")', a => ifStats(a, 'max')),
    'MINIFS': def('Minimum value matching all criteria.', '=MINIFS(C2:C6,A2:A6,"East")', a => ifStats(a, 'min')),
    'AVERAGEA': def('Average including referenced text as zero; direct nonnumeric text is an error.', '=AVERAGEA(A1:A4)', a => { for (const item of a)
        if (typeof item === 'string' && item !== '' && !Number.isFinite(Number(item)))
            err('#VALUE!', 'Literal text is not numeric.'); const z = vectors(a).map(x => typeof x === 'number' ? x : typeof x === 'boolean' ? (x ? 1 : 0) : typeof x === 'string' ? 0 : null).filter((x) => x !== null); return avg(z); }),
    'MAXA': def('Largest numeric or logical value.', '=MAXA(A1:A4)', a => (() => { const z = vectors(a).filter(x => x !== null && x !== '').map(x => typeof x === 'number' ? x : x === true ? 1 : 0); return z.length ? Math.max(...z) : 0; })()),
    'MINA': def('Smallest numeric or logical value.', '=MINA(A1:A4)', a => (() => { const z = vectors(a).filter(x => x !== null && x !== '').map(x => typeof x === 'number' ? x : x === true ? 1 : 0); return z.length ? Math.min(...z) : 0; })()),
    'STDEV.S': def('Sample standard deviation.', '=STDEV.S(A1:A6)', a => Math.sqrt(meanVar(nsList(a), true))),
    'STDEV.P': def('Population standard deviation.', '=STDEV.P(A1:A6)', a => Math.sqrt(meanVar(nsList(a), false))),
    'VAR.S': def('Sample variance.', '=VAR.S(A1:A6)', a => meanVar(nsList(a), true)),
    'VAR.P': def('Population variance.', '=VAR.P(A1:A6)', a => meanVar(nsList(a), false)),
    'MODE.SNGL': def('Most frequent numeric value.', '=MODE.SNGL(A1:A8)', a => { const x = nsList(a), counts = new Map(); for (const z of x)
        counts.set(z, (counts.get(z) ?? 0) + 1); const max = Math.max(0, ...counts.values()); if (max < 2)
        err('#N/A', 'No repeated numeric value.'); return x.find(z => counts.get(z) === max); }),
    'PERCENTILE.INC': def('Inclusive percentile interpolation.', '=PERCENTILE.INC(A1:A8,0.75)', a => { arg(a, 2); return percentile(ns(a[0]), n(a, 1)); }),
    'PERCENTILE.EXC': def('Exclusive percentile interpolation.', '=PERCENTILE.EXC(A1:A8,0.75)', a => { arg(a, 2); return percentile(ns(a[0]), n(a, 1), true); }),
    'PERCENTRANK.INC': def('Inclusive percentile rank.', '=PERCENTRANK.INC(A1:A8,5)', a => pRank(a, false)),
    'PERCENTRANK.EXC': def('Exclusive percentile rank.', '=PERCENTRANK.EXC(A1:A8,5)', a => pRank(a, true)),
    'QUARTILE.INC': def('Inclusive quartile.', '=QUARTILE.INC(A1:A8,3)', a => { arg(a, 2); const q = n(a, 1); if (q < 0 || q > 4 || !Number.isInteger(q))
        err('#NUM!', 'Quartile must be 0–4.'); return percentile(ns(a[0]), q / 4); }),
    'QUARTILE.EXC': def('Exclusive quartile.', '=QUARTILE.EXC(A1:A8,3)', a => { arg(a, 2); const q = n(a, 1); if (q < 1 || q > 3 || !Number.isInteger(q))
        err('#NUM!', 'Quartile must be 1–3.'); return percentile(ns(a[0]), q / 4, true); }),
    'RANK.EQ': def('Rank in numeric list.', '=RANK.EQ(4,A1:A5,0)', a => { arg(a, 2, 3); return rank(n(a, 0), ns(a[1]), a.length === 3 ? n(a, 2) !== 0 : false); }),
    'RANK.AVG': def('Average rank in ties.', '=RANK.AVG(4,A1:A5,0)', a => { arg(a, 2, 3); const x = ns(a[1]), target = n(a, 0), asc = a.length === 3 ? n(a, 2) !== 0 : false, k = x.filter(y => y === target).length; return rank(target, x, asc) + (k - 1) / 2; }),
    'FORECAST.LINEAR': def('Least-squares forecast.', '=FORECAST.LINEAR(5,B1:B4,A1:A4)', a => forecast(a)),
    'SKEW.P': def('Population skewness.', '=SKEW.P(A1:A8)', a => { const x = nsList(a), sd = Math.sqrt(meanVar(x, false)); return sd === 0 ? err('#DIV/0!', 'Standard deviation is zero.') : avg(x.map(y => ((y - avg(x)) / sd) ** 3)); }),
    'FISHER': def('Fisher z transform.', '=FISHER(0.5)', a => { arg(a, 1); const x = n(a, 0); if (Math.abs(x) >= 1)
        err('#NUM!', 'Value must be between -1 and 1.'); return Math.log((1 + x) / (1 - x)) / 2; }),
    'FISHERINV': def('Inverse Fisher transform.', '=FISHERINV(0.5)', a => { arg(a, 1); return Math.tanh(n(a, 0)); }),
    'BINOM.DIST': def('Binomial probability mass or cumulative probability.', '=BINOM.DIST(2,5,0.5,FALSE)', a => { arg(a, 4); const k = Math.floor(n(a, 0)), m = n(a, 1), p = n(a, 2); if (!Number.isInteger(m) || m < 0 || p < 0 || p > 1)
        err('#NUM!', 'Invalid binomial parameters.'); const prob = (z) => comb(m, z) * p ** z * (1 - p) ** (m - z); return sc(a[3]) ? finite(sum(Array.from({ length: Math.min(k, m) + 1 }, (_, i) => prob(i)))) : k < 0 || k > m ? 0 : finite(prob(k)); }),
    'POISSON.DIST': def('Poisson probability.', '=POISSON.DIST(2,4,FALSE)', a => { arg(a, 3); const k = Math.floor(n(a, 0)), lambda = n(a, 1); if (lambda <= 0 || k < 0)
        err('#NUM!', 'Invalid Poisson parameters.'); const p = (i) => Math.exp(-lambda + i * Math.log(lambda) - logGamma(i + 1)); return sc(a[2]) ? sum(Array.from({ length: k + 1 }, (_, i) => p(i))) : p(k); }),
    'EXPON.DIST': def('Exponential distribution probability.', '=EXPON.DIST(2,0.5,TRUE)', a => { arg(a, 3); const x = n(a, 0), lambda = n(a, 1); if (x < 0 || lambda <= 0)
        err('#NUM!', 'Invalid exponential parameters.'); return sc(a[2]) ? 1 - Math.exp(-lambda * x) : lambda * Math.exp(-lambda * x); }),
    'WEIBULL.DIST': def('Weibull distribution probability.', '=WEIBULL.DIST(2,1.5,3,TRUE)', a => { arg(a, 4); const x = n(a, 0), alpha = n(a, 1), beta = n(a, 2); if (x < 0 || alpha <= 0 || beta <= 0)
        err('#NUM!', 'Invalid Weibull parameters.'); return sc(a[3]) ? 1 - Math.exp(-((x / beta) ** alpha)) : alpha / beta * (x / beta) ** (alpha - 1) * Math.exp(-((x / beta) ** alpha)); }),
    'NORM.DIST': def('Normal probability density or cumulative distribution.', '=NORM.DIST(1,0,1,TRUE)', a => { arg(a, 4); const z = (n(a, 0) - n(a, 1)) / n(a, 2); if (n(a, 2) <= 0)
        err('#NUM!', 'Standard deviation must be positive.'); return sc(a[3]) ? normal(z) : Math.exp(-z * z / 2) / (n(a, 2) * Math.sqrt(2 * Math.PI)); }),
    'NORM.INV': def('Inverse normal cumulative distribution.', '=NORM.INV(0.95,0,1)', a => { arg(a, 3); if (n(a, 2) <= 0)
        err('#NUM!', 'Invalid standard deviation.'); return n(a, 1) + n(a, 2) * invNormal(n(a, 0)); }),
    'NORM.S.DIST': def('Standard normal probability density or CDF.', '=NORM.S.DIST(1,TRUE)', a => { arg(a, 2); const x = n(a, 0); return sc(a[1]) ? normal(x) : Math.exp(-x * x / 2) / Math.sqrt(2 * Math.PI); }),
    'NORM.S.INV': def('Standard normal quantile.', '=NORM.S.INV(0.95)', a => { arg(a, 1); return invNormal(n(a, 0)); }),
    'LOGNORM.DIST': def('Lognormal probability density or CDF.', '=LOGNORM.DIST(2,0,1,TRUE)', a => { arg(a, 4); const x = n(a, 0), mean = n(a, 1), sd = n(a, 2); if (x <= 0 || sd <= 0)
        err('#NUM!', 'Invalid lognormal parameters.'); const z = (Math.log(x) - mean) / sd; return sc(a[3]) ? normal(z) : Math.exp(-z * z / 2) / (x * sd * Math.sqrt(2 * Math.PI)); }),
    'LOGNORM.INV': def('Inverse lognormal CDF.', '=LOGNORM.INV(0.95,0,1)', a => { arg(a, 3); if (n(a, 2) <= 0)
        err('#NUM!', 'Invalid deviation.'); return Math.exp(n(a, 1) + n(a, 2) * invNormal(n(a, 0))); }),
    'CONFIDENCE.NORM': def('Normal confidence interval margin.', '=CONFIDENCE.NORM(0.05,2,100)', a => { arg(a, 3); const alpha = n(a, 0), sigma = n(a, 1), size = n(a, 2); if (alpha <= 0 || alpha >= 1 || sigma <= 0 || size <= 0)
        err('#NUM!', 'Invalid confidence parameters.'); return invNormal(1 - alpha / 2) * sigma / Math.sqrt(size); }),
    'PHI': def('Standard normal probability density.', '=PHI(0)', a => { arg(a, 1); return Math.exp(-(n(a, 0) ** 2) / 2) / Math.sqrt(2 * Math.PI); }),
    'GAUSS': def('Normal probability from mean to x.', '=GAUSS(1)', a => { arg(a, 1); return normal(n(a, 0)) - .5; }),
    'Z.TEST': def('One-tailed z-test probability.', '=Z.TEST(A1:A10,5)', a => { arg(a, 2, 3); const x = ns(a[0]), target = n(a, 1), sigma = a.length === 3 ? n(a, 2) : Math.sqrt(meanVar(x, true)); if (sigma <= 0)
        err('#DIV/0!', 'Standard deviation must be positive.'); return 1 - normal((avg(x) - target) * Math.sqrt(x.length) / sigma); }),
    // 4 widely usable text functions
    'ENCODEURL': def('Percent-encode text for URLs.', '=ENCODEURL("a b&c")', a => { arg(a, 1); return encodeURIComponent(s(a, 0)); }),
    'ARRAYTOTEXT': def('Convert a range or value to text.', '=ARRAYTOTEXT(A1:B2,1)', a => { arg(a, 1, 2); return a.length === 2 && n(a, 1) === 1 ? `{${toMatrix(a[0]).map(r => r.map(x => typeof x === 'string' ? `"${x.replaceAll('"', '""')}"` : text(x)).join(',')).join(';')}}` : v(a[0]).map(text).join(', '); }),
    'VALUETOTEXT': def('Convert one value to text, optionally with quotes.', '=VALUETOTEXT("hello",1)', a => { arg(a, 1, 2); const x = sc(a[0]); return a.length === 2 && n(a, 1) === 1 && typeof x === 'string' ? `"${x.replaceAll('"', '""')}"` : text(x); }),
    'REGEXEXTRACT': def('Extract the first match, all matches, or capture groups with optional case-insensitive matching.', '=REGEXEXTRACT("A12 B34","[A-Z]\\d+",1,1)', a => { arg(a, 2, 4); const input = s(a, 0), pattern = s(a, 1), returnMode = a.length >= 3 ? n(a, 2) : 0, caseSensitivity = a.length >= 4 ? n(a, 3) : 0; if (input.length > 2000)
        err('#VALUE!', 'Regex input exceeds 2,000 characters.'); if (!Number.isInteger(returnMode) || returnMode < 0 || returnMode > 2)
        err('#VALUE!', 'return_mode must be 0, 1, or 2.'); if (caseSensitivity !== 0 && caseSensitivity !== 1)
        err('#VALUE!', 'case_sensitivity must be 0 or 1.'); if (pattern.length > 256 || /\\[1-9]/.test(pattern) || /\(\?/.test(pattern) || /\([^)]*[+*{][^)]*\)[+*{]/.test(pattern))
        err('#VALUE!', 'Unsupported or potentially expensive regex pattern.'); const flags = caseSensitivity === 1 ? 'i' : ''; let first; try {
        first = new RegExp(pattern, flags).exec(input);
    }
    catch {
        return err('#VALUE!', 'Invalid regular expression.');
    } if (!first)
        return err('#N/A', 'No regex match.'); if (returnMode === 0)
        return first[0]; if (returnMode === 2) {
        if (first.length < 2)
            return err('#N/A', 'No capture groups were found.');
        return [first.slice(1).map(value => value ?? '')];
    } let all; try {
        all = [...input.matchAll(new RegExp(pattern, `${flags}g`))];
    }
    catch {
        return err('#VALUE!', 'Invalid regular expression.');
    } return [all.map(match => match[0])]; }),
    // 14 logical/reference/information functions; reference-sensitive ones resolved in workbook.ts
    'IFNA': def('Only use fallback for #N/A; short-circuit other errors.', '=IFNA(XMATCH(9,A1:A5),0)', testFunction),
    'XMATCH': def('Exact, wildcard, or approximate matching with search direction.', '=XMATCH("A",A1:A5,0)', a => xmatch(a)),
    'ROW': def('1-based row number of a reference.', '=ROW(A7)', testFunction),
    'COLUMN': def('1-based column number of a reference.', '=COLUMN(B7)', testFunction),
    'FORMULATEXT': def('Return source formula from a referenced cell.', '=FORMULATEXT(A1)', testFunction),
    'ISFORMULA': def('TRUE when reference contains a formula.', '=ISFORMULA(A1)', testFunction),
    'INDIRECT': def('Evaluate an A1 address held in text.', '=INDIRECT("B7")', testFunction),
    'OFFSET': def('Return scalar value at offset from a referenced cell.', '=OFFSET(A1,1,2)', testFunction),
    'HYPERLINK': def('Return display text and activate safe direct HTTP/HTTPS/mailto or simple in-workbook cell targets.', '=HYPERLINK("https://example.com","Site")', a => { arg(a, 1, 2); return a.length === 2 ? s(a, 1) : s(a, 0); }),
    'AREAS': def('Count reference areas (single range supported).', '=AREAS(A1:C4)', a => { arg(a, 1); return 1; }),
    'NA': def('Return #N/A intentionally.', '=NA()', a => { arg(a, 0); return err('#N/A', 'Not available.'); }),
    'ERROR.TYPE': def('Excel-style numeric code for a formula error.', '=ERROR.TYPE(NA())', testFunction),
    'ISREF': def('TRUE when argument is a valid reference.', '=ISREF(A1)', testFunction),
    'CELL': def('Return address or type information for a reference.', '=CELL("address",B2)', testFunction),
    // 10 financial functions
    'CUMIPMT': def('Cumulative interest paid over periods.', '=CUMIPMT(0.01,12,1000,1,3,0)', a => cumulative(a, true)),
    'CUMPRINC': def('Cumulative principal paid over periods.', '=CUMPRINC(0.01,12,1000,1,3,0)', a => cumulative(a, false)),
    'ISPMT': def('Straight-line interest over period.', '=ISPMT(0.05,2,12,1000)', a => { arg(a, 4); const r = n(a, 0), k = n(a, 1), periods = n(a, 2), pv = n(a, 3); if (periods <= 0)
        err('#NUM!', 'Period count must be positive.'); return -pv * r * (periods - k) / periods; }),
    'DB': def('Fixed declining-balance depreciation.', '=DB(10000,1000,5,1)', a => { arg(a, 4, 5); const cost = n(a, 0), salvage = n(a, 1), life = n(a, 2), period = Math.trunc(n(a, 3)), months = a.length === 5 ? n(a, 4) : 12; if (cost <= 0 || salvage < 0 || salvage > cost || life <= 0 || period < 1 || months < 1 || months > 12)
        err('#NUM!', 'Invalid depreciation.'); const rate = Math.round((1 - (salvage / cost) ** (1 / life)) * 1000) / 1000; let balance = cost, dep = 0; for (let i = 1; i <= period; i++) {
        dep = balance * rate * (i === 1 ? months / 12 : i === life + 1 ? (12 - months) / 12 : 1);
        balance -= dep;
    } return finite(dep); }),
    'VDB': def('Variable declining-balance depreciation across range.', '=VDB(10000,1000,5,0,2)', a => { arg(a, 5, 7); const cost = n(a, 0), salvage = n(a, 1), life = n(a, 2), start = n(a, 3), end = n(a, 4), factor = a.length >= 6 ? n(a, 5) : 2; if (life <= 0 || start < 0 || end < start || end > life || factor <= 0)
        err('#NUM!', 'Invalid depreciation interval.'); let balance = cost, out = 0; for (let k = 0; k < Math.ceil(end); k++) {
        const dep = Math.min(Math.max(0, balance - salvage), balance * factor / life);
        const overlap = Math.max(0, Math.min(k + 1, end) - Math.max(k, start));
        out += dep * overlap;
        balance -= dep;
    } return finite(out); }),
    'PDURATION': def('Periods required for investment to reach future value.', '=PDURATION(0.1,1000,2000)', a => { arg(a, 3); if (n(a, 0) <= 0 || n(a, 1) <= 0 || n(a, 2) <= 0)
        err('#NUM!', 'Values must be positive.'); return Math.log(n(a, 2) / n(a, 1)) / Math.log(1 + n(a, 0)); }),
    'DOLLARDE': def('Convert fractional dollar quotation to decimal.', '=DOLLARDE(1.02,16)', a => { arg(a, 2); const dollar = n(a, 0), base = Math.trunc(n(a, 1)); if (base < 1)
        err('#DIV/0!', 'Fraction denominator must be positive.'); const whole = Math.trunc(dollar), fraction = dollar - whole, digits = String(base).length; return whole + Math.round(fraction * 10 ** digits) / base; }),
    'DOLLARFR': def('Convert decimal dollar quotation to fractional.', '=DOLLARFR(1.125,16)', a => { arg(a, 2); const dollar = n(a, 0), base = Math.trunc(n(a, 1)); if (base < 1)
        err('#DIV/0!', 'Fraction denominator must be positive.'); const whole = Math.trunc(dollar); return whole + Math.round((dollar - whole) * base) / 10 ** String(base).length; }),
    'PRICEDISC': def('Price of discounted security per 100 par.', '=PRICEDISC(45000,45100,0.05,100)', a => { arg(a, 3, 4); const days = usThirty360(n(a, 0), n(a, 1)), redemption = a.length === 4 ? n(a, 3) : 100; if (days <= 0 || n(a, 2) <= 0)
        err('#NUM!', 'Invalid security dates or discount.'); return redemption * (1 - n(a, 2) * days / 360); }),
    'DISC': def('Discount rate for a short discounted security.', '=DISC(45000,45100,98,100)', a => { arg(a, 4); const days = usThirty360(n(a, 0), n(a, 1)); if (days <= 0 || n(a, 2) <= 0 || n(a, 3) <= 0)
        err('#NUM!', 'Invalid discount-security parameters.'); return (n(a, 3) - n(a, 2)) / n(a, 3) * 360 / days; }),
    // 2 modern dotted date-function names for existing date capabilities
    'WORKDAY.INTL': def('Working-day offset with chosen weekend days.', '=WORKDAY.INTL(DATE(2026,1,1),5,1)', a => { arg(a, 2, 4); const mask = weekendCode(a.length >= 3 ? sc(a[2]) : 1), holidays = a.length === 4 ? holidaySerials(a[3]) : new Set(); return workday(n(a, 0), n(a, 1), mask, holidays); }),
    'NETWORKDAYS.INTL': def('Working days between dates with chosen weekend days.', '=NETWORKDAYS.INTL(A1,B1,1)', a => { arg(a, 2, 4); const mask = weekendCode(a.length >= 3 ? sc(a[2]) : 1), holidays = a.length === 4 ? holidaySerials(a[3]) : new Set(); let start = Math.trunc(n(a, 0)), end = Math.trunc(n(a, 1)), sign = 1; validateDateSerial(start, 'Start date'); validateDateSerial(end, 'End date'); if (start > end) {
        [start, end] = [end, start];
        sign = -1;
    } if (end - start > 366000)
        err('#NUM!', 'Date span too large.'); let z = 0; for (let i = start; i <= end; i++) {
        const dow = (excelWeekday(i) + 6) % 7;
        if (mask[dow] === '0' && !holidays.has(i))
            z++;
    } return sign * z; })
};
function vectors(a) { return flatten(a); }
const nsList = (a) => flatten(a).filter((x) => typeof x === 'number' && Number.isFinite(x));
function bit(x, y, fn) { for (const t of [x, y])
    if (!Number.isSafeInteger(t) || t < 0 || t >= 2 ** 48)
        err('#NUM!', 'Bitwise inputs require integers 0–2^48−1.'); const out = fn(BigInt(x), BigInt(y)); if (out < 0n || out >= 2n ** 48n)
    err('#NUM!', 'Bitwise result exceeds 48 bits.'); return Number(out); }
function shift(x, count) { if (!Number.isInteger(x) || x < 0 || x >= 2 ** 48 || !Number.isInteger(count) || Math.abs(count) > 53)
    err('#NUM!', 'Invalid shift parameters.'); const output = count >= 0 ? BigInt(x) << BigInt(count) : BigInt(x) >> BigInt(-count); if (output >= 2n ** 48n)
    err('#NUM!', 'Shift exceeds supported range.'); return Number(output); }
function ceilMath(a, down) { arg(a, 1, 3); const x = n(a, 0), step = a.length >= 2 ? Math.abs(n(a, 1)) : 1, mode = a.length === 3 ? n(a, 2) : 0; if (step === 0)
    return 0; const rounded = down ? Math.floor(x / step) : Math.ceil(x / step); if (x < 0 && mode !== 0)
    return (down ? Math.ceil(x / step) : Math.floor(x / step)) * step; return rounded * step; }
function precise(a, up) { arg(a, 1, 2); const step = a.length === 2 ? Math.abs(n(a, 1)) : 1; if (step === 0)
    return 0; return (up ? Math.ceil(n(a, 0) / step) : Math.floor(n(a, 0) / step)) * step; }
function ifStats(a, kind) { arg(a, 3, 255); const dimensions = (value) => { if (!Array.isArray(value))
    return [1, 1]; const width = value[0]?.length ?? 0; if (value.some(row => row.length !== width))
    err('#VALUE!', 'Ranges must be rectangular.'); return [value.length, width]; }, target = dimensions(a[0]), x = v(a[0]), criteria = criterionPairs(a, 1); if (x.length !== criteria.length)
    err('#VALUE!', 'Result and criteria ranges have different sizes.'); for (let i = 1; i < a.length; i += 2) {
    const current = dimensions(a[i]);
    if (current[0] !== target[0] || current[1] !== target[1])
        err('#VALUE!', 'Result and criteria ranges must have the same dimensions.');
} const hits = x.filter((value, i) => matchAll(criteria, i) && typeof value === 'number'); return hits.length ? (kind === 'min' ? Math.min(...hits) : Math.max(...hits)) : 0; }
function pRank(a, exclusive) { arg(a, 2, 3); const x = sorted(ns(a[0])), target = n(a, 1); if (x.length < 2 || target < x[0] || target > x[x.length - 1])
    err('#N/A', 'Value lies outside data.'); const i = x.findIndex(z => z >= target); const step = x[i] === target ? 0 : (target - x[i - 1]) / (x[i] - x[i - 1]); const position = x[i] === target ? i : i - 1 + step; const raw = exclusive ? (position + 1) / (x.length + 1) : position / (x.length - 1); const precision = a.length === 3 ? Math.trunc(n(a, 2)) : 3; return Number(raw.toFixed(Math.max(1, Math.min(12, precision)))); }
function forecast(a) { arg(a, 3); const x = ns(a[2]), y = ns(a[1]); if (x.length !== y.length || x.length < 2)
    err('#N/A', 'Forecast vectors must contain equal numbers of observations.'); const xm = avg(x), ym = avg(y), ss = sum(x.map(z => (z - xm) ** 2)); if (ss === 0)
    err('#DIV/0!', 'Predictors are constant.'); return ym + sum(x.map((z, i) => (z - xm) * (y[i] - ym))) / ss * (n(a, 0) - xm); }
const xmatchCompare = (left, right) => {
    const rank = (x) => x === null ? 0 : typeof x === 'number' ? 1 : typeof x === 'string' ? 2 : 3;
    const lr = rank(left), rr = rank(right);
    if (lr !== rr)
        return lr - rr;
    if (typeof left === 'number' && typeof right === 'number')
        return left < right ? -1 : left > right ? 1 : 0;
    if (typeof left === 'boolean' && typeof right === 'boolean')
        return Number(left) - Number(right);
    return text(left).toLowerCase().localeCompare(text(right).toLowerCase());
};
const xmatchWildcard = (pattern) => {
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
function xmatch(a) {
    arg(a, 2, 4);
    const target = sc(a[0]), data = v(a[1]), mode = a.length >= 3 ? n(a, 2) : 0, search = a.length === 4 ? n(a, 3) : 1;
    if (![-2, -1, 1, 2].includes(search))
        err('#VALUE!', 'XMATCH search_mode must be 1, -1, 2 or -2.');
    if (mode !== 0 && mode !== 1 && mode !== -1 && mode !== 2)
        err('#VALUE!', 'Invalid match mode.');
    const binary = search === 2 || search === -2, descending = search === -2;
    if (!data.length)
        return err('#N/A', 'XMATCH lookup array is empty.');
    if (binary && data.length > 1_048_576)
        err('#VALUE!', 'Binary XMATCH is limited to one Excel column of values.');
    if (binary) {
        for (let i = 1; i < data.length; i++) {
            const order = xmatchCompare(data[i - 1], data[i]);
            if (descending ? order < 0 : order > 0)
                err('#N/A', `XMATCH binary search requires ${descending ? 'descending' : 'ascending'} sorted data.`);
        }
        if (mode === 2) {
            // Arbitrary wildcard patterns cannot generally be narrowed to a contiguous
            // lexical interval, so scan once in the requested direction after validation.
            const rx = xmatchWildcard(text(target));
            if (descending) {
                for (let i = data.length - 1; i >= 0; i--)
                    if (rx.test(text(data[i])))
                        return i + 1;
            }
            else
                for (let i = 0; i < data.length; i++)
                    if (rx.test(text(data[i])))
                        return i + 1;
            return err('#N/A', 'XMATCH found no wildcard match.');
        }
        let lo = 0, hi = data.length;
        // Lower/upper bounds keep binary lookup logarithmic after the linear order check.
        if (mode === 0) {
            while (lo < hi) {
                const mid = lo + Math.floor((hi - lo) / 2), cmp = xmatchCompare(data[mid], target);
                if (descending ? cmp > 0 : cmp < 0)
                    lo = mid + 1;
                else
                    hi = mid;
            }
            return lo < data.length && xmatchCompare(data[lo], target) === 0 ? lo + 1 : err('#N/A', 'XMATCH found no exact match.');
        }
        if (!descending) {
            if (mode === 1) {
                while (lo < hi) {
                    const mid = lo + Math.floor((hi - lo) / 2);
                    if (xmatchCompare(data[mid], target) < 0)
                        lo = mid + 1;
                    else
                        hi = mid;
                }
                return lo < data.length ? lo + 1 : err('#N/A', 'XMATCH found no approximate match.');
            }
            while (lo < hi) {
                const mid = lo + Math.floor((hi - lo) / 2);
                if (xmatchCompare(data[mid], target) <= 0)
                    lo = mid + 1;
                else
                    hi = mid;
            }
            return lo > 0 ? lo : err('#N/A', 'XMATCH found no approximate match.');
        }
        if (mode === 1) {
            while (lo < hi) {
                const mid = lo + Math.floor((hi - lo) / 2);
                if (xmatchCompare(data[mid], target) >= 0)
                    lo = mid + 1;
                else
                    hi = mid;
            }
            return lo > 0 ? lo : err('#N/A', 'XMATCH found no approximate match.');
        }
        while (lo < hi) {
            const mid = lo + Math.floor((hi - lo) / 2);
            if (xmatchCompare(data[mid], target) > 0)
                lo = mid + 1;
            else
                hi = mid;
        }
        return lo < data.length ? lo + 1 : err('#N/A', 'XMATCH found no approximate match.');
    }
    const entries = data.map((val, i) => ({ val, i }));
    if (search === -1)
        entries.reverse();
    if (mode === 0 || mode === 2) {
        const rx = mode === 2 ? xmatchWildcard(text(target)) : null;
        const hit = entries.find(e => rx ? rx.test(text(e.val)) : typeof target === 'number' && typeof e.val === 'number' ? target === e.val : text(target).toLowerCase() === text(e.val).toLowerCase());
        return hit ? hit.i + 1 : err('#N/A', 'XMATCH found no exact match.');
    }
    const eligible = entries.filter(e => { const c = typeof target === 'number' && typeof e.val === 'number' ? e.val - target : text(e.val).localeCompare(text(target)); return mode === 1 ? c >= 0 : c <= 0; });
    eligible.sort((x, y) => { const cmp = typeof x.val === 'number' && typeof y.val === 'number' ? x.val - y.val : text(x.val).localeCompare(text(y.val)); return (mode === 1 ? cmp : -cmp) || x.i - y.i; });
    return eligible.length ? eligible[0].i + 1 : err('#N/A', 'XMATCH found no approximate match.');
}
function cumulative(a, interest) { arg(a, 6); const rate = n(a, 0), periods = Math.trunc(n(a, 1)), pv = n(a, 2), start = Math.trunc(n(a, 3)), end = Math.trunc(n(a, 4)), type = n(a, 5); if (rate <= 0 || periods <= 0 || pv <= 0 || start < 1 || end < start || end > periods || ![0, 1].includes(type))
    err('#NUM!', 'Invalid cumulative loan arguments.'); const pay = pm(rate, periods, pv, 0, type); let balance = pv, total = 0; for (let k = 1; k <= end; k++) {
    const i = type === 1 && k === 1 ? 0 : -balance * rate;
    const principal = pay - i;
    if (k >= start)
        total += interest ? i : principal;
    balance += pay - i;
} return finite(total); }
export const v6FunctionNames = Object.keys(v6Functions).sort();
export const v6Categories = {
    'Math & Engineering': 'BIN2DEC BIN2HEX BIN2OCT DEC2BIN DEC2HEX DEC2OCT HEX2BIN HEX2DEC HEX2OCT OCT2BIN OCT2DEC OCT2HEX SQRTPI SUMX2MY2 SUMX2PY2 SUMXMY2 COMBINA PERMUTATIONA BITAND BITOR BITXOR BITLSHIFT BITRSHIFT DELTA GESTEP ERF ERFC CEILING.MATH FLOOR.MATH CEILING.PRECISE FLOOR.PRECISE ISO.CEILING GAMMA GAMMALN GAMMALN.PRECISE SERIESSUM'.split(' '),
    'Statistical': 'MAXIFS MINIFS AVERAGEA MAXA MINA STDEV.S STDEV.P VAR.S VAR.P MODE.SNGL PERCENTILE.INC PERCENTILE.EXC PERCENTRANK.INC PERCENTRANK.EXC QUARTILE.INC QUARTILE.EXC RANK.EQ RANK.AVG FORECAST.LINEAR SKEW.P FISHER FISHERINV BINOM.DIST POISSON.DIST EXPON.DIST WEIBULL.DIST NORM.DIST NORM.INV NORM.S.DIST NORM.S.INV LOGNORM.DIST LOGNORM.INV CONFIDENCE.NORM PHI GAUSS Z.TEST'.split(' '),
    'Text': 'ENCODEURL ARRAYTOTEXT VALUETOTEXT REGEXEXTRACT'.split(' '),
    'Lookup & Reference': 'XMATCH ROW COLUMN FORMULATEXT INDIRECT OFFSET HYPERLINK AREAS CELL'.split(' '),
    'Information': 'ISFORMULA NA ERROR.TYPE ISREF'.split(' '),
    'Logical': ['IFNA'],
    'Financial': 'CUMIPMT CUMPRINC ISPMT DB VDB PDURATION DOLLARDE DOLLARFR PRICEDISC DISC'.split(' '),
    'Date & Time': 'WORKDAY.INTL NETWORKDAYS.INTL'.split(' ')
};
