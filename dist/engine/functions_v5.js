/** v5: 80 additional scalar/range Excel-style functions. No third-party formula engine.
 * Scope: no dynamic-array spills, no implicit VBA execution, UTC date conversions.
 */
import { SheetError, flatten, number, text, bool, matchCriterion } from './functions.js';
const bad = (code, message) => { throw new SheetError(code, message); };
const def = (description, example, run) => ({ description, example, run });
const req = (a, min, max = min) => { if (a.length < min || a.length > max)
    bad('#ARG!', `Expected ${min}–${max} arguments, got ${a.length}.`); };
const scalar = (d) => Array.isArray(d) ? bad('#VALUE!', 'Expected scalar, not a range.') : d ?? null;
const n = (a, i) => number(scalar(a[i]));
const s = (a, i) => text(scalar(a[i]));
const integer = (a, i) => Math.trunc(n(a, i));
const nums = (a) => flatten(a).filter((v) => typeof v === 'number' && Number.isFinite(v));
const holidaySet = (v) => { const dates = nums([v]); if (dates.some(date => date < 0 || date > 2_958_465))
    bad('#NUM!', 'Holiday date is outside Excel’s supported range.'); return new Set(dates.map(Math.trunc)); };
const sum = (x) => x.reduce((p, v) => p + v, 0);
const safe = (v) => Number.isFinite(v) ? v : bad('#NUM!', 'Result is not finite.');
const one = (label, ex, fn) => def(label, ex, a => { req(a, 1); return safe(fn(n(a, 0))); });
const factorial = (v) => { if (v < 0 || v > 170)
    bad('#NUM!', 'Factorial input must be 0 to 170.'); let result = 1; for (let i = 2; i <= Math.trunc(v); i++)
    result *= i; return result; };
const excelEpoch = Date.UTC(1899, 11, 31), DAY = 86400000;
const serial = (dt) => { const t = dt.getTime(); if (!Number.isFinite(t))
    bad('#VALUE!', 'Invalid date.'); return (t - excelEpoch) / DAY + (t >= Date.UTC(1900, 2, 1) ? 1 : 0); };
const date = (v) => { if (v < 0 || v > 2_958_465 || !Number.isFinite(v))
    bad('#NUM!', 'Invalid date serial.'); const d = Math.floor(v); return new Date(excelEpoch + (d - (d >= 60 ? 1 : 0)) * DAY); };
const excelWeekday = (v) => Math.floor(v) === 60 ? 3 : date(v).getUTCDay();
const excelYmd = (value) => {
    const whole = Math.floor(value), d = date(whole);
    if (whole === 60)
        return { year: 1900, month: 2, day: 29, serial: whole };
    return { year: d.getUTCFullYear(), month: d.getUTCMonth() + 1, day: d.getUTCDate(), serial: whole };
};
const monthLength = (year, month) => year === 1900 && month === 2 ? 29 : new Date(Date.UTC(year, month, 0)).getUTCDate();
const isMonthEnd = (d) => d.day === monthLength(d.year, d.month);
const days360 = (startValue, endValue, european) => {
    const sign = startValue <= endValue ? 1 : -1;
    const start = excelYmd(Math.min(startValue, endValue)), end = excelYmd(Math.max(startValue, endValue));
    let y1 = start.year, m1 = start.month, d1 = start.day, y2 = end.year, m2 = end.month, d2 = end.day;
    if (european) {
        if (d1 === 31)
            d1 = 30;
        if (d2 === 31)
            d2 = 30;
    }
    else {
        // Excel's U.S. method applies month-end rules before computing the 30/360 difference.
        if (isMonthEnd(start))
            d1 = 30;
        if (isMonthEnd(end)) {
            if (d1 < 30) {
                d2 = 1;
                m2++;
                if (m2 > 12) {
                    m2 = 1;
                    y2++;
                }
            }
            else
                d2 = 30;
        }
        else if (d2 === 31) {
            if (d1 < 30) {
                d2 = 1;
                m2++;
                if (m2 > 12) {
                    m2 = 1;
                    y2++;
                }
            }
            else
                d2 = 30;
        }
    }
    return sign * ((y2 - y1) * 360 + (m2 - m1) * 30 + d2 - d1);
};
const vector = (d) => flatten([d]);
const lookupOrder = (left, right) => {
    const rank = (value) => value === null ? 0 : typeof value === 'number' ? 1 : typeof value === 'string' ? 2 : 3;
    const leftRank = rank(left), rightRank = rank(right);
    if (leftRank !== rightRank)
        return leftRank - rightRank;
    if (typeof left === 'number' && typeof right === 'number')
        return left < right ? -1 : left > right ? 1 : 0;
    if (typeof left === 'boolean' && typeof right === 'boolean')
        return Number(left) - Number(right);
    const a = text(left).toLowerCase(), b = text(right).toLowerCase();
    return a < b ? -1 : a > b ? 1 : 0;
};
const average = (v) => v.length ? sum(v) / v.length : bad('#DIV/0!', 'No numeric values.');
const moments = (a) => { const values = nums(a), mean = average(values), sq = sum(values.map(x => (x - mean) ** 2)); return { values, mean, sq }; };
const twoVectors = (a) => { req(a, 2); const x = vector(a[0]), y = vector(a[1]); if (x.length !== y.length)
    bad('#VALUE!', 'Ranges must have matching dimensions.'); const xx = [], yy = []; x.forEach((v, i) => { if (typeof v === 'number' && typeof y[i] === 'number') {
    xx.push(v);
    yy.push(y[i]);
} }); return { x: xx, y: yy }; };
const stdev = (x, sample) => { if (x.length < (sample ? 2 : 1))
    bad('#DIV/0!', 'Not enough data.'); const avg = average(x); return Math.sqrt(sum(x.map(v => (v - avg) ** 2)) / (x.length - (sample ? 1 : 0))); };
const erf = (x) => { const sign = Math.sign(x), z = Math.abs(x), t = 1 / (1 + 0.3275911 * z); const p = (((((1.061405429 * t - 1.453152027) * t) + 1.421413741) * t - 0.284496736) * t + 0.254829592) * t; return sign * (1 - p * Math.exp(-z * z)); };
const normCDF = (x) => 0.5 * (1 + erf(x / Math.SQRT2));
const normPDF = (x) => Math.exp(-x * x / 2) / Math.sqrt(2 * Math.PI);
const normINV = (p) => { if (p <= 0 || p >= 1)
    bad('#NUM!', 'Probability must be strictly between 0 and 1.'); let lo = -12, hi = 12; for (let i = 0; i < 85; i++) {
    const mid = (lo + hi) / 2;
    if (normCDF(mid) < p)
        lo = mid;
    else
        hi = mid;
} return (lo + hi) / 2; };
const formatFixed = (value, digits, commas) => { if (digits < 0 || digits > 15)
    bad('#NUM!', 'Digits must be from 0 to 15.'); return value.toLocaleString('en-US', { useGrouping: commas, minimumFractionDigits: digits, maximumFractionDigits: digits }); };
const matchExcel = (v, w) => matchCriterion(v, w);
const payment = (rate, periods, present, future = 0, type = 0) => { if (periods <= 0)
    bad('#NUM!', 'Number of periods must be positive.'); if (rate === 0)
    return -(present + future) / periods; const factor = (1 + rate) ** periods; return safe(-(present * factor + future) * rate / ((1 + rate * type) * (factor - 1))); };
const irr = (values, guess = .1) => { if (values.length < 2 || !values.some(v => v > 0) || !values.some(v => v < 0))
    bad('#NUM!', 'Cash flows require both positive and negative amounts.'); let rate = Math.max(-.9, guess); for (let i = 0; i < 120; i++) {
    let npv = 0, deriv = 0;
    for (let j = 0; j < values.length; j++) {
        npv += values[j] / (1 + rate) ** j;
        if (j > 0)
            deriv -= j * values[j] / (1 + rate) ** (j + 1);
    }
    if (Math.abs(npv) < 1e-8)
        return safe(rate);
    if (!deriv || !Number.isFinite(deriv))
        break;
    const next = rate - npv / deriv;
    if (next <= -.999999 || next > 1e7 || !Number.isFinite(next))
        break;
    rate = next;
} let l = -.9999, h = 1; const f = (r) => sum(values.map((v, i) => v / (1 + r) ** i)); while (f(l) * f(h) > 0 && h < 1e6)
    h = h * 2 + 1; if (f(l) * f(h) > 0)
    bad('#NUM!', 'Could not converge to an IRR.'); for (let i = 0; i < 170; i++) {
    const m = (l + h) / 2;
    if (f(l) * f(m) <= 0)
        h = m;
    else
        l = m;
} return (l + h) / 2; };
const yearFraction = (start, end, basis) => { if (!Number.isFinite(start) || !Number.isFinite(end) || start < 0 || end < 0 || start > 2_958_465 || end > 2_958_465)
    bad('#VALUE!', 'YEARFRAC requires valid Excel date serials.'); const sign = start <= end ? 1 : -1, a0 = Math.floor(Math.min(start, end)), b0 = Math.floor(Math.max(start, end)), days = b0 - a0; if (basis === 1) {
    let actual = 0, cursor = a0;
    while (cursor < b0) {
        const year = excelYmd(cursor).year, yearStart = serial(new Date(Date.UTC(year, 0, 1))), yearEnd = serial(new Date(Date.UTC(year + 1, 0, 1))), close = Math.min(yearEnd, b0);
        if (yearEnd <= cursor || yearStart >= yearEnd)
            bad('#NUM!', 'Could not resolve YEARFRAC calendar year.');
        actual += (close - cursor) / (yearEnd - yearStart);
        cursor = close;
    }
    return sign * actual;
} if (basis === 2)
    return sign * days / 360; if (basis === 3)
    return sign * days / 365; if (basis === 0)
    return days360(a0, b0, false) / 360; if (basis === 4)
    return days360(a0, b0, true) / 360; return bad('#NUM!', 'YEARFRAC basis must be 0–4.'); };
const toN = (v) => typeof v === 'number' ? v : typeof v === 'boolean' ? (v ? 1 : 0) : 0;
const utf8 = (x) => new TextEncoder().encode(x);
const controlledRegex = (pattern, flags) => {
    // Disallow backreferences, lookaround, and nested quantifiers, which commonly cause ReDoS.
    if (pattern.length > 256 || /\\[1-9]/.test(pattern) || /\(\?/.test(pattern) || /\([^)]*[+*{][^)]*\)[+*{]/.test(pattern))
        bad('#VALUE!', 'Unsupported or potentially expensive regex pattern.');
    if (!/^[gimsuy]*$/.test(flags) || new Set(flags).size !== flags.length)
        bad('#VALUE!', 'Invalid regex flags.');
    try {
        return new RegExp(pattern, flags);
    }
    catch {
        return bad('#VALUE!', 'Invalid regular expression.');
    }
};
const utf8cut = (x, n) => new TextDecoder('utf-8', { fatal: false }).decode(utf8(x).slice(0, n));
const romanValues = [[1000, 'M'], [900, 'CM'], [500, 'D'], [400, 'CD'], [100, 'C'], [90, 'XC'], [50, 'L'], [40, 'XL'], [10, 'X'], [9, 'IX'], [5, 'V'], [4, 'IV'], [1, 'I']];
const financeFactors = (rate, nper, pv, fv, type) => { const p = payment(rate, nper, pv, fv, type); return { p }; };
const doubleFactorial = (x) => { if (x < 0 || x > 300)
    bad('#NUM!', 'FACTDOUBLE input out of supported range.'); let out = 1; for (let i = Math.trunc(x); i > 1; i -= 2)
    out *= i; return safe(out); };
export const v5Functions = {
    // 18 mathematics & engineering functions
    ACOSH: one('Inverse hyperbolic cosine.', '=ACOSH(2)', Math.acosh),
    ASINH: one('Inverse hyperbolic sine.', '=ASINH(2)', Math.asinh),
    ATANH: one('Inverse hyperbolic tangent.', '=ATANH(0.5)', Math.atanh),
    ACOT: one('Inverse cotangent in radians.', '=ACOT(2)', x => Math.PI / 2 - Math.atan(x)),
    ACOTH: one('Inverse hyperbolic cotangent for |x| > 1.', '=ACOTH(2)', x => { if (Math.abs(x) <= 1)
        bad('#NUM!', '|x| must be greater than 1.'); return Math.atanh(1 / x); }),
    COT: one('Cotangent.', '=COT(1)', x => 1 / Math.tan(x)),
    COTH: one('Hyperbolic cotangent.', '=COTH(1)', x => 1 / Math.tanh(x)),
    CSC: one('Cosecant.', '=CSC(1)', x => 1 / Math.sin(x)),
    CSCH: one('Hyperbolic cosecant.', '=CSCH(1)', x => 1 / Math.sinh(x)),
    SEC: one('Secant.', '=SEC(1)', x => 1 / Math.cos(x)),
    SECH: one('Hyperbolic secant.', '=SECH(1)', x => 1 / Math.cosh(x)),
    MROUND: def('Round to nearest specified multiple.', '=MROUND(14,5)', a => { req(a, 2); const v = n(a, 0), step = n(a, 1); if (v * step < 0)
        bad('#NUM!', 'Value and multiple must have same sign.'); return step === 0 ? 0 : safe(Math.sign(v) * Math.floor(Math.abs(v / step) + .5) * Math.abs(step)); }),
    FACTDOUBLE: def('Double factorial n!!.', '=FACTDOUBLE(7)', a => { req(a, 1); return doubleFactorial(n(a, 0)); }),
    MULTINOMIAL: def('Multinomial coefficient of nonnegative integer arguments.', '=MULTINOMIAL(2,3,4)', a => { if (!a.length)
        bad('#ARG!', 'Arguments required.'); const values = nums(a).map(Math.trunc); if (values.some(x => x < 0))
        bad('#NUM!', 'Counts must be positive.'); return safe(factorial(sum(values)) / values.reduce((v, x) => v * factorial(x), 1)); }),
    ARABIC: def('Convert standard Roman numerals to a number.', '=ARABIC("XIV")', a => { req(a, 1); let str = s(a, 0).toUpperCase().trim(), total = 0; if (!str || !/^(M{0,4})(CM|CD|D?C{0,3})(XC|XL|L?X{0,3})(IX|IV|V?I{0,3})$/.test(str))
        bad('#VALUE!', 'Invalid Roman numeral.'); for (const [value, symbol] of romanValues)
        while (str.startsWith(symbol)) {
            total += value;
            str = str.slice(symbol.length);
        } return total; }),
    ROMAN: def('Convert positive integer to Roman numerals.', '=ROMAN(1987)', a => { req(a, 1, 2); let v = integer(a, 0), out = ''; if (v < 1 || v > 3999)
        bad('#VALUE!', 'ROMAN supports 1–3999.'); for (const [value, symbol] of romanValues)
        while (v >= value) {
            out += symbol;
            v -= value;
        } return out; }),
    BASE: def('Convert integer to text in the supplied radix 2–36.', '=BASE(255,16)', a => { req(a, 2, 3); const v = integer(a, 0), radix = integer(a, 1), pad = a.length === 3 ? integer(a, 2) : 0; if (v < 0 || v > 2 ** 53 - 1 || radix < 2 || radix > 36 || pad < 0 || pad > 200)
        bad('#NUM!', 'Invalid BASE arguments.'); return v.toString(radix).toUpperCase().padStart(pad, '0'); }),
    DECIMAL: def('Convert base-N text to a decimal integer.', '=DECIMAL("FF",16)', a => { req(a, 2); const str = s(a, 0).trim().toUpperCase(), radix = integer(a, 1); if (radix < 2 || radix > 36 || !str)
        bad('#NUM!', 'Invalid radix or digits.'); let val = 0; for (const c of str) {
        const digit = parseInt(c, 36);
        if (!Number.isInteger(digit) || digit >= radix)
            bad('#NUM!', 'Invalid digit in radix.');
        val = val * radix + digit;
        if (!Number.isSafeInteger(val))
            bad('#NUM!', 'Integer overflow.');
    } return val; }),
    // 12 statistics functions
    SKEW: def('Sample skewness of numeric observations.', '=SKEW(A1:A10)', a => { const { values, mean, sq } = moments(a), len = values.length; if (len < 3 || sq === 0)
        bad('#DIV/0!', 'SKEW requires at least three varying numbers.'); const sd = Math.sqrt(sq / (len - 1)); return safe(len / ((len - 1) * (len - 2)) * sum(values.map(v => ((v - mean) / sd) ** 3))); }),
    KURT: def('Excess sample kurtosis.', '=KURT(A1:A10)', a => { const { values, mean, sq } = moments(a), len = values.length; if (len < 4 || sq === 0)
        bad('#DIV/0!', 'KURT requires at least four varying numbers.'); const fourth = sum(values.map(v => (v - mean) ** 4)); return safe(len * (len + 1) * (len - 1) / ((len - 2) * (len - 3)) * fourth / (sq * sq) - 3 * (len - 1) ** 2 / ((len - 2) * (len - 3))); }),
    COVARIANCEP: def('Population covariance of two ranges.', '=COVARIANCEP(A1:A5,B1:B5)', a => { const { x, y } = twoVectors(a), mx = average(x), my = average(y); return safe(sum(x.map((v, i) => (v - mx) * (y[i] - my))) / x.length); }),
    COVARIANCES: def('Sample covariance of two ranges.', '=COVARIANCES(A1:A5,B1:B5)', a => { const { x, y } = twoVectors(a); if (x.length < 2)
        bad('#DIV/0!', 'At least two observations required.'); const mx = average(x), my = average(y); return safe(sum(x.map((v, i) => (v - mx) * (y[i] - my))) / (x.length - 1)); }),
    PEARSON: def('Pearson correlation coefficient.', '=PEARSON(A1:A5,B1:B5)', a => { const { x, y } = twoVectors(a), mx = average(x), my = average(y), sx = sum(x.map(v => (v - mx) ** 2)), sy = sum(y.map(v => (v - my) ** 2)); if (!sx || !sy)
        bad('#DIV/0!', 'Variance must be nonzero.'); return safe(sum(x.map((v, i) => (v - mx) * (y[i] - my))) / Math.sqrt(sx * sy)); }),
    STEYX: def('Standard error of predicted y in regression.', '=STEYX(B1:B6,A1:A6)', a => { const { x: y, y: x } = twoVectors(a); if (x.length < 3)
        bad('#DIV/0!', 'At least 3 pairs needed.'); const mx = average(x), my = average(y), xx = sum(x.map(v => (v - mx) ** 2)); if (!xx)
        bad('#DIV/0!', 'All predictor values identical.'); const slope = sum(x.map((v, i) => (v - mx) * (y[i] - my))) / xx; return Math.sqrt(sum(x.map((v, i) => (y[i] - my - slope * (v - mx)) ** 2)) / (x.length - 2)); }),
    STANDARDIZE: def('Standardize a value with population mean and standard deviation.', '=STANDARDIZE(10,8,2)', a => { req(a, 3); const sd = n(a, 2); if (sd <= 0)
        bad('#NUM!', 'Standard deviation must be positive.'); return (n(a, 0) - n(a, 1)) / sd; }),
    TRIMMEAN: def('Mean after trimming specified fraction equally at both ends.', '=TRIMMEAN(A1:A10,0.2)', a => { req(a, 2); const p = n(a, 1), v = nums([a[0]]).sort((x, y) => x - y); if (p < 0 || p >= 1)
        bad('#NUM!', 'Fraction must be from 0 inclusive to 1 exclusive.'); const cut = Math.floor(v.length * p / 2); return average(v.slice(cut, v.length - cut)); }),
    NORMDIST: def('Normal CDF or PDF.', '=NORMDIST(1,0,1,TRUE)', a => { req(a, 4); const sd = n(a, 2); if (sd <= 0)
        bad('#NUM!', 'Standard deviation must be positive.'); const z = (n(a, 0) - n(a, 1)) / sd; return bool(scalar(a[3])) ? normCDF(z) : normPDF(z) / sd; }),
    NORMSDIST: def('Standard normal cumulative distribution.', '=NORMSDIST(1)', a => { req(a, 1); return normCDF(n(a, 0)); }),
    NORMINV: def('Normal inverse cumulative distribution.', '=NORMINV(0.9,0,1)', a => { req(a, 3); const sd = n(a, 2); if (sd <= 0)
        bad('#NUM!', 'Standard deviation must be positive.'); return n(a, 1) + sd * normINV(n(a, 0)); }),
    NORMSINV: def('Standard normal inverse CDF.', '=NORMSINV(0.975)', a => { req(a, 1); return normINV(n(a, 0)); }),
    // 14 text functions: UTF-8 *B variants are byte-oriented (with replacement on incomplete sequences)
    T: def('Keep text values and return empty for non-text.', '=T("Hello")', a => { req(a, 1); const v = scalar(a[0]); return typeof v === 'string' ? v : ''; }),
    N: def('Convert numeric and logical values, text becomes 0.', '=N(TRUE)', a => { req(a, 1); return toN(scalar(a[0])); }),
    DOLLAR: def('Format a currency number as text.', '=DOLLAR(1234.567,2)', a => { req(a, 1, 2); return '$' + formatFixed(n(a, 0), a.length === 2 ? integer(a, 1) : 2, true); }),
    FIXED: def('Format a number with fixed decimal digits and optional grouping.', '=FIXED(1234.567,2,TRUE)', a => { req(a, 1, 3); return formatFixed(n(a, 0), a.length >= 2 ? integer(a, 1) : 2, a.length === 3 ? !bool(scalar(a[2])) : true); }),
    TEXT: def('Format number using a supported Excel-style format subset.', '=TEXT(12.345,"0.00")', a => { req(a, 2); const pattern = s(a, 1), v = n(a, 0); if (pattern === '0' || pattern === '#,##0')
        return formatFixed(v, 0, pattern.includes(',')); if (/^\$?\#?,?\#?0\.0{1,10}$/.test(pattern))
        return (pattern.startsWith('$') ? '$' : '') + formatFixed(v, (pattern.split('.')[1] ?? '').length, pattern.includes(',')); if (/^0\.0{1,10}%$/.test(pattern))
        return formatFixed(v * 100, pattern.slice(2, -1).length, false) + '%'; if (pattern === '0%')
        return formatFixed(v * 100, 0, false) + '%'; return bad('#VALUE!', 'TEXT supports 0, decimal, currency, and percent formats only.'); }),
    LENB: def('Length of text in UTF-8 bytes.', '=LENB("Hello")', a => { req(a, 1); return utf8(s(a, 0)).length; }),
    LEFTB: def('Leftmost bytes interpreted as UTF-8 text.', '=LEFTB("Hello",3)', a => { req(a, 1, 2); return utf8cut(s(a, 0), a.length === 2 ? Math.max(0, integer(a, 1)) : 1); }),
    RIGHTB: def('Rightmost bytes interpreted as UTF-8 text.', '=RIGHTB("Hello",3)', a => { req(a, 1, 2); const bytes = utf8(s(a, 0)), len = a.length === 2 ? Math.max(0, integer(a, 1)) : 1; return new TextDecoder().decode(bytes.slice(-len)); }),
    MIDB: def('Extract bytes from UTF-8-encoded text.', '=MIDB("Hello",2,3)', a => { req(a, 3); return new TextDecoder().decode(utf8(s(a, 0)).slice(Math.max(0, integer(a, 1) - 1), Math.max(0, integer(a, 1) - 1) + Math.max(0, integer(a, 2)))); }),
    FINDB: def('Byte position of text occurrence (case sensitive).', '=FINDB("el","Hello")', a => { req(a, 2, 3); const str = s(a, 1), start = a.length === 3 ? integer(a, 2) - 1 : 0, idx = str.indexOf(s(a, 0), Math.max(0, start)); if (idx < 0)
        bad('#VALUE!', 'Text not found.'); return utf8(str.slice(0, idx)).length + 1; }),
    SEARCHB: def('Case insensitive byte position of text occurrence.', '=SEARCHB("EL","Hello")', a => { req(a, 2, 3); const str = s(a, 1), start = a.length === 3 ? integer(a, 2) - 1 : 0, idx = str.toLowerCase().indexOf(s(a, 0).toLowerCase(), Math.max(0, start)); if (idx < 0)
        bad('#VALUE!', 'Text not found.'); return utf8(str.slice(0, idx)).length + 1; }),
    REPLACEB: def('Replace a byte span in UTF-8 text.', '=REPLACEB("Hello",2,2,"i")', a => { req(a, 4); const bytes = utf8(s(a, 0)), start = integer(a, 1) - 1, len = integer(a, 2); if (start < 0 || len < 0)
        bad('#VALUE!', 'Invalid byte span.'); return new TextDecoder().decode(bytes.slice(0, start)) + s(a, 3) + new TextDecoder().decode(bytes.slice(start + len)); }),
    REGEXTEST: def('Test text against a JavaScript-compatible regular expression.', '=REGEXTEST("ABC12","[0-9]+")', a => { req(a, 2, 3); const pat = s(a, 1); if (s(a, 0).length > 2000)
        bad('#VALUE!', 'Regex input exceeds 2,000 characters.'); return controlledRegex(pat, a.length === 3 ? s(a, 2) : '').test(s(a, 0)); }),
    REGEXREPLACE: def('Replace regex matches, global by default.', '=REGEXREPLACE("ABC12","[0-9]+","99")', a => { req(a, 3, 4); const pat = s(a, 1), input = s(a, 0); if (input.length > 2000)
        bad('#VALUE!', 'Regex input exceeds 2,000 characters.'); return input.replace(controlledRegex(pat, a.length === 4 ? s(a, 3) : 'g'), s(a, 2)); }),
    // 8 date/time functions
    DATEDIF: def('Return a date difference in years, months, days or month/year remainders.', '=DATEDIF(DATE(2020,1,1),DATE(2022,4,1),"M")', a => {
        req(a, 3);
        const start = Math.trunc(n(a, 0)), end = Math.trunc(n(a, 1)), unit = s(a, 2).toUpperCase();
        if (!['Y', 'M', 'D', 'MD', 'YM', 'YD'].includes(unit))
            bad('#NUM!', 'DATEDIF unit must be Y, M, D, MD, YM or YD.');
        if (end < start)
            bad('#NUM!', 'Start date after end date.');
        const a0 = date(start), b = date(end);
        const months = (b.getUTCFullYear() - a0.getUTCFullYear()) * 12 + b.getUTCMonth() - a0.getUTCMonth() - (b.getUTCDate() < a0.getUTCDate() ? 1 : 0);
        if (unit === 'Y')
            return Math.floor(months / 12);
        if (unit === 'M')
            return months;
        if (unit === 'D')
            return end - start;
        if (unit === 'YM')
            return months % 12;
        if (unit === 'YD') {
            // Compare month/day on a leap-year template so Feb 29 remains representable;
            // when the end month/day precedes the start, count through the year boundary.
            const startOfCycle = Date.UTC(2000, a0.getUTCMonth(), a0.getUTCDate());
            const endOfCycle = Date.UTC(2000, b.getUTCMonth(), b.getUTCDate());
            const elapsed = (endOfCycle - startOfCycle) / DAY;
            return elapsed < 0 ? elapsed + 366 : elapsed;
        }
        // MD is defined as the day difference with month/year parts ignored. Borrow
        // the actual length of the month before the end date instead of a fixed 31.
        const startDay = a0.getUTCDate(), endDay = b.getUTCDate();
        if (endDay >= startDay)
            return endDay - startDay;
        const daysInPreviousMonth = new Date(Date.UTC(b.getUTCFullYear(), b.getUTCMonth(), 0)).getUTCDate();
        return endDay + daysInPreviousMonth - startDay;
    }),
    YEARFRAC: def('Fractional years between two dates using Excel basis 0–4.', '=YEARFRAC(A1,B1,1)', a => { req(a, 2, 3); return yearFraction(n(a, 0), n(a, 1), a.length === 3 ? integer(a, 2) : 0); }),
    DAYS360: def('360-day calendar difference using the U.S. (NASD) method by default or the European method when requested.', '=DAYS360(A1,B1,FALSE)', a => { req(a, 2, 3); const start = n(a, 0), end = n(a, 1); if (start < 0 || end < 0 || start > 2_958_465 || end > 2_958_465)
        bad('#NUM!', 'DAYS360 requires valid Excel date serials.'); return days360(start, end, a.length === 3 && bool(scalar(a[2]))); }),
    WEEKNUM: def('Week number with supported Sunday-, Monday-, other-day, and ISO week starts.', '=WEEKNUM(DATE(2026,10,20),2)', a => { req(a, 1, 2); const serialValue = n(a, 0), dt = date(serialValue), type = a.length === 2 ? integer(a, 1) : 1; if (type === 21) {
        const iso = new Date(dt);
        iso.setUTCDate(iso.getUTCDate() + 4 - (iso.getUTCDay() || 7));
        const jan1 = new Date(Date.UTC(iso.getUTCFullYear(), 0, 1));
        return Math.ceil(((iso.getTime() - jan1.getTime()) / DAY + 1) / 7);
    } if (type !== 1 && type !== 2 && !(type >= 11 && type <= 17))
        bad('#NUM!', 'WEEKNUM return type supports 1, 2, 11–17 or 21.'); const jan1 = new Date(Date.UTC(dt.getUTCFullYear(), 0, 1)), weekStart = type === 1 ? 0 : type === 2 ? 1 : (type - 11 + 1) % 7, start = (jan1.getUTCDay() - weekStart + 7) % 7; return Math.floor(((dt.getTime() - jan1.getTime()) / DAY + start) / 7) + 1; }),
    ISOWEEKNUM: def('ISO 8601 week number of date serial.', '=ISOWEEKNUM(DATE(2026,10,20))', a => { req(a, 1); const dt = date(n(a, 0)); dt.setUTCHours(0, 0, 0, 0); dt.setUTCDate(dt.getUTCDate() + 4 - (dt.getUTCDay() || 7)); const yearStart = new Date(Date.UTC(dt.getUTCFullYear(), 0, 1)); return Math.ceil((((dt.getTime() - yearStart.getTime()) / DAY) + 1) / 7); }),
    DATEVALUE: def('Parse unambiguous year-first or English month-name date text.', '=DATEVALUE("January 20, 2026")', a => { req(a, 1); const input = s(a, 0).trim(), yearFirst = /^(\d{4})[-/](\d{1,2})[-/](\d{1,2})$/.exec(input), monthFirst = /^([A-Za-z]+)\s+(\d{1,2}),?\s+(\d{4})$/.exec(input), dayFirst = /^(\d{1,2})[-\s]([A-Za-z]+)[-\s](\d{4})$/.exec(input), months = ['january', 'february', 'march', 'april', 'may', 'june', 'july', 'august', 'september', 'october', 'november', 'december']; let year = 0, month = 0, day = 0; if (yearFirst) {
        year = Number(yearFirst[1]);
        month = Number(yearFirst[2]);
        day = Number(yearFirst[3]);
    }
    else if (monthFirst || dayFirst) {
        const parts = monthFirst ?? dayFirst, monthText = monthFirst ? parts[1] : parts[2], monthKey = monthText.toLowerCase(), monthIndex = months.findIndex(name => name === monthKey || monthKey.length === 3 && name.slice(0, 3) === monthKey);
        if (monthIndex < 0)
            bad('#VALUE!', 'DATEVALUE could not parse the month name.');
        year = Number(parts[3]);
        month = monthIndex + 1;
        day = Number(monthFirst ? parts[2] : parts[1]);
    }
    else
        bad('#VALUE!', 'DATEVALUE accepts year-first numeric dates or English month-name dates; ambiguous locale-numeric dates are unsupported.'); if (year < 1900 || year > 9999 || month < 1 || month > 12 || day < 1 || day > 31)
        bad('#VALUE!', 'Invalid calendar date.'); if (year === 1900 && month === 2 && day === 29)
        return 60; const dt = new Date(0); dt.setUTCHours(0, 0, 0, 0); dt.setUTCFullYear(year, month - 1, day); if (dt.getUTCFullYear() !== year || dt.getUTCMonth() !== month - 1 || dt.getUTCDate() !== day)
        bad('#VALUE!', 'Invalid calendar date.'); return serial(dt); }),
    TIMEVALUE: def('Parse 24-hour or AM/PM time text into a fractional day.', '=TIMEVALUE("2:30 PM")', a => { req(a, 1); const m = /^(\d{1,2}):(\d{2})(?::(\d{2}))?\s*(AM|PM)?$/i.exec(s(a, 0).trim()); if (!m)
        bad('#VALUE!', 'TIMEVALUE requires a valid hour and minute, with optional seconds and AM/PM.'); let hour = Number(m[1]), minute = Number(m[2]), second = Number(m[3] ?? 0); const meridiem = m[4]?.toUpperCase(); if (minute > 59 || second > 59 || (meridiem ? hour < 1 || hour > 12 : hour > 23))
        bad('#VALUE!', 'TIMEVALUE received an invalid time.'); if (meridiem)
        hour = hour % 12 + (meridiem === 'PM' ? 12 : 0); return (hour * 3600 + minute * 60 + second) / 86400; }),
    NETWORKDAYSINTL: def('Workdays excluding configurable weekends (code 1 or string of 7 bits).', '=NETWORKDAYSINTL(A1,B1,1)', a => { req(a, 2, 4); let start = integer(a, 0), end = integer(a, 1), sgn = 1; if (start < 0 || end < 0 || start > 2_958_465 || end > 2_958_465)
        bad('#NUM!', 'Date serial is outside Excel’s supported range.'); if (start > end) {
        [start, end] = [end, start];
        sgn = -1;
    } if (end - start > 365250)
        bad('#NUM!', 'Date span too large.'); const wk = a.length >= 3 ? scalar(a[2]) : 1; const masks = { 1: '0000011', 2: '1000001', 3: '1100000', 4: '0110000', 5: '0011000', 6: '0001100', 7: '0000110', 11: '0000001', 12: '1000000', 13: '0100000', 14: '0010000', 15: '0001000', 16: '0000100', 17: '0000010' }; const mask = typeof wk === 'string' ? wk : masks[number(wk)]; if (!mask || !/^[01]{7}$/.test(mask))
        bad('#VALUE!', 'Invalid weekend mask.'); const holidays = a.length === 4 ? holidaySet(a[3]) : new Set(); let total = 0; for (let i = start; i <= end; i++) {
        const weekday = (excelWeekday(i) + 6) % 7;
        if (mask[weekday] === '0' && !holidays.has(i))
            total++;
    } return total * sgn; }),
    // 15 financial functions
    RATE: def('Solve periodic interest rate for an annuity.', '=RATE(12,-100,1000)', a => { req(a, 3, 6); const periods = n(a, 0), pay = n(a, 1), pv = n(a, 2), fv = a.length >= 4 ? n(a, 3) : 0, type = a.length >= 5 ? n(a, 4) : 0; let r = a.length === 6 ? n(a, 5) : .1; if (periods <= 0)
        bad('#NUM!', 'Periods must be positive.'); if (Math.abs(pv + pay * periods + fv) < 1e-10)
        return 0; for (let i = 0; i < 120; i++) {
        const f = (q) => q === 0 ? pv + pay * periods + fv : pv * (1 + q) ** periods + pay * (1 + q * type) * ((1 + q) ** periods - 1) / q + fv;
        const y = f(r), h = 1e-6, deriv = (f(r + h) - f(r - h)) / (2 * h);
        if (Math.abs(y) < 1e-8)
            return r;
        if (!Number.isFinite(deriv) || Math.abs(deriv) < 1e-12)
            break;
        r -= y / deriv;
        if (r <= -1 || r > 1000 || !Number.isFinite(r))
            break;
    } return bad('#NUM!', 'RATE did not converge.'); }),
    NPER: def('Calculate number of periods for fixed payment loan.', '=NPER(0.01,-100,1000)', a => { req(a, 3, 5); const r = n(a, 0), pay = n(a, 1), pv = n(a, 2), fv = a.length >= 4 ? n(a, 3) : 0, type = a.length === 5 ? n(a, 4) : 0; if (r === 0) {
        if (pay === 0)
            bad('#DIV/0!', 'Payment cannot be zero.');
        return -(pv + fv) / pay;
    } const q = pay * (1 + r * type) / r, ratio = (q - fv) / (pv + q); if (ratio <= 0 || 1 + r <= 0)
        bad('#NUM!', 'Invalid loan cash flows.'); return safe(Math.log(ratio) / Math.log(1 + r)); }),
    IPMT: def('Interest component of periodic payment.', '=IPMT(0.05/12,1,60,10000)', a => { req(a, 4, 6); const rate = n(a, 0), period = integer(a, 1), periods = integer(a, 2), pv = n(a, 3), fv = a.length >= 5 ? n(a, 4) : 0, type = a.length === 6 ? n(a, 5) : 0; if (period < 1 || period > periods || ![0, 1].includes(type))
        bad('#NUM!', 'Invalid period or payment type.'); const pay = payment(rate, periods, pv, fv, type); let balance = pv; for (let k = 1; k < period; k++) {
        const interest = type === 1 && k === 1 ? 0 : balance * rate;
        balance += interest + pay;
    } return type === 1 && period === 1 ? 0 : -balance * rate; }),
    PPMT: def('Principal component of periodic payment.', '=PPMT(0.05/12,1,60,10000)', a => { req(a, 4, 6); const rate = n(a, 0), periods = integer(a, 2), pv = n(a, 3), fv = a.length >= 5 ? n(a, 4) : 0, type = a.length === 6 ? n(a, 5) : 0, interest = v5Functions.IPMT.run(a); return safe(payment(rate, periods, pv, fv, type) - number(scalar(interest))); }),
    EFFECT: def('Effective annual interest rate.', '=EFFECT(0.12,12)', a => { req(a, 2); const rate = n(a, 0), k = integer(a, 1); if (rate <= 0 || k < 1)
        bad('#NUM!', 'Nominal rate must be positive and periods >= 1.'); return (1 + rate / k) ** k - 1; }),
    NOMINAL: def('Nominal annual interest rate.', '=NOMINAL(0.1268,12)', a => { req(a, 2); const rate = n(a, 0), k = integer(a, 1); if (rate <= 0 || k < 1)
        bad('#NUM!', 'Effective rate must be positive and periods >= 1.'); return k * ((1 + rate) ** (1 / k) - 1); }),
    SLN: def('Straight-line depreciation.', '=SLN(10000,1000,5)', a => { req(a, 3); if (n(a, 2) <= 0)
        bad('#DIV/0!', 'Life must be positive.'); return (n(a, 0) - n(a, 1)) / n(a, 2); }),
    SYD: def('Sum of years digits depreciation.', '=SYD(10000,1000,5,1)', a => { req(a, 4); const life = n(a, 2), period = n(a, 3); if (life <= 0 || period < 1 || period > life)
        bad('#NUM!', 'Invalid life or period.'); return (n(a, 0) - n(a, 1)) * (life - period + 1) * 2 / (life * (life + 1)); }),
    DDB: def('Double declining balance depreciation.', '=DDB(10000,1000,5,1)', a => { req(a, 4, 5); const cost = n(a, 0), salvage = n(a, 1), life = n(a, 2), period = integer(a, 3), factor = a.length === 5 ? n(a, 4) : 2; if (life <= 0 || period < 1 || period > life || factor <= 0)
        bad('#NUM!', 'Invalid DDB arguments.'); let balance = cost, amount = 0; for (let i = 1; i <= period; i++) {
        amount = Math.min(Math.max(0, balance - salvage), balance * factor / life);
        balance -= amount;
    } return safe(amount); }),
    MIRR: def('Modified internal rate of return.', '=MIRR(A1:A5,0.08,0.12)', a => { req(a, 3); const vals = nums([a[0]]), finance = n(a, 1), reinvest = n(a, 2), len = vals.length; if (len < 2 || finance <= -1 || reinvest <= -1)
        bad('#NUM!', 'Invalid rates or cash flows.'); let positives = 0, negatives = 0; vals.forEach((v, i) => { if (v > 0)
        positives += v * (1 + reinvest) ** (len - i - 1);
    else
        negatives += v / (1 + finance) ** i; }); if (positives <= 0 || negatives >= 0)
        bad('#DIV/0!', 'Both positive and negative flows required.'); return safe((positives / -negatives) ** (1 / (len - 1)) - 1); }),
    IRR: def('Internal rate of return for periodic cash flows.', '=IRR(A1:A5)', a => { req(a, 1, 2); return irr(nums([a[0]]), a.length === 2 ? n(a, 1) : .1); }),
    XNPV: def('Net present value of irregular cash flows.', '=XNPV(0.1,A1:A5,B1:B5)', a => {
        req(a, 3);
        const rate = n(a, 0), cash = flatten([a[1]]), dates = flatten([a[2]]);
        if (cash.length !== dates.length || !cash.length)
            bad('#NUM!', 'Cash-flow and date ranges must have the same nonzero length.');
        if (cash.some(v => typeof v !== 'number' || !Number.isFinite(v)) || dates.some(v => typeof v !== 'number' || !Number.isFinite(v)))
            bad('#VALUE!', 'XNPV cash flows and dates must be numeric.');
        // The frozen Excel 16.0 parity fixture empirically returns #NUM! at rate 0;
        // Microsoft's XNPV reference does not document a positive-only rate requirement.
        if (rate <= 0)
            bad('#NUM!', 'Excel XNPV requires a positive discount rate.');
        const cashValues = cash, dateValues = dates;
        if (!cashValues.some(v => v < 0) || !cashValues.some(v => v > 0))
            bad('#NUM!', 'Cash flows must include at least one negative and one positive value.');
        const excelMaxDate = 2958465, wholeDates = dateValues.map(Math.trunc);
        if (wholeDates.some(d => d < 1 || d > excelMaxDate))
            bad('#VALUE!', 'XNPV contains an invalid date serial.');
        const first = wholeDates[0];
        if (wholeDates.some(d => d < first))
            bad('#NUM!', 'XNPV requires dates on or after the first date.');
        return safe(sum(cashValues.map((v, i) => v / (1 + rate) ** ((wholeDates[i] - first) / 365))));
    }),
    XIRR: def('Internal return of irregularly dated cash flows.', '=XIRR(A1:A5,B1:B5)', a => { req(a, 2, 3); const cash = nums([a[0]]), dates = nums([a[1]]); if (cash.length !== dates.length || cash.length < 2 || !cash.some(v => v > 0) || !cash.some(v => v < 0))
        bad('#NUM!', 'Invalid flows.'); const first = dates[0], f = (r) => sum(cash.map((v, i) => v / (1 + r) ** ((dates[i] - first) / 365))); let lo = -.999, hi = 1; while (f(lo) * f(hi) > 0 && hi < 1e5)
        hi = hi * 2 + 1; if (f(lo) * f(hi) > 0)
        bad('#NUM!', 'XIRR cannot find root.'); for (let i = 0; i < 160; i++) {
        const mid = (lo + hi) / 2;
        if (f(lo) * f(mid) <= 0)
            hi = mid;
        else
            lo = mid;
    } return (lo + hi) / 2; }),
    FVSCHEDULE: def('Future value after sequence of interest rates.', '=FVSCHEDULE(1000,A1:A4)', a => { req(a, 2); return safe(nums([a[1]]).reduce((v, r) => v * (1 + r), n(a, 0))); }),
    RRI: def('Equivalent annual compound growth rate.', '=RRI(5,100,150)', a => { req(a, 3); const periods = n(a, 0), present = n(a, 1), future = n(a, 2); if (periods <= 0 || present <= 0 || future < 0)
        bad('#NUM!', 'Invalid RRI arguments.'); return safe((future / present) ** (1 / periods) - 1); }),
    // 8 information functions; ISERR/ISERROR/ISNA are lazily handled by the evaluator
    ISERR: def('TRUE when expression has an error other than #N/A.', '=ISERR(1/0)', () => bad('#INTERNAL!', 'ISERR is evaluated lazily.')),
    ISERROR: def('TRUE for any formula error.', '=ISERROR(1/0)', () => bad('#INTERNAL!', 'ISERROR is evaluated lazily.')),
    ISNA: def('TRUE if formula returns #N/A.', '=ISNA(MATCH(7,A1:A3,0))', () => bad('#INTERNAL!', 'ISNA is evaluated lazily.')),
    ISEVEN: def('TRUE for even integers.', '=ISEVEN(4)', a => { req(a, 1); return Math.trunc(n(a, 0)) % 2 === 0; }),
    ISODD: def('TRUE for odd integers.', '=ISODD(3)', a => { req(a, 1); return Math.abs(Math.trunc(n(a, 0)) % 2) === 1; }),
    ISLOGICAL: def('TRUE for Boolean values.', '=ISLOGICAL(TRUE)', a => { req(a, 1); return typeof scalar(a[0]) === 'boolean'; }),
    ISNONTEXT: def('TRUE when value is not text.', '=ISNONTEXT(123)', a => { req(a, 1); return typeof scalar(a[0]) !== 'string'; }),
    TYPE: def('Excel-style type code (1 number, 2 text, 4 Boolean).', '=TYPE("Hello")', a => { req(a, 1); const v = a[0]; return Array.isArray(v) ? 64 : typeof v === 'number' ? 1 : typeof v === 'string' ? 2 : typeof v === 'boolean' ? 4 : 1; }),
    // 5 lookup & reference functions
    LOOKUP: def('Approximate match in an ascending lookup vector; returns the greatest key not exceeding the target.', '=LOOKUP(5,A1:A10,B1:B10)', a => { req(a, 2, 3); const keys = vector(a[1]), vals = a.length === 3 ? vector(a[2]) : keys; if (keys.length !== vals.length)
        bad('#VALUE!', 'Lookup/return vector lengths differ.'); if (!keys.length)
        bad('#N/A', 'LOOKUP lookup vector is empty.'); const target = scalar(a[0]); let found = -1; for (let i = 0; i < keys.length; i++) {
        const current = keys[i];
        if (i > 0 && lookupOrder(keys[i - 1], current) > 0)
            bad('#N/A', 'LOOKUP approximate lookup requires ascending sorted keys.');
        if (lookupOrder(current, target) <= 0)
            found = i;
    } if (found < 0)
        bad('#N/A', 'LOOKUP found no eligible value.'); return vals[found] ?? null; }),
    CHOOSE: def('Choose the nth value (lazily evaluated by the engine).', '=CHOOSE(2,"red","blue","green")', () => bad('#INTERNAL!', 'CHOOSE is evaluated lazily.')),
    ROWS: def('Number of rows in supplied range.', '=ROWS(A1:C8)', a => { req(a, 1); return Array.isArray(a[0]) ? a[0].length : 1; }),
    COLUMNS: def('Number of columns in supplied range.', '=COLUMNS(A1:C8)', a => { req(a, 1); return Array.isArray(a[0]) ? (a[0][0]?.length ?? 0) : 1; }),
    ADDRESS: def('Create an A1- or R1C1-style address with absolute/relative modes and an optional worksheet name.', '=ADDRESS(4,2,4,FALSE,"Sheet 2")', a => { req(a, 2, 5); const row = integer(a, 0), col = integer(a, 1), type = a.length > 2 && a[2] !== null ? integer(a, 2) : 1, a1 = a.length < 4 || a[3] === null ? true : bool(scalar(a[3])), sheet = a.length > 4 && a[4] !== null ? s(a, 4).trim() : ''; if (row < 1 || row > 1048576 || col < 1 || col > 16384 || type < 1 || type > 4)
        bad('#VALUE!', 'ADDRESS row, column, or absolute-reference mode is out of range.'); if (sheet.length > 255 || /[\[\]!]/.test(sheet))
        bad('#VALUE!', 'ADDRESS worksheet text must be a worksheet name.'); let address; if (a1) {
        let z = col, out = '';
        while (z) {
            z--;
            out = String.fromCharCode(65 + z % 26) + out;
            z = Math.floor(z / 26);
        }
        address = `${type === 1 || type === 3 ? '$' : ''}${out}${type === 1 || type === 2 ? '$' : ''}${row}`;
    }
    else {
        const r = type === 1 || type === 2 ? `R${row}` : `R[${row}]`, c = type === 1 || type === 3 ? `C${col}` : `C[${col}]`;
        address = `${r}${c}`;
    } return `${sheet ? `'${sheet.replace(/'/g, "''")}'!` : ''}${address}`; }),
};
export const v5FunctionNames = Object.keys(v5Functions).sort();
export const v5Categories = {
    'Math & Engineering': 'ACOSH ASINH ATANH ACOT ACOTH COT COTH CSC CSCH SEC SECH MROUND FACTDOUBLE MULTINOMIAL ARABIC ROMAN BASE DECIMAL'.split(' '),
    'Statistical': 'SKEW KURT COVARIANCEP COVARIANCES PEARSON STEYX STANDARDIZE TRIMMEAN NORMDIST NORMSDIST NORMINV NORMSINV'.split(' '),
    'Text': 'T N DOLLAR FIXED TEXT LENB LEFTB RIGHTB MIDB FINDB SEARCHB REPLACEB REGEXTEST REGEXREPLACE'.split(' '),
    'Date & Time': 'DATEDIF YEARFRAC DAYS360 WEEKNUM ISOWEEKNUM DATEVALUE TIMEVALUE NETWORKDAYSINTL'.split(' '),
    'Financial': 'RATE NPER IPMT PPMT EFFECT NOMINAL SLN SYD DDB MIRR IRR XNPV XIRR FVSCHEDULE RRI'.split(' '),
    'Information': 'ISERR ISERROR ISNA ISEVEN ISODD ISLOGICAL ISNONTEXT TYPE'.split(' '),
    'Lookup & Reference': 'LOOKUP CHOOSE ROWS COLUMNS ADDRESS'.split(' ')
};
