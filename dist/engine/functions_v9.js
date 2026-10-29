/** v9: 100 additional Excel-compatible functions.
 * Implemented here, independently of external spreadsheet calculation engines.
 * Array-returning functions return rectangular results to the workbook spill engine.
 */
import { SheetError, flatten, matchCriterion, number, text } from './functions.js';
const fail = (code, message) => { throw new SheetError(code, message); };
const def = (description, example, run) => ({ description, example, run });
const check = (a, min, max = min) => { if (a.length < min || a.length > max)
    fail('#ARG!', `Expected ${min}${min === max ? '' : `–${max}`} arguments, received ${a.length}.`); };
const val = (a, i) => Array.isArray(a[i]) ? fail('#VALUE!', 'Expected scalar input.') : a[i] ?? null;
const num = (a, i) => number(val(a, i));
const str = (a, i) => text(val(a, i));
const values = (a) => flatten([a]);
const nums = (a) => values(a).filter((v) => typeof v === 'number' && Number.isFinite(v));
const sum = (a) => a.reduce((p, v) => p + v, 0);
const mean = (a) => a.length ? sum(a) / a.length : fail('#DIV/0!', 'Expected numeric observations.');
const variance = (a, sample = false) => a.length > (sample ? 1 : 0) ? sum(a.map(x => (x - mean(a)) ** 2)) / (a.length - (sample ? 1 : 0)) : fail('#DIV/0!', 'Insufficient observations.');
const good = (n) => Number.isFinite(n) ? n : fail('#NUM!', 'Result outside numeric domain.');
const integer = (x) => Number.isFinite(x) && Math.trunc(x) === x;
const fl = (a, i, d) => a.length > i ? num(a, i) : d;
const matrix = (d) => Array.isArray(d) ? d.map(row => row.map(x => Array.isArray(x) ? fail('#VALUE!', 'Nested matrix unsupported.') : x)) : [[d]];
const mustMatrix = (d) => Array.isArray(d) ? matrix(d) : fail('#VALUE!', 'Expected cell range.');
const fit = (x, lo, hi) => Math.max(lo, Math.min(hi, x));
const lgamma = (z) => {
    const c = [0.99999999999980993, 676.5203681218851, -1259.1392167224028, 771.32342877765313, -176.6150291621406, 12.507343278686905, -.13857109526572012, 9.9843695780195716e-6, 1.5056327351493116e-7];
    if (z <= 0)
        return fail('#NUM!', 'Gamma input must be positive.');
    if (z < .5)
        return Math.log(Math.PI) - Math.log(Math.sin(Math.PI * z)) - lgamma(1 - z);
    let y = z - 1, x = c[0];
    for (let i = 1; i < c.length; i++)
        x += c[i] / (y + i);
    const t = y + 7.5;
    return .5 * Math.log(2 * Math.PI) + (y + .5) * Math.log(t) - t + Math.log(x);
};
const comb = (n, k) => n >= 0 && k >= 0 && k <= n && integer(n) && integer(k) ? Math.round(Math.exp(lgamma(n + 1) - lgamma(k + 1) - lgamma(n - k + 1))) : 0;
const logBeta = (a, b) => lgamma(a) + lgamma(b) - lgamma(a + b);
const betacf = (a, b, x) => {
    let c = 1, d = 1 - (a + b) * x / (a + 1);
    if (Math.abs(d) < 1e-30)
        d = 1e-30;
    d = 1 / d;
    let h = d;
    for (let m = 1; m < 300; m++) {
        const m2 = 2 * m;
        let aa = m * (b - m) * x / ((a + m2 - 1) * (a + m2));
        d = 1 + aa * d;
        if (Math.abs(d) < 1e-30)
            d = 1e-30;
        c = 1 + aa / c;
        if (Math.abs(c) < 1e-30)
            c = 1e-30;
        d = 1 / d;
        h *= d * c;
        aa = -(a + m) * (a + b + m) * x / ((a + m2) * (a + m2 + 1));
        d = 1 + aa * d;
        if (Math.abs(d) < 1e-30)
            d = 1e-30;
        c = 1 + aa / c;
        if (Math.abs(c) < 1e-30)
            c = 1e-30;
        d = 1 / d;
        const delta = d * c;
        h *= delta;
        if (Math.abs(delta - 1) < 1e-12)
            break;
    }
    return h;
};
const betaCDF = (x, a, b) => {
    if (a <= 0 || b <= 0)
        fail('#NUM!', 'Beta parameters must be positive.');
    if (x <= 0)
        return 0;
    if (x >= 1)
        return 1;
    const bt = Math.exp(lgamma(a + b) - lgamma(a) - lgamma(b) + a * Math.log(x) + b * Math.log1p(-x));
    return fit(x < (a + 1) / (a + b + 2) ? bt * betacf(a, b, x) / a : 1 - bt * betacf(b, a, 1 - x) / b, 0, 1);
};
const inverse = (p, cdf, lo, hi) => {
    if (p < 0 || p > 1)
        fail('#NUM!', 'Probability must lie between zero and one.');
    if (p === 0)
        return lo;
    if (p === 1)
        return hi;
    for (let i = 0; i < 110; i++) {
        const mid = (lo + hi) / 2;
        if (cdf(mid) < p)
            lo = mid;
        else
            hi = mid;
    }
    return (lo + hi) / 2;
};
const gammaP = (a, x) => {
    if (a <= 0 || x < 0)
        fail('#NUM!', 'Invalid gamma distribution parameters.');
    if (x === 0)
        return 0;
    const gl = lgamma(a);
    if (x < a + 1) {
        let sum0 = 1 / a, del = sum0, ap = a;
        for (let n = 1; n < 500; n++) {
            ap++;
            del *= x / ap;
            sum0 += del;
            if (Math.abs(del) < Math.abs(sum0) * 1e-13)
                break;
        }
        return fit(sum0 * Math.exp(-x + a * Math.log(x) - gl), 0, 1);
    }
    let b = x + 1 - a, c = 1 / 1e-30, d = 1 / b, h = d;
    for (let i = 1; i < 500; i++) {
        const an = -i * (i - a);
        b += 2;
        d = an * d + b;
        if (Math.abs(d) < 1e-30)
            d = 1e-30;
        c = b + an / c;
        if (Math.abs(c) < 1e-30)
            c = 1e-30;
        d = 1 / d;
        const delta = d * c;
        h *= delta;
        if (Math.abs(delta - 1) < 1e-13)
            break;
    }
    return fit(1 - Math.exp(-x + a * Math.log(x) - gl) * h, 0, 1);
};
const normal = (z) => .5 * (1 + erf(z / Math.SQRT2));
function erf(x) { const sign = x < 0 ? -1 : 1; const z = Math.abs(x), t = 1 / (1 + .3275911 * z); return sign * (1 - (((((1.061405429 * t - 1.453152027) * t) + 1.421413741) * t - .284496736) * t + .254829592) * t * Math.exp(-z * z)); }
const chiCDF = (x, df) => df > 0 ? gammaP(df / 2, x / 2) : fail('#NUM!', 'Degrees of freedom must be positive.');
const fCDF = (x, d1, d2) => x <= 0 ? 0 : betaCDF((d1 * x) / (d1 * x + d2), d1 / 2, d2 / 2);
const tCDF = (x, df) => df > 0 ? (x >= 0 ? 1 - .5 * betaCDF(df / (df + x * x), df / 2, .5) : .5 * betaCDF(df / (df + x * x), df / 2, .5)) : fail('#NUM!', 'Invalid degrees of freedom.');
const binomCDF = (k, n, p) => { if (n < 0 || p < 0 || p > 1 || !integer(n))
    fail('#NUM!', 'Invalid binomial parameters.'); let out = 0; for (let i = 0; i <= Math.min(Math.floor(k), n); i++)
    out += comb(n, i) * p ** i * (1 - p) ** (n - i); return fit(out, 0, 1); };
const matNum = (d) => mustMatrix(d).map(r => r.map(x => number(x)));
const determinant = (arr) => {
    const n = arr.length;
    if (!n || arr.some(r => r.length !== n))
        fail('#VALUE!', 'Square matrix required.');
    const a = arr.map(r => [...r]);
    let det = 1;
    for (let i = 0; i < n; i++) {
        let pivot = i;
        for (let r = i + 1; r < n; r++)
            if (Math.abs(a[r][i]) > Math.abs(a[pivot][i]))
                pivot = r;
        if (Math.abs(a[pivot][i]) < 1e-13)
            return 0;
        if (pivot !== i) {
            [a[i], a[pivot]] = [a[pivot], a[i]];
            det = -det;
        }
        const t = a[i][i];
        det *= t;
        for (let r = i + 1; r < n; r++) {
            const f = a[r][i] / t;
            for (let k = i + 1; k < n; k++)
                a[r][k] -= f * a[i][k];
        }
    }
    return good(det);
};
const invert = (arr) => {
    const n = arr.length;
    if (!n || arr.some(r => r.length !== n))
        fail('#VALUE!', 'Square matrix required.');
    const a = arr.map((r, i) => [...r, ...Array.from({ length: n }, (_, j) => i === j ? 1 : 0)]);
    for (let i = 0; i < n; i++) {
        let p = i;
        for (let r = i + 1; r < n; r++)
            if (Math.abs(a[r][i]) > Math.abs(a[p][i]))
                p = r;
        if (Math.abs(a[p][i]) < 1e-13)
            fail('#NUM!', 'Matrix is singular.');
        [a[i], a[p]] = [a[p], a[i]];
        const d = a[i][i];
        for (let j = 0; j < 2 * n; j++)
            a[i][j] /= d;
        for (let r = 0; r < n; r++)
            if (r !== i) {
                const f = a[r][i];
                for (let j = 0; j < 2 * n; j++)
                    a[r][j] -= f * a[i][j];
            }
    }
    return a.map(r => r.slice(n));
};
const regression = (ys, xs) => {
    if (xs.length !== ys.length || xs.length < 2)
        fail('#N/A', 'Regression ranges must have same length and at least two values.');
    const mx = mean(xs), my = mean(ys), den = sum(xs.map(x => (x - mx) ** 2));
    if (den === 0)
        fail('#DIV/0!', 'Zero variation in X values.');
    const slope = sum(xs.map((x, i) => (x - mx) * (ys[i] - my))) / den;
    return { slope, intercept: my - slope * mx };
};
const complex = (v) => {
    if (typeof v === 'number')
        return { re: v, im: 0 };
    const raw = text(v).trim().replace(/j$/i, 'i');
    if (!raw.includes('i')) {
        const re = Number(raw);
        if (Number.isFinite(re))
            return { re, im: 0 };
        return fail('#NUM!', 'Invalid complex value.');
    }
    const z = raw.slice(0, -1);
    let split = -1;
    for (let i = 1; i < z.length; i++)
        if ((z[i] === '+' || z[i] === '-') && z[i - 1] !== 'e' && z[i - 1] !== 'E')
            split = i;
    const re = split === -1 ? 0 : Number(z.slice(0, split)), s = split === -1 ? z : z.slice(split), im = s === '' || s === '+' ? 1 : s === '-' ? -1 : Number(s);
    return Number.isFinite(re) && Number.isFinite(im) ? { re, im } : fail('#NUM!', 'Invalid complex value.');
};
const excelNumber = (value) => {
    if (value === 0)
        return '0';
    let s = Number(value.toPrecision(15)).toString();
    if (/e/i.test(s)) {
        const [mant, exp] = s.toLowerCase().split('e');
        s = mant.replace(/\.0$/, '') + 'E' + (Number(exp) >= 0 ? '+' : '-') + String(Math.abs(Number(exp))).padStart(2, '0');
    }
    return s;
};
const cfmt = (z) => {
    const re = Number(excelNumber(z.re).replace('E', 'e')), im = Number(excelNumber(z.im).replace('E', 'e'));
    if (im === 0)
        return excelNumber(re);
    if (re === 0)
        return `${im === 1 ? '' : im === -1 ? '-' : excelNumber(im)}i`;
    return `${excelNumber(re)}${im < 0 ? '' : '+'}${im === 1 ? '' : im === -1 ? '-' : excelNumber(im)}i`;
};
const cmul = (a, b) => ({ re: a.re * b.re - a.im * b.im, im: a.re * b.im + a.im * b.re });
const cdiv = (a, b) => { const d = b.re * b.re + b.im * b.im; if (!d)
    fail('#NUM!', 'Cannot divide complex value by zero.'); return { re: (a.re * b.re + a.im * b.im) / d, im: (a.im * b.re - a.re * b.im) / d }; };
const clog = (z) => ({ re: Math.log(Math.hypot(z.re, z.im)), im: Math.atan2(z.im, z.re) });
const cexp = (z) => ({ re: Math.exp(z.re) * Math.cos(z.im), im: Math.exp(z.re) * Math.sin(z.im) });
const cpow = (a, b) => cexp(cmul(clog(a), b));
const csin = (z) => ({ re: Math.sin(z.re) * Math.cosh(z.im), im: Math.cos(z.re) * Math.sinh(z.im) });
const ccos = (z) => ({ re: Math.cos(z.re) * Math.cosh(z.im), im: -Math.sin(z.re) * Math.sinh(z.im) });
const ca = (a, i) => complex(val(a, i));
const radix = (value, from, to, places) => {
    let x = value.trim().toUpperCase();
    if (!/^[0-9A-F]+$/.test(x) || x.length > (from === 2 ? 10 : from === 8 ? 10 : 10))
        fail('#NUM!', 'Invalid engineering radix number.');
    let v = parseInt(x, from);
    if (!Number.isFinite(v))
        fail('#NUM!', 'Invalid radix digits.');
    // Excel interprets 10-bit, 30-bit and 40-bit all-one prefixes as negative numbers.
    const bits = from === 2 ? 10 : from === 8 ? 30 : 40;
    if (x.length === (from === 2 ? 10 : 10) && BigInt('0x' + v.toString(16)) >= (1n << BigInt(bits - 1)))
        v -= 2 ** bits;
    if (to === 10)
        return String(v);
    const max = to === 2 ? 512 : to === 8 ? 2 ** 29 : 2 ** 39;
    if (v >= max || v < -max)
        fail('#NUM!', 'Output outside radix range.');
    if (v < 0)
        return (BigInt(2) ** BigInt(to === 2 ? 10 : to === 8 ? 30 : 40) + BigInt(v)).toString(to).toUpperCase();
    const s = v.toString(to).toUpperCase();
    if (places !== undefined) {
        if (!integer(places) || places < 1 || places > 10 || s.length > places)
            fail('#NUM!', 'Invalid output places.');
        return s.padStart(places, '0');
    }
    return s;
};
const database = (args, operation) => {
    check(args, 3);
    const rows = mustMatrix(args[0]), crit = mustMatrix(args[2]);
    if (rows.length < 2 || crit.length < 2)
        fail('#VALUE!', 'Database and criteria must contain headers and data.');
    const headers = rows[0].map(text), headCrit = crit[0].map(text);
    const field = val(args, 1), idx = typeof field === 'number' ? field - 1 : headers.findIndex(h => h.toLowerCase() === text(field).toLowerCase());
    if (!integer(idx) || idx < 0 || idx >= headers.length)
        fail('#VALUE!', 'Unknown database field.');
    const selection = rows.slice(1).filter(row => crit.slice(1).some(cr => headCrit.every((head, j) => { const c = cr[j]; if (!head || c === null || c === '')
        return true; const ci = headers.findIndex(h => h.toLowerCase() === head.toLowerCase()); if (ci < 0)
        fail('#VALUE!', 'Criteria header absent from database.'); return matchCriterion(row[ci] ?? null, c ?? null); }))).map(row => row[idx] ?? null);
    const n = selection.filter((x) => typeof x === 'number' && Number.isFinite(x));
    const statVar = (sample) => variance(n, sample);
    switch (operation) {
        case 'DAVERAGE': return mean(n);
        case 'DCOUNT': return n.length;
        case 'DCOUNTA': return selection.filter(x => x !== null && x !== '').length;
        case 'DMAX': return n.length ? Math.max(...n) : 0;
        case 'DMIN': return n.length ? Math.min(...n) : 0;
        case 'DPRODUCT': return n.reduce((acc, x) => acc * x, 1);
        case 'DSUM': return sum(n);
        case 'DSTDEV': return Math.sqrt(statVar(true));
        case 'DSTDEVP': return Math.sqrt(statVar(false));
        case 'DVAR': return statVar(true);
        case 'DVARP': return statVar(false);
        case 'DGET':
            if (selection.length !== 1)
                fail(selection.length ? '#NUM!' : '#VALUE!', 'DGET requires exactly one matching row.');
            return selection[0] ?? null;
        default: return fail('#NAME?', 'Unknown database aggregate.');
    }
};
const excelDate = (serial) => new Date(Date.UTC(1899, 11, 30) + Math.trunc(serial) * 86400000);
const excelSerial = (d) => Math.round((d.getTime() - Date.UTC(1899, 11, 30)) / 86400000);
const lastFeb = (d) => d.getUTCMonth() === 1 && d.getUTCDate() === new Date(Date.UTC(d.getUTCFullYear(), 2, 0)).getUTCDate();
const dayCount = (a, b, basis = 0) => {
    if (![0, 1, 2, 3, 4].includes(basis))
        fail('#NUM!', 'Day count basis must be 0..4.');
    if (basis === 1 || basis === 2 || basis === 3)
        return b - a;
    const left = excelDate(a), right = excelDate(b);
    let d1 = left.getUTCDate(), d2 = right.getUTCDate();
    if (basis === 0) {
        if (d1 === 31 || lastFeb(left))
            d1 = 30;
        if (d2 === 31 && (d1 === 30 || d1 === 31))
            d2 = 30;
        if (lastFeb(left) && lastFeb(right))
            d2 = 30;
    }
    else {
        d1 = Math.min(d1, 30);
        d2 = Math.min(d2, 30);
    }
    return 360 * (right.getUTCFullYear() - left.getUTCFullYear()) + 30 * (right.getUTCMonth() - left.getUTCMonth()) + d2 - d1;
};
const years = (a, b, basis = 0) => { if (b <= a)
    fail('#NUM!', 'Maturity must follow settlement.'); return dayCount(a, b, basis) / (basis === 3 ? 365 : basis === 1 ? 365 : 360); };
const addMonths = (serial, change, anchorDay) => {
    const d = excelDate(serial), out = new Date(Date.UTC(d.getUTCFullYear(), d.getUTCMonth() + change, 1));
    out.setUTCDate(Math.min(anchorDay, new Date(Date.UTC(out.getUTCFullYear(), out.getUTCMonth() + 1, 0)).getUTCDate()));
    return excelSerial(out);
};
const schedule = (settlement, maturity, frequency, basis) => {
    if (![1, 2, 4].includes(frequency) || settlement >= maturity)
        fail('#NUM!', 'Invalid coupon schedule.');
    const span = 12 / frequency, anchorDay = excelDate(maturity).getUTCDate();
    let next = maturity, prev = maturity, count = 0;
    while (prev > settlement && count < 600) {
        next = prev;
        prev = addMonths(maturity, -(count + 1) * span, anchorDay);
        count++;
    }
    if (prev > settlement || count >= 600)
        fail('#NUM!', 'Coupon schedule exceeds supported range.');
    const E = basis === 0 || basis === 2 || basis === 4 ? 360 / frequency : dayCount(prev, next, basis);
    const A = dayCount(prev, settlement, basis), DSC = dayCount(settlement, next, basis);
    return { n: count, E, A, DSC, fraction: DSC / E, previous: prev, next };
};
const bondPrice = (settlement, maturity, coupon, yld, redemption, frequency, basis) => {
    if (coupon < 0 || yld < 0 || redemption <= 0)
        fail('#NUM!', 'Invalid yield or redemption.');
    const q = schedule(settlement, maturity, frequency, basis), k = 1 + yld / frequency;
    if (q.n === 1) {
        return redemption / (1 + yld / frequency * q.fraction) + 100 * coupon / frequency / (1 + yld / frequency * q.fraction) - 100 * coupon / frequency * q.A / q.E;
    }
    const dirty = redemption / k ** (q.n - 1 + q.fraction) + sum(Array.from({ length: q.n }, (_, i) => 100 * coupon / frequency / k ** (i + q.fraction)));
    return good(dirty - 100 * coupon / frequency * q.A / q.E);
};
const security = (a, mode) => {
    const start = num(a, 0), end = num(a, 1);
    if (mode === 'TBILLPRICE' || mode === 'TBILLYIELD' || mode === 'TBILLEQ') {
        check(a, 3);
        const v = num(a, 2);
        if (end - start > 365 || v <= 0)
            fail('#NUM!', 'Treasury bill term or rate invalid.');
        if (mode === 'TBILLPRICE')
            return good(100 * (1 - v * (end - start) / 360));
        if (mode === 'TBILLYIELD')
            return good((100 - v) / v * 360 / (end - start));
        return good(365 * v / (360 - v * (end - start)));
    }
    if (mode === 'RECEIVED') {
        check(a, 4, 5);
        const investment = num(a, 2), discount = num(a, 3), basis = fl(a, 4, 0), d = years(start, end, basis), den = 1 - discount * d;
        if (den <= 0 || investment <= 0)
            fail('#NUM!', 'Invalid investment and discount.');
        return investment / den;
    }
    if (mode === 'PRICEMAT') {
        check(a, 5, 6);
        const issue = num(a, 2), rate = num(a, 3), yld = num(a, 4), basis = fl(a, 5, 0), DSM = years(start, end, basis), DIM = years(issue, end, basis), A = years(issue, start, basis);
        return (100 + 100 * rate * DIM) / (1 + yld * DSM) - 100 * rate * A;
    }
    if (mode === 'PRICE') {
        check(a, 6, 7);
        return bondPrice(start, end, num(a, 2), num(a, 3), num(a, 4), num(a, 5), fl(a, 6, 0));
    }
    if (mode === 'YIELD') {
        check(a, 6, 7);
        const coupon = num(a, 2), price = num(a, 3), redeem = num(a, 4), frequency = num(a, 5), basis = fl(a, 6, 0);
        if (price <= 0)
            fail('#NUM!', 'Price must be positive.');
        let lo = 0, hi = 1;
        while (bondPrice(start, end, coupon, hi, redeem, frequency, basis) > price && hi < 100)
            hi *= 2;
        for (let j = 0; j < 130; j++) {
            const mid = (lo + hi) / 2;
            if (bondPrice(start, end, coupon, mid, redeem, frequency, basis) > price)
                lo = mid;
            else
                hi = mid;
        }
        return (lo + hi) / 2;
    }
    if (mode === 'YIELDDISC') {
        check(a, 4, 5);
        const price = num(a, 2), redeem = num(a, 3), basis = fl(a, 4, 0);
        if (price <= 0 || redeem <= 0)
            fail('#NUM!', 'Invalid price or redemption.');
        return (redeem / price - 1) / years(start, end, basis);
    }
    if (mode === 'YIELDMAT') {
        check(a, 5, 6);
        const issue = num(a, 2), rate = num(a, 3), price = num(a, 4), basis = fl(a, 5, 0);
        const A = years(issue, start, basis), DSM = years(start, end, basis), DIM = years(issue, end, basis);
        return ((100 + 100 * rate * DIM) / (price + 100 * rate * A) - 1) / DSM;
    }
    return fail('#NAME?', 'Unknown security function.');
};
const definitions = {};
const add = (names, category, description, formula, example) => {
    for (const name of names) {
        definitions[name] = def(description, example ?? `=${name}(A1:A5)`, a => formula(a, name));
        v9Categories[category] ??= [];
        v9Categories[category].push(name);
    }
};
export const v9Categories = {};
// Historical Excel names that retain equivalent semantics.
const compat = { BINOMDIST: 'BINOM.DIST', CONFIDENCE: 'CONFIDENCE.NORM', COVAR: 'COVARIANCEP', 'ECMA.CEILING': 'CEILING', EXPONDIST: 'EXPON.DIST', LOGINV: 'LOGNORM.INV', LOGNORMDIST: 'LOGNORM.DIST', POISSON: 'POISSON.DIST', WEIBULL: 'WEIBULL.DIST', ZTEST: 'Z.TEST' };
export const v9AliasNames = Object.keys(compat);
add(['FALSE', 'TRUE'], 'Logical', 'Excel logical constant.', (_, name) => name === 'TRUE', '=TRUE()');
add(['BINOM.INV'], 'Statistical', 'Smallest number of successful trials for a given binomial cumulative probability.', a => { check(a, 3); const trials = num(a, 0), p = num(a, 1), alpha = num(a, 2); if (!integer(trials) || trials < 0 || p < 0 || p > 1 || alpha < 0 || alpha > 1)
    fail('#NUM!', 'Invalid binomial inverse inputs.'); for (let k = 0; k <= trials; k++)
    if (binomCDF(k, trials, p) >= alpha)
        return k; return trials; }, '=BINOM.INV(10,0.5,0.5)');
add(['RAND', 'RANDBETWEEN'], 'Math & Engineering', 'Random numeric values. Recalculates on cell update, not on a live clock.', (a, name) => { if (name === 'RAND') {
    check(a, 0);
    return Math.random();
} check(a, 2); const lo = Math.ceil(num(a, 0)), hi = Math.floor(num(a, 1)); if (lo > hi)
    fail('#NUM!', 'Bottom must not exceed top.'); return lo + Math.floor(Math.random() * (hi - lo + 1)); }, '=RANDBETWEEN(1,6)');
add(['SUBTOTAL'], 'Math & Engineering', 'Aggregate a selected range by Excel subtotal operation code.', a => { check(a, 2, 255); const code = Math.trunc(num(a, 0)); if (!((code >= 1 && code <= 11) || (code >= 101 && code <= 111)))
    fail('#VALUE!', 'Unsupported SUBTOTAL function code.'); const all = a.slice(1), obs = all.flatMap(d => nums(d)); const xs = all.flatMap(d => values(d)); switch (code % 100) {
    case 1: return mean(obs);
    case 2: return obs.length;
    case 3: return xs.filter(x => x !== null && x !== '').length;
    case 4: return Math.max(...obs);
    case 5: return Math.min(...obs);
    case 6: return obs.reduce((p, x) => p * x, 1);
    case 7: return Math.sqrt(variance(obs, true));
    case 8: return Math.sqrt(variance(obs, false));
    case 9: return sum(obs);
    case 10: return variance(obs, true);
    case 11: return variance(obs, false);
    default: return fail('#VALUE!', 'Unsupported SUBTOTAL operation code.');
} }, '=SUBTOTAL(9,A1:A10)');
add(['PROB'], 'Statistical', 'Probability of values inside lower and upper limits.', a => { check(a, 3, 4); const xs = nums(a[0]), ps = nums(a[1]); if (xs.length !== ps.length || !xs.length || ps.some(p => p < 0 || p > 1) || Math.abs(sum(ps) - 1) > 1e-7)
    fail('#NUM!', 'Probabilities must sum to one and match values.'); const lo = num(a, 2), hi = fl(a, 3, lo); return good(sum(xs.map((x, i) => x >= lo && x <= hi ? ps[i] : 0))); }, '=PROB(A1:A3,B1:B3,1,2)');
add(['FREQUENCY'], 'Statistical', 'Frequency histogram (array result; use INDEX to extract a bin).', a => { check(a, 2); const xs = nums(a[0]), bins = nums(a[1]).sort((x, y) => x - y); return bins.concat([Infinity]).map((boundary, i) => [xs.filter(v => v <= boundary && (i === 0 || v > bins[i - 1])).length]); }, '=INDEX(FREQUENCY(A1:A5,B1:B2),1,1)');
add(['TRANSPOSE'], 'Lookup & Reference', 'Transpose a rectangular range; use INDEX for a single output.', a => { check(a, 1); const m = matrix(a[0]); return Array.from({ length: m[0].length }, (_, c) => m.map(r => r[c] ?? null)); }, '=INDEX(TRANSPOSE(A1:B2),2,1)');
add(['MDETERM', 'MINVERSE', 'MMULT'], 'Math & Engineering', 'Matrix determinant, matrix inverse or matrix product.', (a, name) => { check(a, name === 'MMULT' ? 2 : 1); if (name === 'MDETERM')
    return determinant(matNum(a[0])); if (name === 'MINVERSE')
    return invert(matNum(a[0])); const m = matNum(a[0]), n = matNum(a[1]); if (!m.length || !n.length || m[0].length !== n.length || m.some(r => r.length !== n.length) || n.some(r => r.length !== n[0].length))
    fail('#VALUE!', 'Matrix dimensions do not match.'); return m.map(r => n[0].map((_, j) => sum(r.map((v, k) => v * n[k][j])))); }, '=MDETERM(A1:B2)');
add(['TREND', 'GROWTH', 'LINEST', 'LOGEST'], 'Statistical', 'Linear or exponential regression forecast; multi-output results can be used with INDEX.', (a, name) => {
    check(a, 1, 4);
    const ys = nums(a[0]);
    const xs = a.length > 1 && Array.isArray(a[1]) ? nums(a[1]) : ys.map((_, i) => i + 1);
    const exp = name === 'GROWTH' || name === 'LOGEST';
    if (exp && ys.some(y => y <= 0))
        fail('#NUM!', 'Exponential regression requires positive Y.');
    const reg = regression(exp ? ys.map(Math.log) : ys, xs);
    if (name === 'LINEST')
        return [[reg.slope, reg.intercept]];
    if (name === 'LOGEST')
        return [[Math.exp(reg.slope), Math.exp(reg.intercept)]];
    const wantX = a.length > 2 ? nums(a[2]) : xs;
    return wantX.map(x => [exp ? Math.exp(reg.slope * x + reg.intercept) : reg.slope * x + reg.intercept]);
}, '=INDEX(TREND(B1:B5,A1:A5,A6),1,1)');
add(['STDEVA', 'STDEVPA', 'VARA', 'VARPA'], 'Statistical', 'Sample/population variance or deviation; text is zero and logical values are numeric.', (a, name) => { const all = a.flatMap(x => values(x)); const xs = all.map(x => typeof x === 'number' ? x : x === true ? 1 : x === false || typeof x === 'string' ? 0 : NaN).filter(x => Number.isFinite(x)); const v = variance(xs, name === 'STDEVA' || name === 'VARA'); return name.startsWith('STD') ? Math.sqrt(v) : v; }, '=STDEVA(1,2,TRUE)');
add(['BETADIST', 'BETAINV', 'CHIDIST', 'CHIINV', 'CRITBINOM', 'FDIST', 'FINV', 'GAMMADIST', 'GAMMAINV', 'HYPGEOMDIST', 'NEGBINOMDIST', 'TDIST', 'TINV'], 'Statistical', 'Probability distributions and legacy Excel compatibility functions.', (a, name) => {
    if (name === 'BETADIST' || name === 'BETAINV') {
        check(a, 3, 5);
        const p = num(a, 0), alpha = num(a, 1), beta = num(a, 2), lower = fl(a, 3, 0), upper = fl(a, 4, 1);
        if (upper <= lower)
            fail('#NUM!', 'Upper bound must exceed lower bound.');
        return name === 'BETADIST' ? betaCDF((p - lower) / (upper - lower), alpha, beta) : lower + inverse(p, x => betaCDF(x, alpha, beta), 0, 1) * (upper - lower);
    }
    if (name === 'CHIDIST' || name === 'CHIINV') {
        check(a, 2);
        const p = num(a, 0), df = num(a, 1);
        return name === 'CHIDIST' ? 1 - chiCDF(p, df) : inverse(1 - p, x => chiCDF(x, df), 0, 1e4);
    }
    if (name === 'CRITBINOM') {
        check(a, 3);
        const n = num(a, 0), p = num(a, 1), threshold = num(a, 2);
        if (threshold < 0 || threshold > 1)
            fail('#NUM!', 'Invalid threshold.');
        for (let k = 0; k <= n; k++)
            if (binomCDF(k, n, p) >= threshold)
                return k;
        return n;
    }
    if (name === 'FDIST' || name === 'FINV') {
        check(a, 3);
        const x = num(a, 0), d1 = num(a, 1), d2 = num(a, 2);
        return name === 'FDIST' ? 1 - fCDF(x, d1, d2) : inverse(1 - x, k => fCDF(k, d1, d2), 0, 1e6);
    }
    if (name === 'GAMMADIST' || name === 'GAMMAINV') {
        check(a, 3, 4);
        const x = num(a, 0), alpha = num(a, 1), beta = num(a, 2);
        if (alpha <= 0 || beta <= 0)
            fail('#NUM!', 'Gamma parameters must be positive.');
        return name === 'GAMMADIST' ? (a.length >= 4 && !Boolean(val(a, 3)) ? Math.exp((alpha - 1) * Math.log(x) - x / beta - lgamma(alpha) - alpha * Math.log(beta)) : gammaP(alpha, x / beta)) : inverse(x, k => gammaP(alpha, k / beta), 0, Math.max(100, alpha * beta * 100));
    }
    if (name === 'HYPGEOMDIST') {
        check(a, 4);
        const hits = num(a, 0), draws = num(a, 1), success = num(a, 2), population = num(a, 3);
        if (draws > population || success > population)
            fail('#NUM!', 'Sample exceeds population.');
        return comb(success, hits) * comb(population - success, draws - hits) / comb(population, draws);
    }
    if (name === 'NEGBINOMDIST') {
        check(a, 3);
        const failures = num(a, 0), successes = num(a, 1), p = num(a, 2);
        if (!integer(failures) || !integer(successes) || failures < 0 || successes < 1 || p < 0 || p > 1)
            fail('#NUM!', 'Invalid negative binomial inputs.');
        return comb(failures + successes - 1, failures) * p ** successes * (1 - p) ** failures;
    }
    if (name === 'TDIST') {
        check(a, 3);
        const tails = num(a, 2);
        if (![1, 2].includes(tails))
            fail('#NUM!', 'Tails must be 1 or 2.');
        return tails * (1 - tCDF(Math.abs(num(a, 0)), num(a, 1)));
    }
    if (name === 'TINV') {
        check(a, 2);
        return inverse(1 - num(a, 0) / 2, x => tCDF(x, num(a, 1)), 0, 1e4);
    }
    return fail('#NAME?', 'Unsupported statistical name.');
}, '=BETADIST(0.5,2,2)');
add(['CHITEST', 'FTEST', 'TTEST'], 'Statistical', 'Chi-square, F or Student t-test probability.', (a, name) => {
    check(a, name === 'TTEST' ? 4 : 2);
    const x = nums(a[0]), y = nums(a[1]);
    if (!x.length || !y.length)
        fail('#DIV/0!', 'No observations.');
    if (name === 'CHITEST') {
        if (x.length !== y.length)
            fail('#N/A', 'Observed and expected arrays differ.');
        if (y.some(v => v <= 0))
            fail('#DIV/0!', 'Expected frequencies must be positive.');
        const chi = sum(x.map((v, i) => (v - y[i]) ** 2 / y[i]));
        return 1 - chiCDF(chi, Math.max(1, x.length - 1));
    }
    if (name === 'FTEST') {
        const f = variance(x, true) / variance(y, true), df1 = x.length - 1, df2 = y.length - 1;
        return fit(2 * Math.min(fCDF(f, df1, df2), 1 - fCDF(f, df1, df2)), 0, 1);
    }
    const tails = num(a, 2), type = num(a, 3);
    if (![1, 2].includes(tails) || ![1, 2, 3].includes(type))
        fail('#NUM!', 'Invalid t-test tails/type.');
    let t, df;
    if (type === 1) {
        if (x.length !== y.length)
            fail('#N/A', 'Paired ranges differ.');
        const ds = x.map((v, i) => v - y[i]);
        t = Math.abs(mean(ds)) / Math.sqrt(variance(ds, true) / ds.length);
        df = ds.length - 1;
    }
    else {
        const vx = variance(x, true), vy = variance(y, true);
        if (type === 2) {
            df = x.length + y.length - 2;
            const pool = ((x.length - 1) * vx + (y.length - 1) * vy) / df;
            t = Math.abs(mean(x) - mean(y)) / Math.sqrt(pool * (1 / x.length + 1 / y.length));
        }
        else {
            const a0 = vx / x.length, b0 = vy / y.length;
            t = Math.abs(mean(x) - mean(y)) / Math.sqrt(a0 + b0);
            df = (a0 + b0) ** 2 / (a0 * a0 / (x.length - 1) + b0 * b0 / (y.length - 1));
        }
    }
    return fit(tails * (1 - tCDF(t, df)), 0, 1);
}, '=FTEST(A1:A6,B1:B6)');
add(['BIN2DEC', 'BIN2HEX', 'BIN2OCT', 'DEC2BIN', 'DEC2HEX', 'DEC2OCT', 'HEX2BIN', 'HEX2DEC', 'HEX2OCT', 'OCT2BIN', 'OCT2DEC', 'OCT2HEX'], 'Math & Engineering', 'Convert binary, octal, decimal and hexadecimal engineering numbers.', (a, name) => { check(a, 1, 2); const src = name.slice(0, 3), dst = name.slice(4), from = src === 'BIN' ? 2 : src === 'OCT' ? 8 : src === 'HEX' ? 16 : 10, to = dst === 'BIN' ? 2 : dst === 'OCT' ? 8 : dst === 'HEX' ? 16 : 10; const v = from === 10 ? num(a, 0) : val(a, 0); if (from === 10 && (!integer(v) || Math.abs(v) > 2 ** 39))
    fail('#NUM!', 'Engineering decimal outside range.'); const converted = from === 10 ? String(v) : text(v); if (from === 10 && to !== 10) {
    const limit = 2 ** (to === 2 ? 9 : to === 8 ? 29 : 39);
    if (Number(v) < -limit || Number(v) >= limit)
        fail('#NUM!', 'Engineering output is outside Excel signed radix range.');
} const out = from === 10 && Number(v) < 0 && to !== 10 ? (2 ** (to === 2 ? 10 : to === 8 ? 30 : 40) + Number(v)).toString(to).toUpperCase() : radix(converted, from, to, a.length > 1 ? num(a, 1) : undefined); return to === 10 ? Number(out) : out; }, '=DEC2HEX(255)');
add(['COMPLEX', 'IMABS', 'IMAGINARY', 'IMARGUMENT', 'IMCONJUGATE', 'IMCOS', 'IMDIV', 'IMEXP', 'IMLN', 'IMLOG10', 'IMLOG2', 'IMPOWER', 'IMPRODUCT', 'IMREAL', 'IMSIN', 'IMSQRT', 'IMSUB', 'IMSUM'], 'Math & Engineering', 'Complex-number engineering arithmetic.', (a, name) => {
    if (name === 'COMPLEX') {
        check(a, 2, 3);
        const unit = a.length === 3 ? str(a, 2) : 'i';
        if (unit !== 'i' && unit !== 'j')
            fail('#VALUE!', 'Complex suffix must be i or j.');
        return cfmt({ re: num(a, 0), im: num(a, 1) }).replace('i', unit);
    }
    check(a, ['IMDIV', 'IMPOWER', 'IMSUB'].includes(name) ? 2 : 1, name === 'IMPRODUCT' || name === 'IMSUM' ? 255 : ['IMDIV', 'IMPOWER', 'IMSUB'].includes(name) ? 2 : 1);
    const x = ca(a, 0), m = Math.hypot(x.re, x.im);
    switch (name) {
        case 'IMABS': return m;
        case 'IMREAL': return x.re;
        case 'IMAGINARY': return x.im;
        case 'IMARGUMENT': return m ? Math.atan2(x.im, x.re) : fail('#DIV/0!', 'Argument undefined at zero.');
        case 'IMCONJUGATE': return cfmt({ re: x.re, im: -x.im });
        case 'IMEXP': return cfmt(cexp(x));
        case 'IMLN': return cfmt(clog(x));
        case 'IMLOG10': {
            const z = clog(x);
            return cfmt({ re: z.re / Math.LN10, im: z.im / Math.LN10 });
        }
        case 'IMLOG2': {
            const z = clog(x);
            return cfmt({ re: z.re / Math.LN2, im: z.im / Math.LN2 });
        }
        case 'IMCOS': return cfmt(ccos(x));
        case 'IMSIN': return cfmt(csin(x));
        case 'IMSQRT': return cfmt(cpow(x, { re: 0.5, im: 0 }));
        case 'IMPOWER': return cfmt(cpow(x, ca(a, 1)));
        case 'IMDIV': return cfmt(cdiv(x, ca(a, 1)));
        case 'IMSUB': {
            const y = ca(a, 1);
            return cfmt({ re: x.re - y.re, im: x.im - y.im });
        }
        case 'IMSUM': {
            const zs = a.flatMap(d => values(d).map(complex));
            return cfmt({ re: sum(zs.map(z => z.re)), im: sum(zs.map(z => z.im)) });
        }
        case 'IMPRODUCT': return cfmt(a.flatMap(d => values(d).map(complex)).reduce(cmul, { re: 1, im: 0 }));
    }
    return fail('#NAME?', 'Unknown complex operation.');
}, '=IMABS("3+4i")');
add(['DAVERAGE', 'DCOUNT', 'DCOUNTA', 'DGET', 'DMAX', 'DMIN', 'DPRODUCT', 'DSTDEV', 'DSTDEVP', 'DSUM', 'DVAR', 'DVARP'], 'Database', 'Database aggregates using an Excel-style field and criteria table.', database, '=DSUM(A1:C5,"Sales",E1:E2)');
add(['TBILLEQ', 'TBILLPRICE', 'TBILLYIELD', 'PRICE', 'PRICEMAT', 'RECEIVED', 'YIELD', 'YIELDDISC', 'YIELDMAT'], 'Financial', 'Treasury and fixed-income calculations (simplified conventions for coupon schedules).', security, '=TBILLPRICE(45000,45100,0.05)');
add(['DURATION', 'MDURATION'], 'Financial', 'Macaulay or modified duration using periodic coupon periods.', (a, name) => { check(a, 5, 6); const settlement = num(a, 0), maturity = num(a, 1), coupon = num(a, 2), yld = num(a, 3), frequency = num(a, 4); if (![1, 2, 4].includes(frequency))
    fail('#NUM!', 'Frequency must be 1, 2 or 4.'); const basis = fl(a, 5, 0), q = schedule(settlement, maturity, frequency, basis), cash = 100 * coupon / frequency, discount = 1 + yld / frequency; let pv = 0, weighted = 0; for (let t = 1; t <= q.n; t++) {
    const periods = t - 1 + q.fraction, p = (cash + (t === q.n ? 100 : 0)) / discount ** periods;
    pv += p;
    weighted += (periods / frequency) * p;
} const duration = weighted / pv; return name === 'MDURATION' ? duration / discount : duration; }, '=DURATION(45000,47000,0.05,0.06,2)');
add(['CONVERT'], 'Math & Engineering', 'Convert a measurement between standard metric and imperial units.', a => { check(a, 3); const u = { m: ['length', 1], km: ['length', 1000], cm: ['length', .01], mm: ['length', .001], mi: ['length', 1609.344], ft: ['length', .3048], in: ['length', .0254], yd: ['length', .9144], g: ['mass', .001], kg: ['mass', 1], lbm: ['mass', .45359237], ozm: ['mass', .028349523125], s: ['time', 1], sec: ['time', 1], mn: ['time', 60], hr: ['time', 3600], day: ['time', 86400], l: ['volume', 1], ml: ['volume', .001], gal: ['volume', 3.785411784] }; const from = u[str(a, 1)], to = u[str(a, 2)]; if (!from || !to || from[0] !== to[0])
    fail('#N/A', 'Unknown or incompatible conversion units.'); return num(a, 0) * from[1] / to[1]; }, '=CONVERT(1,"mi","km")');
/** Resolve legacy names against the existing registry after it has been created. */
export function v9Aliases(existing) { const out = {}; for (const [old, current] of Object.entries(compat)) {
    const f = existing[current];
    if (!f)
        throw Error(`Missing compatibility target ${current}`);
    out[old] = { ...f, description: `Legacy Excel alias for ${current}. ${f.description}` };
    v9Categories['Compatibility'] ??= [];
    v9Categories['Compatibility'].push(old);
} return out; }
export const v9Functions = definitions;
export const v9FunctionNames = [...Object.keys(definitions), ...Object.keys(compat)].sort();
if (v9FunctionNames.length !== 100 || new Set(v9FunctionNames).size !== 100)
    throw Error(`Expected 100 new unique v9 functions; received ${v9FunctionNames.length}.`);
