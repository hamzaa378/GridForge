/** GridForge v10 Excel formula compatibility, independently implemented.
 * Matrix-returning functions are exposed to the workbook's dynamic-array spill engine.
 * Lambda-taking functions support explicit safe reducer names instead of executable lambda expressions.
 */
import { SheetError, number, text, flatten } from './functions.js';
import { excelDateSerial, excelSerialUtcDate } from './date_serial.js';
const fail = (c, m) => { throw new SheetError(c, m); };
const one = (d) => Array.isArray(d) ? fail('#VALUE!', 'Expected one cell rather than a range.') : d ?? null;
const n = (a, i) => number(one(a[i]));
const s = (a, i) => text(one(a[i]));
const arr = (d) => Array.isArray(d) ? d.map(r => r.map(v => one(v))) : [[one(d)]];
const nums = (d) => flatten([d]).filter((v) => typeof v === 'number' && Number.isFinite(v));
const sum = (a) => a.reduce((p, v) => p + v, 0);
const avg = (a) => a.length ? sum(a) / a.length : fail('#DIV/0!', 'No numeric data.');
const req = (a, lo, hi = lo) => { if (a.length < lo || a.length > hi)
    fail('#ARG!', `Expected ${lo}–${hi} arguments.`); };
const safe = (x) => Number.isFinite(x) ? x : fail('#NUM!', 'Result is not finite.');
const bool = (v) => v === true || typeof v === 'number' && v !== 0 || typeof v === 'string' && v !== '' && v.toUpperCase() !== 'FALSE';
const matrix = (a, i) => arr(a[i]);
const positiveInt = (x, max = 20000) => Number.isSafeInteger(x) && x > 0 && x <= max ? x : fail('#NUM!', 'Expected a positive integer inside the supported size limit.');
const guard = (r, c) => { if (r * c > 20000)
    fail('#NUM!', 'Array results are limited to 20,000 elements.'); };
const scalar = (v) => v === null ? '' : String(v);
const forecastSeries = (valuesArg, timelineArg, dataCompletion = 1, aggregation = 0) => {
    const values = flatten([valuesArg ?? null]), timeline = flatten([timelineArg ?? null]);
    if (values.length !== timeline.length)
        fail('#N/A', 'FORECAST.ETS values and timeline must have the same number of points.');
    if (values.length < 2)
        fail('#NUM!', 'FORECAST.ETS needs at least two observations.');
    if (timeline.some(value => typeof value !== 'number' || !Number.isFinite(value) || value === 0))
        fail('#VALUE!', 'FORECAST.ETS timeline must contain nonzero numeric values.');
    const pairs = timeline.map((x, index) => {
        const value = values[index];
        if (value === null)
            return { x: x, y: null, index };
        if (typeof value !== 'number' && typeof value !== 'string' && !(typeof value === 'boolean' && aggregation === 3))
            fail('#VALUE!', 'FORECAST.ETS values must be numbers, or text/booleans for COUNTA aggregation, or blank cells.');
        if (typeof value === 'number' && !Number.isFinite(value))
            fail('#VALUE!', 'FORECAST.ETS values must be finite numbers.');
        return { x: x, y: value, index };
    }).sort((a, b) => a.x - b.x || a.index - b.index);
    const grouped = [];
    for (let start = 0; start < pairs.length;) {
        let end = start + 1;
        while (end < pairs.length && pairs[end].x === pairs[start].x)
            end++;
        const bucket = pairs.slice(start, end), numbers = bucket.flatMap(point => typeof point.y === 'number' ? [point.y] : []);
        let value = null;
        if (bucket.length === 1)
            value = bucket[0].y;
        else
            switch (aggregation) {
                case 0:
                    value = numbers.length ? avg(numbers) : null;
                    break;
                case 1:
                    value = sum(numbers);
                    break;
                case 2:
                    value = numbers.length;
                    break;
                case 3:
                    value = bucket.filter(point => point.y !== null).length;
                    break;
                case 4:
                    value = numbers.length ? numbers.reduce((best, value) => Math.min(best, value), Infinity) : null;
                    break;
                case 5:
                    value = numbers.length ? numbers.reduce((best, value) => Math.max(best, value), -Infinity) : null;
                    break;
                case 6: {
                    const sorted = [...numbers].sort((a, b) => a - b);
                    value = !sorted.length ? null : sorted.length % 2 ? sorted[(sorted.length - 1) / 2] : (sorted[sorted.length / 2 - 1] + sorted[sorted.length / 2]) / 2;
                    break;
                }
                default: fail('#VALUE!', 'FORECAST.ETS aggregation must be from zero to six.');
            }
        if (typeof value === 'string' || typeof value === 'boolean')
            fail('#VALUE!', 'FORECAST.ETS cannot use text or booleans with the selected aggregation option.');
        grouped.push({ x: pairs[start].x, y: value });
        start = end;
    }
    if (grouped.length < 2)
        fail('#NUM!', 'FORECAST.ETS needs at least two unique timeline points.');
    const completedPairs = grouped;
    const missing = completedPairs.filter(point => point.y === null).length;
    if (missing / completedPairs.length > 0.3)
        fail('#NUM!', 'FORECAST.ETS can complete no more than 30% missing values.');
    if (missing) {
        const known = completedPairs.flatMap((point, index) => point.y === null ? [] : [{ x: point.x, index, y: point.y }]);
        for (let index = 0; index < completedPairs.length; index++)
            if (completedPairs[index].y === null) {
                if (dataCompletion === 0) {
                    completedPairs[index].y = 0;
                    continue;
                }
                let before, after;
                for (const point of known) {
                    if (point.index < index)
                        before = point;
                    else if (point.index > index) {
                        after = point;
                        break;
                    }
                }
                const point = completedPairs[index];
                point.y = !before ? after.y : !after ? before.y : before.y + (after.y - before.y) * (point.x - before.x) / (after.x - before.x);
            }
    }
    const timelineNumbers = completedPairs.map(point => point.x), valuesNumbers = completedPairs.map(point => point.y);
    if (valuesNumbers.length < 2)
        fail('#NUM!', 'FORECAST.ETS needs at least two numeric observations.');
    const pairsWithValues = completedPairs.map(point => ({ x: point.x, y: point.y }));
    const differences = pairsWithValues.slice(1).map((point, index) => point.x - pairsWithValues[index].x);
    const step = differences[0];
    if (!(step > 0) || differences.some(value => Math.abs(value - step) > Math.max(1, Math.abs(step)) * 1e-9))
        fail('#NUM!', 'FORECAST.ETS timeline must use a consistent nonzero step.');
    const x = timelineNumbers, y = valuesNumbers, meanX = avg(x), meanY = avg(y), variance = sum(x.map(value => (value - meanX) ** 2));
    if (variance === 0)
        fail('#DIV/0!', 'FORECAST.ETS timeline has no variation.');
    const slope = sum(x.map((value, index) => (value - meanX) * (y[index] - meanY))) / variance, intercept = meanY - slope * meanX, predicted = x.map(value => intercept + slope * value), residual = y.map((value, index) => value - predicted[index]);
    return { x, y, step, slope, intercept, xMean: meanX, xVariance: variance, residual, mae: avg(residual.map(Math.abs)), rmse: Math.sqrt(avg(residual.map(value => value * value))) };
};
const validateForecastOptions = (args, seasonalityIndex, dataCompletionIndex, aggregationIndex) => {
    if (args.length > seasonalityIndex) {
        const seasonality = n(args, seasonalityIndex);
        if (!Number.isInteger(seasonality) || seasonality < 0 || seasonality > 8760)
            fail('#NUM!', 'FORECAST.ETS seasonality must be zero or a whole number up to 8,760.');
    }
    if (args.length > dataCompletionIndex) {
        const completion = n(args, dataCompletionIndex);
        if (completion !== 0 && completion !== 1)
            fail('#VALUE!', 'FORECAST.ETS data_completion must be zero or one.');
    }
    if (args.length > aggregationIndex) {
        const aggregation = n(args, aggregationIndex);
        if (!Number.isInteger(aggregation) || aggregation < 0 || aggregation > 6)
            fail('#VALUE!', 'FORECAST.ETS aggregation must be an integer from zero to six.');
    }
};
const forecastSeasonality = (values) => {
    if (values.length < 6)
        return 0;
    const differences = values.slice(1).map((value, index) => value - values[index]);
    let bestLag = 0, bestCorrelation = -Infinity;
    for (let lag = 2; lag <= Math.floor(values.length / 2); lag++) {
        const left = differences.slice(lag), right = differences.slice(0, -lag), leftMean = avg(left), rightMean = avg(right), cov = sum(left.map((value, index) => (value - leftMean) * (right[index] - rightMean))), den = Math.sqrt(sum(left.map(value => (value - leftMean) ** 2)) * sum(right.map(value => (value - rightMean) ** 2)));
        const correlation = den ? cov / den : 0;
        if (correlation > bestCorrelation) {
            bestCorrelation = correlation;
            bestLag = lag;
        }
    }
    return bestCorrelation >= 0.7 ? bestLag : 0;
};
const forecastEtsModel = (series, seasonalityOption) => {
    const { x, y, step, slope, intercept } = series, period = seasonalityOption === 0 ? 0 : seasonalityOption === 1 ? forecastSeasonality(y) : seasonalityOption;
    const linearFit = y.map((_, index) => intercept + slope * x[index]), linearResidual = y.map((value, index) => value - linearFit[index]);
    const linear = () => ({ period: 0, startIndex: 0, alpha: null, beta: null, gamma: null, fitted: linearFit, residual: linearResidual, mae: avg(linearResidual.map(Math.abs)), rmse: Math.sqrt(avg(linearResidual.map(value => value * value))), forecastAt: steps => safe(intercept + slope * (x.at(-1) + steps * step)) });
    if (period < 2 || period * 2 > y.length || y.length > 4000)
        return linear();
    let best;
    const firstLevel = avg(y.slice(0, period)), secondLevel = avg(y.slice(period, period * 2));
    for (const alpha of [.1, .3, .5, .7, .9])
        for (const beta of [.1, .3, .5, .7, .9])
            for (const gamma of [.1, .3, .5, .7, .9]) {
                let level = firstLevel, trend = (secondLevel - firstLevel) / period, seasonals = y.slice(0, period).map(value => value - firstLevel), sse = 0;
                const fitted = Array(y.length).fill(NaN), residual = [];
                for (let index = period; index < y.length; index++) {
                    const phase = index % period, estimate = level + trend + seasonals[phase], error = y[index] - estimate;
                    fitted[index] = estimate;
                    residual.push(error);
                    sse += error * error;
                    const nextLevel = alpha * (y[index] - seasonals[phase]) + (1 - alpha) * (level + trend), nextTrend = beta * (nextLevel - level) + (1 - beta) * trend;
                    seasonals[phase] = gamma * (y[index] - nextLevel) + (1 - gamma) * seasonals[phase];
                    level = nextLevel;
                    trend = nextTrend;
                }
                if (!best || sse < best.sse)
                    best = { alpha, beta, gamma, level, trend, seasonals, fitted, residual, sse };
            }
    if (!best || !Number.isFinite(best.sse))
        return linear();
    const fit = best, phaseAt = (steps) => ((y.length + Math.max(1, Math.ceil(steps)) - 1) % period + period) % period;
    return { period, startIndex: period, alpha: fit.alpha, beta: fit.beta, gamma: fit.gamma, fitted: fit.fitted, residual: fit.residual, mae: avg(fit.residual.map(Math.abs)), rmse: Math.sqrt(avg(fit.residual.map(value => value * value))), forecastAt: steps => safe(fit.level + fit.trend * steps + fit.seasonals[phaseAt(steps)]) };
};
// Acklam's rational approximation to the standard-normal quantile. This is
// used to scale uncertainty around GridForge's bounded ETS approximation.
const normalQuantile = (p) => {
    if (!(p > 0 && p < 1))
        fail('#NUM!', 'Probability must be strictly between zero and one.');
    const a = [-3.969683028665376e1, 2.209460984245205e2, -2.759285104469687e2, 1.383577518672690e2, -3.066479806614716e1, 2.506628277459239];
    const b = [-5.447609879822406e1, 1.615858368580409e2, -1.556989798598866e2, 6.680131188771972e1, -1.328068155288572e1];
    const c = [-7.784894002430293e-3, -3.223964580411365e-1, -2.400758277161838, -2.549732539343734, 4.374664141464968, 2.938163982698783];
    const d = [7.784695709041462e-3, 3.224671290700398e-1, 2.445134137142996, 3.754408661907416];
    const low = 0.02425, high = 1 - low;
    if (p < low) {
        const q = Math.sqrt(-2 * Math.log(p));
        return (((((c[0] * q + c[1]) * q + c[2]) * q + c[3]) * q + c[4]) * q + c[5]) / ((((d[0] * q + d[1]) * q + d[2]) * q + d[3]) * q + 1);
    }
    if (p > high) {
        const q = Math.sqrt(-2 * Math.log(1 - p));
        return -(((((c[0] * q + c[1]) * q + c[2]) * q + c[3]) * q + c[4]) * q + c[5]) / ((((d[0] * q + d[1]) * q + d[2]) * q + d[3]) * q + 1);
    }
    const q = p - 0.5, r = q * q;
    return (((((a[0] * r + a[1]) * r + a[2]) * r + a[3]) * r + a[4]) * r + a[5]) * q / (((((b[0] * r + b[1]) * r + b[2]) * r + b[3]) * r + b[4]) * r + 1);
};
const filterXml = (source, path) => {
    if (source.length > 1_000_000 || path.length > 1024)
        fail('#VALUE!', 'FILTERXML input exceeds the supported size.');
    if (typeof DOMParser !== 'undefined') {
        const document = new DOMParser().parseFromString(source, 'application/xml');
        if (document.getElementsByTagName('parsererror').length)
            fail('#VALUE!', 'FILTERXML source is not well-formed XML.');
        let result;
        try {
            result = document.evaluate(path, document, document.createNSResolver(document.documentElement), XPathResult.ANY_TYPE, null);
        }
        catch {
            fail('#VALUE!', 'FILTERXML XPath expression is invalid or unsupported.');
        }
        if (result.resultType === XPathResult.NUMBER_TYPE)
            return result.numberValue;
        if (result.resultType === XPathResult.STRING_TYPE)
            return result.stringValue;
        if (result.resultType === XPathResult.BOOLEAN_TYPE)
            return result.booleanValue;
        const values = [];
        try {
            let node;
            while ((node = result.iterateNext())) {
                if (values.length >= 20_000)
                    fail('#VALUE!', 'FILTERXML result is limited to 20,000 values.');
                values.push(node.textContent ?? '');
            }
        }
        catch {
            fail('#VALUE!', 'FILTERXML XPath expression could not be evaluated.');
        }
        if (!values.length)
            fail('#VALUE!', 'FILTERXML XPath expression did not match any values.');
        return values.map(value => [value]);
    }
    const tag = path.replace(/^\/\/?/, '').replace(/[^\w:-]/g, '');
    if (!tag)
        fail('#VALUE!', 'Use an XPath such as //item.');
    const match = new RegExp(`<${tag}(?:\\s[^>]*)?>([^<]*)<\\/${tag}>`, 'i').exec(source);
    if (!match)
        return fail('#VALUE!', 'Matching XML tag not found.');
    return match[1];
};
const unixToSerial = (dt) => excelDateSerial(dt.toISOString()) ?? NaN;
const date = (serial) => excelSerialUtcDate(serial);
const months = (serial) => { const d = date(serial); return d.getUTCFullYear() * 12 + d.getUTCMonth(); };
const dayCount = (a, b, basis = 0) => {
    if (![0, 1, 2, 3, 4].includes(basis))
        fail('#NUM!', 'Day-count basis must be 0..4.');
    if (basis === 1 || basis === 2 || basis === 3)
        return b - a;
    const start = date(a), end = date(b);
    let y1 = start.getUTCFullYear(), m1 = start.getUTCMonth() + 1, d1 = start.getUTCDate();
    let y2 = end.getUTCFullYear(), m2 = end.getUTCMonth() + 1, d2 = end.getUTCDate();
    const lastFeb = (dt) => dt.getUTCMonth() === 1 && dt.getUTCDate() === new Date(Date.UTC(dt.getUTCFullYear(), 2, 0)).getUTCDate();
    if (basis === 0) {
        if (d1 === 31 || lastFeb(start))
            d1 = 30;
        if (d2 === 31 && (d1 === 30 || d1 === 31))
            d2 = 30;
        if (lastFeb(start) && lastFeb(end))
            d2 = 30;
    }
    else {
        d1 = Math.min(d1, 30);
        d2 = Math.min(d2, 30);
    }
    return 360 * (y2 - y1) + 30 * (m2 - m1) + d2 - d1;
};
const yearFrac = (a, b, basis = 0) => dayCount(a, b, basis) / (basis === 3 ? 365 : basis === 1 ? 365 : 360);
const present = (x) => x !== null && x !== '';
const distNormal = (x) => Math.exp(-x * x / 2) / Math.sqrt(2 * Math.PI);
const lgamma = (z) => { const p = [0.9999999999998099, 676.5203681218851, -1259.1392167224028, 771.32342877765313, -176.6150291621406, 12.507343278686905, -.13857109526572012, 9.9843695780195716e-6, 1.5056327351493116e-7]; if (z <= 0)
    fail('#NUM!', 'Positive shape required.'); if (z < .5)
    return Math.log(Math.PI) - Math.log(Math.sin(Math.PI * z)) - lgamma(1 - z); z -= 1; let x = p[0]; for (let i = 1; i < p.length; i++)
    x += p[i] / (z + i); let t = z + 7.5; return .5 * Math.log(2 * Math.PI) + (z + .5) * Math.log(t) - t + Math.log(x); };
// Regularized lower gamma P(a,x). Series for x<a+1; stable Lentz upper-tail
// continued fraction otherwise. Avoid overflow from the old forward series.
const gammaP = (a, x) => {
    if (a <= 0 || x < 0)
        fail('#NUM!', 'Invalid gamma domain.');
    if (x === 0)
        return 0;
    const logPref = a * Math.log(x) - x - lgamma(a);
    if (x < a + 1) {
        let term = 1 / a, total = term;
        for (let j = 1; j < 2000; j++) {
            term *= x / (a + j);
            total += term;
            if (Math.abs(term) < Math.abs(total) * 1e-15)
                break;
        }
        return Math.min(1, Math.max(0, total * Math.exp(logPref)));
    }
    let b = x + 1 - a, c = 1e300, d = 1 / Math.max(1e-300, b), h = d;
    for (let i = 1; i < 2000; i++) {
        const an = -i * (i - a);
        b += 2;
        d = b + an * d;
        if (Math.abs(d) < 1e-300)
            d = 1e-300;
        c = b + an / c;
        if (Math.abs(c) < 1e-300)
            c = 1e-300;
        d = 1 / d;
        const delta = d * c;
        h *= delta;
        if (Math.abs(delta - 1) < 1e-15)
            break;
    }
    return Math.min(1, Math.max(0, 1 - Math.exp(logPref) * h));
};
const approximateCDF = (f, lo, hi, p) => { if (p < 0 || p > 1)
    fail('#NUM!', 'Probability must be 0 to 1.'); for (let i = 0; i < 90; i++) {
    const mid = (lo + hi) / 2;
    if (f(mid) < p)
        lo = mid;
    else
        hi = mid;
} return (lo + hi) / 2; };
const erf = (x) => { if (x === 0)
    return 0; const sign = x < 0 ? -1 : 1, z = Math.abs(x); if (z >= 6)
    return sign; let term = z, total = z; for (let k = 1; k < 160; k++) {
    term *= -(z * z) / k;
    const add = term / (2 * k + 1);
    total += add;
    if (Math.abs(add) < Math.abs(total) * 2e-16)
        break;
} return sign * 2 / Math.sqrt(Math.PI) * total; };
const choose = (nn, kk) => { if (!Number.isInteger(nn) || !Number.isInteger(kk) || nn < 0 || kk < 0 || kk > nn)
    return 0; let x = 1; for (let i = 1; i <= kk; i++)
    x = x * (nn - kk + i) / i; return x; };
const betaCF = (a, b, x) => { let c = 1, d = 1 - (a + b) * x / (a + 1); if (Math.abs(d) < 1e-30)
    d = 1e-30; d = 1 / d; let h = d; for (let m = 1; m < 250; m++) {
    let k = 2 * m, aa = m * (b - m) * x / ((a + k - 1) * (a + k));
    d = 1 + aa * d;
    if (Math.abs(d) < 1e-30)
        d = 1e-30;
    c = 1 + aa / c;
    if (Math.abs(c) < 1e-30)
        c = 1e-30;
    d = 1 / d;
    h *= d * c;
    aa = -(a + m) * (a + b + m) * x / ((a + k) * (a + k + 1));
    d = 1 + aa * d;
    if (Math.abs(d) < 1e-30)
        d = 1e-30;
    c = 1 + aa / c;
    if (Math.abs(c) < 1e-30)
        c = 1e-30;
    d = 1 / d;
    let delta = d * c;
    h *= delta;
    if (Math.abs(delta - 1) < 1e-12)
        break;
} return h; };
const betaCDF = (x, a, b) => { if (a <= 0 || b <= 0)
    fail('#NUM!', 'Invalid beta shape.'); if (x <= 0)
    return 0; if (x >= 1)
    return 1; const bt = Math.exp(lgamma(a + b) - lgamma(a) - lgamma(b) + a * Math.log(x) + b * Math.log1p(-x)); return Math.max(0, Math.min(1, x < (a + 1) / (a + b + 2) ? bt * betaCF(a, b, x) / a : 1 - bt * betaCF(b, a, 1 - x) / b)); };
const chiCDF = (x, df) => { if (df <= 0 || x < 0)
    fail('#NUM!', 'Invalid chi-square arguments.'); return gammaP(df / 2, x / 2); };
const fCDF = (x, d1, d2) => { if (d1 <= 0 || d2 <= 0 || x < 0)
    fail('#NUM!', 'Invalid F distribution.'); return betaCDF(d1 * x / (d1 * x + d2), d1 / 2, d2 / 2); };
const tCDF = (x, df) => { if (df <= 0)
    fail('#NUM!', 'Invalid degrees of freedom.'); return x >= 0 ? 1 - betaCDF(df / (df + x * x), df / 2, .5) / 2 : betaCDF(df / (df + x * x), df / 2, .5) / 2; };
const stock = (start, end, freq, settlement) => { if (!(end > start) || ![1, 2, 4].includes(freq))
    fail('#NUM!', 'Invalid coupon dates or frequency.'); const last = settlement ?? start; const span = 12 / freq; let dt = months(end), count = 0, previous = end; while (count < 2000) {
    const m = months(end) - span * count;
    const year = Math.floor(m / 12), month = ((m % 12) + 12) % 12;
    const d = date(end);
    const test = unixToSerial(new Date(Date.UTC(year, month, Math.min(d.getUTCDate(), new Date(Date.UTC(year, month + 1, 0)).getUTCDate()))));
    if (test <= last) {
        previous = test;
        break;
    }
    count++;
} let d = date(previous), next = unixToSerial(new Date(Date.UTC(d.getUTCFullYear(), d.getUTCMonth() + span, d.getUTCDate()))); return { previous, next, count, period: next - previous }; };
export const v10Categories = {};
export function buildV10Functions(existing) {
    const f = {};
    const add = (category, names, description, impl, example) => { for (const name of names) {
        if (existing[name] || f[name])
            throw Error(`Duplicate v10 function ${name}`);
        (v10Categories[category] ??= []).push(name);
        f[name] = { description, example: example ?? `=${name}(A1:A5)`, run: a => impl(name, a) };
    } };
    const call = (name, args) => { const item = existing[name]; if (!item)
        return fail('#NAME?', `Missing implementation ${name}`); return item.run(args); };
    // Full statistical formula names; use the same arguments as their Excel counterparts.
    add('Statistical', ['BETA.DIST', 'BETA.INV'], 'Beta probability and inverse.', (name, a) => { req(a, name === 'BETA.INV' ? 3 : 4, 6); const v = n(a, 0), alpha = n(a, 1), beta = n(a, 2), lo = a.length > (name === 'BETA.INV' ? 3 : 4) ? n(a, name === 'BETA.INV' ? 3 : 4) : 0, hi = a.length > (name === 'BETA.INV' ? 4 : 5) ? n(a, name === 'BETA.INV' ? 4 : 5) : 1; if (hi <= lo)
        fail('#NUM!', 'Upper limit must be greater.'); return name === 'BETA.INV' ? lo + (hi - lo) * approximateCDF(z => betaCDF(z, alpha, beta), 0, 1, v) : n(a, 3) === 0 ? Math.exp((alpha - 1) * Math.log((v - lo) / (hi - lo)) + (beta - 1) * Math.log(1 - (v - lo) / (hi - lo)) - lgamma(alpha) - lgamma(beta) + lgamma(alpha + beta)) / (hi - lo) : betaCDF((v - lo) / (hi - lo), alpha, beta); }, '=BETA.DIST(0.5,2,2,TRUE)');
    add('Statistical', ['BINOM.DIST.RANGE'], 'Probability of a range of binomial successes.', (_, a) => { req(a, 3, 4); const nn = positiveInt(n(a, 0), 1000), p = n(a, 1), start = n(a, 2), end = a.length > 3 ? n(a, 3) : start; if (p < 0 || p > 1 || start < 0 || end < start || end > nn)
        fail('#NUM!', 'Invalid binomial domain.'); let total = 0; for (let k = start; k <= end; k++)
        total += choose(nn, k) * p ** k * (1 - p) ** (nn - k); return total; });
    add('Statistical', ['CHISQ.DIST', 'CHISQ.DIST.RT', 'CHISQ.INV', 'CHISQ.INV.RT', 'CHISQ.TEST'], 'Chi-square distribution and hypothesis test.', (name, a) => { if (name === 'CHISQ.TEST')
        return call('CHITEST', a); req(a, name === 'CHISQ.DIST' ? 3 : 2); const x = n(a, 0), df = n(a, 1); if (name === 'CHISQ.DIST')
        return bool(one(a[2])) ? chiCDF(x, df) : safe(Math.exp((df / 2 - 1) * Math.log(x) - x / 2 - lgamma(df / 2) - df / 2 * Math.LN2)); if (name === 'CHISQ.DIST.RT')
        return 1 - chiCDF(x, df); return approximateCDF(t => chiCDF(t, df), 0, Math.max(1e4, df * 100), name === 'CHISQ.INV.RT' ? 1 - x : x); });
    add('Statistical', ['CONFIDENCE.T'], 'Student t confidence interval width.', (_, a) => { req(a, 3); const p = n(a, 0), stdev = n(a, 1), size = n(a, 2); if (p <= 0 || p >= 1 || size < 2 || stdev <= 0)
        fail('#NUM!', 'Invalid confidence parameters.'); return approximateCDF(x => tCDF(x, size - 1), 0, 100, 1 - p / 2) * stdev / Math.sqrt(size); });
    add('Statistical', ['COVARIANCE.P', 'COVARIANCE.S'], 'Population or sample covariance.', (name, a) => { req(a, 2); const x = nums(a[0]), y = nums(a[1]); if (x.length !== y.length || x.length < (name.endsWith('.S') ? 2 : 1))
        fail('#N/A', 'Covariance ranges must be equally sized.'); const mx = avg(x), my = avg(y); return sum(x.map((v, i) => (v - mx) * (y[i] - my))) / (x.length - (name.endsWith('.S') ? 1 : 0)); });
    add('Statistical', ['ERF.PRECISE', 'ERFC.PRECISE'], 'Error function variants.', (name, a) => { req(a, 1); return name === 'ERF.PRECISE' ? erf(n(a, 0)) : 1 - erf(n(a, 0)); });
    add('Statistical', ['F.DIST', 'F.DIST.RT', 'F.INV', 'F.INV.RT', 'F.TEST'], 'F-distribution and variance test.', (name, a) => { if (name === 'F.TEST')
        return call('FTEST', a); req(a, name === 'F.DIST' ? 4 : 3); const x = n(a, 0), d1 = n(a, 1), d2 = n(a, 2); if (name === 'F.DIST')
        return bool(one(a[3])) ? fCDF(x, d1, d2) : safe(Math.exp((d1 / 2) * Math.log(d1 / d2) + (d1 / 2 - 1) * Math.log(x) - (d1 + d2) / 2 * Math.log(1 + d1 * x / d2) - lgamma(d1 / 2) - lgamma(d2 / 2) + lgamma((d1 + d2) / 2))); if (name === 'F.DIST.RT')
        return 1 - fCDF(x, d1, d2); return approximateCDF(y => fCDF(y, d1, d2), 0, 100000, name === 'F.INV.RT' ? 1 - x : x); });
    add('Statistical', ['GAMMA.DIST', 'GAMMA.INV'], 'Gamma distribution and inverse.', (name, a) => { req(a, name === 'GAMMA.DIST' ? 4 : 3); const x = n(a, 0), alpha = n(a, 1), beta = n(a, 2); if (alpha <= 0 || beta <= 0)
        fail('#NUM!', 'Gamma shape and scale must be positive.'); if (name === 'GAMMA.INV')
        return approximateCDF(y => chiCDF(2 * y / beta, 2 * alpha), 0, Math.max(100, alpha * beta * 100), x); return bool(one(a[3])) ? chiCDF(2 * x / beta, 2 * alpha) : Math.exp((alpha - 1) * Math.log(x) - x / beta - lgamma(alpha) - alpha * Math.log(beta)); });
    add('Statistical', ['HYPGEOM.DIST'], 'Hypergeometric probability.', (_, a) => { req(a, 5); const hits = n(a, 0), draws = n(a, 1), success = n(a, 2), population = n(a, 3), cumulative = bool(one(a[4])); const pmf = (i) => choose(success, i) * choose(population - success, draws - i) / choose(population, draws); return cumulative ? sum(Array.from({ length: hits + 1 }, (_, i) => pmf(i))) : pmf(hits); });
    add('Statistical', ['MODE.MULT'], 'All modes (array result; use INDEX).', (_, a) => { const x = a.flatMap(nums), counts = new Map(); x.forEach(v => counts.set(v, (counts.get(v) ?? 0) + 1)); const max = Math.max(...counts.values()); if (max <= 1)
        fail('#N/A', 'No repeated values.'); return [...counts].filter(([, c]) => c === max).map(([v]) => [v]); });
    add('Statistical', ['T.DIST', 'T.DIST.2T', 'T.DIST.RT', 'T.INV', 'T.INV.2T', 'T.TEST'], 'Student t distribution and tests.', (name, a) => { if (name === 'T.TEST')
        return call('TTEST', a); req(a, name === 'T.DIST' ? 3 : 2); const x = n(a, 0), df = n(a, 1); if (name === 'T.DIST')
        return bool(one(a[2])) ? tCDF(x, df) : Math.exp(lgamma((df + 1) / 2) - lgamma(df / 2)) / (Math.sqrt(df * Math.PI)) * (1 + x * x / df) ** (-(df + 1) / 2); if (name === 'T.DIST.RT')
        return 1 - tCDF(x, df); if (name === 'T.DIST.2T')
        return 2 * (1 - tCDF(Math.abs(x), df)); if (name === 'T.INV' && x === 0.5)
        return 0; return approximateCDF(z => tCDF(z, df), -1000, 1000, name === 'T.INV.2T' ? 1 - x / 2 : x); });
    // Array transformations return rectangular matrices to the workbook's spill engine.
    add('Lookup & Reference', ['SEQUENCE', 'RANDARRAY', 'MUNIT'], 'Generate numeric arrays.', (name, a) => {
        req(a, 1, name === 'MUNIT' ? 1 : name === 'RANDARRAY' ? 5 : 4);
        const rows = positiveInt(n(a, 0), 200), cols = a.length > 1 ? positiveInt(n(a, 1), 200) : name === 'MUNIT' ? rows : 1;
        guard(rows, cols);
        if (name === 'MUNIT')
            return Array.from({ length: rows }, (_, i) => Array.from({ length: cols }, (_, j) => i === j ? 1 : 0));
        const start = a.length > 2 ? n(a, 2) : name === 'RANDARRAY' ? 0 : 1, step = a.length > 3 ? n(a, 3) : 1;
        if (name === 'RANDARRAY') {
            if (a.length > 5)
                fail('#ARG!', 'RANDARRAY takes at most 5 arguments.');
            const low = start, high = a.length > 3 ? n(a, 3) : 1, whole = a.length > 4 ? bool(one(a[4])) : false;
            if (low > high)
                fail('#NUM!', 'RANDARRAY min must not exceed max.');
            return Array.from({ length: rows }, () => Array.from({ length: cols }, () => whole ? Math.floor(Math.random() * (Math.floor(high) - Math.ceil(low) + 1)) + Math.ceil(low) : Math.random() * (high - low) + low));
        }
        return Array.from({ length: rows }, (_, i) => Array.from({ length: cols }, (_, j) => start + (i * cols + j) * step));
    }, '=INDEX(SEQUENCE(3,2),2,2)');
    add('Lookup & Reference', ['TAKE', 'DROP'], 'Take or drop rows and columns of an array with Excel empty-array errors.', (name, a) => { req(a, 2, 3); const source = matrix(a, 0), height = source.length, width = Math.max(...source.map(row => row.length)), omittedRows = a[1] === null || a[1] === undefined, omittedColumns = a.length < 3 || a[2] === null || a[2] === undefined, rows = omittedRows ? (name === 'TAKE' ? height : 0) : n(a, 1), columns = omittedColumns ? (name === 'TAKE' ? width : 0) : n(a, 2); if (!Number.isInteger(rows) || !Number.isInteger(columns))
        fail('#VALUE!', 'Rows and columns must be whole numbers.'); if (name === 'TAKE' && (rows === 0 || columns === 0) || name === 'DROP' && rows === 0 && columns === 0)
        fail('#CALC!', 'An empty array cannot be returned.'); if (name === 'TAKE' && (Math.abs(rows) > height || Math.abs(columns) > width))
        fail('#NUM!', 'Requested size is larger than the input array.'); const slice = (length, count, mode) => { if (mode === 'take')
        return count > 0 ? Array.from({ length: count }, (_, i) => i) : Array.from({ length: -count }, (_, i) => length + count + i); if (count === 0)
        return Array.from({ length }, (_, i) => i); return count > 0 ? Array.from({ length: Math.max(0, length - count) }, (_, i) => i + count) : Array.from({ length: Math.max(0, length + count) }, (_, i) => i); }, rowIndexes = slice(height, rows, name === 'TAKE' ? 'take' : 'drop'), columnIndexes = slice(width, columns, name === 'TAKE' ? 'take' : 'drop'); if (!rowIndexes.length || !columnIndexes.length)
        fail('#CALC!', 'DROP removed the entire input array.'); return rowIndexes.map(row => columnIndexes.map(column => source[row]?.[column] ?? null)); });
    add('Lookup & Reference', ['CHOOSECOLS', 'CHOOSEROWS'], 'Select columns or rows by 1-based positions.', (name, a) => { req(a, 2, 50); const m = matrix(a, 0), size = name === 'CHOOSECOLS' ? m[0].length : m.length; const selected = a.slice(1).map((_, i) => { const x = n(a, i + 1); const idx = x > 0 ? x - 1 : size + x; if (!Number.isInteger(idx) || idx < 0 || idx >= size)
        fail('#VALUE!', 'Selected row or column is outside array.'); return idx; }); return name === 'CHOOSECOLS' ? m.map(row => selected.map(j => row[j] ?? null)) : selected.map(i => m[i]); });
    add('Lookup & Reference', ['HSTACK', 'VSTACK'], 'Join rectangular array ranges.', (name, a) => { req(a, 1, 25); const tables = a.map((_, i) => matrix(a, i)); const widths = tables.map(t => Math.max(...t.map(r => r.length))), heights = tables.map(t => t.length); let r = name === 'VSTACK' ? sum(heights) : Math.max(...heights), c = name === 'HSTACK' ? sum(widths) : Math.max(...widths); guard(r, c); if (name === 'VSTACK')
        return tables.flatMap(t => t.map(row => Array.from({ length: c }, (_, j) => row[j] ?? null))); return Array.from({ length: r }, (_, i) => tables.flatMap((t, index) => Array.from({ length: widths[index] }, (_, j) => t[i]?.[j] ?? null))); });
    add('Lookup & Reference', ['TOCOL', 'TOROW'], 'Flatten an array with blank filtering and row- or column-major scan order.', (name, a) => { req(a, 1, 3); const source = matrix(a, 0), ignore = a.length > 1 ? n(a, 1) : 0, byColumn = a.length > 2 ? bool(one(a[2])) : false; if (!Number.isInteger(ignore) || ignore < 0 || ignore > 3)
        fail('#VALUE!', 'Ignore must be 0, 1, 2, or 3.'); const values = byColumn ? source[0].map((_, column) => source.map(row => row[column] ?? null)).flat() : source.flat(), out = values.filter(value => !(ignore % 2 === 1 && !present(value))); return name === 'TOCOL' ? out.map(value => [value]) : [out]; });
    add('Lookup & Reference', ['WRAPCOLS', 'WRAPROWS'], 'Wrap one-dimensional array into rows or columns.', (name, a) => { req(a, 2, 3); const values = matrix(a, 0).flat(), width = positiveInt(n(a, 1), 20000), pad = a.length > 2 ? one(a[2]) : null; const count = Math.ceil(values.length / width); guard(width, count); if (name === 'WRAPROWS')
        return Array.from({ length: count }, (_, i) => Array.from({ length: width }, (_, j) => { const index = i * width + j; return index < values.length ? values[index] : pad; })); return Array.from({ length: width }, (_, i) => Array.from({ length: count }, (_, j) => { const index = j * width + i; return index < values.length ? values[index] : pad; })); });
    add('Lookup & Reference', ['EXPAND'], 'Extend a matrix to new dimensions, applying pad only outside source bounds.', (_, a) => { req(a, 2, 4); const m = matrix(a, 0), sourceRows = m.length, sourceColumns = Math.max(...m.map(row => row.length)), r = positiveInt(n(a, 1)), c = a.length > 2 ? positiveInt(n(a, 2)) : sourceColumns, pad = a.length > 3 ? one(a[3]) : null; if (r < sourceRows || c < sourceColumns)
        fail('#VALUE!', 'EXPAND cannot shrink input.'); guard(r, c); return Array.from({ length: r }, (_, i) => Array.from({ length: c }, (_, j) => i < sourceRows && j < sourceColumns ? (m[i][j] ?? null) : pad)); });
    add('Lookup & Reference', ['FILTER'], 'Filter rows or columns using a matching Boolean include vector.', (_, a) => { req(a, 2, 3); const rows = matrix(a, 0), includeMatrix = matrix(a, 1), ifEmpty = a.length === 3 ? arr(a[2]) : null, height = rows.length, width = Math.max(...rows.map(row => row.length)); if (includeMatrix.every(row => row.length === 1) && includeMatrix.length === height) {
        const include = includeMatrix.map(row => row[0] ?? null), kept = rows.filter((_, i) => bool(include[i]));
        if (kept.length)
            return kept;
        if (ifEmpty)
            return ifEmpty;
        return fail('#CALC!', 'FILTER found no rows.');
    } if (includeMatrix.length === 1 && includeMatrix[0].length === width) {
        const include = includeMatrix[0], selected = include.map((value, index) => bool(value) ? index : -1).filter(index => index >= 0);
        if (selected.length)
            return rows.map(row => selected.map(column => row[column] ?? null));
        if (ifEmpty)
            return ifEmpty;
        return fail('#CALC!', 'FILTER found no columns.');
    } return fail('#VALUE!', 'FILTER include must be a one-dimensional vector matching the array rows or columns.'); });
    add('Lookup & Reference', ['SORT', 'SORTBY'], 'Sort rows or columns by a selected index or by multiple key arrays.', (name, a) => {
        req(a, name === 'SORTBY' ? 2 : 1, name === 'SORTBY' ? 255 : 4);
        const source = matrix(a, 0).map(row => [...row]), byColumns = name === 'SORT' && a.length > 3 ? bool(one(a[3])) : false;
        const transposed = (rows) => rows[0].map((_, column) => rows.map(row => row[column] ?? null));
        const restore = (items) => byColumns ? transposed(items) : items;
        const items = byColumns ? transposed(source) : source, itemCount = items.length;
        const compare = (left, right) => {
            if (left === right)
                return 0;
            if (left === null)
                return -1;
            if (right === null)
                return 1;
            if (typeof left === 'number' && typeof right === 'number')
                return left - right;
            const rank = (value) => typeof value === 'number' ? 0 : typeof value === 'string' ? 1 : 2;
            if (rank(left) !== rank(right))
                return rank(left) - rank(right);
            if (typeof left === 'boolean' && typeof right === 'boolean')
                return Number(left) - Number(right);
            return scalar(left).toLocaleLowerCase().localeCompare(scalar(right).toLocaleLowerCase(), undefined, { numeric: true, sensitivity: 'base' });
        };
        const sortKeys = [];
        if (name === 'SORT') {
            const rawIndex = a.length > 1 ? n(a, 1) : 1, index = rawIndex - 1, order = a.length > 2 ? n(a, 2) : 1;
            if (!Number.isInteger(rawIndex) || rawIndex < 1 || rawIndex > (byColumns ? source.length : source[0].length))
                fail('#VALUE!', 'SORT index must select a row or column in the array.');
            if (order !== 1 && order !== -1)
                fail('#VALUE!', 'SORT order must be 1 (ascending) or -1 (descending).');
            sortKeys.push({ values: items.map(item => item[index] ?? null), order });
        }
        else {
            for (let i = 1; i < a.length;) {
                const keyMatrix = matrix(a, i), key = keyMatrix.flat(), isVector = keyMatrix.length === itemCount && keyMatrix.every(row => row.length === 1) || keyMatrix.length === 1 && keyMatrix[0].length === itemCount, orderArg = a[i + 1], order = orderArg === undefined ? 1 : number(one(orderArg));
                if (!isVector)
                    fail('#VALUE!', 'Each SORTBY key must be a one-dimensional array with one value for every sorted row.');
                if (order !== 1 && order !== -1)
                    fail('#VALUE!', 'SORTBY order must be 1 (ascending) or -1 (descending).');
                sortKeys.push({ values: key, order });
                i += orderArg === undefined ? 1 : 2;
            }
        }
        const ordered = items.map((item, index) => ({ item, index })).sort((left, right) => {
            for (const key of sortKeys) {
                const compared = compare(key.values[left.index], key.values[right.index]) * key.order;
                if (compared)
                    return compared;
            }
            return left.index - right.index;
        }).map(entry => entry.item);
        return restore(ordered);
    });
    add('Lookup & Reference', ['UNIQUE'], 'Return distinct rows or columns using Excel case-insensitive text comparisons.', (_, a) => { req(a, 1, 3); let m = matrix(a, 0), byColumn = a.length > 1 ? bool(one(a[1])) : false, once = a.length > 2 ? bool(one(a[2])) : false; if (byColumn)
        m = m[0].map((_, i) => m.map(row => row[i] ?? null)); const keyOf = (row) => JSON.stringify(row.map(value => typeof value === 'string' ? ['text', value.toLocaleLowerCase()] : value === null ? ['blank'] : typeof value === 'number' ? ['number', value] : ['boolean', value])), counts = new Map(); m.forEach(row => { const key = keyOf(row); counts.set(key, (counts.get(key) ?? 0) + 1); }); const seen = new Set(), unique = m.filter(row => { const key = keyOf(row); if (seen.has(key) || once && counts.get(key) !== 1)
        return false; seen.add(key); return true; }); if (!unique.length)
        fail('#CALC!', 'UNIQUE produced an empty array.'); return byColumn ? unique[0].map((_, j) => unique.map(row => row[j] ?? null)) : unique; });
    add('Lookup & Reference', ['TRIMRANGE'], 'Remove completely empty bounding rows and columns.', (_, a) => { req(a, 1); let m = matrix(a, 0); while (m.length && m[0].every(v => !present(v)))
        m = m.slice(1); while (m.length && m.at(-1).every(v => !present(v)))
        m = m.slice(0, -1); if (!m.length)
        return [[null]]; while (m[0].length && m.every(r => !present(r[0])))
        m = m.map(r => r.slice(1)); while (m[0].length && m.every(r => !present(r.at(-1))))
        m = m.map(r => r.slice(0, -1)); return m; });
    add('Text', ['TEXTSPLIT'], 'Split text using one or more row/column delimiters, with Excel match-mode and padding options.', (_, a) => {
        req(a, 2, 6);
        const source = s(a, 0), delimiters = (index) => matrix(a, index).flat().map(value => scalar(value)), columnDelimiters = a[1] === null || a[1] === undefined ? [] : delimiters(1), rowDelimiters = a.length > 2 ? delimiters(2) : [], ignoreEmpty = a.length > 3 ? bool(one(a[3])) : false, matchMode = a.length > 4 ? n(a, 4) : 0, pad = a.length > 5 ? one(a[5]) : null;
        if (matchMode !== 0 && matchMode !== 1)
            fail('#VALUE!', 'TEXTSPLIT match_mode must be 0 (case-sensitive) or 1 (case-insensitive).');
        const split = (value, separators) => {
            if (!separators.length)
                return [value];
            if (separators.length === 1 && separators[0] === '')
                return value.length ? Array.from(value) : [''];
            const usable = separators.filter(separator => separator.length > 0).sort((left, right) => right.length - left.length);
            if (!usable.length)
                return [value];
            const escaped = usable.map(separator => separator.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'));
            return value.split(new RegExp(escaped.join('|'), matchMode === 1 ? 'gi' : 'g'));
        };
        let rows = rowDelimiters.some(delimiter => delimiter.length > 0) ? split(source, rowDelimiters) : [source];
        if (ignoreEmpty)
            rows = rows.filter(row => row.length > 0);
        if (!rows.length)
            rows = [''];
        const columns = rows.map(value => { const parts = split(value, columnDelimiters); return ignoreEmpty ? parts.filter(part => part.length > 0) : parts; }), width = Math.max(1, ...columns.map(row => row.length));
        guard(columns.length, width);
        return columns.map(row => Array.from({ length: width }, (_, index) => index < row.length ? row[index] : pad));
    }, '=INDEX(TEXTSPLIT("a,b",","),1,2)');
    // Financial securities: US/30 and actual-day coupon conventions are supported; no holiday calendars.
    add('Financial', ['COUPDAYBS', 'COUPDAYS', 'COUPDAYSNC', 'COUPNCD', 'COUPNUM', 'COUPPCD'], 'Coupon-date calculations for regular coupon schedules.', (name, a) => { req(a, 3, 4); const settlement = n(a, 0), maturity = n(a, 1), frequency = n(a, 2), basis = a.length > 3 ? n(a, 3) : 0; if (settlement >= maturity || ![0, 1, 2, 3, 4].includes(basis))
        fail('#NUM!', 'Invalid settlement/maturity or day-count basis.'); const cp = stock(settlement, maturity, frequency, settlement); const total = basis === 0 || basis === 4 ? 360 / frequency : cp.period; const passed = (settlement - cp.previous); return { COUPDAYBS: passed, COUPDAYS: total, COUPDAYSNC: total - passed, COUPNCD: cp.next, COUPNUM: cp.count, COUPPCD: cp.previous }[name]; }, '=COUPNUM(DATE(2026,1,1),DATE(2028,1,1),2)');
    add('Financial', ['INTRATE'], 'Simple annualized rate from redemption and investment.', (_, a) => { req(a, 4, 5); const issue = n(a, 0), maturity = n(a, 1), investment = n(a, 2), redemption = n(a, 3), basis = a.length > 4 ? n(a, 4) : 0; if (maturity <= issue || investment <= 0)
        fail('#NUM!', 'Invalid security dates or investment.'); return (redemption - investment) / investment / yearFrac(issue, maturity, basis); });
    add('Financial', ['ACCRINTM'], 'Accrued interest paid at maturity.', (_, a) => { req(a, 3, 5); const issue = n(a, 0), settlement = n(a, 1), rate = n(a, 2), par = a.length > 3 ? n(a, 3) : 1000, basis = a.length > 4 ? n(a, 4) : 0; if (settlement <= issue || rate < 0 || par <= 0)
        fail('#NUM!', 'Invalid accrued interest dates.'); return par * rate * yearFrac(issue, settlement, basis); });
    add('Financial', ['ACCRINT'], 'Accrued interest on regular-coupon securities.', (_, a) => { req(a, 4, 7); const issue = n(a, 0), first = n(a, 1), settlement = n(a, 2), rate = n(a, 3), par = a.length > 4 ? n(a, 4) : 1000, freq = a.length > 5 ? n(a, 5) : 2, basis = a.length > 6 ? n(a, 6) : 0; if (first <= issue || settlement <= issue || rate < 0 || par <= 0 || ![1, 2, 4].includes(freq))
        fail('#NUM!', 'Invalid accrued interest inputs.'); return par * rate * yearFrac(issue, settlement, basis); });
    add('Financial', ['AMORLINC', 'AMORDEGRC'], 'Depreciation for a period; period zero begins at purchase.', (name, a) => { req(a, 6, 7); const cost = n(a, 0), purchase = n(a, 1), first = n(a, 2), salvage = n(a, 3), period = n(a, 4), rate = n(a, 5); if (cost < 0 || salvage < 0 || cost < salvage || period < 0 || rate <= 0 || first < purchase)
        fail('#NUM!', 'Invalid depreciation data.'); let value = cost; for (let p = 0; p <= period; p++) {
        const multiplier = name === 'AMORDEGRC' ? (rate <= 1 / 6 ? 2.5 : rate <= 1 / 4 ? 2 : 1.5) : 1;
        const loss = Math.min(value - salvage, (p === 0 ? yearFrac(purchase, first) * cost : name === 'AMORLINC' ? cost : value) * rate * multiplier);
        const booked = name === 'AMORDEGRC' ? Math.floor(Math.max(0, loss)) : Math.max(0, loss);
        if (p === period)
            return booked;
        value -= booked;
    } return 0; });
    add('Financial', ['ODDFPRICE', 'ODDLPRICE'], 'Approximate odd-coupon price; invalid date orders and out-of-range rates give Excel-style #NUM!.', (name, a) => {
        const first = name === 'ODDFPRICE';
        req(a, first ? 8 : 7, first ? 9 : 8);
        const settlement = n(a, 0), maturity = n(a, 1), issue = n(a, 2), rate = n(a, first ? 4 : 3), yieldRate = n(a, first ? 5 : 4), redemption = n(a, first ? 6 : 5), freq = n(a, first ? 7 : 6);
        if (maturity <= settlement || rate < 0 || rate > 1 || yieldRate < 0 || redemption <= 0 || ![1, 2, 4].includes(freq) || issue > settlement || (first && n(a, 3) <= issue))
            fail('#NUM!', 'Invalid odd-coupon security inputs.');
        const periods = Math.ceil(yearFrac(settlement, maturity) * freq);
        let pv = redemption / (1 + yieldRate / freq) ** periods;
        for (let i = 1; i <= periods; i++)
            pv += 100 * rate / freq / (1 + yieldRate / freq) ** i;
        return pv;
    });
    add('Financial', ['ODDFYIELD', 'ODDLYIELD'], 'Approximate yield for an odd-coupon security with Excel-compatible input validation.', (name, a) => {
        const first = name === 'ODDFYIELD';
        req(a, first ? 8 : 7, first ? 9 : 8);
        const settlement = n(a, 0), maturity = n(a, 1), issue = n(a, 2), rate = n(a, first ? 4 : 3), price = n(a, first ? 5 : 4), redemption = n(a, first ? 6 : 5), freq = n(a, first ? 7 : 6);
        if (maturity <= settlement || issue > settlement || rate < 0 || rate > 1 || price <= 0 || redemption <= 0 || ![1, 2, 4].includes(freq) || (first && n(a, 3) <= issue))
            fail('#NUM!', 'Invalid odd-coupon security inputs.');
        return (100 * rate + (redemption - price) / yearFrac(settlement, maturity)) / ((redemption + price) / 2);
    });
    // Text & locale. Excel locale-dependent phonetics and baht rendering are limited to deterministic fallback behavior.
    add('Text', ['ASC', 'JIS'], 'Convert full-width and half-width Latin text.', (name, a) => { req(a, 1); const source = s(a, 0); return [...source].map(c => { const cp = c.charCodeAt(0); if (name === 'ASC')
        return c; return cp >= 0x21 && cp <= 0x7E ? String.fromCharCode(cp + 0xFEE0) : c === ' ' ? '　' : c; }).join(''); });
    add('Text', ['PHONETIC'], 'Return stored phonetic text, if available; plain text is returned otherwise.', (_, a) => { req(a, 1); return s(a, 0); });
    const thaiDigits = ['ศูนย์', 'หนึ่ง', 'สอง', 'สาม', 'สี่', 'ห้า', 'หก', 'เจ็ด', 'แปด', 'เก้า'];
    const thaiNum = (v) => { if (!Number.isInteger(v) || v < 0 || v >= 1e6)
        fail('#NUM!', 'BAHTTEXT supports positive integer components up to 999,999.'); if (v === 0)
        return thaiDigits[0]; let out = ''; const units = ['', 'สิบ', 'ร้อย', 'พัน', 'หมื่น', 'แสน']; for (let i = 5; i >= 0; i--) {
        const div = 10 ** i, d = Math.floor(v / div) % 10;
        if (!d)
            continue;
        out += (i === 1 && d === 2 ? 'ยี่' : i === 1 && d === 1 ? '' : i === 0 && d === 1 && out ? 'เอ็ด' : thaiDigits[d]) + (units[i] ?? '');
    } return out; };
    add('Text', ['BAHTTEXT'], 'Write nonnegative baht and satang amounts in Thai words (six digits max).', (_, a) => { req(a, 1); const x = n(a, 0), baht = Math.floor(x), satang = Math.round((x - baht) * 100); if (x < 0)
        fail('#NUM!', 'BAHTTEXT requires a nonnegative amount.'); return thaiNum(baht) + 'บาท' + (satang === 0 ? 'ถ้วน' : thaiNum(satang) + 'สตางค์'); });
    add('Information', ['INFO'], 'Return the modeled Excel-compatible runtime information for this workbook.', (_, a) => { req(a, 1); const type = s(a, 0).toLowerCase(); if (type === 'system') {
        const agent = typeof navigator === 'undefined' ? '' : navigator.userAgent;
        if (/Windows/i.test(agent))
            return 'pcdos';
        if (/Macintosh|Mac OS X/i.test(agent))
            return 'mac';
        return 'web';
    } if (type === 'release')
        return '16.0'; if (type === 'numfile')
        return 1; if (type === 'recalc')
        return 'Automatic'; return fail('#VALUE!', 'Unsupported INFO type.'); });
    add('Information', ['SHEET', 'SHEETS'], 'Return a worksheet position or the number of worksheets in the workbook; reference arguments support one worksheet at a time.', (name, a) => { req(a, 0, 1); return 1; });
    // Bessel series and numerical integration, bounded for browser responsiveness.
    const factorial = (k) => Math.exp(lgamma(k + 1));
    const besselJ = (order, x, modified) => { if (!Number.isInteger(order) || order < 0 || order > 30 || Math.abs(x) > 30)
        fail('#NUM!', 'Bessel J/I supports integer orders 0..30 and |x|<=30.'); let result = 0; for (let k = 0; k < 90; k++) {
        const term = (modified ? 1 : (k % 2 ? -1 : 1)) * (x / 2) ** (2 * k + order) / (factorial(k) * factorial(k + order));
        result += term;
        if (Math.abs(term) < 1e-14)
            break;
    } return result; };
    add('Math & Engineering', ['BESSELI', 'BESSELJ'], 'Modified or ordinary Bessel functions of the first kind.', (name, a) => { req(a, 2); return besselJ(n(a, 1), n(a, 0), name === 'BESSELI'); });
    add('Math & Engineering', ['BESSELK', 'BESSELY'], 'Modified and ordinary Bessel functions of the second kind, numerical integral.', (name, a) => {
        req(a, 2);
        const x = n(a, 0), order = n(a, 1);
        if (x <= 0 || x > 30 || order < 0 || order > 12 || !Number.isInteger(order))
            fail('#NUM!', 'Positive argument and order 0..12 required.');
        if (name === 'BESSELK') {
            const slices = 1500, hi = 20, h = hi / slices;
            let total = 0;
            for (let i = 0; i <= slices; i++) {
                const t = i * h, term = Math.exp(-x * Math.cosh(t)) * Math.cosh(order * t);
                total += (i === 0 || i === slices ? 1 : i % 2 ? 4 : 2) * term;
            }
            return total * h / 3;
        }
        const h = Math.PI / 2400;
        let first = 0;
        for (let i = 0; i <= 2400; i++) {
            const t = i * h;
            first += (i === 0 || i === 2400 ? 1 : i % 2 ? 4 : 2) * Math.sin(x * Math.sin(t) - order * t);
        }
        first *= h / (3 * Math.PI);
        const ht = 15 / 3000;
        let tail = 0;
        for (let i = 0; i <= 3000; i++) {
            const t = i * ht, ex = -x * Math.sinh(t);
            const v = Math.exp(ex + order * t) + (order % 2 ? -1 : 1) * Math.exp(ex - order * t);
            tail += (i === 0 || i === 3000 ? 1 : i % 2 ? 4 : 2) * v;
        }
        return first - tail * ht / (3 * Math.PI);
    });
    add('Math & Engineering', ['PERCENTOF'], 'Percentage of a subtotal of a numeric range.', (_, a) => { req(a, 2); const subset = sum(nums(a[0])), all = sum(nums(a[1])); if (all === 0)
        fail('#DIV/0!', 'Cannot divide by zero total.'); return subset / all; });
    add('Other', ['AGGREGATE'], 'Aggregate codes 1–19 with reference-form options for hidden rows, errors, and nested SUBTOTAL/AGGREGATE formulas; computed array expressions remain limited.', (_, a) => {
        req(a, 3, 4);
        const code = n(a, 0), option = n(a, 1);
        if (!Number.isInteger(code) || code < 1 || code > 19)
            fail('#VALUE!', 'AGGREGATE function_num must be a supported integer from 1 to 19.');
        if (!Number.isInteger(option) || option < 0 || option > 7)
            fail('#VALUE!', 'AGGREGATE option must be an integer from 0 to 7.');
        const takesRank = code >= 14;
        req(a, takesRank ? 4 : 3, takesRank ? 4 : 3);
        const values = nums(a[2]);
        if (code <= 11) {
            if (!values.length)
                fail('#VALUE!', 'AGGREGATE input contains no numeric values.');
            const mapping = { 1: 1, 2: 2, 3: 3, 4: 4, 5: 5, 6: 6, 7: 7, 8: 8, 9: 9, 10: 10, 11: 11 };
            return call('SUBTOTAL', [mapping[code], a[2]]);
        }
        if (code === 12)
            return call('MEDIAN', [a[2]]);
        if (code === 13)
            return call('MODE.SNGL', [a[2]]);
        if (code === 14 || code === 15) {
            if (!values.length)
                fail('#VALUE!', 'AGGREGATE input contains no numeric values.');
            const k = positiveInt(n(a, 3)), sorted = [...values].sort((x, y) => code === 14 ? y - x : x - y);
            if (k > sorted.length)
                fail('#NUM!', 'Rank exceeds data size.');
            return sorted[k - 1];
        }
        const percentileName = code === 16 ? 'PERCENTILE.INC' : code === 17 ? 'QUARTILE.INC' : code === 18 ? 'PERCENTILE.EXC' : 'QUARTILE.EXC';
        return call(percentileName, [a[2], a[3]]);
    });
    // Workbook evaluates LAMBDA expressions in a scoped formula context; these
    // definitions validate non-lambda calls and provide metadata for the library.
    add('Lookup & Reference', ['BYROW', 'BYCOL'], 'Apply a scoped LAMBDA to each row or column and return one result per row or column.', (_, a) => { req(a, 2); return fail('#VALUE!', 'BYROW and BYCOL require a LAMBDA expression.'); }, '=BYROW(A1:B3,LAMBDA(row,SUM(row)))');
    add('Lookup & Reference', ['MAKEARRAY'], 'Generate an array by invoking a two-parameter LAMBDA for each row and column.', (_, a) => { req(a, 3); return fail('#VALUE!', 'MAKEARRAY requires a LAMBDA expression.'); }, '=MAKEARRAY(3,2,LAMBDA(row,col,row*col))');
    add('Lookup & Reference', ['MAP'], 'Map one or more arrays with a LAMBDA and return a same-sized result array.', (_, a) => { req(a, 2); return fail('#VALUE!', 'MAP requires a LAMBDA expression as its final argument.'); }, '=MAP(A1:A3,LAMBDA(value,value*2))');
    add('Lookup & Reference', ['REDUCE', 'SCAN'], 'Reduce an array or return intermediate accumulator values using a LAMBDA.', (name, a) => { req(a, 3); return fail('#VALUE!', `${name} requires a LAMBDA expression as its final argument.`); }, '=SCAN(0,A1:A3,LAMBDA(total,value,total+value))');
    add('Lookup & Reference', ['LET'], 'Evaluate local name/value bindings and a final calculation; bindings are resolved by the workbook formula evaluator.', () => fail('#INTERNAL!', 'LET is evaluated with local workbook formula bindings.'), '=LET(rate,0.08,amount,100,amount*rate)');
    add('Lookup & Reference', ['LAMBDA'], 'Create a scoped calculation for supported lambda-aware functions or a worksheet-local defined name.', (_, a) => { req(a, 2, 254); return fail('#CALC!', 'A LAMBDA must be invoked or passed to a lambda-aware function.'); }, '=MAP(A1:A3,LAMBDA(value,value*2))');
    add('Information', ['ISOMITTED'], 'Check whether an optional parameter was omitted when a LAMBDA was invoked.', (_, a) => { req(a, 1); return fail('#VALUE!', 'ISOMITTED requires a LAMBDA parameter.'); }, '=LAMBDA(value,optional,ISOMITTED(optional))(1,)');
    add('Statistical', ['FORECAST.ETS'], 'Bounded additive triple-exponential smoothing for seasonal series; nonseasonal, short, or oversized data fall back to linear regression. Blank values honor data_completion.', (_, a) => { req(a, 3, 6); validateForecastOptions(a, 3, 4, 5); const target = n(a, 0), series = forecastSeries(a[1], a[2], a.length >= 5 ? n(a, 4) : 1, a.length >= 6 ? n(a, 5) : 0); if (target < series.x.at(-1))
        fail('#NUM!', 'FORECAST.ETS target date cannot precede the timeline end.'); const model = forecastEtsModel(series, a.length >= 4 ? n(a, 3) : 1); return model.forecastAt((target - series.x.at(-1)) / series.step); }, '=FORECAST.ETS(14,B1:B12,A1:A12)');
    add('Statistical', ['FORECAST.ETS.CONFINT'], 'Confidence-scaled interval estimate for the bounded additive ETS model; linear fallback and approximate uncertainty remain. Blank values honor data_completion.', (_, a) => { req(a, 3, 7); validateForecastOptions(a, 4, 5, 6); const target = n(a, 0), series = forecastSeries(a[1], a[2], a.length >= 6 ? n(a, 5) : 1, a.length >= 7 ? n(a, 6) : 0); if (target < series.x.at(-1))
        fail('#NUM!', 'FORECAST.ETS.CONFINT target date cannot precede the timeline end.'); const confidence = a.length >= 4 ? n(a, 3) : 0.95; if (!(confidence > 0 && confidence < 1))
        fail('#NUM!', 'Confidence level must be greater than zero and less than one.'); const model = forecastEtsModel(series, a.length >= 5 ? n(a, 4) : 1), steps = Math.max(1, (target - series.x.at(-1)) / series.step), leverage = 1 + 1 / series.x.length + ((target - series.xMean) ** 2 / series.xVariance); return safe(normalQuantile((1 + confidence) / 2) * model.rmse * Math.sqrt(steps * leverage)); }, '=FORECAST.ETS.CONFINT(14,B1:B12,A1:A12)');
    add('Statistical', ['FORECAST.ETS.SEASONALITY'], 'Estimate a repeated cycle length from equally spaced observations; returns zero when no strong cycle is detected. Blank values honor data_completion.', (_, a) => { req(a, 2, 4); if (a.length >= 3 && ![0, 1].includes(n(a, 2)))
        fail('#VALUE!', 'Data completion must be zero or one.'); if (a.length >= 4) {
        const aggregation = n(a, 3);
        if (!Number.isInteger(aggregation) || aggregation < 0 || aggregation > 6)
            fail('#VALUE!', 'Aggregation must be an integer from zero to six.');
    } const forecast = forecastSeries(a[0], a[1], a.length >= 3 ? n(a, 2) : 1, a.length >= 4 ? n(a, 3) : 0); return forecastSeasonality(forecast.y); }, '=FORECAST.ETS.SEASONALITY(B1:B24,A1:A24)');
    add('Statistical', ['FORECAST.ETS.STAT'], 'Return statistics from the bounded additive ETS approximation; smoothing coefficients are exposed when a seasonal model fits. Blank values honor data_completion.', (_, a) => { req(a, 3, 6); validateForecastOptions(a, 3, 4, 5); const statistic = n(a, 2); if (!Number.isInteger(statistic) || statistic < 1 || statistic > 8)
        fail('#NUM!', 'Statistic type must be an integer from 1 to 8.'); const series = forecastSeries(a[0], a[1], a.length >= 5 ? n(a, 4) : 1, a.length >= 6 ? n(a, 5) : 0), model = forecastEtsModel(series, a.length >= 4 ? n(a, 3) : 1); switch (statistic) {
        case 1: return model.alpha ?? fail('#N/A', 'No seasonal smoothing coefficient is available for the linear fallback.');
        case 2: return model.beta ?? fail('#N/A', 'No seasonal smoothing coefficient is available for the linear fallback.');
        case 3: return model.gamma ?? fail('#N/A', 'No seasonal smoothing coefficient is available for the linear fallback.');
        case 4: {
            const scale = avg(series.y.slice(1).map((value, index) => Math.abs(value - series.y[index])));
            return scale ? model.mae / scale : fail('#DIV/0!', 'MASE is unavailable for a constant series.');
        }
        case 5: {
            const actuals = series.y.slice(model.startIndex);
            return avg(model.residual.map((error, index) => { const actual = actuals[index], predicted = actual - error, den = Math.abs(actual) + Math.abs(predicted); return den ? 2 * Math.abs(error) / den : 0; }));
        }
        case 6: return model.mae;
        case 7: return model.rmse;
        default: return series.step;
    } }, '=FORECAST.ETS.STAT(B1:B12,A1:A12,7)');
    add('Database', ['GROUPBY', 'PIVOTBY'], 'Formula-based grouping, aggregation, filtering, sorting, and grand totals; the workbook evaluator handles their array and LAMBDA arguments.', (name, a) => { req(a, name === 'GROUPBY' ? 3 : 4); return fail('#VALUE!', `${name} requires workbook-context array evaluation.`); }, '=GROUPBY(A2:A8,B2:B8,SUM)');
    add('Lookup & Reference', ['FLATTEN'], 'Flatten a rectangular array into one column (newer Excel beta-style convenience).', (_, a) => { req(a, 1); return matrix(a, 0).flat().map(v => [v]); });
    add('Logical', ['ISBETWEEN'], 'Test whether a number falls within inclusive or exclusive bounds.', (_, a) => { req(a, 3, 4); const value = n(a, 0), lo = n(a, 1), hi = n(a, 2), inclusive = a.length > 3 ? bool(one(a[3])) : true; return inclusive ? value >= lo && value <= hi : value > lo && value < hi; });
    add('Financial', ['EUROCONVERT'], 'Convert historic euro fixed-rate currencies (legacy Excel feature).', (_, a) => { req(a, 3, 5); const amount = n(a, 0), from = s(a, 1).toUpperCase(), to = s(a, 2).toUpperCase(); const rates = { EUR: 1, ATS: 13.7603, BEF: 40.3399, DEM: 1.95583, ESP: 166.386, FIM: 5.94573, FRF: 6.55957, IEP: .787564, ITL: 1936.27, LUF: 40.3399, NLG: 2.20371, PTE: 200.482, GRD: 340.750 }; if (!(from in rates) || !(to in rates))
        fail('#VALUE!', 'Unsupported historical euro currency.'); return amount / rates[from] * rates[to]; });
    add('Information', ['FIELDVALUE'], 'Read a field from a linked Excel data type (linked data types are not available in GridForge).', (_, a) => { req(a, 2); return fail('#FIELD!', 'The value does not contain a supported linked data type.'); });
    add('Text', ['FILTERXML'], 'Evaluate bounded XPath queries against offline XML; node selections spill as a vertical list when the host provides a DOM parser.', (_, a) => { req(a, 2); return filterXml(s(a, 0), s(a, 1)); });
    add('Information', ['DETECTLANGUAGE'], 'Detect a few languages offline from written characters; unsupported text returns und.', (_, a) => { req(a, 1); const v = s(a, 0); if (/[\u0900-\u097F]/.test(v))
        return 'hi'; if (/[\u0E00-\u0E7F]/.test(v))
        return 'th'; if (/[\u3040-\u30FF]/.test(v))
        return 'ja'; if (/[\u4E00-\u9FFF]/.test(v))
        return 'zh'; if (/[A-Za-z]/.test(v))
        return 'en'; return 'und'; });
    add('Text', ['TRANSLATE'], 'Offline translation dictionary for limited common greetings.', (_, a) => { req(a, 3); const source = s(a, 0).trim().toLowerCase(), from = s(a, 1).toLowerCase(), to = s(a, 2).toLowerCase(); const dictionary = { hello: { es: 'hola', fr: 'bonjour', hi: 'नमस्ते' }, thankyou: { es: 'gracias', fr: 'merci', hi: 'धन्यवाद' }, goodbye: { es: 'adiós', fr: 'au revoir', hi: 'अलविदा' } }; const key = source.replace(/\s+/g, ''); if (from !== 'en' || !dictionary[key]?.[to])
        fail('#N/A', 'This offline dictionary cannot translate the requested text or language.'); return dictionary[key][to]; });
    add('Information', ['IMAGE'], 'Return safe image URL as a text placeholder (GridForge cell image objects not supported).', (_, a) => { req(a, 1, 5); const value = s(a, 0); if (!/^https:\/\//i.test(value))
        fail('#VALUE!', 'IMAGE requires an HTTPS URL.'); return a.length > 1 ? s(a, 1) || value : value; });
    // External/network Excel functions are deliberately not registered: no false claim of live data services.
    return f;
}
export const v10FunctionNames = Object.values(v10Categories).flat();
