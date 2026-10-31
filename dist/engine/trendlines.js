/** Fit a polynomial in a centered/scaled x domain to avoid huge powers of date-like X values. */
const polynomialFit = (xs, ys, order) => {
    const min = Math.min(...xs), max = Math.max(...xs), range = max - min;
    if (!Number.isFinite(range))
        return null;
    const center = min + range / 2, scale = range / 2;
    if (scale === 0)
        return null;
    const degree = order + 1, vectors = Array.from({ length: degree }, (_, power) => xs.map(x => ((x - center) / scale) ** power)), q = [], r = Array.from({ length: degree }, () => Array(degree).fill(0));
    const dot = (a, b) => a.reduce((sum, value, index) => sum + value * b[index], 0);
    for (let col = 0; col < degree; col++) {
        const v = [...vectors[col]];
        for (let prior = 0; prior < col; prior++) {
            const projection = dot(q[prior], v);
            r[prior][col] = projection;
            for (let row = 0; row < v.length; row++)
                v[row] = v[row] - projection * q[prior][row];
        }
        const norm = Math.sqrt(dot(v, v));
        if (!Number.isFinite(norm) || norm <= Number.EPSILON * xs.length * 32)
            return null;
        r[col][col] = norm;
        q.push(v.map(value => value / norm));
    }
    const rhs = q.map(column => dot(column, ys)), coefficients = Array(degree).fill(0);
    for (let row = degree - 1; row >= 0; row--) {
        let value = rhs[row];
        for (let col = row + 1; col < degree; col++)
            value -= r[row][col] * coefficients[col];
        coefficients[row] = value / r[row][row];
    }
    if (coefficients.some(value => !Number.isFinite(value)))
        return null;
    const predict = (x) => { const z = (x - center) / scale; let value = 0; for (let i = coefficients.length - 1; i >= 0; i--)
        value = value * z + coefficients[i]; return value; };
    const terms = coefficients.map((value, power) => `${value.toPrecision(5)}${power ? 'z' + (power > 1 ? `^${power}` : '') : ''}`).reverse().join(' + ');
    return { predict, equation: `y = ${terms}; z = (x − ${center.toPrecision(5)}) / ${scale.toPrecision(5)}` };
};
const compensatedSum = (values) => { let sum = 0, correction = 0; for (const value of values) {
    const adjusted = value - correction, next = sum + adjusted;
    correction = (next - sum) - adjusted;
    sum = next;
} return sum; };
/** Center observations before forming moments so date-like X values do not
 * lose slope precision to cancellation in n*sum(x²)-sum(x)². */
const linearFit = (xs, ys) => {
    if (xs.length !== ys.length || xs.length < 2)
        return null;
    const n = xs.length, meanX = compensatedSum(xs) / n, meanY = compensatedSum(ys) / n;
    const dx = xs.map(x => x - meanX), dy = ys.map(y => y - meanY), variance = compensatedSum(dx.map(value => value * value)), covariance = compensatedSum(dx.map((value, index) => value * dy[index]));
    if (!Number.isFinite(variance) || !Number.isFinite(covariance) || variance <= 0)
        return null;
    const slope = covariance / variance, intercept = meanY - slope * meanX;
    if (!Number.isFinite(slope) || !Number.isFinite(intercept))
        return null;
    return { slope, intercept, predict: (x) => slope * x + intercept };
};
export function fitTrendline(values, spec, xValues) {
    const points = values.map((y, i) => ({ x: xValues ? xValues[i] : i + 1, y })).filter((p) => typeof p.y === 'number' && Number.isFinite(p.y) && typeof p.x === 'number' && Number.isFinite(p.x)).sort((a, b) => a.x - b.x);
    const period = Math.max(2, Math.min(100, Math.trunc(spec.period ?? 2))), order = Math.max(2, Math.min(6, Math.trunc(spec.order ?? 2))), forward = Math.max(0, Math.min(50, Math.trunc(spec.forward ?? 0))), backward = Math.max(0, Math.min(50, Math.trunc(spec.backward ?? 0)));
    if (spec.type === 'movingAverage') {
        if (points.length < period)
            return { points: [], equation: null, rSquared: null, error: 'Moving-average period exceeds the available data.' };
        const step = points.length > 1 ? (points.at(-1).x - points[0].x) / (points.length - 1) : 1, averages = points.map((_, i) => i + 1 < period ? null : points.slice(i - period + 1, i + 1).reduce((s, x) => s + x.y, 0) / period), first = averages.find((v) => v !== null), last = averages.at(-1);
        const out = [];
        for (let i = -backward; i < points.length + forward; i++) {
            const x = i < 0 ? points[0].x + i * step : i >= points.length ? points.at(-1).x + (i - points.length + 1) * step : points[i].x, y = i < 0 ? first : i >= points.length ? last : averages[i] ?? null;
            out.push({ x, y });
        }
        return { points: out, equation: null, rSquared: null };
    }
    let predict = null, equation = null;
    if (spec.type === 'linear') {
        const fit = linearFit(points.map(p => p.x), points.map(p => p.y));
        if (fit) {
            predict = fit.predict;
            equation = `y = ${fit.slope.toPrecision(5)}x ${fit.intercept < 0 ? '−' : '+'} ${Math.abs(fit.intercept).toPrecision(5)}`;
        }
    }
    else if (spec.type === 'exponential' || spec.type === 'power') {
        const transformed = points.filter(p => p.y > 0 && (spec.type !== 'power' || p.x > 0)), fit = linearFit(transformed.map(p => spec.type === 'power' ? Math.log(p.x) : p.x), transformed.map(p => Math.log(p.y)));
        if (fit) {
            predict = (x) => spec.type === 'power' && x <= 0 ? null : Math.exp(fit.intercept + fit.slope * (spec.type === 'power' ? Math.log(x) : x));
            equation = spec.type === 'power' ? `y = ${Math.exp(fit.intercept).toPrecision(5)}x^${fit.slope.toPrecision(5)}` : `y = ${Math.exp(fit.intercept).toPrecision(5)}e^(${fit.slope.toPrecision(5)}x)`;
            if (transformed.length < points.length)
                return { points: [], equation: null, rSquared: null, error: `${spec.type} trendlines require ${spec.type === 'power' ? 'positive x and y' : 'positive y'} values.` };
        }
    }
    else if (spec.type === 'logarithmic') {
        const valid = points.filter(p => p.x > 0), fit = linearFit(valid.map(p => Math.log(p.x)), valid.map(p => p.y));
        if (valid.length < points.length)
            return { points: [], equation: null, rSquared: null, error: 'Logarithmic trendlines require positive x values.' };
        if (fit) {
            predict = x => x <= 0 ? null : fit.predict(Math.log(x));
            equation = `y = ${fit.slope.toPrecision(5)} ln(x) ${fit.intercept < 0 ? '−' : '+'} ${Math.abs(fit.intercept).toPrecision(5)}`;
        }
    }
    else if (spec.type === 'polynomial') {
        if (points.length <= order)
            return { points: [], equation: null, rSquared: null, error: 'Polynomial order must be lower than the number of valid points.' };
        const fit = polynomialFit(points.map(p => p.x), points.map(p => p.y), order);
        if (fit) {
            predict = fit.predict;
            equation = fit.equation;
        }
    }
    if (!predict)
        return { points: [], equation: null, rSquared: null, error: 'The data is degenerate for this trendline type.' };
    const mean = points.reduce((s, p) => s + p.y, 0) / points.length, ssTot = points.reduce((s, p) => s + (p.y - mean) ** 2, 0), ssRes = points.reduce((s, p) => s + (p.y - (predict(p.x) ?? NaN)) ** 2, 0), rSquared = ssTot === 0 ? null : 1 - ssRes / ssTot;
    const step = points.length > 1 ? (points.at(-1).x - points[0].x) / (points.length - 1) : 1, lo = points[0].x - step * backward, hi = points.at(-1).x + step * forward;
    return { points: Array.from({ length: Math.max(1, Math.round((hi - lo) / (step || 1)) + 1) }, (_, i) => { const x = lo + i * (step || 1), y = predict(x); return { x, y: y !== null && Number.isFinite(y) ? y : null }; }), equation, rSquared };
}
