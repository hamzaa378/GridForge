/** Native, editable ECMA-376 spreadsheet charts. All chart data stays offline; no external relationships. */
import { display } from './workbook.js';
import { colToName } from './parser.js';
import { chartRange, extractChartData, stockData, waterfallTotalPointIndexes, CHART_PALETTES } from './charts.js';
import { nativeSlicerAnchorXml } from './ooxml_slicers.js';
const C = 'http://schemas.openxmlformats.org/drawingml/2006/chart';
const A = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const D = 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const P = 'http://schemas.openxmlformats.org/package/2006/relationships';
const CX = 'http://schemas.microsoft.com/office/drawing/2014/chartex';
const CHART_EX_REL = 'http://schemas.microsoft.com/office/2014/relationships/chartEx';
export const isNativeChartEx = (kind) => ['histogram', 'waterfall', 'treemap', 'sunburst', 'funnel', 'pareto', 'box-whisker'].includes(kind);
const escapeXml = (s) => s.replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&apos;' }[c]));
const chartExSeriesShape = (chart) => { const color = chart.seriesOptions?.[0]?.color; if (!color || !/^#[0-9a-f]{6}$/i.test(color))
    return ''; return `<cx:spPr><a:solidFill><a:srgbClr val="${color.slice(1).toUpperCase()}"/></a:solidFill></cx:spPr>`; };
// ChartEx visibility defines value/categoryName/seriesName; it has no percentage visibility flag.
const chartExDataLabels = (chart) => chart.dataLabels || chart.dataLabelCategoryName || chart.dataLabelSeriesName ? `<cx:dataLabels><cx:visibility value="${chart.dataLabels ? '1' : '0'}" categoryName="${chart.dataLabelCategoryName ? '1' : '0'}" seriesName="${chart.dataLabelSeriesName ? '1' : '0'}"/></cx:dataLabels>` : '';
const colRange = (sheet, c, start, end) => `'${sheet.replace(/'/g, "''")}'!$${colToName(c)}$${start + 1}:$${colToName(c)}$${end + 1}`;
const rowRange = (sheet, row, start, end) => `'${sheet.replace(/'/g, "''")}'!$${colToName(start)}$${row + 1}:$${colToName(end)}$${row + 1}`;
const pal = ['2563EB', '12A594', 'DF8B32', '9866E6', 'DC5B73', '4D97CF', 'A99329', '36948A'];
const rich = (name) => `<c:${name}><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US"/><a:t>${escapeXml(name)}</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/></c:${name}>`;
const titleNode = (value) => `<c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US"/><a:t>${escapeXml(value)}</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/></c:title>`;
const numRef = (ref, values) => `<c:numRef><c:f>${escapeXml(ref)}</c:f><c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="${values.length}"/>${values.map((v, i) => v === null || !Number.isFinite(v) ? '' : `<c:pt idx="${i}"><c:v>${v}</c:v></c:pt>`).join('')}</c:numCache></c:numRef>`;
const labelsRef = (ref, labels) => ref ? `<c:strRef><c:f>${escapeXml(ref)}</c:f><c:strCache><c:ptCount val="${labels.length}"/>${labels.map((s, i) => `<c:pt idx="${i}"><c:v>${escapeXml(s)}</c:v></c:pt>`).join('')}</c:strCache></c:strRef>` : `<c:strLit><c:ptCount val="${labels.length}"/>${labels.map((s, i) => `<c:pt idx="${i}"><c:v>${escapeXml(s)}</c:v></c:pt>`).join('')}</c:strLit>`;
const legend = (chart) => chart.legend === 'none' ? '' : `<c:legend><c:legendPos val="${chart.legend === 'right' ? 'r' : chart.legend === 'top' ? 't' : 'b'}"/><c:layout/><c:overlay val="0"/></c:legend>`;
const chartExLegend = (chart) => chart.legend === 'none' ? '' : `<cx:legend pos="${chart.legend === 'right' ? 'r' : chart.legend === 'top' ? 't' : 'b'}" align="ctr" overlay="0"/>`;
const axisBounds = (minimum, maximum) => `<c:scaling><c:orientation val="minMax"/>${minimum !== undefined ? `<c:min val="${minimum}"/>` : ''}${maximum !== undefined ? `<c:max val="${maximum}"/>` : ''}</c:scaling>`;
const numberFormat = (format, code) => code ?? (format === 'currency' ? '$#,##0.00' : format === 'percent' ? '0%' : format === 'number' ? '0.00' : 'General');
const gridlinesFor = (chart) => `${chart.majorGridlines === false ? '' : '<c:majorGridlines/>'}${chart.minorGridlines ? '<c:minorGridlines/>' : ''}`;
const axes = (chart, scatter) => {
    const catTitle = chart.xAxisTitle ? titleNode(chart.xAxisTitle) : '', valTitle = chart.yAxisTitle ? titleNode(chart.yAxisTitle) : '';
    const cax = scatter ? `<c:valAx><c:axId val="401"/>${axisBounds(chart.xAxisMin, chart.xAxisMax)}<c:axPos val="b"/>${catTitle}${gridlinesFor(chart)}<c:numFmt formatCode="${escapeXml(numberFormat(chart.xAxisNumberFormat, chart.xAxisFormatCode))}" sourceLinked="0"/>${chart.xAxisMajorUnit !== undefined ? `<c:majorUnit val="${chart.xAxisMajorUnit}"/>` : ''}<c:tickLblPos val="nextTo"/><c:crossAx val="402"/><c:crosses val="autoZero"/><c:crossBetween val="between"/></c:valAx>` : `<c:catAx><c:axId val="401"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:axPos val="b"/>${catTitle}<c:tickLblPos val="nextTo"/><c:crossAx val="402"/><c:crosses val="autoZero"/><c:auto val="1"/><c:lblAlgn val="ctr"/>${chart.xAxisLabelInterval ? `<c:tickLblSkip val="${chart.xAxisLabelInterval}"/><c:tickMarkSkip val="${chart.xAxisLabelInterval}"/>` : ''}${chart.xAxisLabelAngle ? `<c:txPr><a:bodyPr rot="${chart.xAxisLabelAngle * 60000}"/><a:lstStyle/><a:p><a:pPr><a:defRPr/></a:pPr><a:endParaRPr/></a:p></c:txPr>` : ''}</c:catAx>`;
    const vax = `<c:valAx><c:axId val="402"/>${axisBounds(chart.yAxisMin, chart.yAxisMax)}<c:axPos val="l"/>${valTitle}${gridlinesFor(chart)}<c:numFmt formatCode="${escapeXml(numberFormat(chart.yAxisNumberFormat, chart.yAxisFormatCode))}" sourceLinked="0"/>${chart.yAxisMajorUnit !== undefined ? `<c:majorUnit val="${chart.yAxisMajorUnit}"/>` : ''}<c:tickLblPos val="nextTo"/><c:crossAx val="401"/><c:crosses val="autoZero"/><c:crossBetween val="between"/></c:valAx>`;
    return cax + vax;
};
export function nativeChartXml(book, sheetName, chart, pivotSourceName) {
    const pivotSource = pivotSourceName ? `<c:pivotSource><c:name>${escapeXml(pivotSourceName)}</c:name><c:fmtId val="0"/></c:pivotSource>` : '';
    if (chart.kind === 'histogram') {
        const bounds = chartRange(chart.range, book.rows, book.cols), data = extractChartData(book, chart.range, false), series = data.series[0];
        if (data.series.length !== 1 || !series)
            throw Error('Native Histogram export needs exactly one numeric value series so its original worksheet range stays editable.');
        const values = series.values, nonBlankValues = values.filter((value) => value !== null && Number.isFinite(value));
        if (!nonBlankValues.length)
            throw Error('Native Histogram export needs at least one numeric value.');
        const underflow = chart.histogramUnderflow, overflow = chart.histogramOverflow;
        if (underflow !== undefined && (!Number.isFinite(underflow) || Math.abs(underflow) > 1e15) || overflow !== undefined && (!Number.isFinite(overflow) || Math.abs(overflow) > 1e15) || underflow !== undefined && overflow !== undefined && underflow >= overflow)
            throw Error('Histogram underflow threshold must be less than the overflow threshold, and thresholds must be finite.');
        const bins = chart.histogramBins ?? Math.max(1, Math.min(20, Math.ceil(Math.log2(nonBlankValues.length) + 1))), binning = chart.histogramBinWidth === undefined ? `<cx:binCount>${bins}</cx:binCount>` : `<cx:binSize>${chart.histogramBinWidth}</cx:binSize>`, binningAttributes = `${underflow === undefined ? '' : ` underflow="${underflow}"`}${overflow === undefined ? '' : ` overflow="${overflow}"`}`, start = bounds.r2 - data.rows + 1, valueCol = bounds.c2;
        const formula = colRange(sheetName, valueCol, start, bounds.r2), points = values.map((value, index) => value === null || !Number.isFinite(value) ? '' : `<cx:pt idx="${index}">${value}</cx:pt>`).join(''), seriesShape = chartExSeriesShape(chart), dataLabels = chartExDataLabels(chart), title = chart.title ? `<cx:title pos="t" align="ctr" overlay="0"><cx:tx><cx:rich><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US"/><a:t>${escapeXml(chart.title)}</a:t></a:r><a:endParaRPr lang="en-US"/></a:p></a:rich></cx:tx></cx:title>` : '';
        return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><cx:chartSpace xmlns:cx="${CX}" xmlns:a="${A}" version="1.2"><cx:chartData><cx:data id="0"><cx:numDim type="val"><cx:f>${escapeXml(formula)}</cx:f><cx:lvl ptCount="${values.length}" name="${escapeXml(series.name)}" formatCode="General">${points}</cx:lvl></cx:numDim></cx:data></cx:chartData><cx:chart>${title}<cx:plotArea><cx:plotAreaRegion><cx:series layoutId="clusteredColumn" uniqueId="{histogram-gridforge}">${seriesShape}${dataLabels}<cx:dataId val="0"/><cx:layoutPr><cx:binning${binningAttributes}>${binning}</cx:binning></cx:layoutPr></cx:series></cx:plotAreaRegion></cx:plotArea>${chartExLegend(chart)}</cx:chart></cx:chartSpace>`;
    }
    const chartExLayouts = { waterfall: 'waterfall', treemap: 'treemap', sunburst: 'sunburst', funnel: 'funnel', pareto: 'paretoLine', 'box-whisker': 'boxWhisker' };
    const chartExLayout = chartExLayouts[chart.kind];
    if (chartExLayout) {
        const bounds = chartRange(chart.range, book.rows, book.cols), hierarchy = chart.kind === 'treemap' || chart.kind === 'sunburst', data = hierarchy ? undefined : extractChartData(book, chart.range);
        if (!hierarchy && (!data || data.series.length !== 1 || data.labels.length === 0))
            throw Error(`${chart.kind} ChartEx export needs one numeric value column and category labels.`);
        const start = bounds.r1 + 1, end = bounds.r2, categoryCol = bounds.c1, valueCol = bounds.c2;
        const levelCount = hierarchy ? valueCol - categoryCol : 1, count = hierarchy ? end - start + 1 : data.labels.length;
        const categoryFormula = hierarchy
            ? `'${sheetName.replace(/'/g, "''")}'!$${colToName(categoryCol)}$${start + 1}:$${colToName(valueCol - 1)}$${end + 1}`
            : colRange(sheetName, categoryCol, start, end);
        const categories = hierarchy
            ? Array.from({ length: levelCount }, (_, level) => {
                const col = categoryCol + level, header = display(book.getValue(`${colToName(col)}${bounds.r1 + 1}`)).trim() || `Level ${level + 1}`;
                const points = Array.from({ length: count }, (_, index) => {
                    const label = display(book.getValue(`${colToName(col)}${start + index + 1}`));
                    return `<cx:pt idx="${index}"><cx:v>${escapeXml(label)}</cx:v></cx:pt>`;
                }).join('');
                return `<cx:lvl ptCount="${count}" name="${escapeXml(header)}">${points}</cx:lvl>`;
            }).join('')
            : `<cx:lvl ptCount="${data.labels.length}" name="Categories">${data.labels.map((label, index) => `<cx:pt idx="${index}"><cx:v>${escapeXml(label)}</cx:v></cx:pt>`).join('')}</cx:lvl>`;
        const sourceValues = hierarchy
            ? Array.from({ length: count }, (_, index) => { const result = book.getValue(`${colToName(valueCol)}${start + index + 1}`); return result.error || typeof result.value !== 'number' || !Number.isFinite(result.value) ? null : result.value; })
            : data.series[0].values;
        const subtotalIndexes = chart.kind === 'waterfall' ? waterfallTotalPointIndexes(chart, count) : [];
        if (subtotalIndexes.some(index => sourceValues[index] === null || sourceValues[index] === undefined))
            throw Error('Every Waterfall total point must have a numeric source value.');
        const numbers = sourceValues.map((value, index) => value === null ? '' : `<cx:pt idx="${index}"><cx:v>${value}</cx:v></cx:pt>`).join('');
        const valueFormula = colRange(sheetName, valueCol, start, end);
        const seriesName = hierarchy ? display(book.getValue(`${colToName(valueCol)}${bounds.r1 + 1}`)).trim() || 'Values' : data.series[0].name;
        const seriesShape = chartExSeriesShape(chart), dataLabels = chartExDataLabels(chart), subtotal = subtotalIndexes.length ? `<cx:layoutPr><cx:subtotals>${subtotalIndexes.map(index => `<cx:idx val="${index}"/>`).join('')}</cx:subtotals></cx:layoutPr>` : '', title = chart.title ? `<cx:title pos="t" align="ctr" overlay="0"><cx:tx><cx:rich><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US"/><a:t>${escapeXml(chart.title)}</a:t></a:r><a:endParaRPr lang="en-US"/></a:p></a:rich></cx:tx></cx:title>` : '';
        return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><cx:chartSpace xmlns:cx="${CX}" xmlns:a="${A}" version="1.2"><cx:chartData><cx:data id="0"><cx:strDim type="cat"><cx:f>${escapeXml(categoryFormula)}</cx:f>${categories}</cx:strDim><cx:numDim type="val"><cx:f>${escapeXml(valueFormula)}</cx:f><cx:lvl ptCount="${count}" name="${escapeXml(seriesName)}" formatCode="General">${numbers}</cx:lvl></cx:numDim></cx:data></cx:chartData><cx:chart>${title}<cx:plotArea><cx:plotAreaRegion><cx:series layoutId="${chartExLayout}" uniqueId="{${chartExLayout}-gridforge}">${seriesShape}${dataLabels}<cx:dataId val="0"/>${subtotal}</cx:series></cx:plotAreaRegion></cx:plotArea>${chartExLegend(chart)}</cx:chart></cx:chartSpace>`;
    }
    // The GridForge waterfall, Pareto, box-plot, funnel, treemap and sunburst previews use ChartEx above. Keep any remaining unsupported specialized chart data and placement editable by
    // exporting the selected source as an ordinary clustered column chart.
    if (chart.kind === 'stock') {
        try {
            const bounds = chartRange(chart.range, book.rows, book.cols), data = stockData(book, chart.range), headers = new Map();
            for (let col = bounds.c1; col <= bounds.c2; col++) {
                const text = String(book.getValue(`${colToName(col)}${bounds.r1 + 1}`).value ?? '').trim().toLowerCase();
                for (const field of ['open', 'high', 'low', 'close'])
                    if (text === field || text.startsWith(`${field} `) || text === `${field} price`)
                        headers.set(field, col);
            }
            const fields = data.series.map(item => item.name.toLowerCase());
            if (fields.some(field => !headers.has(field)))
                throw Error('Stock chart needs High, Low and Close columns, with optional Open.');
            const start = bounds.r1 + 1, end = bounds.r2, fieldColumns = new Set(headers.values()), labelCol = Array.from({ length: bounds.c2 - bounds.c1 + 1 }, (_, i) => bounds.c1 + i).find(col => !fieldColumns.has(col)), catFormula = labelCol === undefined ? undefined : colRange(sheetName, labelCol, start, end), stockSeries = fields.map((field, index) => {
                const col = headers.get(field), values = data.series.find(item => item.name.toLowerCase() === field).values;
                const headerRef = `'${sheetName.replace(/'/g, "''")}'!$${colToName(col)}$${bounds.r1 + 1}`;
                const valuesRef = colRange(sheetName, col, start, end);
                return `<c:ser><c:idx val="${index}"/><c:order val="${index}"/><c:tx><c:strRef><c:f>${escapeXml(headerRef)}</c:f><c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>${field[0].toUpperCase() + field.slice(1)}</c:v></c:pt></c:strCache></c:strRef></c:tx><c:cat>${labelsRef(catFormula, data.labels)}</c:cat><c:val>${numRef(valuesRef, values)}</c:val></c:ser>`;
            }).join('');
            chart = { ...chart, kind: 'stock', seriesOptions: undefined };
            const plot = `<c:stockChart>${stockSeries}<c:hiLowLines/><c:upDownBars><c:gapWidth val="150"/><c:upBars><c:spPr><a:solidFill><a:srgbClr val="70AD47"/></a:solidFill></c:spPr></c:upBars><c:downBars><c:spPr><a:solidFill><a:srgbClr val="ED7D31"/></a:solidFill></c:spPr></c:downBars></c:upDownBars><c:axId val="401"/><c:axId val="402"/></c:stockChart>`;
            return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><c:chartSpace xmlns:c="${C}" xmlns:a="${A}" xmlns:r="${R}"><c:date1904 val="0"/><c:lang val="en-US"/>${pivotSource}<c:chart>${titleNode(chart.title || 'Stock Chart')}<c:autoTitleDeleted val="0"/><c:plotArea><c:layout/>${plot}${axes(chart, false)}</c:plotArea>${legend(chart)}<c:plotVisOnly val="1"/><c:dispBlanksAs val="gap"/></c:chart></c:chartSpace>`;
        }
        catch {
            chart = { ...chart, kind: 'column', seriesOptions: undefined };
        }
    }
    const range = chartRange(chart.range, book.rows, book.cols), baseData = extractChartData(book, chart.range), data = extractChartData(book, chart.range, chart.switchRowsColumns ?? false);
    const start = range.r2 - baseData.rows + 1, firstNumber = range.c2 - baseData.series.length + 1, hasLabels = firstNumber > range.c1;
    const labelFormula = hasLabels ? (chart.switchRowsColumns ? rowRange(sheetName, range.r1, firstNumber, range.c2) : colRange(sheetName, range.c1, start, range.r2)) : undefined;
    const scatter = chart.kind === 'scatter', pie = chart.kind === 'pie', doughnut = chart.kind === 'doughnut', noAxes = pie || doughnut;
    const series = scatter && data.series.length > 1 ? data.series.slice(1) : noAxes ? data.series.slice(0, 1) : data.series.slice(0, 8);
    const palette = CHART_PALETTES[chart.theme ?? 'office'];
    const chunks = series.map((seriesItem, index) => {
        const column = firstNumber + (scatter ? index + 1 : index);
        const seriesRef = chart.switchRowsColumns ? rowRange(sheetName, start + index, firstNumber, range.c2) : colRange(sheetName, column, start, range.r2);
        const color = (chart.seriesOptions?.[index]?.color ?? palette[index % palette.length]).slice(1).toUpperCase();
        const nameRow = chart.switchRowsColumns ? start + index : range.r1, nameCol = chart.switchRowsColumns ? range.c1 : firstNumber + (scatter ? index + 1 : index), nameRef = `'${sheetName.replace(/'/g, "''")}'!$${colToName(nameCol)}$${nameRow + 1}`, nameCell = display(book.getValue(`${colToName(nameCol)}${nameRow + 1}`)), nameXml = nameCell === seriesItem.name ? `<c:tx><c:strRef><c:f>${escapeXml(nameRef)}</c:f><c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>${escapeXml(seriesItem.name)}</c:v></c:pt></c:strCache></c:strRef></c:tx>` : `<c:tx><c:v>${escapeXml(seriesItem.name)}</c:v></c:tx>`;
        const pointColors = pie || doughnut ? (chart.dataPointColors ?? []).map((pointColor, pointIndex) => `<c:dPt><c:idx val="${pointIndex}"/><c:spPr><a:solidFill><a:srgbClr val="${pointColor.slice(1).toUpperCase()}"/></a:solidFill></c:spPr></c:dPt>`).join('') : '';
        const shared = `<c:idx val="${index}"/><c:order val="${index}"/>${nameXml}<c:spPr><a:solidFill><a:srgbClr val="${color}"/></a:solidFill></c:spPr>${pointColors}`;
        const payload = scatter ? `<c:xVal>${numRef(colRange(sheetName, firstNumber, start, range.r2), data.series[0].values)}</c:xVal><c:yVal>${numRef(seriesRef, seriesItem.values)}</c:yVal>` : `<c:cat>${labelsRef(labelFormula, data.labels)}</c:cat><c:val>${numRef(seriesRef, seriesItem.values)}</c:val>`;
        const trend = chart.seriesOptions?.[index]?.trendline, trendXml = trend ? `<c:trendline><c:name>${escapeXml(`${trend.type} trendline`)}</c:name><c:trendlineType val="${trend.type === 'movingAverage' ? 'movingAvg' : trend.type}"/>${trend.type === 'polynomial' ? `<c:order val="${trend.order ?? 2}"/>` : ''}${trend.type === 'movingAverage' ? `<c:period val="${trend.period ?? 2}"/>` : ''}${trend.forward ? `<c:forward val="${trend.forward}"/>` : ''}${trend.backward ? `<c:backward val="${trend.backward}"/>` : ''}${trend.showEquation ? '<c:dispEq val="1"/>' : ''}${trend.showR2 ? '<c:dispRSqr val="1"/>' : ''}</c:trendline>` : '';
        const error = chart.seriesOptions?.[index]?.errorBars, errorXml = error ? `<c:errBars><c:errDir val="y"/><c:errBarType val="both"/><c:errValType val="${error.type === 'fixed' ? 'fixedVal' : error.type === 'percentage' ? 'percentage' : 'stdDev'}"/><c:val val="${error.value}"/><c:noEndCap val="0"/></c:errBars>` : '';
        const marker = chart.seriesOptions?.[index]?.marker === undefined ? '' : `<c:marker><c:symbol val="${chart.seriesOptions[index].marker ? 'circle' : 'none'}"/><c:size val="5"/></c:marker>`;
        return `<c:ser>${shared}${marker}${payload}${errorXml}${trendXml}</c:ser>`;
    }).join('');
    const axisIds = noAxes ? '' : '<c:axId val="401"/><c:axId val="402"/>';
    let plot;
    if (scatter)
        plot = `<c:scatterChart><c:scatterStyle val="lineMarker"/><c:varyColors val="0"/>${chunks}${axisIds}</c:scatterChart>`;
    else if (pie)
        plot = `<c:pieChart><c:varyColors val="1"/>${chunks}<c:firstSliceAng val="${chart.firstSliceAngle ?? 0}"/></c:pieChart>`;
    else if (doughnut)
        plot = `<c:doughnutChart><c:varyColors val="1"/>${chunks}<c:firstSliceAng val="${chart.firstSliceAngle ?? 0}"/><c:holeSize val="${chart.doughnutHoleSize ?? 57}"/></c:doughnutChart>`;
    else if (chart.kind === 'radar')
        plot = `<c:radarChart><c:radarStyle val="marker"/><c:varyColors val="0"/>${chunks}${axisIds}</c:radarChart>`;
    else if (chart.kind === 'combo' || chart.seriesOptions?.some(o => o.axis === 'secondary')) {
        const seriesChunks = chunks.match(/<c:ser>[\s\S]*?<\/c:ser>/g) ?? [], options = chart.seriesOptions ?? [], hasSecondary = options.some(o => o.axis === 'secondary');
        const group = (items, kind, axis) => kind === 'column' ? `<c:barChart><c:barDir val="col"/><c:grouping val="clustered"/>${items.join('')}<c:axId val="401"/><c:axId val="${axis === 'secondary' ? '404' : '402'}"/></c:barChart>` : `<c:lineChart><c:grouping val="standard"/><c:marker val="1"/>${items.join('')}<c:axId val="401"/><c:axId val="${axis === 'secondary' ? '404' : '402'}"/></c:lineChart>`;
        const groups = [];
        seriesChunks.forEach((item, index) => { const axis = options[index]?.axis === 'secondary' ? 'secondary' : 'primary', kind = options[index]?.kind ?? (chart.kind === 'combo' && !hasSecondary && index === seriesChunks.length - 1 ? 'line' : axis === 'secondary' ? 'line' : 'column'); let target = groups.find(candidate => candidate.axis === axis && candidate.kind === kind); if (!target) {
            target = { axis, kind, items: [] };
            groups.push(target);
        } target.items.push(item); });
        plot = groups.map(candidate => group(candidate.items, candidate.kind, candidate.axis)).join('');
    }
    else if (chart.kind === 'column' || chart.kind === 'bar' || chart.kind === 'stacked-column' || chart.kind === 'stacked-bar' || chart.kind === 'percent-column' || chart.kind === 'percent-bar')
        plot = `<c:barChart><c:barDir val="${chart.kind.includes('bar') ? 'bar' : 'col'}"/><c:grouping val="${chart.kind.startsWith('percent') ? 'percentStacked' : chart.kind.startsWith('stacked') ? 'stacked' : 'clustered'}"/><c:varyColors val="0"/>${chunks}${(chart.kind.startsWith('stacked') || chart.kind.startsWith('percent')) ? '<c:overlap val="100"/>' : ''}${axisIds}</c:barChart>`;
    else if (chart.kind === 'area')
        plot = `<c:areaChart><c:grouping val="standard"/>${chunks}${axisIds}</c:areaChart>`;
    else
        plot = `<c:lineChart><c:grouping val="standard"/><c:varyColors val="0"/>${chunks}<c:marker val="1"/>${axisIds}</c:lineChart>`;
    const labelPosition = { center: 'ctr', insideEnd: 'inEnd', outsideEnd: 'outEnd', bestFit: 'bestFit' }[chart.dataLabelPosition ?? 'outsideEnd'];
    const labels = chart.dataLabels || chart.dataLabelCategoryName || chart.dataLabelSeriesName || chart.dataLabelPercent ? `<c:dLbls><c:showVal val="${chart.dataLabels ? '1' : '0'}"/><c:showCatName val="${chart.dataLabelCategoryName ? '1' : '0'}"/><c:showSerName val="${chart.dataLabelSeriesName ? '1' : '0'}"/><c:showPercent val="${chart.dataLabelPercent && (chart.kind === 'pie' || chart.kind === 'doughnut') ? '1' : '0'}"/>${chart.dataLabelSeparator ? `<c:separator>${escapeXml(chart.dataLabelSeparator)}</c:separator>` : ""}<c:dLblPos val="${labelPosition}"/></c:dLbls>` : '';
    if (labels)
        plot = noAxes ? plot.replace('<c:firstSliceAng', `${labels}<c:firstSliceAng`) : plot.replace('<c:axId', `${labels}<c:axId`);
    const secondaryAxis = chart.seriesOptions?.some(o => o.axis === 'secondary') ? `<c:valAx><c:axId val="404"/>${axisBounds(chart.secondaryYAxisMin, chart.secondaryYAxisMax)}<c:axPos val="r"/>${chart.secondaryYAxisTitle ? titleNode(chart.secondaryYAxisTitle) : ''}${gridlinesFor(chart)}<c:numFmt formatCode="${escapeXml(numberFormat(chart.secondaryYAxisNumberFormat ?? chart.yAxisNumberFormat, chart.secondaryYAxisFormatCode ?? chart.yAxisFormatCode))}" sourceLinked="0"/>${chart.secondaryYAxisMajorUnit !== undefined ? `<c:majorUnit val="${chart.secondaryYAxisMajorUnit}"/>` : ''}<c:tickLblPos val="nextTo"/><c:crossAx val="401"/><c:crosses val="max"/><c:crossBetween val="between"/></c:valAx>` : '';
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><c:chartSpace xmlns:c="${C}" xmlns:a="${A}" xmlns:r="${R}"><c:date1904 val="0"/><c:lang val="en-US"/>${pivotSource}<c:chart>${titleNode(chart.title || 'Chart')}<c:autoTitleDeleted val="0"/><c:plotArea><c:layout/>${plot}${noAxes ? '' : axes(chart, scatter) + secondaryAxis}</c:plotArea>${legend(chart)}<c:plotVisOnly val="1"/><c:dispBlanksAs val="gap"/></c:chart></c:chartSpace>`;
}
export function nativeDrawingXml(charts, slicers = []) {
    const frames = charts.map((chart, index) => {
        const x = Math.round(chart.x * 9525), y = Math.round(chart.y * 9525), width = Math.round(chart.width * 9525), height = Math.round(chart.height * 9525);
        const chartEx = isNativeChartEx(chart.kind), chartNs = chartEx ? CX : C, chartPrefix = chartEx ? 'cx' : 'c';
        return `<xdr:absoluteAnchor><xdr:pos x="${x}" y="${y}"/><xdr:ext cx="${width}" cy="${height}"/><xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="${index + 2}" name="Chart ${index + 1}"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr><xdr:xfrm><a:off x="${x}" y="${y}"/><a:ext cx="${width}" cy="${height}"/></xdr:xfrm><a:graphic><a:graphicData uri="${chartNs}"><${chartPrefix}:chart xmlns:${chartPrefix}="${chartNs}" xmlns:r="${R}" r:id="rId${index + 1}"/></a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/></xdr:absoluteAnchor>`;
    }).join('') + slicers.map((slicer, index) => nativeSlicerAnchorXml(slicer, charts.length + index + 2)).join('');
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><xdr:wsDr xmlns:xdr="${D}" xmlns:a="${A}" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:sle="http://schemas.microsoft.com/office/drawing/2010/slicer">${frames}</xdr:wsDr>`;
}
export function nativeDrawingRelationships(charts, startChartIndex) {
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="${P}">${charts.map((chart, i) => `<Relationship Id="rId${i + 1}" Type="${isNativeChartEx(chart.kind) ? CHART_EX_REL : `${R}/chart`}" Target="../charts/chart${startChartIndex + i}.xml"/>`).join('')}</Relationships>`;
}
export const chartRelationshipType = `${R}/chart`;
export const drawingRelationshipType = `${R}/drawing`;
