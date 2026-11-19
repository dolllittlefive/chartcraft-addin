// F109: native stacked bars/columns + XY segments. No shapes, VBA or refresh handler.
import {UserError} from '../core/errors.js';

const FONT = 'SimSun';
const SERIES_COUNT = 12;
const COLORS = {wide: '#40A1ED', actual: '#0A1C2C', reference: '#404040', text: '#404040', baseline: '#DDDDDD', bands: ['#A6A6A6', '#BFBFBF', '#D9D9D9', '#EDEDED']};
const MAX_ROWS = 100;
const quote = name => `'${String(name).replace(/'/g, "''")}'`;
const literal = value => "'" + String(value);
function col(index) { let out = '', n = index + 1; while (n) { const m = (n - 1) % 26; out = String.fromCharCode(65 + m) + out; n = Math.floor((n - 1) / 26); } return out; }
const blank = x => x === '' || x === null || x === undefined;
const numeric = x => typeof x === 'number' && Number.isFinite(x) && x >= 0;

// Keep native failures visible, with the exact recipe stage for host diagnosis.
// A failed request must never be treated as successful chart creation.
async function syncStage(context, stage) {
  try { await context.sync(); }
  catch (error) { error.message = `子弹图[${stage}]：${error.message}`; throw error; }
}

function validate(values) {
  if (!Array.isArray(values) || values.length < 2 || !Array.isArray(values[0])) throw new UserError('请选择含表头和至少 1 行数据的区域。');
  const width = values[0].length, count = values.length - 1;
  if (width < 4 || width > 7) throw new UserError('需要 4–7 列：分类、窄条值、参考刻线（可空）、1–4 个递增背景上限。');
  if (count > MAX_ROWS) throw new UserError(`最多支持 ${MAX_ROWS} 行数据。`);
  if (values[0].some(x => typeof x !== 'string' || !x.trim() || x.length > 255)) throw new UserError('各列表头必须是 1–255 字符的非空文字。');
  values.slice(1).forEach((row, i) => {
    if (!Array.isArray(row) || row.length !== width) throw new UserError(`第 ${i + 2} 行列数不一致。`);
    if (typeof row[0] !== 'string' || !row[0].trim()) throw new UserError(`第 ${i + 2} 行分类必须是非空文字。`);
    for (let j = 1; j < width; j++) {
      if (j === 2 && blank(row[j])) continue;
      if (!numeric(row[j])) throw new UserError(`第 ${i + 2} 行第 ${j + 1} 列需要非负有限数字；仅参考刻线允许留空。`);
      if (j > 3 && row[j] < row[j - 1]) throw new UserError(`第 ${i + 2} 行的背景上限必须从小到大排列。`);
    }
  });
  return {count, rangeCount: width - 3, headers: values[0].map(String)};
}

function settings(options) {
  const o = {subtitle: '', footnote: '', direction: 'horizontal', scaleMode: 'shared', labelMode: 'both', summaryRows: 0, showScale: false, showReferenceLabels: false, decimals: 1, ...options};
  if (!['horizontal', 'vertical'].includes(o.direction) || !['shared', 'row'].includes(o.scaleMode) || !['both', 'actual', 'none'].includes(o.labelMode)) throw new UserError('子弹图选项无效。');
  if (!Number.isInteger(o.summaryRows) || o.summaryRows < 0 || o.summaryRows > MAX_ROWS - 1 || !Number.isInteger(o.decimals) || o.decimals < 0 || o.decimals > 3) throw new UserError('汇总行数或小数位数无效。');
  if (o.scaleMode === 'row' && o.showScale) throw new UserError('逐行尺度不能显示共用数值刻度；请关闭数值刻度。');
  return o;
}

function layout(count, gap, vertical) {
  const slots = count + gap;
  // At nine rows + a gap: 625×380 source image scaled to 500×304 pt.
  return vertical
    ? {width: Math.max(500, slots * 36 + 72), height: 360, plot: {left: 42, top: 74, width: Math.max(428, slots * 36), height: 210}}
    : {width: 500, height: Math.max(170, 92 + slots * 21.2), plot: {left: 137, top: 76, width: 350, height: slots * 21.2}};
}

async function writeHelper({context, helper, chart, source, options, previous}) {
  const o = settings(options), n = source.count, rangeCount = source.rangeCount;
  if (previous) {
    const schema = helper.getRange('B1'); schema.load('values'); chart.series.load('items'); await context.sync();
    if (schema.values[0][0] !== rangeCount) throw new UserError('背景区间列数发生变化。请新建实例；原图及其手工格式已保留。');
    if (chart.series.items.length !== SERIES_COUNT) throw new UserError('此图使用旧版子弹图结构或系列已被增删。请新建实例；原图、辅助数据及手工格式未修改。');
  }
  const summaryRows = Math.min(o.summaryRows, n - 1), gap = summaryRows > 0 ? 1 : 0;
  const slots = n + gap, last = n + 2, geometryLast = 3 * n + 1;
  const vertical = o.direction === 'vertical', xy = (value, lane) => vertical ? [lane, value] : [value, lane];
  // Clear only recipe cells below the metadata row. The separator rows must be
  // truly empty cells, never formula-empty strings or NA points between lines.
  helper.getRange(`A2:BA${Math.max(15, geometryLast, 3 * (previous?.count || 0) + 1)}`).clear(Excel.ClearApplyTo.contents);
  const agg = `AGGREGATE(4,6,$B$3:$G$${last})`;
  helper.getRange('B1:H1').formulas = [[String(rangeCount), `=IFERROR(IF(${agg}>0,${agg},1),1)`, String(slots), String(summaryRows), literal(o.direction), literal(o.scaleMode), String(o.decimals)]];
  const s = quote(source.sheet), ref = (j, i) => `${s}!$${col(source.columnIndex + j)}$${source.rowIndex + 2 + i}`;
  const clean = r => `=IFERROR(IF(AND(ISNUMBER(${r}),${r}>=0),${r},NA()),NA())`;
  const fmt = o.decimals ? '0.' + '0'.repeat(o.decimals) : '0';
  const header = ['分类', '窄条原值', '参考原值', '上限1', '上限2', '上限3', '上限4', '行尺度', '槽位', '位置', '窄条坐标', '参考坐标', '背景坐标', '窄条标签', '背景标签', '参考标签', '区间状态', '段1', '段2', '段3', '段4'];
  helper.getRange('A2:U2').formulas = [header.map(literal)];
  const rows = [];
  for (let i = 0; i < n; i++) {
    const r = i + 3, slot = i + 1 + (gap && i >= n - summaryRows ? 1 : 0);
    const bounds = Array.from({length: 4}, (_, j) => j < rangeCount ? clean(ref(j + 3, i)) : '');
    const rowAgg = `AGGREGATE(4,6,B${r}:G${r})`;
    const validRanges = ['ISNUMBER(D' + r + ')', ...Array.from({length: rangeCount - 1}, (_, j) => `${col(j + 4)}${r}>=${col(j + 3)}${r}`)].join(',');
    rows.push([
      `=IFERROR(${ref(0, i)}&"","")`, clean(ref(1, i)), `=IFERROR(IF(${ref(2, i)}="","",IF(AND(ISNUMBER(${ref(2, i)}),${ref(2, i)}>=0),${ref(2, i)},NA())),NA())`, ...bounds,
      o.scaleMode === 'row' ? `=IFERROR(IF(${rowAgg}>0,${rowAgg},1),1)` : '=$C$1', String(slot), vertical ? `=(I${r}-0.5)/$D$1` : `=($D$1-I${r}+0.5)/$D$1`,
      `=B${r}/H${r}`, `=IF(ISNUMBER(C${r}),C${r}/H${r},NA())`, `=IF(Q${r}="OK",${col(rangeCount + 2)}${r}/H${r},NA())`,
      `=IFERROR(TEXT(B${r},"${fmt}"),"")`, `=IFERROR(IF(Q${r}="OK",TEXT(${col(rangeCount + 2)}${r},"${fmt}"),""),"")`, `=IFERROR(IF(ISNUMBER(C${r}),TEXT(C${r},"${fmt}"),""),"")`,
      `=IFERROR(IF(AND(${validRanges}),"OK","INVALID_RANGE"),"INVALID_RANGE")`,
      ...Array.from({length: 4}, (_, j) => j < rangeCount ? `=IF(Q${r}="OK",(${col(j + 3)}${r}-${j ? col(j + 2) + r : '0'})/H${r},NA())` : '0'),
    ]);
  }
  helper.getRange(`A3:U${last}`).formulas = rows;
  const bars = [];
  for (let i = 0; i < n; i++) {
    if (gap && i === n - summaryRows) bars.push(['', '0', '0', '0', '0']);
    const r = i + 3; bars.push([`=A${r}`, ...['R', 'S', 'T', 'U'].map(c => `=${c}${r}`)]);
    // Two-point segments in each series, with an unwritten row between each pair.
    const g = 3 + i * 3;
    helper.getRange(`AC${g}:AD${g + 1}`).formulas = [xy('0', `=J${r}`), xy(`=K${r}`, `=J${r}`)];
    helper.getRange(`AF${g}:AG${g + 1}`).formulas = [xy(`=L${r}`, `=J${r}-0.40/$D$1`), xy(`=L${r}`, `=J${r}+0.40/$D$1`)];
    helper.getRange(`AI${r}:AJ${r}`).formulas = [xy(`=K${r}`, `=J${r}`)];
    helper.getRange(`AL${r}:AM${r}`).formulas = [xy(`=M${r}`, `=J${r}`)];
    helper.getRange(`AO${r}:AP${r}`).formulas = [xy(`=L${r}`, `=J${r}`)];
  }
  // Native horizontal bars put the first category at the bottom by default.
  // Reverse only their helper rows; XY points keep their source order and J
  // positions, so the visual order is source order without reversing any axis.
  helper.getRange(`W3:AA${slots + 2}`).formulas = vertical ? bars : bars.reverse();
  helper.getRange('AR3:AT7').formulas = Array.from({length: 5}, (_, i) => [...xy(String(i / 4), '0'), `=TEXT($C$1*${i / 4},"${fmt}")`]);
  helper.getRange('AV3:AW4').formulas = [xy('0', '0'), xy('0', '1')];
  // Fixed native label points: six legend slots (key + linked header) and a
  // footnote. Their point identities do not change when data rows are rebound.
  const legendColumns = [...Array.from({length: rangeCount}, (_, i) => i + 3), 1, 2];
  const headerRef = j => `${s}!$${col(source.columnIndex + j)}$${source.rowIndex + 1}`;
  helper.getRange('AY3:BA15').formulas = Array.from({length: 13}, (_, i) => {
    const entry = Math.floor(i / 2), column = legendColumns[entry];
    const text = i === 12 ? literal(o.footnote) : column === undefined ? '' : i % 2 ? `=IFERROR(${headerRef(column)}&"","")` : literal(column === 2 ? '│' : '■');
    return ['0', '0', text];
  });
  return {count: n, last, slots, geometryLast, rangeCount, options: o, layout: layout(n, gap, vertical)};
}

function bindSeries(chart, helper, b) {
  const s = i => chart.series.getItemAt(i), range = c => helper.getRange(`${c}3:${c}${b.last}`);
  for (let i = 0; i < 4; i++) { s(i).setXAxisValues(helper.getRange(`W3:W${b.slots + 2}`)); s(i).setValues(helper.getRange(`${col(23 + i)}3:${col(23 + i)}${b.slots + 2}`)); }
  for (const [index, x, y, end] of [[4,'AC','AD',b.geometryLast],[5,'AF','AG',b.geometryLast]]) { s(index).setXAxisValues(helper.getRange(`${x}3:${x}${end}`)); s(index).setValues(helper.getRange(`${y}3:${y}${end}`)); }
  for (const [index, x, y] of [[6,'AI','AJ'],[7,'AL','AM'],[8,'AO','AP']]) { s(index).setXAxisValues(range(x)); s(index).setValues(range(y)); }
  s(9).setXAxisValues(helper.getRange('AR3:AR7')); s(9).setValues(helper.getRange('AS3:AS7'));
  s(10).setXAxisValues(helper.getRange('AV3:AV4')); s(10).setValues(helper.getRange('AW3:AW4'));
  s(11).setXAxisValues(helper.getRange('AY3:AY15')); s(11).setValues(helper.getRange('AZ3:AZ15'));
}

async function createChart({context, sourceSheet, helper, source, instance, binding}) {
  const vertical = binding.options.direction === 'vertical';
  const chart = sourceSheet.charts.add(vertical ? Excel.ChartType.columnStacked : Excel.ChartType.barStacked, helper.getRange(`X3:X${binding.slots + 2}`), Excel.ChartSeriesBy.columns);
  chart.name = instance.chartName;
  // A values-only one-column seed has exactly one series; all other series are explicit.
  chart.series.getItemAt(0).name = source.headers[3];
  for (let i = 1; i < SERIES_COUNT; i++) chart.series.add(i < 4 ? source.headers[i + 3] || `未使用区间${i + 1}` : i === 4 ? source.headers[1] : i === 5 ? source.headers[2] : i === 11 ? '图例与脚注' : `辅助${i}`);
  bindSeries(chart, helper, binding);
  const anchor = sourceSheet.getRangeByIndexes(source.rowIndex, source.columnIndex + source.columnCount, 1, 1); anchor.load('left,top'); await context.sync();
  chart.left = anchor.left + 12; chart.top = anchor.top; chart.width = binding.layout.width; chart.height = binding.layout.height;
  chart.plotVisibleOnly = false; chart.displayBlanksAs = 'NotPlotted';
  return chart;
}

function linkLabels(chart, helperName, b, previousCount = 0, initial = false) {
  const o = b.options, vertical = o.direction === 'vertical';
  const configs = [[6, 'N', o.labelMode !== 'none'], [7, 'O', o.labelMode === 'both'], [8, 'P', o.showReferenceLabels]];
  for (const [index, c, enabled] of configs) {
    for (let i = 0; i < b.count; i++) {
      const point = chart.series.getItemAt(index).points.getItemAt(i);
      if (initial || i >= previousCount) {
        point.hasDataLabel = Boolean(enabled);
        if (enabled) {
          point.dataLabel.position = index === 8 ? (vertical ? 'Right' : 'Top') : (vertical ? 'Top' : 'Right');
          point.dataLabel.format.font.name = '微软雅黑'; point.dataLabel.format.font.size = 10;
          point.dataLabel.format.font.color = '#000000';
        }
      }
      // No hasDataLabel/position/font assignment to existing points during rebind.
      if (enabled) point.dataLabel.formula = `=${quote(helperName)}!$${c}$${i + 3}`;
    }
  }
  if (initial && o.showScale) for (let i = 0; i < 5; i++) {
    const point = chart.series.getItemAt(9).points.getItemAt(i); point.hasDataLabel = true;
    point.dataLabel.formula = `=${quote(helperName)}!$AT$${i + 3}`;
    point.dataLabel.position = vertical ? 'Left' : 'Bottom';
    point.dataLabel.format.font.name = FONT; point.dataLabel.format.font.size = 8;
  }
  const entryCount = b.rangeCount + 2;
  for (let i = 0; i < 13; i++) {
    const enabled = i === 12 ? Boolean(o.footnote) : i < entryCount * 2;
    const point = chart.series.getItemAt(11).points.getItemAt(i);
    if (initial) {
      point.hasDataLabel = enabled;
      if (enabled) {
        const entry = Math.floor(i / 2), key = i % 2 === 0;
        const font = point.dataLabel.format.font;
        font.name = FONT; font.size = i === 12 ? 6 : key ? (entry === entryCount - 1 ? 10 : 7) : 8;
        font.color = i === 12 ? '#999999' : !key ? COLORS.text : entry < b.rangeCount ? (b.rangeCount === 1 ? COLORS.wide : COLORS.bands[entry]) : entry === b.rangeCount ? COLORS.actual : COLORS.reference;
        point.dataLabel.position = 'Right';
      }
    }
    if (enabled) point.dataLabel.formula = `=${quote(helperName)}!$BA$${i + 3}`;
  }
}

function placeAnnotations(chart, b) {
  const count = b.rangeCount + 2, start = (b.layout.width - (count * 44 - 8)) / 2;
  for (let i = 0; i < count * 2; i++) {
    const label = chart.series.getItemAt(11).points.getItemAt(i).dataLabel;
    label.left = start + Math.floor(i / 2) * 44 + (i % 2 ? 8 : 0);
    label.top = i % 2 === 0 && i / 2 === count - 1 ? 53 : 55;
  }
  if (b.options.footnote) {
    const label = chart.series.getItemAt(11).points.getItemAt(12).dataLabel;
    label.left = 6; label.top = b.layout.height - 13;
  }
}

async function clearReferenceFromActualLabels(context, chart, helper, b, x, y) {
  if (b.options.direction !== 'horizontal' || b.options.labelMode === 'none') return {moved: 0, visualVerified: false};
  const actual = chart.series.getItemAt(6), coordinates = helper.getRange(`J3:L${b.last}`), labels = [];
  // linkLabels has already applied YaHei 10pt and synced. Use native text size,
  // but derive its anchor from the same formula coordinates as the actual bar.
  // Windows may return transient automatic Right positions after plot layout;
  // using those positions for collision detection can select unrelated rows.
  coordinates.load('values');
  for (let index = 0; index < b.count; index++) {
    const label = actual.points.getItemAt(index).dataLabel;
    label.load('left,top,width,height'); labels.push(label);
  }
  await syncStage(context, '数值标签与参考刻线边界');
  const p = chart.plotArea, moves = [], positions = [];
  for (let index = 0; index < b.count; index++) {
    const [lane, value, reference] = coordinates.values[index], label = labels[index];
    if (![lane, value, label.width, label.height, p.insideLeft, p.insideTop, p.insideWidth, p.insideHeight, x.minimum, x.maximum, y.minimum, y.maximum].every(Number.isFinite) || label.width <= 0 || label.height <= 0 || x.maximum <= x.minimum || y.maximum <= y.minimum)
      throw new UserError(`子弹图[标签留白]：第 ${index + 1} 行未返回可用原生标签/刻线边界。`);
    const actualX = p.insideLeft + (value - x.minimum) / (x.maximum - x.minimum) * p.insideWidth;
    const rowY = p.insideTop + (y.maximum - lane) / (y.maximum - y.minimum) * p.insideHeight;
    const position = {index, fromLeft: label.left, fromTop: label.top, left: actualX + 3, top: rowY - label.height / 2, width: label.width, height: label.height, actualX, rowY};
    // A missing reference is the existing #N/A coordinate. It does not prevent
    // positioning its formula-backed actual-value label on the correct row.
    if (Number.isFinite(reference)) {
      const markerX = p.insideLeft + (reference - x.minimum) / (x.maximum - x.minimum) * p.insideWidth;
      position.referenceX = markerX;
      // The reference stroke is 0.8pt; both it and the label share rowY.
      if (position.left <= markerX + 0.4 && position.left + label.width >= markerX - 0.4) {
        position.left = markerX + 3; moves.push(position);
      }
    }
    // Set both coordinates: assigning left alone can retain an automatic Y
    // anchor while the host converts the label to a custom native position.
    label.left = position.left; label.top = position.top; positions.push(position);
  }
  if (positions.length) {
    actual.showLeaderLines = false;
    await syncStage(context, '数值标签避让参考刻线');
    for (const {index} of positions) labels[index].load('left,top,width,height');
    actual.load('showLeaderLines');
    await syncStage(context, '数值标签避让回读');
    const failed = positions.filter(position => ['left', 'top', 'width', 'height'].some(key => !Number.isFinite(labels[position.index][key]) || Math.abs(labels[position.index][key] - position[key]) > 0.25));
    if (actual.showLeaderLines !== false || failed.length) {
      const detail = {showLeaderLines: actual.showLeaderLines, failed: failed.map(position => ({row: position.index + 1, requested: position, actual: Object.fromEntries(['left', 'top', 'width', 'height'].map(key => [key, labels[position.index][key]]))}))};
      throw new UserError(`子弹图[标签留白]：宿主未保留同一行向右避让或关闭引导线，未认定布局通过。${JSON.stringify(detail)}`);
    }
  }
  return {moved: moves.length, moves, positioned: positions.length, positions, visualVerified: false};
}

async function style({context, chart, helper, instance, options, binding}) {
  const b = binding, o = b.options, vertical = o.direction === 'vertical';
  chart.format.fill.setSolidColor('#FFFFFF'); chart.format.border.lineStyle = 'None';
  chart.format.font.name = FONT; chart.format.font.size = 8; chart.format.font.color = COLORS.text;
  chart.title.visible = true; chart.title.text = options.title + (o.subtitle ? '\n' + o.subtitle : '');
  // ChartTitle.width is read-only. R4's 15pt title wrapped after 25 of 27
  // characters in the host's automatic title box; 13.5pt leaves fit margin.
  chart.title.format.font.name = FONT; chart.title.format.font.size = 14; chart.title.format.font.bold = false; chart.title.horizontalAlignment = 'Left';
  // Native data labels give the source's square / square / vertical-tick keys.
  // Excel's automatic scatter-series legend always substitutes horizontal lines.
  chart.legend.visible = false;
  const thickness = (vertical ? b.layout.plot.width : b.layout.plot.height) / b.slots * 0.26;
  for (let i = 0; i < SERIES_COUNT; i++) {
    const s = chart.series.getItemAt(i); s.hasDataLabels = false;
    if (i < 4) {
      s.chartType = vertical ? Excel.ChartType.columnStacked : Excel.ChartType.barStacked; s.axisGroup = 'Primary';
      s.gapWidth = 62; s.overlap = 100; s.format.fill.setSolidColor(b.rangeCount === 1 ? COLORS.wide : COLORS.bands[i]); s.format.line.lineStyle = 'None';
    } else {
      s.chartType = i === 4 || i === 5 || i === 10 ? Excel.ChartType.xyscatterLinesNoMarkers : Excel.ChartType.xyscatter;
      s.axisGroup = 'Secondary'; s.markerStyle = Excel.ChartMarkerStyle.none;
      s.format.line.lineStyle = i === 4 || i === 5 || i === 10 ? 'Continuous' : 'None';
      if (i === 4) { s.format.line.color = COLORS.actual; s.format.line.weight = thickness; }
      if (i === 5) { s.format.line.color = COLORS.reference; s.format.line.weight = 0.8; }
      if (i === 10) { s.format.line.color = COLORS.baseline; s.format.line.weight = 0.5; }
    }
  }
  await syncStage(context, '系列与坐标轴分组');
  if (o.subtitle) chart.title.getSubstring(options.title.length + 1, o.subtitle.length).font.size = 9;
  const category = chart.axes.getItem('Category', 'Primary'), value = chart.axes.getItem('Value', 'Primary');
  category.isBetweenCategories = true; category.tickLabelPosition = 'Low';
  category.textOrientation = vertical ? 45 : 0; category.format.font.name = FONT; category.format.font.size = 8; category.majorTickMark = 'None'; category.minorTickMark = 'None'; category.format.line.lineStyle = 'None';
  await syncStage(context, '主分类轴');
  value.minimum = 0; value.maximum = 1.1; value.majorUnit = 0.25; value.tickLabelPosition = 'None'; value.majorTickMark = 'None'; value.minorTickMark = 'None'; value.format.line.lineStyle = 'None'; value.majorGridlines.visible = false; value.minorGridlines.visible = false;
  await syncStage(context, '主数值轴');
  const x = chart.axes.getItem('Category', 'Secondary'), y = chart.axes.getItem('Value', 'Secondary');
  const secondaryAxes = [];
  for (const [axis, maximum, stage] of [[x, vertical ? 1 : 1.1, '次散点X轴'], [y, vertical ? 1.1 : 1, '次散点Y轴']]) {
    axis.minimum = 0; axis.maximum = maximum;
    await syncStage(context, stage + '范围');
    // The Mac interrupted workbook already has a hidden secondary X axis.
    // Its line-format object rejects lineStyle. Hide the native axis directly;
    // never touch secondary gridlines (Excel exposes them only on primary axes).
    axis.visible = false;
    await syncStage(context, stage + '隐藏');
    axis.load('minimum,maximum,visible');
    await syncStage(context, stage + '回读');
    if (axis.visible !== false || !Number.isFinite(axis.minimum) || !Number.isFinite(axis.maximum) || Math.abs(axis.minimum) > 1e-9 || Math.abs(axis.maximum - maximum) > 1e-9) {
      throw new UserError(`子弹图[${stage}回读]：隐藏次轴后未保留所需范围 0–${maximum}，原生图未验收。`);
    }
    secondaryAxes.push({axis: stage, minimum: axis.minimum, maximum: axis.maximum, visible: axis.visible});
  }
  linkLabels(chart, instance.helperName, b, 0, true);
  await syncStage(context, '图内图例与脚注标签');
  chart.title.left = 6; chart.title.top = 2;
  const p = chart.plotArea, wanted = b.layout.plot;
  for (let k = 0; k < 2; k++) { p.insideLeft = wanted.left; p.insideTop = wanted.top; p.insideWidth = wanted.width; p.insideHeight = wanted.height; await context.sync(); }
  placeAnnotations(chart, b);
  const annotations = chart.series.getItemAt(11);
  // Moving labels away from their (0,0) anchors enables native leader lines by
  // default. Disable at series level after positioning, without line-format IO.
  annotations.showLeaderLines = false;
  await syncStage(context, '图内图例与脚注定位');
  annotations.load('showLeaderLines'); chart.title.load('width,height');
  p.load('insideLeft,insideTop,insideWidth,insideHeight'); await syncStage(context, '标题与引导线回读');
  if (annotations.showLeaderLines !== false) throw new UserError('子弹图[图内图例引导线]：宿主未保留关闭状态，原生图未验收。');
  const labelClearance = await clearReferenceFromActualLabels(context, chart, helper, b, x, y);
  return {notes: ['原生堆积条/柱与散点线组合；显示数据标签来自辅助公式。源图参考刻线缺数时保留空白。', ...(o.scaleMode === 'row' ? ['逐行尺度：不同行的长度不可直接比较。'] : [])], secondaryAxes, annotations: {showLeaderLines: annotations.showLeaderLines}, title: {fontSize: 13.5, width: chart.title.width, height: chart.title.height}, plot: {requested: wanted, actual: {left: p.insideLeft, top: p.insideTop, width: p.insideWidth, height: p.insideHeight}}, labelClearance};
}

function rebind({chart, helper, source, instance, previous, binding}) {
  bindSeries(chart, helper, binding);
  // Legend names are data. Rebinding a reordered source must not retain the
  // old series meanings; the existing series objects and their formats remain.
  for (let i = 0; i < binding.rangeCount; i++) chart.series.getItemAt(i).name = source.headers[i + 3];
  chart.series.getItemAt(4).name = source.headers[1];
  chart.series.getItemAt(5).name = source.headers[2];
  linkLabels(chart, instance.helperName, binding, previous.count, false);
  return {notes: ['仅更新数据范围与链接标签；图表、轴、系列、字体和已有标签格式未重设。']};
}

const inputSpec = {
  category: {id: 'comparison', title: '比较'}, minVersion: '1.9', columns: 4, minRows: 1, maxRows: MAX_ROWS,
  description: '选择含表头的 4–7 列：分类、窄条值、参考刻线（可空）、1–4 个递增背景上限。原例按“类别、前2年、前10年、疫情后”排列；百分比按 6.0 表示 6.0%。',
  options: [
    {id: 'title', label: '标题', type: 'text', default: '子弹图', maxLength: 80},
    {id: 'subtitle', label: '副标题', type: 'text', default: '', maxLength: 120},
    {id: 'footnote', label: '脚注', type: 'text', default: '', maxLength: 200},
    {id: 'direction', label: '方向', type: 'select', default: 'horizontal', choices: [{value: 'horizontal', label: '横向'}, {value: 'vertical', label: '纵向'}]},
    {id: 'scaleMode', label: '数值尺度', type: 'select', default: 'shared', choices: [{value: 'shared', label: '统一尺度'}, {value: 'row', label: '逐行尺度（长度不可跨行比较）'}]},
    {id: 'labelMode', label: '数值标签', type: 'select', default: 'both', choices: [{value: 'both', label: '窄条与背景'}, {value: 'actual', label: '仅窄条'}, {value: 'none', label: '不显示'}]},
    {id: 'showReferenceLabels', label: '显示参考刻线数值', type: 'checkbox', default: false},
    {id: 'showScale', label: '显示数值刻度', type: 'checkbox', default: false},
    {id: 'summaryRows', label: '末尾汇总行数（之前留一行间距）', type: 'number', default: 0, min: 0, max: 99, step: 1},
    {id: 'decimals', label: '标签小数位', type: 'number', default: 1, min: 0, max: 3, step: 1},
  ],
  storage: {marker: 'ChartCraft.Bullet.v1', helperPrefix: 'CC_子弹_', chartPrefix: 'CC_Bullet_'},
};

export default {id: 'bullet', title: '子弹图', inputSpec, validate, writeHelper, createChart, style, rebind};
