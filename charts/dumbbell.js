// 景歌“滑珠/哑铃图”图族。三种原生结构，用“样式”切换：
//   horizontal  水平哑铃：分类自上而下，两组标记点 + 灰色连线（散点）
//   vertical    纵向滑珠：分类自左向右，两组标记点 + 连线，可选差值/涨幅标签（散点）
//   progress    进度滑珠柱：完成部分 + 未完成底 + 顶端滑珠标签（堆积柱形 + 带标记折线）
// 连线用一个 XYScatterLines 系列，每个分类三点 (A端, B端, #N/A)，#N/A 断开分类之间的线段；
// 不用误差线、形状或图片。全部定位与标签文本都由辅助表公式算出，源数据改动后无需插件即可联动。
import {UserError} from '../core/errors.js';

const MARK = 'ChartCraft.Dumbbell.v1';
const HELPER_PREFIX = 'CC_哑铃_';
const CHART_PREFIX = 'CC_Dumbbell_';
const MAX_ROWS = 40;
const FONT = '微软雅黑';
const COLORS = {a: '#E5475A', b: '#2D4E9E', link: '#BFBFBF', text: '#595959', grid: '#EEEEEE', positive: '#404040', negative: '#E02020', done: '#D47272', rest: '#F5DEDE'};
const VARIANTS = ['auto', 'horizontal', 'vertical', 'progress'];
// 每个分类的序列：0 连线，1 A 端，2 B 端，3 正变化锚，4 负变化锚，5 分类锚，6 上/A 值锚，7 下/B 值锚。
const PAIR_SERIES = 8;
const PROGRESS_SERIES = 3;

const quote = name => `'${String(name).replace(/'/g, "''")}'`;
const literal = text => `"${String(text).replace(/"/g, '""')}"`;
function columnLetter(index) {
  let result = '', n = index + 1;
  while (n > 0) { const m = (n - 1) % 26; result = String.fromCharCode(65 + m) + result; n = Math.floor((n - 1) / 26); }
  return result;
}

function validate(values) {
  if (!Array.isArray(values) || !Array.isArray(values[0]) || ![2, 3].includes(values[0].length))
    throw new UserError('请选择含表头的两列（分类、完成率）或三列（分类、系列A、系列B）。');
  const width = values[0].length;
  if (values.length < 3) throw new UserError('至少需要 2 行数据（不含表头）。');
  if (values.length - 1 > MAX_ROWS) throw new UserError(`最多支持 ${MAX_ROWS} 行数据。`);
  values.slice(1).forEach((row, i) => {
    if (!Array.isArray(row) || row.length !== width) throw new UserError(`第 ${i + 2} 行列数不一致。`);
    if (row[0] == null || String(row[0]).trim() === '') throw new UserError(`第 ${i + 2} 行分类为空。`);
    for (let j = 1; j < width; j++)
      if (typeof row[j] !== 'number' || !Number.isFinite(row[j])) throw new UserError(`第 ${i + 2} 行第 ${j + 1} 列必须是数字。`);
  });
  return {
    count: values.length - 1, inputKind: width === 2 ? 'progress' : 'pair',
    categoryHeader: String(values[0][0]), valueHeader: String(values[0][1]), secondHeader: width === 3 ? String(values[0][2]) : ''
  };
}

function resolvedVariant(source, options = {}) {
  const requested = options.variant ?? 'auto';
  if (!VARIANTS.includes(requested)) throw new UserError('请选择有效的滑珠/哑铃样式。');
  const columns = source.columnCount ?? (source.inputKind === 'progress' ? 2 : 3);
  if (requested === 'progress' && columns !== 2) throw new UserError('进度滑珠柱需要两列：分类、完成率（小数，86% 输入为 0.86）。');
  if ((requested === 'horizontal' || requested === 'vertical') && columns !== 3) throw new UserError('哑铃/滑珠对比需要三列：分类、系列A、系列B。两列数据请选“进度滑珠柱”。');
  if (requested !== 'auto') return requested;
  return columns === 2 ? 'progress' : 'horizontal';
}

function numberMask(options = {}) {
  const d = Math.max(0, Math.min(3, Math.round(Number(options.decimals ?? 0)) || 0));
  const tail = d ? '.' + '0'.repeat(d) : '';
  return options.valueFormat === 'percent' ? `0${tail}%` : `#,##0${tail}`;
}

// 纯函数：返回需要写入辅助表的公式块，不访问 Office。单测用 formula-evaluator 逐格求值。
function helperRecipe(source, options = {}) {
  const variant = resolvedVariant(source, options);
  const n = source.count, last = n + 2, sheet = quote(source.sheet);
  const col = k => columnLetter(source.columnIndex + k);
  const src = (k, i) => `${sheet}!$${col(k)}$${source.rowIndex + i + 2}`;
  const mask = numberMask(options);
  if (variant === 'progress') {
    const rows = [];
    for (let i = 0; i < n; i++) {
      const r = i + 3;
      rows.push([
        `=${src(0, i)}&""`,
        `=IF(ISNUMBER(${src(1, i)}),${src(1, i)},NA())`,
        `=IF(ISNUMBER(B${r}),MEDIAN(0,B${r},1),NA())`,
        `=IF(ISNUMBER(B${r}),1-C${r},NA())`,
        `=B${r}`,
        `=IF(ISNUMBER(B${r}),TEXT(B${r},${literal(mask)}),"")`
      ]);
    }
    return {variant, n, last, mask, seriesCount: PROGRESS_SERIES, blocks: [
      {address: 'A2:F2', formulas: [['分类', '完成率', '完成部分', '未完成底', '滑珠', '滑珠标签']]},
      {address: `A3:F${last}`, formulas: rows}
    ]};
  }
  const horizontal = variant === 'horizontal', change = ['diff', 'growth'].includes(options.changeLabel) ? options.changeLabel : 'none';
  const suffix = String(options.changeSuffix ?? '');
  const both = r => `AND(ISNUMBER(B${r}),ISNUMBER(C${r}))`;
  const rows = [], links = [];
  for (let i = 0; i < n; i++) {
    const r = i + 3;
    const changeValue = change === 'diff' ? `=IF(${both(r)},C${r}-B${r},NA())`
      : change === 'growth' ? `=IF(${both(r)},IF(B${r}=0,NA(),(C${r}-B${r})/ABS(B${r})),NA())` : '=NA()';
    const changeText = change === 'diff' ? `=IF(ISNUMBER(E${r}),IF(E${r}>0,"+","")&TEXT(E${r},${literal(mask)})&${literal(suffix)},"")`
      : change === 'growth' ? `=IF(ISNUMBER(E${r}),IF(E${r}>=0,"▲","▼")&TEXT(ABS(E${r}),"0.0%")&${literal(suffix)},"")` : '=""';
    const hi = horizontal ? `=B${r}` : `=IF(ISNUMBER(B${r}),IF(ISNUMBER(C${r}),MAX(B${r},C${r}),B${r}),IF(ISNUMBER(C${r}),C${r},NA()))`;
    const lo = horizontal ? `=C${r}` : `=IF(${both(r)},MIN(B${r},C${r}),NA())`;
    rows.push([
      `=${src(0, i)}&""`,
      `=IF(ISNUMBER(${src(1, i)}),${src(1, i)},NA())`,
      `=IF(ISNUMBER(${src(2, i)}),${src(2, i)},NA())`,
      String(horizontal ? n - i : i + 1),
      changeValue,
      `=IF(${both(r)},(B${r}+C${r})/2,NA())`,
      `=IF(ISNUMBER(E${r}),IF(E${r}>=0,F${r},NA()),NA())`,
      `=IF(ISNUMBER(E${r}),IF(E${r}<0,F${r},NA()),NA())`,
      changeText,
      '=$B$1',
      hi, lo,
      `=IF(ISNUMBER(K${r}),TEXT(K${r},${literal(mask)}),"")`,
      `=IF(ISNUMBER(L${r}),TEXT(L${r},${literal(mask)}),"")`
    ]);
    // Value axis coordinate first in horizontal (x=value, y=position), swapped in vertical.
    for (const v of [`=B${r}`, `=C${r}`]) links.push(horizontal ? [v, `=D${r}`] : [`=D${r}`, v]);
    links.push(['=NA()', '=NA()']);
  }
  const values = `B3:C${last}`;
  const lo = `AGGREGATE(5,6,${values})`, hi = `AGGREGATE(4,6,${values})`;
  const base = options.startAtZero === false ? lo : `MIN(0,${lo})`;
  // 水平：分类锚就在数值轴起点，标签向左伸进预留的绘图区左边距；纵向：锚点放在最低值下方 12% 跨度处。
  const floor = horizontal ? `=IFERROR(${base},0)` : `=IFERROR(${base}-IF(${hi}>${base},${hi}-${base},MAX(ABS(${hi}),1))*0.12,0)`;
  return {variant, horizontal, change, n, last, mask, linkLast: 3 * n + 2, seriesCount: PAIR_SERIES, blocks: [
    {address: 'B1', formulas: [[floor]]},
    {address: 'A2:N2', formulas: [['分类', '系列A', '系列B', '位置', '变化值', '中点', '正变化锚', '负变化锚', '变化标签', '分类锚', '上端/A值', '下端/B值', '上端标签', '下端标签']]},
    {address: `A3:N${last}`, formulas: rows},
    {address: 'R2:S2', formulas: [[horizontal ? '连线X' : '连线X', '连线Y']]},
    {address: `R3:S${3 * n + 2}`, formulas: links}
  ]};
}

// 每个系列的 [X 列, Y 列]；水平时数值在 X，纵向时数值在 Y。连线系列单独用 R:S。
function pairColumns(horizontal) {
  const value = ['B', 'C', 'G', 'H', 'J', 'K', 'L'];
  return value.map(v => horizontal ? [v, 'D'] : ['D', v]);
}

async function readStoredVariant(context, helper) {
  const cell = helper.getRange('N1'); cell.load('values'); await context.sync();
  return String(cell.values[0][0] ?? '');
}

async function writeHelper(args) {
  const {context, helper, source, options, previous} = args;
  const binding = helperRecipe(source, options);
  if (previous) {
    const stored = await readStoredVariant(context, helper);
    if (stored && stored !== binding.variant)
      throw new UserError(`该图是“${label(stored)}”结构，不能改成“${label(binding.variant)}”。请保留原图，用新选区另建。`);
    if (previous.count > source.count) {
      helper.getRange(`A${source.count + 3}:N${previous.count + 2}`).clear(Excel.ClearApplyTo.contents);
      if (binding.variant !== 'progress') helper.getRange(`R${3 * source.count + 3}:S${3 * previous.count + 2}`).clear(Excel.ClearApplyTo.contents);
    }
  }
  helper.getRange('N1').values = [[binding.variant]];
  for (const block of binding.blocks) helper.getRange(block.address).formulas = block.formulas;
  return binding;
}

function label(variant) {
  return {horizontal: '水平哑铃', vertical: '纵向滑珠', progress: '进度滑珠柱'}[variant] ?? variant;
}

async function syncStage(context, stage) {
  try { await context.sync(); }
  catch (error) { error.message = `滑珠/哑铃 [${stage}]：${error.message}`; throw error; }
}

function bindSeries(chart, helper, binding) {
  const last = binding.last;
  if (binding.variant === 'progress') {
    const categories = helper.getRange(`A3:A${last}`);
    ['C', 'D', 'E'].forEach((c, i) => {
      const series = chart.series.getItemAt(i);
      series.setValues(helper.getRange(`${c}3:${c}${last}`));
      series.setXAxisValues(categories);
    });
    return;
  }
  const link = chart.series.getItemAt(0);
  link.setXAxisValues(helper.getRange(`R3:R${binding.linkLast}`));
  link.setValues(helper.getRange(`S3:S${binding.linkLast}`));
  pairColumns(binding.horizontal).forEach(([x, y], i) => {
    const series = chart.series.getItemAt(i + 1);
    series.setXAxisValues(helper.getRange(`${x}3:${x}${last}`));
    series.setValues(helper.getRange(`${y}3:${y}${last}`));
  });
}

async function createChart(args) {
  const {context, helper, sourceSheet, source, instance, binding} = args;
  const progress = binding.variant === 'progress';
  const chart = progress
    ? sourceSheet.charts.add(Excel.ChartType.columnStacked, helper.getRange(`C3:C${binding.last}`), Excel.ChartSeriesBy.columns)
    : sourceSheet.charts.add(Excel.ChartType.xyscatter, helper.getRange(`S3:S${binding.linkLast}`), Excel.ChartSeriesBy.columns);
  chart.name = instance.chartName;
  const names = progress ? [source.valueHeader, '未完成', '滑珠']
    : ['连线', source.valueHeader, source.secondHeader, '正变化', '负变化', '分类标签', '上端标签', '下端标签'];
  chart.series.getItemAt(0).name = names[0];
  for (const name of names.slice(1)) chart.series.add(name);
  const anchor = sourceSheet.getRangeByIndexes(source.rowIndex, source.columnIndex + source.columnCount + 1, 1, 1);
  anchor.load('left,top'); await syncStage(context, 'create-chart-and-anchor');
  chart.left = anchor.left; chart.top = anchor.top;
  const size = {horizontal: [560, 460], vertical: [640, 420], progress: [600, 380]}[binding.variant];
  chart.width = size[0]; chart.height = size[1];
  chart.plotVisibleOnly = false; chart.displayBlanksAs = 'NotPlotted';
  bindSeries(chart, helper, binding);
  return chart;
}

function formatLabel(dataLabel, position, color, size = 10, bold = false) {
  dataLabel.position = position;
  dataLabel.format.font.size = size; dataLabel.format.font.color = color; dataLabel.format.font.bold = bold;
}

// 逐点链接标签文本到辅助表；重绑时只给新增点设默认格式，保留用户已改的标签格式与显隐。
function linkLabels({chart, instance, source, options, binding, previous}) {
  const sheet = quote(instance.helperName), oldCount = previous?.count ?? 0, P = Excel.ChartDataLabelPosition;
  const groups = binding.variant === 'progress'
    ? [{series: 2, column: 'F', position: P.center, color: '#FFFFFF', size: 8, bold: true, show: options.showValues !== false}]
    : [
      {series: 3, column: 'I', position: binding.horizontal ? P.top : P.right, color: COLORS.positive, bold: true, show: binding.change !== 'none'},
      {series: 4, column: 'I', position: binding.horizontal ? P.top : P.right, color: COLORS.negative, bold: true, show: binding.change !== 'none'},
      {series: 5, column: 'A', position: binding.horizontal ? P.left : P.bottom, color: COLORS.text, show: true},
      {series: 6, column: 'M', position: P.top, color: binding.horizontal ? COLORS.a : COLORS.text, show: options.showValues !== false},
      {series: 7, column: 'N', position: binding.horizontal ? P.top : P.bottom, color: binding.horizontal ? COLORS.b : COLORS.text, show: options.showValues !== false}
    ];
  for (const g of groups) {
    if (!g.show) continue;
    for (let i = 0; i < source.count; i++) {
      const point = chart.series.getItemAt(g.series).points.getItemAt(i);
      if (i >= oldCount) { point.hasDataLabel = true; formatLabel(point.dataLabel, g.position, g.color, g.size, g.bold); }
      point.dataLabel.formula = `=${sheet}!$${g.column}$${i + 3}`;
    }
  }
}

function hideAxis(axis) {
  axis.majorTickMark = 'None'; axis.minorTickMark = 'None'; axis.tickLabelPosition = 'None';
  axis.format.line.lineStyle = 'None'; axis.majorGridlines.visible = false;
}

function setPairBounds(chart, binding) {
  const positionAxis = binding.horizontal ? chart.axes.valueAxis : chart.axes.categoryAxis;
  if (binding.horizontal) { positionAxis.minimum = 0; positionAxis.maximum = binding.n + 1; positionAxis.majorUnit = 1; }
  else { positionAxis.minimum = 0.5; positionAxis.maximum = binding.n + 0.5; positionAxis.majorUnit = 1; }
}

function labelWidth(values) {
  // 估算分类文字宽度（10pt：中文≈10.5pt、西文≈6pt），给水平哑铃左侧预留空间。
  const width = text => [...String(text)].reduce((w, ch) => w + (/[⺀-￿]/.test(ch) ? 10.5 : 6), 0);
  const widest = Math.max(0, ...values.slice(1).map(row => width(row[0])));
  return Math.min(170, Math.max(48, Math.round(widest + 18)));
}

// 源数据全为非负时，纵向分类锚可能把轴拉到负数段；该段只放标签，刻度不显示负值
// （景歌麦肯锡滑珠“40 以下不显示”同一手法）。有负值时照常显示。
function axisNumberFormat(mask, values) {
  const nonNegative = values.slice(1).every(row => row.slice(1).every(v => typeof v !== 'number' || v >= 0));
  return nonNegative ? `${mask};;${mask.replace(/[#,]/g, '') || '0'}` : mask;
}

async function showPairLegend(context, chart) {
  chart.legend.visible = true; chart.legend.position = 'Top'; chart.legend.format.font.size = 10;
  await syncStage(context, 'legend');
  try {
    const entries = chart.legend.legendEntries; entries.load('items'); await context.sync();
    entries.items.forEach((entry, i) => { entry.visible = i === 1 || i === 2; });
    await context.sync();
    return true;
  } catch {
    // 宿主不支持逐项图例时退回隐藏图例，A/B 颜色由标签颜色区分。
    chart.legend.visible = false; await syncStage(context, 'legend-fallback');
    return false;
  }
}

async function style(args) {
  const {context, chart, options, binding, source} = args;
  await syncStage(context, 'bind-series-data');
  chart.format.font.name = FONT; chart.format.font.size = 10; chart.format.font.color = COLORS.text;
  chart.format.fill.setSolidColor('#FFFFFF'); chart.format.border.lineStyle = 'None';
  chart.title.visible = true; chart.title.text = options.title;
  chart.title.format.font.size = 14; chart.title.format.font.bold = true;
  await syncStage(context, 'chart-area-title');
  if (binding.variant === 'progress') return styleProgress(args);

  const link = chart.series.getItemAt(0);
  link.chartType = Excel.ChartType.xyscatterLinesNoMarkers; link.axisGroup = Excel.ChartAxisGroup.primary;
  link.format.line.lineStyle = 'Continuous'; link.format.line.color = COLORS.link; link.format.line.weight = binding.horizontal ? 2.25 : 1.5;
  for (const [i, color] of [[1, COLORS.a], [2, COLORS.b]]) {
    const s = chart.series.getItemAt(i);
    s.chartType = Excel.ChartType.xyscatter; s.axisGroup = Excel.ChartAxisGroup.primary;
    s.markerStyle = Excel.ChartMarkerStyle.circle; s.markerSize = 11;
    s.markerBackgroundColor = color; s.markerForegroundColor = color; s.format.line.lineStyle = 'None';
  }
  for (let i = 3; i < PAIR_SERIES; i++) {
    const s = chart.series.getItemAt(i);
    s.chartType = Excel.ChartType.xyscatter; s.axisGroup = Excel.ChartAxisGroup.primary;
    s.markerStyle = Excel.ChartMarkerStyle.none; s.format.line.lineStyle = 'None';
  }
  await syncStage(context, 'series');
  const valueAxis = binding.horizontal ? chart.axes.categoryAxis : chart.axes.valueAxis;
  const positionAxis = binding.horizontal ? chart.axes.valueAxis : chart.axes.categoryAxis;
  hideAxis(positionAxis);
  if (binding.horizontal) { positionAxis.majorGridlines.visible = true; positionAxis.majorGridlines.format.line.color = COLORS.grid; positionAxis.majorGridlines.format.line.weight = 0.5; }
  setPairBounds(chart, binding);
  valueAxis.majorTickMark = 'None'; valueAxis.minorTickMark = 'None'; valueAxis.format.line.lineStyle = 'None';
  valueAxis.tickLabelPosition = 'Low'; valueAxis.majorGridlines.visible = true;
  valueAxis.majorGridlines.format.line.color = COLORS.grid; valueAxis.majorGridlines.format.line.weight = 0.5;
  valueAxis.numberFormat = axisNumberFormat(binding.mask, source.values);
  await syncStage(context, 'axes');
  const legend = options.showLegend !== false ? await showPairLegend(context, chart) : (chart.legend.visible = false, await syncStage(context, 'legend-off'), false);
  linkLabels(args); await syncStage(context, 'linked-labels');
  let reserved = null;
  if (binding.horizontal) {
    // 分类锚落在数值轴起点，标签向左；给绘图区左侧固定留白，避免文字被图表边缘裁掉。
    reserved = labelWidth(source.values);
    chart.plotArea.load('insideLeft,insideWidth'); await syncStage(context, 'plot-area-read');
    const right = chart.plotArea.insideLeft + chart.plotArea.insideWidth;
    if (chart.plotArea.insideLeft < reserved) {
      chart.plotArea.insideLeft = reserved; chart.plotArea.insideWidth = Math.max(120, right - reserved);
      await syncStage(context, 'plot-area-reserve');
    }
  }
  return {variant: binding.variant, legendEntries: legend, reservedLeft: reserved, notes: [`${label(binding.variant)}：原生散点 ${PAIR_SERIES} 系列，连线用 #N/A 断开。`]};
}

async function styleProgress(args) {
  const {context, chart, options} = args;
  const done = chart.series.getItemAt(0), rest = chart.series.getItemAt(1), bead = chart.series.getItemAt(2);
  done.gapWidth = 150; done.overlap = 100;
  done.format.fill.setSolidColor(COLORS.done); done.format.line.lineStyle = 'None';
  rest.format.fill.setSolidColor(COLORS.rest); rest.format.line.lineStyle = 'None';
  await syncStage(context, 'progress-columns');
  bead.chartType = Excel.ChartType.lineMarkers; bead.axisGroup = Excel.ChartAxisGroup.primary;
  bead.format.line.lineStyle = 'None';
  bead.markerStyle = Excel.ChartMarkerStyle.circle; bead.markerSize = 22;
  bead.markerBackgroundColor = COLORS.done; bead.markerForegroundColor = '#FFFFFF';
  await syncStage(context, 'progress-bead');
  const valueAxis = chart.axes.valueAxis, categoryAxis = chart.axes.categoryAxis;
  valueAxis.minimum = 0;
  if (!args.source.values.slice(1).some(row => row[1] > 1)) valueAxis.maximum = 1;
  valueAxis.majorUnit = 0.1; valueAxis.numberFormat = '0%';
  valueAxis.majorGridlines.visible = false; valueAxis.format.line.lineStyle = 'None'; valueAxis.majorTickMark = 'None';
  categoryAxis.majorTickMark = 'None'; categoryAxis.format.line.color = '#D9D9D9';
  chart.legend.visible = false;
  await syncStage(context, 'progress-axes');
  linkLabels(args); await syncStage(context, 'progress-labels');
  return {variant: 'progress', notes: ['进度滑珠柱：原生堆积柱形 2 系列 + 带标记折线滑珠。']};
}

async function rebind(args) {
  const {context, chart, binding, previous} = args;
  chart.series.load('items'); await syncStage(context, 'rebind-read-series');
  if (chart.series.items.length !== binding.seriesCount) throw new UserError('原图系列数已被改动，未重绑。请另建新图。');
  bindSeries(chart, args.helper, binding);
  if (binding.variant !== 'progress' && previous?.count !== binding.n) setPairBounds(chart, binding);
  await syncStage(context, 'rebind-series-data');
  linkLabels(args); await syncStage(context, 'rebind-linked-labels');
  return {variant: binding.variant};
}

const inputSpec = {
  category: {id: 'comparison', title: '比较'},
  description: '含表头：三列（分类、系列A、系列B）生成水平哑铃或纵向滑珠，两点连线比较两组数据；两列（分类、完成率小数）生成进度滑珠柱。改源数据自动联动。',
  minVersion: '1.9', columns: 3, minRows: 2, maxRows: MAX_ROWS,
  options: [
    {id: 'title', label: '标题', type: 'text', default: '哑铃图', maxLength: 80},
    {id: 'variant', label: '样式', type: 'select', default: 'auto', choices: [
      {value: 'auto', label: '自动识别'}, {value: 'horizontal', label: '水平哑铃'},
      {value: 'vertical', label: '纵向滑珠'}, {value: 'progress', label: '进度滑珠柱'}]},
    {id: 'valueFormat', label: '数值格式', type: 'select', default: 'number', choices: [{value: 'number', label: '数字'}, {value: 'percent', label: '百分比'}]},
    {id: 'decimals', label: '小数位', type: 'number', default: 0, min: 0, max: 3, step: 1},
    {id: 'changeLabel', label: '变化标签', type: 'select', default: 'none', choices: [
      {value: 'none', label: '不显示'}, {value: 'diff', label: '差值（B−A）'}, {value: 'growth', label: '涨幅（B/A−1）'}]},
    {id: 'changeSuffix', label: '变化标签后缀', type: 'text', default: '', maxLength: 12},
    {id: 'startAtZero', label: '数值轴从 0 开始', type: 'checkbox', default: true},
    {id: 'showValues', label: '显示数值', type: 'checkbox', default: true},
    {id: 'showLegend', label: '显示图例', type: 'checkbox', default: true}
  ],
  storage: {marker: MARK, helperPrefix: HELPER_PREFIX, chartPrefix: CHART_PREFIX}
};

export {validate, helperRecipe, resolvedVariant, numberMask, pairColumns, labelWidth, axisNumberFormat, COLORS};
export default {id: 'dumbbell', title: '滑珠/哑铃图', inputSpec, validate, writeHelper, createChart, style, rebind};
