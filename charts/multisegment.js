// 景歌图族“多层柱形/多段条形”：每个系列占一段（条形）或一层（柱形），同轴对齐。
// 原生堆积条形/柱形；每系列三条堆积系列：数值、补全（可着浅色底条）、段间空白（透明）。
// 绘图值按公式归一化，使数值轴固定为 0..系列数，改数后无需改轴即可保持分段对齐。
// 来源文章、数据与已知差异见 references/multisegment/。
import {UserError} from '../core/errors.js';
import {setChartTitleTypography} from '../core/presentation.js';

export const id = 'multisegment', title = '多段条形/多层柱形';
export const FONT = '微软雅黑';
export const MAX_ROWS = 60, MIN_SERIES = 2, MAX_SERIES = 8, ROLES = 3;
export const COLORS = Object.freeze(['#156FA4', '#2FBFD6', '#E2B66C', '#72C9C3', '#D77A7A', '#7DC855', '#F39200', '#8C75C9']);
export const VARIANTS = Object.freeze(['bar-equal', 'bar-fit', 'column-layers']);
export const LAYOUT = Object.freeze({
  bar: {width: 600, height: 400, gapWidth: 70},
  column: {width: 600, height: 480, gapWidth: 60},
});
const quote = name => `'${String(name).replace(/'/g, "''")}'`;
const textLiteral = value => `"${String(value).replace(/"/g, '""')}"`;
const errorText = value => /^#(?:REF!|VALUE!|N\/A|DIV\/0!|NAME\?|NUM!|SPILL!|CALC!|NULL!|GETTING_DATA)/.test(String(value));
const goodText = value => ['string', 'number'].includes(typeof value) && String(value).trim() !== '' && !errorText(value) && !(typeof value === 'number' && !Number.isFinite(value));

export function columnLetter(index) {
  if (!Number.isInteger(index) || index < 0 || index > 16383) throw new UserError('列索引无效。');
  let n = index + 1, name = '';
  while (n) { const digit = (n - 1) % 26; name = String.fromCharCode(65 + digit) + name; n = Math.floor((n - 1) / 26); }
  return name;
}

// 浅色补全条：系列色向白色混合。
export function tint(hex, amount = 0.75) {
  const m = /^#?([0-9a-f]{6})$/i.exec(hex);
  if (!m) throw new UserError('颜色须为 #RRGGBB。');
  const parts = m[1].match(/../g).map(h => parseInt(h, 16)).map(c => Math.round(c + (255 - c) * amount));
  return '#' + parts.map(c => c.toString(16).padStart(2, '0')).join('').toUpperCase();
}

export function validate(values) {
  if (!Array.isArray(values) || values.length < 2 || values.length > MAX_ROWS + 1) throw new UserError(`请选择含表头的数据：第一列分类，其后 ${MIN_SERIES}–${MAX_SERIES} 列系列数值，1–${MAX_ROWS} 行。`);
  const width = values[0]?.length;
  if (!Number.isInteger(width) || width < MIN_SERIES + 1 || width > MAX_SERIES + 1 || values.some(row => !Array.isArray(row) || row.length !== width)) {
    throw new UserError(`需要 ${MIN_SERIES + 1}–${MAX_SERIES + 1} 列：分类 + ${MIN_SERIES}–${MAX_SERIES} 个系列。`);
  }
  if (values[0].slice(1).some(v => !goodText(v) || String(v).length > 255)) throw new UserError('各系列列须有非空表头（最多255字）。');
  values.slice(1).forEach((row, i) => {
    if (!goodText(row[0])) throw new UserError(`第${i + 2}行分类为空或无效。`);
    row.slice(1).forEach((v, j) => {
      if (typeof v !== 'number' || !Number.isFinite(v) || v < 0) throw new UserError(`第${i + 2}行第${j + 2}列须为不小于0的数值；多段/多层图不支持负数，缺失值请勿留空。`);
    });
  });
  return {count: values.length - 1, series: width - 1, headers: values[0].map(String)};
}

export function resolveOptions(input = {}) {
  const o = {title: '多系列并列条形图', variant: 'bar-equal', showTrack: false, labelPosition: 'outside', gap: 0.3, decimals: 0, showLegend: true, ...input};
  if (typeof o.title !== 'string' || o.title.length > 160) throw new UserError('标题最多160字。');
  if (!VARIANTS.includes(o.variant)) throw new UserError('图形变体无效。');
  if (!['outside', 'inside', 'none'].includes(o.labelPosition)) throw new UserError('标签位置无效。');
  for (const key of ['showTrack', 'showLegend']) if (typeof o[key] !== 'boolean') throw new UserError(`${key} 须为开关。`);
  if (!Number.isFinite(o.gap) || o.gap < 0.05 || o.gap > 2) throw new UserError('段间空白须为最大值的 5%–200%。');
  if (!Number.isInteger(o.decimals) || o.decimals < 0 || o.decimals > 4) throw new UserError('小数位须为0至4的整数。');
  const layout = o.variant === 'column-layers' ? LAYOUT.column : LAYOUT.bar;
  if (o.width === undefined) o.width = layout.width;
  if (o.height === undefined) o.height = layout.height;
  for (const [key, min, max] of [['width', 300, 1400], ['height', 220, 2000]]) if (!Number.isFinite(o[key]) || o[key] < min || o[key] > max) throw new UserError('图表尺寸超出范围。');
  return o;
}

// 辅助表列布局（m 个系列，零基列号）：
//   A 分类；1+3j / 2+3j / 3+3j：系列 j 的归一化 数值 / 补全 / 空白；
//   L0+j 标签文本；R0+j 原值；V0+j 有效标记；P0.. 参数块（每系列一行）。
// 第1行：B1 行数 C1 系列数 D1 全局最大值 E1 归一化比例 F1 空白比例 G1 无效单元格数 H1 各段宽度合计。
export function layoutColumns(m) {
  const L0 = 1 + ROLES * m, R0 = L0 + m, V0 = R0 + m, P0 = V0 + m + 1;
  return {L0, R0, V0, P0, last: P0 + 4};
}

export function formulaPlan(source, input = {}) {
  const o = resolveOptions(input), info = validate(source.values), n = info.count, m = info.series, last = n + 2;
  if (source.count !== n || source.rowCount !== n + 1 || source.columnCount !== m + 1 || !Number.isInteger(source.rowIndex) || source.rowIndex < 0
    || source.rowIndex + source.rowCount > 1048576 || !Number.isInteger(source.columnIndex) || source.columnIndex < 0
    || source.columnIndex + source.columnCount + 1 > 16383 || typeof source.sheet !== 'string' || !source.sheet) throw new UserError('源区域尺寸或地址不一致。');
  const {L0, R0, V0, P0} = layoutColumns(m), c = columnLetter;
  const src = (col, row) => `${quote(source.sheet)}!$${c(source.columnIndex + col)}$${row}`;
  const fmt = '0' + (o.decimals ? '.' + '0'.repeat(o.decimals) : '');
  const rawRange = `$${c(R0)}$3:$${c(R0 + m - 1)}$${last}`, validRange = `$${c(V0)}$3:$${c(V0 + m - 1)}$${last}`;
  const fit = o.variant === 'bar-fit';
  const header = ['分类'];
  for (let j = 0; j < m; j++) header.push(`=${src(j + 1, source.rowIndex + 1)}&""`, '补全', '空白');
  for (let j = 0; j < m; j++) header.push('标签');
  for (let j = 0; j < m; j++) header.push('原值');
  for (let j = 0; j < m; j++) header.push('有效');
  header.push('', '系列', '系列最大值', '参考满格', '段宽', '段起点');
  const rows = [];
  for (let i = 0; i < n; i++) {
    const r = i + 3, sr = source.rowIndex + i + 2, row = [`=IFERROR(IF(${src(0, sr)}&""="","数据无效",${src(0, sr)}&""),"数据无效")`];
    for (let j = 0; j < m; j++) {
      const raw = `$${c(R0 + j)}${r}`, ref = `$${c(P0 + 2)}$${j + 3}`;
      row.push(`=${raw}*$E$1`, `=MAX(0,${ref}-${raw})*$E$1`, `=$F$1*$D$1*$E$1`);
    }
    for (let j = 0; j < m; j++) row.push(`=IF(${c(V0 + j)}${r},TEXT(${c(R0 + j)}${r},"${fmt}"),"无效")`);
    for (let j = 0; j < m; j++) { const s = src(j + 1, sr); row.push(`=IFERROR(IF(AND(ISNUMBER(${s}),${s}>=0),${s},0),0)`); }
    for (let j = 0; j < m; j++) { const s = src(j + 1, sr); row.push(`=IFERROR(AND(ISNUMBER(${s}),${s}>=0),FALSE)`); }
    rows.push(row);
  }
  const params = [];
  for (let j = 0; j < m; j++) {
    const r = j + 3, col = `$${c(R0 + j)}$3:$${c(R0 + j)}$${last}`;
    params.push([
      `=${c(1 + ROLES * j)}2`,
      `=MAX(${col})`,
      fit ? `=${c(P0 + 1)}${r}` : '=$D$1',
      `=${c(P0 + 2)}${r}+$F$1*$D$1`,
      j === 0 ? '=0' : `=${c(P0 + 4)}${r - 1}+${c(P0 + 3)}${r - 1}*$E$1`,
    ]);
  }
  const scale = fit ? '=$C$1/$H$1' : '=1/($D$1*(1+$F$1))';
  return {options: o, count: n, series: m, last, numberFormat: fmt, columns: layoutColumns(m), seriesCount: ROLES * m, headers: info.headers,
    writes: [
      {range: 'B1:H1', formulas: [[`=${n}`, `=${m}`, `=IF(MAX(${rawRange})>0,MAX(${rawRange}),1)`, scale, `=${o.gap}`, `=COUNTIF(${validRange},FALSE)`,
        `=SUM($${c(P0 + 3)}$3:$${c(P0 + 3)}$${m + 2})`]]},
      {range: 'N1', formulas: [[`=${textLiteral(o.variant)}`]]},
      {range: `A2:${c(P0 + 4)}2`, formulas: [header]},
      {range: `A3:${c(V0 + m - 1)}${last}`, formulas: rows},
      {range: `${c(P0)}3:${c(P0 + 4)}${m + 2}`, formulas: params},
    ]};
}

// 纯计算：与辅助表公式一致的归一化结果，供单测和验收期望值生成使用。
export function computeSegments(matrix, input = {}) {
  const o = resolveOptions(input), m = matrix[0].length, n = matrix.length;
  const all = matrix.flat(), M = Math.max(...all) > 0 ? Math.max(...all) : 1;
  const seriesMax = Array.from({length: m}, (_, j) => Math.max(...matrix.map(r => r[j])));
  const ref = seriesMax.map(v => o.variant === 'bar-fit' ? v : M);
  const widths = ref.map(v => v + o.gap * M), total = widths.reduce((a, b) => a + b, 0);
  const scale = o.variant === 'bar-fit' ? m / total : 1 / (M * (1 + o.gap));
  const starts = widths.map((_, j) => widths.slice(0, j).reduce((a, b) => a + b, 0) * scale);
  const series = [];
  for (let j = 0; j < m; j++) {
    series.push(matrix.map(r => r[j] * scale), matrix.map(r => Math.max(0, ref[j] - r[j]) * scale), Array(n).fill(o.gap * M * scale));
  }
  return {max: M, scale, seriesMax, ref, widths, total, starts, series, axisMax: m};
}

function chartType(o) { return o.variant === 'column-layers' ? 'ColumnStacked' : 'BarStacked'; }

async function writeHelper({context, helper, source, options, previous, chart}) {
  const plan = formulaPlan(source, options);
  if (previous) {
    chart.series.load('items/chartType,items/axisGroup'); await context.sync();
    const type = chartType(plan.options);
    if (chart.series.items.length !== plan.seriesCount || chart.series.items.some(s => s.chartType !== type || s.axisGroup !== 'Primary')) {
      throw new UserError('系列数（每个系列3条堆积）或图表类型已改变；按选区更新只支持增删行，系列列数须与原图一致。辅助表未改写。');
    }
    const labelSeries = labelRole(plan.options);
    if (labelSeries !== null) {
      const items = Array.from({length: plan.series}, (_, j) => chart.series.getItemAt(ROLES * j + labelSeries));
      for (const s of items) s.points.load('items/hasDataLabel');
      await context.sync();
      plan.visibility = items.map(s => s.points.items.map(p => p.hasDataLabel === true));
    }
    const rows = Math.max(previous.count, plan.count) + 2;
    helper.getRange(`A3:${columnLetter(plan.columns.V0 + plan.series - 1)}${rows}`).clear('Contents');
  }
  for (const w of plan.writes) helper.getRange(w.range).formulas = w.formulas;
  return plan;
}

function bind(chart, helper, p) {
  const categories = helper.getRange(`A3:A${p.last}`);
  for (let k = 0; k < p.seriesCount; k++) {
    const s = chart.series.getItemAt(k), col = columnLetter(1 + k);
    s.setXAxisValues(categories); s.setValues(helper.getRange(`${col}3:${col}${p.last}`));
  }
}

// outside：贴在补全段起点（InsideBase），即数值段末端外侧；inside：数值段居中。
function labelRole(o) { return o.labelPosition === 'outside' ? 1 : o.labelPosition === 'inside' ? 0 : null; }

async function bindLabels({context, chart, instance, binding: p}, initial) {
  const role = labelRole(p.options);
  if (role === null) return;
  const {L0} = p.columns;
  for (let j = 0; j < p.series; j++) {
    const series = chart.series.getItemAt(ROLES * j + role);
    for (let i = 0; i < p.count; i++) {
      const known = !initial && p.visibility?.[j]?.[i] !== undefined;
      if (known && !p.visibility[j][i]) continue; // 用户手动隐藏的标签保持隐藏
      const point = series.points.getItemAt(i);
      if (!known) {
        point.hasDataLabel = true;
        const label = point.dataLabel;
        label.position = role === 1 ? 'InsideBase' : 'Center';
        label.format.font.name = FONT; label.format.font.size = 10;
        label.format.font.color = role === 1 ? COLORS[j % COLORS.length] : '#FFFFFF';
      }
      point.dataLabel.formula = `=${quote(instance.helperName)}!$${columnLetter(L0 + j)}$${i + 3}`;
    }
  }
  await context.sync();
}

async function createChart({context, sourceSheet, helper, source, instance, binding: p}) {
  const chart = sourceSheet.charts.add(chartType(p.options), helper.getRange(`B3:B${p.last}`), 'Columns');
  chart.name = instance.chartName; chart.plotVisibleOnly = false; chart.displayBlanksAs = 'Zero';
  chart.series.load('items'); await context.sync();
  if (chart.series.items.length !== 1) throw new UserError('初始系列数量异常。');
  for (let k = 1; k < p.seriesCount; k++) chart.series.add(`S${k}`);
  await context.sync();
  for (let k = 0; k < p.seriesCount; k++) {
    const s = chart.series.getItemAt(k), j = Math.floor(k / ROLES), role = k % ROLES;
    s.chartType = chartType(p.options); s.axisGroup = 'Primary';
    s.name = role === 0 ? `=${quote(instance.helperName)}!$${columnLetter(1 + k)}$2` : `${p.headers[j + 1]}${role === 1 ? '补全' : '空白'}`;
  }
  bind(chart, helper, p);
  const anchor = sourceSheet.getRangeByIndexes(source.rowIndex, source.columnIndex + source.columnCount + 1, 1, 1);
  anchor.load('left,top'); await context.sync();
  chart.left = anchor.left; chart.top = anchor.top; chart.width = p.options.width; chart.height = p.options.height;
  await context.sync();
  await bindLabels({context, chart, instance, binding: p}, true);
  return chart;
}

async function style({context, chart, binding: p}) {
  const o = p.options, column = o.variant === 'column-layers';
  chart.format.font.name = FONT; chart.format.font.size = 10; chart.format.font.color = '#000000';
  chart.format.fill.setSolidColor('#FFFFFF');
  chart.title.visible = Boolean(o.title); chart.title.overlay = false;
  if (o.title) { chart.title.text = o.title; chart.title.horizontalAlignment = 'Center'; setChartTitleTypography(chart, {text: o.title}); chart.title.format.font.bold = false; }
  for (let k = 0; k < p.seriesCount; k++) {
    const s = chart.series.getItemAt(k), j = Math.floor(k / ROLES), role = k % ROLES, color = COLORS[j % COLORS.length];
    s.overlap = 100; s.gapWidth = (column ? LAYOUT.column : LAYOUT.bar).gapWidth; s.format.line.lineStyle = 'None';
    if (role === 0) s.format.fill.setSolidColor(color);
    else if (role === 1 && o.showTrack) s.format.fill.setSolidColor(tint(color));
    else s.format.fill.clear();
    if (labelRole(o) !== role) s.hasDataLabels = false;
  }
  chart.legend.visible = o.showLegend;
  if (o.showLegend) {
    chart.legend.position = 'Top'; chart.legend.overlay = false;
    await context.sync();
    const entries = chart.legend.legendEntries; entries.load('items'); await context.sync();
    entries.items.forEach((e, i) => { e.visible = i % ROLES === 0; });
  }
  const category = chart.axes.getItem('Category', 'Primary'), value = chart.axes.getItem('Value', 'Primary');
  category.categoryType = 'TextAxis'; category.reversePlotOrder = !column; category.tickLabelPosition = 'Low';
  category.majorTickMark = 'None'; category.minorTickMark = 'None'; category.majorGridlines.visible = false;
  category.format.line.color = '#BFBFBF'; category.format.line.weight = .75;
  category.format.font.name = FONT; category.format.font.size = 10; category.format.font.color = '#000000';
  value.minimum = 0; value.maximum = p.series; value.majorUnit = 1;
  value.tickLabelPosition = 'None'; value.majorTickMark = 'None'; value.minorTickMark = 'None'; value.format.line.lineStyle = 'None';
  // 等宽分段/多层柱形：段宽恒为1，网格线正好落在每段起点（条形的竖线、柱形每层的底线）。
  value.majorGridlines.visible = o.variant !== 'bar-fit';
  if (o.variant !== 'bar-fit') { value.majorGridlines.format.line.color = column ? '#A6A6A6' : '#D9D9D9'; value.majorGridlines.format.line.weight = .75; }
  await context.sync();
  return {nativeVerified: false, visualVerified: false, notes: [
    `${o.variant}：每系列 3 条原生堆积系列（数值/补全/空白），共 ${p.seriesCount} 条；数值轴固定 0–${p.series}。`,
    '分段对齐由公式归一化维持，改源数自动联动；标签为逐点公式链接文本。',
    '原文的平均值箭头线、左侧层名、字体与配色细节未复刻，见 references/multisegment/README.md。',
  ]};
}

async function finalizePresentation({context, chart, binding: p}) {
  // 统一字体会把逐点标签色改黑；恢复外侧标签的系列同色、内侧标签白字。
  const role = labelRole(p.options);
  if (role === null) return;
  const items = Array.from({length: p.series}, (_, j) => chart.series.getItemAt(ROLES * j + role));
  for (const s of items) s.points.load('items/hasDataLabel');
  await context.sync();
  items.forEach((s, j) => s.points.items.forEach(point => {
    if (point.hasDataLabel === true) point.dataLabel.format.font.color = role === 1 ? COLORS[j % COLORS.length] : '#FFFFFF';
  }));
  await context.sync();
}

async function rebind(args) {
  // 只更新引用和新点标签；不改颜色、轴、尺寸、位置。
  bind(args.chart, args.helper, args.binding); await args.context.sync();
  await bindLabels(args, false);
  return {notes: ['仅更新公式、系列引用和新增点标签；已有系列/点格式保留。系列列数不可变。']};
}

const inputSpec = {
  category: {id: 'comparison', title: '比较'}, minVersion: '1.9', columns: 5, minRows: 1, maxRows: MAX_ROWS,
  description: '第一列分类，其后2–8列为各系列数值（含表头，非负）。每个系列占一段（条形）或一层（柱形），同轴对齐。',
  options: [
    {id: 'title', label: '标题', type: 'text', default: '多系列并列条形图', maxLength: 160},
    {id: 'variant', label: '样式', type: 'select', default: 'bar-equal', choices: [
      {value: 'bar-equal', label: '多段式条形（等宽分段）'},
      {value: 'bar-fit', label: '分段式条形（段宽按系列最大值）'},
      {value: 'column-layers', label: '多层柱形（每系列一层）'},
    ]},
    {id: 'showTrack', label: '显示浅色补全底条', type: 'checkbox', default: false},
    {id: 'labelPosition', label: '数值标签', type: 'select', default: 'outside', choices: [
      {value: 'outside', label: '末端外侧'}, {value: 'inside', label: '段内居中'}, {value: 'none', label: '不显示'},
    ]},
    {id: 'gap', label: '段间空白（占最大值）', type: 'number', default: 30, min: 5, max: 200, step: 5, scale: 0.01, suffix: '%'},
    {id: 'decimals', label: '小数位', type: 'number', default: 0, min: 0, max: 4, step: 1},
    {id: 'showLegend', label: '显示图例', type: 'checkbox', default: true},
    {id: 'width', label: '宽度（pt）', type: 'number', default: LAYOUT.bar.width, min: 300, max: 1400, step: 1},
    {id: 'height', label: '高度（pt）', type: 'number', default: LAYOUT.bar.height, min: 220, max: 2000, step: 1},
  ],
  storage: {marker: 'ChartCraft.MultiSegment.v1', helperPrefix: 'CC_多段_', chartPrefix: 'CC_MultiSeg_'},
};
export default {id, title, inputSpec, validate, writeHelper, createChart, style, rebind, finalizePresentation};
