// 景歌图族“达成/偏差/差额柱形”：三种原生组合图，全部由辅助表公式驱动。
// 来源与已知差异见 references/achievement/README.md。只用 Chart API，不用形状或图片。
import {UserError} from '../core/errors.js';
import {setChartTitleTypography} from '../core/presentation.js';

export const id = 'achievement', title = '达成差额柱形图';
export const FONT = '微软雅黑';
export const MAX_ROWS = 60;
export const COLORS = {
  done: '#1B3F6E', notDone: '#BFBFBF', plan: '#7F7F7F', gapUp: '#9BBBB1', gapDown: '#FF0000',
  completed: '#ADB9CA', remaining: '#F6D5D5', remainingLine: '#C00000', excess: '#8E7CC3', rate: '#8B1A1A',
  above: '#1F3FD6', below: '#6CCBF5', negativeText: '#FF0000', grid: '#EDEDED'
};
export const LAYOUT = {width: 560, height: 360};
export const VARIANTS = ['gap', 'progress', 'updown'];
// Series layout per variant: [name, helper column, chartType, axisGroup].
export const SERIES = {
  gap: [['完成', 'F', 'ColumnClustered', 'Primary'], ['未完成', 'G', 'ColumnClustered', 'Primary'], ['计划', 'B', 'Line', 'Primary'],
    ['差额', 'H', 'ColumnClustered', 'Secondary'], ['差额（负）', 'I', 'ColumnClustered', 'Secondary']],
  progress: [['已完成', 'J', 'ColumnStacked', 'Primary'], ['未完成', 'K', 'ColumnStacked', 'Primary'], ['超额完成', 'L', 'ColumnStacked', 'Primary'],
    ['完成率', 'O', 'Line', 'Primary']],
  updown: [['达标', 'M', 'ColumnClustered', 'Primary'], ['未达标', 'N', 'ColumnClustered', 'Primary']]
};
const quote = s => `'${String(s).replace(/'/g, "''")}'`;
const ERROR_TEXT = /^#(REF!|VALUE!|N\/A|DIV\/0!|NAME\?|NUM!|SPILL!|CALC!|NULL!)/;
const validText = v => ['string', 'number'].includes(typeof v) && String(v).trim() !== '' && !(typeof v === 'number' && !Number.isFinite(v)) && !ERROR_TEXT.test(String(v));

export function columnLetter(index) {
  if (!Number.isInteger(index) || index < 0 || index > 16383) throw new UserError('源列索引无效。');
  let s = '', n = index + 1; while (n) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s;
}

export function resolveOptions(input = {}) {
  const o = {title: '月度达成差额分析图', variant: 'gap', showLabels: true, showLegend: true, ...input};
  if (typeof o.title !== 'string' || o.title.length > 80) throw new UserError('标题须为不超过 80 字的文字。');
  if (!VARIANTS.includes(o.variant)) throw new UserError('图表变体无效。');
  for (const k of ['showLabels', 'showLegend']) if (typeof o[k] !== 'boolean') throw new UserError('显示选项必须为开关。');
  return o;
}

export function validate(values) {
  if (!Array.isArray(values) || values.length < 2 || values.length > MAX_ROWS + 1) throw new UserError(`请选择含表头的三列：分类、计划（目标）、实际，1–${MAX_ROWS} 行数据。`);
  if (values.some(r => !Array.isArray(r) || r.length !== 3)) throw new UserError('数据必须是规则的 3 列：分类、计划（目标）、实际。');
  if (values[0].some(v => !validText(v) || String(v).length > 255)) throw new UserError('三列都须有非空文本表头。');
  values.slice(1).forEach((r, i) => {
    if (!validText(r[0])) throw new UserError(`第 ${i + 2} 行分类为空或无效。`);
    if (typeof r[1] !== 'number' || !Number.isFinite(r[1]) || r[1] <= 0) throw new UserError(`第 ${i + 2} 行计划须为大于 0 的数值（完成率以它为分母）。`);
    if (typeof r[2] !== 'number' || !Number.isFinite(r[2]) || r[2] < 0) throw new UserError(`第 ${i + 2} 行实际须为非负数值。`);
  });
  return {count: values.length - 1, categoryHeader: String(values[0][0]), planHeader: String(values[0][1]), actualHeader: String(values[0][2])};
}

// “好看”的刻度上限：1/1.2/1.5/2/2.5/3/4/5/6/8 × 10^k，不小于 x。
export function niceCeil(x) {
  if (!Number.isFinite(x) || x <= 0) return 1;
  const k = Math.floor(Math.log10(x)), base = 10 ** k;
  for (const m of [1, 1.2, 1.5, 2, 2.5, 3, 4, 5, 6, 8, 10]) { const v = +(m * base).toPrecision(12); if (v >= x - 1e-12) return v; }
  return 10 * base;
}

// 轴范围只在创建/“应用新版格式”时按当前数据计算；改数后柱高仍由公式联动。
export function scales(rows, variant) {
  const plan = rows.map(r => r[1]), actual = rows.map(r => r[2]);
  if (variant === 'gap') {
    // 下层实际柱占绘图区约 62%；差额零线在 72% 高度，正差额向上、负差额向下不压到实际柱。
    const top = Math.max(...plan, ...actual), d = rows.map(r => r[2] - r[1]);
    const primaryMax = niceCeil(top / 0.62), up = Math.max(0, ...d), down = Math.max(0, ...d.map(v => -v));
    const span = niceCeil(Math.max(up / 0.238, down / 0.08, 1)), zero = 0.72;
    return {primary: {min: 0, max: primaryMax}, secondary: {min: +(-zero * span).toFixed(6), max: +((1 - zero) * span).toFixed(6)}};
  }
  if (variant === 'progress') {
    // 三角标记放在柱顶上方 pad 处，标签在三角之上。
    const top = Math.max(...rows.map(r => Math.max(r[1], r[2]))), max = niceCeil(top * 1.25);
    return {primary: {min: 0, max}, pad: +(max * 0.08).toPrecision(12)};
  }
  const rates = rows.map(r => r[2] / r[1]);
  const half = Math.max(0.05, Math.ceil(Math.max(...rates.map(v => Math.abs(v - 1))) * 1.08 / 0.05 - 1e-9) * 0.05);
  const unit = half > 0.3 ? 0.1 : 0.05;
  const h = Math.ceil(half / unit - 1e-9) * unit;
  return {primary: {min: +(1 - h).toFixed(4), max: +(1 + h).toFixed(4), unit, cross: 1}};
}

/** 纯公式方案。A1、I1:M1、R1 归 core；本模块用 B1（元数据）、C1（标记间距）。 */
export function buildHelperPlan(source, input = {}) {
  const o = resolveOptions(input), info = validate(source.values), n = info.count, last = n + 2;
  if (source.count !== n || !Number.isInteger(source.rowIndex) || source.rowIndex < 0 || !Number.isInteger(source.columnIndex) || source.columnIndex < 0
    || source.rowIndex + n + 1 > 1048576 || source.columnIndex + 3 > 16384 || typeof source.sheet !== 'string' || !source.sheet) throw new UserError('源选区地址或行数无效。');
  const ref = (c, r) => `${quote(source.sheet)}!$${columnLetter(source.columnIndex + c)}$${r}`;
  const num = x => `IF(ISNUMBER(${x}),${x},NA())`;
  const rows = [];
  for (let i = 0; i < n; i++) {
    const r = i + 3, sr = source.rowIndex + i + 2;
    rows.push([
      `=IFERROR(${ref(0, sr)}&"","")`, `=${num(ref(1, sr))}`, `=${num(ref(2, sr))}`,
      `=IFERROR(C${r}-B${r},NA())`, `=IFERROR(IF(B${r}>0,C${r}/B${r},NA()),NA())`,
      `=IFERROR(IF(C${r}>=B${r},C${r},NA()),NA())`, `=IFERROR(IF(C${r}<B${r},C${r},NA()),NA())`,
      `=IFERROR(IF(D${r}>=0,D${r},NA()),NA())`, `=IFERROR(IF(D${r}<0,D${r},NA()),NA())`,
      `=IFERROR(MIN(B${r},C${r}),NA())`, `=IFERROR(IF(B${r}>C${r},B${r}-C${r},NA()),NA())`, `=IFERROR(IF(C${r}>B${r},C${r}-B${r},NA()),NA())`,
      `=IFERROR(IF(E${r}>=1,E${r},NA()),NA())`, `=IFERROR(IF(E${r}<1,E${r},NA()),NA())`,
      `=IFERROR(MAX(B${r},C${r})+$C$1,NA())`, `=IFERROR(TEXT(E${r},"0.0%"),"")`
    ]);
  }
  const scale = scales(source.values.slice(1), o.variant);
  const metadata = {version: 1, variant: o.variant, count: n, seriesCount: SERIES[o.variant].length};
  return {options: o, info, n, last, rows, scale, metadata,
    headers: ['分类', '计划', '实际', '差额', '完成率', '完成', '未完成', '差额正', '差额负', '已完成', '未完成量', '超额', '达标率', '未达标率', '完成率标记Y', '完成率标签'],
    writes: [
      {address: 'C1', formulas: [[String(scale.pad ?? 0)]]},
      {address: 'A2:P2', formulas: null},
      {address: `A3:P${last}`, formulas: rows}
    ]};
}

export async function writeHelper({context, helper, source, options, previous, chart}) {
  const p = buildHelperPlan(source, options);
  p.writes[1].formulas = [p.headers];
  if (previous) {
    const meta = helper.getRange('B1'); meta.load('values'); chart.series.load('items/name'); await context.sync();
    let old; try { old = JSON.parse(meta.values[0][0]); } catch { throw new UserError('达成图实例记录无效，未重绑。'); }
    if (old.variant !== p.options.variant || old.seriesCount !== p.metadata.seriesCount || chart.series.items.length !== p.metadata.seriesCount)
      throw new UserError('原图系列数量或变体已变化，未重绑；请保留原图并另建。');
    if (p.options.variant === 'progress') {
      const s = chart.series.getItemAt(3); s.points.load('items/hasDataLabel'); await context.sync();
      p.labelVisibility = s.points.items.map(x => x.hasDataLabel === true);
    }
    p.previousCount = old.count;
    helper.getRange(`A3:P${Math.max(old.count, p.n) + 2}`).clear('Contents');
  }
  helper.getRange('B1').values = [[JSON.stringify(p.metadata)]];
  for (const w of p.writes) helper.getRange(w.address).formulas = w.formulas;
  helper.getRange(`E3:E${p.last}`).numberFormat = p.rows.map(() => ['0.0%']);
  return p;
}

function bind(chart, helper, p) {
  SERIES[p.options.variant].forEach(([, col], i) => {
    const s = chart.series.getItemAt(i);
    s.setXAxisValues(helper.getRange(`A3:A${p.last}`));
    s.setValues(helper.getRange(`${col}3:${col}${p.last}`));
  });
}

// 完成率三角的标签逐点链接到 P 列文本，改数自动联动。
function bindRateLabels(chart, helperName, p, initial) {
  if (p.options.variant !== 'progress') return;
  const s = chart.series.getItemAt(3);
  for (let i = 0; i < p.n; i++) {
    const existing = !initial && p.labelVisibility && i < p.labelVisibility.length;
    const point = s.points.getItemAt(i);
    if (!existing) {
      point.hasDataLabel = p.options.showLabels;
      if (!p.options.showLabels) continue;
      point.dataLabel.position = 'Top';
      point.dataLabel.format.font.name = FONT; point.dataLabel.format.font.size = 10; point.dataLabel.format.font.color = '#000000';
    } else if (!p.labelVisibility[i]) continue;
    point.dataLabel.formula = `=${quote(helperName)}!$P$${i + 3}`;
  }
}

export async function createChart({context, sourceSheet, helper, source, instance, binding: p}) {
  const layout = SERIES[p.options.variant];
  const chart = sourceSheet.charts.add(layout[0][2], helper.getRange(`${layout[0][1]}3:${layout[0][1]}${p.last}`), 'Columns');
  chart.name = instance.chartName; chart.plotVisibleOnly = false; chart.displayBlanksAs = 'NotPlotted';
  chart.series.load('items'); await context.sync();
  if (chart.series.items.length !== 1) throw new UserError('新图未生成预期的单个初始系列。');
  chart.series.getItemAt(0).name = layout[0][0];
  for (let i = 1; i < layout.length; i++) chart.series.add(layout[i][0]);
  await context.sync();
  layout.forEach(([, , type, axis], i) => { const s = chart.series.getItemAt(i); s.chartType = type; s.axisGroup = axis; });
  await context.sync();
  bind(chart, helper, p);
  const anchor = sourceSheet.getRangeByIndexes(source.rowIndex, source.columnIndex + source.columnCount + 1, 1, 1);
  anchor.load('left,top'); await context.sync();
  chart.left = anchor.left; chart.top = anchor.top; chart.width = LAYOUT.width; chart.height = LAYOUT.height;
  await context.sync();
  return chart;
}

async function stage(context, name, write) {
  try { await write(); await context.sync(); }
  catch (error) { const e = new Error(`[达成图/${name}] ${error.message || String(error)}`); e.name = error.name || 'Error'; e.code = error.code; e.debugInfo = error.debugInfo; e.cause = error; throw e; }
}
function valueLabels(s, show, format, position, color = '#000000') {
  s.hasDataLabels = show; if (!show) return;
  const d = s.dataLabels; d.showValue = true; d.showSeriesName = false; d.showCategoryName = false; d.showLegendKey = false;
  if (position) d.position = position;
  d.numberFormat = format; d.format.font.name = FONT; d.format.font.size = 10; d.format.font.color = color;
}
function fill(s, color, line) {
  s.format.fill.setSolidColor(color);
  if (line) { s.format.line.lineStyle = 'Continuous'; s.format.line.color = line; s.format.line.weight = 0.75; } else s.format.line.lineStyle = 'None';
}

async function styleGap(context, chart, p) {
  const o = p.options;
  await stage(context, '实际柱', () => {
    const a = chart.series.getItemAt(0), b = chart.series.getItemAt(1);
    a.gapWidth = 80; a.overlap = 100; fill(a, COLORS.done); fill(b, COLORS.notDone);
    valueLabels(a, o.showLabels, '0', 'OutsideEnd'); valueLabels(b, o.showLabels, '0', 'OutsideEnd');
  });
  await stage(context, '计划短横线', () => {
    const s = chart.series.getItemAt(2);
    s.format.line.lineStyle = 'None'; s.markerStyle = 'Dash'; s.markerSize = 14;
    s.markerForegroundColor = COLORS.plan; s.markerBackgroundColor = COLORS.plan; s.hasDataLabels = false;
  });
  await stage(context, '差额柱', () => {
    const a = chart.series.getItemAt(3), b = chart.series.getItemAt(4);
    a.gapWidth = 80; a.overlap = 100; fill(a, COLORS.gapUp); fill(b, COLORS.gapDown);
    valueLabels(a, o.showLabels, '0', 'OutsideEnd'); valueLabels(b, o.showLabels, '0', 'OutsideEnd', COLORS.negativeText);
  });
  const x = chart.axes.getItem('Category', 'Primary'), y = chart.axes.getItem('Value', 'Primary'), y2 = chart.axes.getItem('Value', 'Secondary');
  await stage(context, '坐标轴范围', () => {
    y.minimum = p.scale.primary.min; y.maximum = p.scale.primary.max;
    y2.minimum = p.scale.secondary.min; y2.maximum = p.scale.secondary.max;
    for (const a of [y, y2]) { a.tickLabelPosition = 'None'; a.majorGridlines.visible = false; a.majorTickMark = 'None'; a.format.line.lineStyle = 'None'; }
    x.majorGridlines.visible = true; x.majorGridlines.format.line.color = COLORS.grid; x.majorTickMark = 'None'; x.tickLabelPosition = 'Low';
    x.format.line.color = '#D9D9D9';
  });
  return 4; // 隐藏“差额（负）”图例项
}

async function styleProgress(context, chart, p) {
  const o = p.options;
  await stage(context, '堆积柱', () => {
    const a = chart.series.getItemAt(0), b = chart.series.getItemAt(1), c = chart.series.getItemAt(2);
    a.gapWidth = 60; a.overlap = 100;
    fill(a, COLORS.completed); fill(b, COLORS.remaining, COLORS.remainingLine); fill(c, COLORS.excess);
    valueLabels(a, o.showLabels, '0', 'Center', '#FFFFFF'); valueLabels(b, o.showLabels, '0', 'Center'); valueLabels(c, o.showLabels, '0', 'Center', '#FFFFFF');
  });
  await stage(context, '完成率三角', () => {
    const s = chart.series.getItemAt(3);
    s.format.line.lineStyle = 'None'; s.markerStyle = 'Triangle'; s.markerSize = 7;
    s.markerForegroundColor = COLORS.rate; s.markerBackgroundColor = COLORS.rate; s.hasDataLabels = false;
  });
  const x = chart.axes.getItem('Category', 'Primary'), y = chart.axes.getItem('Value', 'Primary');
  await stage(context, '坐标轴范围', () => {
    y.minimum = 0; y.maximum = p.scale.primary.max; y.tickLabelPosition = 'None'; y.majorGridlines.visible = false; y.majorTickMark = 'None'; y.format.line.lineStyle = 'None';
    x.majorTickMark = 'None'; x.format.line.color = '#D9D9D9';
  });
  return -1;
}

async function styleUpdown(context, chart, p) {
  const o = p.options;
  await stage(context, '上下柱', () => {
    const a = chart.series.getItemAt(0), b = chart.series.getItemAt(1);
    a.gapWidth = 60; a.overlap = 100; fill(a, COLORS.above); fill(b, COLORS.below);
    valueLabels(a, o.showLabels, '0%', 'OutsideEnd'); valueLabels(b, o.showLabels, '0%', 'OutsideEnd');
  });
  const x = chart.axes.getItem('Category', 'Primary'), y = chart.axes.getItem('Value', 'Primary');
  await stage(context, '坐标轴与100%交叉', () => {
    y.minimum = p.scale.primary.min; y.maximum = p.scale.primary.max; y.majorUnit = p.scale.primary.unit; y.numberFormat = '0%';
    y.majorGridlines.visible = false; y.majorTickMark = 'None'; y.format.line.lineStyle = 'None';
    y.setPositionAt(p.scale.primary.cross);
    x.tickLabelPosition = 'Low'; x.majorTickMark = 'None'; x.format.line.color = '#BFBFBF';
  });
  return -1;
}

export async function style({context, chart, instance, binding: p}) {
  const o = p.options;
  await stage(context, '标题与图表区', () => {
    chart.format.font.name = FONT; chart.format.font.size = 10; chart.format.font.color = '#000000'; chart.format.fill.setSolidColor('#FFFFFF');
    chart.title.visible = Boolean(o.title); if (o.title) { chart.title.text = o.title; setChartTitleTypography(chart, {text: o.title}); }
    chart.legend.visible = o.showLegend && o.variant !== 'updown'; if (chart.legend.visible) chart.legend.position = 'Top';
  });
  const hidden = o.variant === 'gap' ? await styleGap(context, chart, p) : o.variant === 'progress' ? await styleProgress(context, chart, p) : await styleUpdown(context, chart, p);
  if (o.variant === 'progress') await stage(context, '完成率标签', () => bindRateLabels(chart, instance.helperName, p, true));
  if (hidden >= 0 && chart.legend.visible) {
    const entries = chart.legend.legendEntries; entries.load('items'); await context.sync();
    if (entries.items[hidden]) { entries.items[hidden].visible = false; await context.sync(); }
  }
  return {variant: o.variant, nativeValidation: 'not-run', notes: [
    '坐标范围按创建时数据计算；改数自动联动柱高和标签，若新值超出原范围，请用“应用新版格式”重算范围。',
    o.variant === 'progress' ? '“未完成”用浅红实色+红框代替原文图案填充（Office.js 无图案填充 API）。' : '颜色、字体为景歌原图近似，未取得原工作簿。'
  ]};
}

export async function rebind({context, chart, helper, instance, binding: p}) {
  bind(chart, helper, p); await context.sync();
  if (p.options.variant === 'progress') { bindRateLabels(chart, instance.helperName, p, false); await context.sync(); }
  return {variant: p.options.variant, preservedFormatting: true};
}

export const inputSpec = {
  category: {id: 'comparison', title: '比较'},
  description: '选择含表头的三列：分类、计划（目标）、实际。差额分析：完成/未完成分色实际柱＋计划短横线＋上层差额柱；完成进度：已完成/未完成/超额堆积柱＋完成率三角；上下对比：完成率以100%为基线上下分色。',
  minVersion: '1.9', columns: 3, minRows: 2, maxRows: MAX_ROWS,
  options: [
    {id: 'title', label: '标题', type: 'text', default: '月度达成差额分析图', maxLength: 80},
    {id: 'variant', label: '变体', type: 'select', default: 'gap', choices: [
      {value: 'gap', label: '差额分析（实际柱＋计划线＋差额柱）'},
      {value: 'progress', label: '完成进度（已完成/未完成/超额＋完成率）'},
      {value: 'updown', label: '上下对比（完成率围绕100%）'}]},
    {id: 'showLabels', label: '显示数值标签', type: 'checkbox', default: true},
    {id: 'showLegend', label: '显示图例', type: 'checkbox', default: true}
  ],
  storage: {marker: 'ChartCraft.Achievement.v1', helperPrefix: 'CC_达成_', chartPrefix: 'CC_Achievement_'}
};
export default {id, title, inputSpec, validate, writeHelper, createChart, style, rebind};
