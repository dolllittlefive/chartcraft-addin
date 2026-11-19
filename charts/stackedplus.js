// 景歌图族“堆积柱形/补全堆积”＋“连体面积柱形图”。原生 Chart API 与公式辅助表；不用形状、图片、VBA。
// 变体：total 堆积柱＋顶部合计；complete 补全堆积条形；connected 连体面积柱形（堆积面积＋日期轴）。
import {UserError} from '../core/errors.js';

export const FONT = '微软雅黑';
export const MAX_ROWS = 60, MAX_SERIES = 10;
export const DATE_ORIGIN = 40000; // 日期轴起点；与 profitloss/variablewidthbar 同一约定
export const BAR_WIDTH = 10, TOP_BAND = 0.14; // 连体柱：每柱 10 天宽，中点为整数日；顶部名称带高（归一化）
export const PALETTES = Object.freeze({
  total: ['#20619A', '#6E9ABF', '#D3D3D3', '#F7A5A5', '#8FB9AA', '#F2C46D', '#4A586E', '#B58BB0', '#7ABFC6', '#E07B54'],
  complete: ['#00A0DC', '#5B8FA8', '#E8202A', '#4A586E', '#8FB9AA', '#F2C46D', '#B58BB0', '#7ABFC6', '#E07B54', '#20619A'],
  connected: ['#A06686', '#0E5470', '#F0644A', '#7ABFC6', '#E0B36B', '#4A586E', '#8FB9AA', '#B58BB0', '#20619A', '#E07B54'],
});
export const TRACK_COLOR = '#F2F2F2';
export const VARIANTS = ['total', 'complete', 'connected'];
const FORMATS = ['0', '0.0', '0.00', '0%', '0.0%'];
const quote = s => `'${String(s).replace(/'/g, "''")}'`;
const literal = s => `"${String(s).replace(/"/g, '""')}"`;
const errorText = v => /^#(?:REF!|VALUE!|N\/A|DIV\/0!|NAME\?|NUM!|SPILL!|CALC!|NULL!)/.test(String(v));
export function columnLetter(index) {
  if (!Number.isInteger(index) || index < 0 || index > 16383) throw new UserError('列索引无效。');
  let n = index + 1, s = '';
  while (n) { const d = (n - 1) % 26; s = String.fromCharCode(65 + d) + s; n = Math.floor((n - 1) / 26); }
  return s;
}
export function lighten(hex, amount = 0.6) {
  const n = parseInt(hex.slice(1), 16), c = [n >> 16, (n >> 8) & 255, n & 255].map(v => Math.round(v + (255 - v) * amount));
  return '#' + c.map(v => v.toString(16).padStart(2, '0')).join('').toUpperCase();
}

export function resolveOptions(input = {}) {
  const o = {title: '带合计标签的堆积柱形图', variant: 'total', showValues: true, showTotals: true, showNames: true, valueFormat: '0',
    capacity: 0, gapPercent: 0.1, bridgePercent: 0.8, align: 'center', ...input};
  if (typeof o.title !== 'string' || o.title.length > 160) throw new UserError('标题须为160字以内文本。');
  if (!VARIANTS.includes(o.variant)) throw new UserError('图表变体无效。');
  if (!FORMATS.includes(o.valueFormat)) throw new UserError('数值格式无效。');
  if (!['center', 'bottom'].includes(o.align)) throw new UserError('连体柱对齐方式无效。');
  for (const k of ['showValues', 'showTotals', 'showNames']) if (typeof o[k] !== 'boolean') throw new UserError(`${k}须为开关。`);
  for (const k of ['capacity', 'gapPercent', 'bridgePercent']) o[k] = Number(o[k]);
  if (!Number.isFinite(o.capacity) || o.capacity < 0) throw new UserError('补全上限须为非负数（0 表示取数据最大值）。');
  if (!Number.isFinite(o.gapPercent) || o.gapPercent < 0 || o.gapPercent > 1) throw new UserError('补全条间距须在0–100%。');
  if (!Number.isFinite(o.bridgePercent) || o.bridgePercent < 0.2 || o.bridgePercent > 3) throw new UserError('连体间距须在20–300%柱宽。');
  return o;
}

export function validate(values) {
  if (!Array.isArray(values) || values.length < 2 || values.length > MAX_ROWS + 1) throw new UserError(`请选择含表头的数据：第一列为分类，其后1–${MAX_SERIES}列为堆积分段，1–${MAX_ROWS}行数据。`);
  const width = values[0]?.length;
  if (!Number.isInteger(width) || width < 2 || width > MAX_SERIES + 1 || values.some(r => !Array.isArray(r) || r.length !== width)) throw new UserError(`需要规则的2–${MAX_SERIES + 1}列表格：分类＋各分段数值。`);
  if (values[0].some(v => !['string', 'number'].includes(typeof v) || !String(v).trim() || String(v).length > 255 || errorText(v))) throw new UserError('每列需要非空表头。');
  values.slice(1).forEach((r, i) => {
    if (!['string', 'number'].includes(typeof r[0]) || !String(r[0]).trim() || errorText(r[0]) || typeof r[0] === 'number' && !Number.isFinite(r[0])) throw new UserError(`第${i + 2}行分类为空或无效。`);
    r.slice(1).forEach((v, j) => {
      if (v === '') return; // 景歌模板允许留空系列；按0处理
      if (typeof v !== 'number' || !Number.isFinite(v) || v < 0) throw new UserError(`第${i + 2}行第${j + 2}列须为非负数值或留空。`);
    });
  });
  return {count: values.length - 1, segments: width - 1, headers: values[0].map(v => String(v))};
}

function checkSource(source, o) {
  const info = validate(source.values);
  if (info.count !== source.count || !Number.isInteger(source.rowIndex) || source.rowIndex < 0 || !Number.isInteger(source.columnIndex) || source.columnIndex < 0
      || source.rowIndex + info.count + 1 > 1048576 || source.columnIndex + info.segments + 1 > 16384 || typeof source.sheet !== 'string' || !source.sheet) throw new UserError('源选区地址或行数无效。');
  if (o.variant === 'connected' && info.count < 2) throw new UserError('连体面积柱形至少需要2行（2根柱）。');
  return info;
}

// 公式计划：纯函数，不触碰 Office。writes 只写第2行及以下与 B1:F1（不碰 core 的 A1、I1:M1、R1）。
export function buildPlan(source, input = {}) {
  const o = resolveOptions(input), info = checkSource(source, o), n = info.count, m = info.segments;
  const ref = (c, r) => `${quote(source.sheet)}!$${columnLetter(source.columnIndex + c)}$${r}`;
  const srcRow = i => source.rowIndex + i + 2, headRow = source.rowIndex + 1;
  const val = (i, j) => `MAX(0,N(${ref(j + 1, srcRow(i))}))`;
  const cat = i => `=IFERROR(${ref(0, srcRow(i))}&"","")`;
  const head = j => `=IFERROR(${ref(j + 1, headRow)}&"","")`;
  const fmt = literal(o.valueFormat);
  const plan = {o, info, n, m, variant: o.variant, writes: []};
  if (o.variant === 'total') {
    const last = n + 2, lastCol = columnLetter(1 + m);
    plan.last = last; plan.seriesCount = m + 1;
    plan.writes.push({address: `A2:${lastCol}2`, formulas: [['="分类"', '="合计"', ...Array.from({length: m}, (_, j) => head(j))]]});
    plan.writes.push({address: `A3:${lastCol}${last}`, formulas: Array.from({length: n}, (_, i) => [cat(i), `=SUM(C${i + 3}:${lastCol}${i + 3})`, ...Array.from({length: m}, (_, j) => `=IFERROR(${val(i, j)},0)`)])});
    plan.series = [...Array.from({length: m}, (_, j) => ({name: j, x: `A3:A${last}`, y: `${columnLetter(2 + j)}3:${columnLetter(2 + j)}${last}`})), {name: '合计', x: `A3:A${last}`, y: `B3:B${last}`}];
    plan.axisMax = null;
  } else if (o.variant === 'complete') {
    const last = n + 2, textBase = 1 + 3 * m, lastCol = columnLetter(textBase + m - 1);
    const range = `${ref(1, srcRow(0))}:${ref(m, srcRow(n - 1)).split('!')[1]}`;
    plan.last = last; plan.seriesCount = 3 * m - 1; plan.textBase = textBase;
    plan.axisMax = m + (m - 1) * o.gapPercent;
    plan.writes.push({address: 'C1:D1', formulas: [[`=IF(${o.capacity}>0,${o.capacity},MAX(${range}))`, `=${o.gapPercent}`]]});
    const header = ['="分类"'];
    for (let j = 0; j < m; j++) header.push(head(j), `=${columnLetter(2 + 3 * j - 1)}2&"补全"`, '="间隔"');
    for (let j = 0; j < m; j++) header.push(`=${columnLetter(1 + 3 * j)}2&"标签"`);
    plan.writes.push({address: `A2:${lastCol}2`, formulas: [header]});
    plan.writes.push({address: `A3:${lastCol}${last}`, formulas: Array.from({length: n}, (_, i) => {
      const r = i + 3, row = [cat(i)];
      for (let j = 0; j < m; j++) { const v = columnLetter(1 + 3 * j); row.push(`=IFERROR(${val(i, j)}/IF($C$1>0,$C$1,1),0)`, `=MAX(0,1-${v}${r})`, j < m - 1 ? '=$D$1' : '=0'); }
      for (let j = 0; j < m; j++) row.push(`=IFERROR(TEXT(${val(i, j)},${fmt}),"")`);
      return row;
    })});
    plan.series = [];
    for (let j = 0; j < m; j++) for (let k = 0; k < (j < m - 1 ? 3 : 2); k++) { const c = columnLetter(1 + 3 * j + k); plan.series.push({name: k === 0 ? j : k === 1 ? `${info.headers[j + 1]}补全` : '间隔', role: ['value', 'track', 'gap'][k], segment: j, x: `A3:A${last}`, y: `${c}3:${c}${last}`}); }
  } else {
    // 连体：period 块在 BA 列起：BA 名称、BB 合计、BC 偏移(归一)、BD.. 原值 m 列、其后归一值 m 列
    const P0 = 52, colName = columnLetter(P0), colTotal = columnLetter(P0 + 1), colOff = columnLetter(P0 + 2);
    const rawCol = j => columnLetter(P0 + 3 + j), normCol = j => columnLetter(P0 + 3 + m + j), pLast = n + 2;
    const pLastCol = columnLetter(P0 + 2 + 2 * m);
    plan.writes.push({address: 'C1', formulas: [[`=MAX(${colTotal}3:${colTotal}${pLast})`]]});
    plan.writes.push({address: `${colName}2:${pLastCol}2`, formulas: [['="期间"', '="合计"', '="偏移"', ...Array.from({length: m}, (_, j) => head(j)), ...Array.from({length: m}, (_, j) => `=${rawCol(j)}2&"归一"`)]]});
    plan.writes.push({address: `${colName}3:${pLastCol}${pLast}`, formulas: Array.from({length: n}, (_, i) => {
      const r = i + 3;
      return [cat(i), `=SUM(${rawCol(0)}${r}:${rawCol(m - 1)}${r})`, o.align === 'center' ? `=IF($C$1>0,($C$1-${colTotal}${r})/2/$C$1,0)` : '=0',
        ...Array.from({length: m}, (_, j) => `=IFERROR(${val(i, j)},0)`), ...Array.from({length: m}, (_, j) => `=IF($C$1>0,${rawCol(j)}${r}/$C$1,0)`)];
    })});
    // 点行：每柱 3 点（左、中、右），柱间 2 点（左柱右缘、右柱左缘）
    const G = Math.max(1, Math.round(BAR_WIDTH * o.bridgePercent)), pitch = BAR_WIDTH + G, points = [];
    for (let i = 0; i < n; i++) {
      for (const dx of [0, BAR_WIDTH / 2, BAR_WIDTH]) points.push({kind: 'bar', period: i, x: i * pitch + dx, mid: dx === BAR_WIDTH / 2});
      if (i < n - 1) points.push({kind: 'bridge', period: i, x: i * pitch + BAR_WIDTH}, {kind: 'bridge', period: i + 1, x: (i + 1) * pitch});
    }
    // 列：A x；B 柱偏移；C 桥偏移；D.. 柱 m；其后桥 m；顶部带 1；柱标签文本 m；顶部文本 1
    const barCol = j => columnLetter(3 + j), bridgeCol = j => columnLetter(3 + m + j), topCol = columnLetter(3 + 2 * m);
    const textCol = j => columnLetter(4 + 2 * m + j), topText = columnLetter(4 + 3 * m), last = points.length + 2;
    const header = ['="X(日期轴)"', '="柱偏移"', '="桥偏移"', ...Array.from({length: m}, (_, j) => `=${rawCol(j)}2`), ...Array.from({length: m}, (_, j) => `=${rawCol(j)}2&"连接"`), '="顶部"', ...Array.from({length: m}, (_, j) => `=${rawCol(j)}2&"标签"`), '="顶部标签"'];
    plan.writes.push({address: `A2:${topText}2`, formulas: [header]});
    plan.writes.push({address: `A3:${topText}${last}`, formulas: points.map(p => {
      const pr = p.period + 3, bar = p.kind === 'bar', row = [`=${DATE_ORIGIN + p.x}`, bar ? `=${colOff}${pr}` : '=0', bar ? '=0' : `=${colOff}${pr}`];
      for (let j = 0; j < m; j++) row.push(bar ? `=${normCol(j)}${pr}` : '=0');
      for (let j = 0; j < m; j++) row.push(bar ? '=0' : `=${normCol(j)}${pr}`);
      row.push(bar ? `=${TOP_BAND}` : '=0');
      for (let j = 0; j < m; j++) row.push(!p.mid ? '=""' : o.showNames && p.period === 0 ? `=IFERROR(${rawCol(j)}$2&"  "&TEXT(${rawCol(j)}${pr},${fmt}),"")` : `=IFERROR(TEXT(${rawCol(j)}${pr},${fmt}),"")`);
      row.push(!p.mid ? '=""' : o.showTotals ? `=${colName}${pr}&"  "&TEXT(${colTotal}${pr},${fmt})` : `=${colName}${pr}&""`);
      return row;
    })});
    plan.points = points; plan.last = last; plan.pLast = pLast; plan.lastX = points.at(-1).x; plan.pitch = pitch;
    plan.midRows = points.map((p, k) => p.mid ? k : -1).filter(k => k >= 0);
    plan.textCol = textCol; plan.topText = topText;
    plan.axisMax = 1 + TOP_BAND; plan.seriesCount = 3 + 2 * m;
    const xr = `A3:A${last}`;
    plan.series = [{name: '柱偏移', role: 'offset', x: xr, y: `B3:B${last}`}, {name: '桥偏移', role: 'offset', x: xr, y: `C3:C${last}`}];
    for (let j = 0; j < m; j++) plan.series.push({name: j, role: 'bar', segment: j, x: xr, y: `${barCol(j)}3:${barCol(j)}${last}`}, {name: `${info.headers[j + 1]}连接`, role: 'bridge', segment: j, x: xr, y: `${bridgeCol(j)}3:${bridgeCol(j)}${last}`});
    plan.series.push({name: '顶部标签', role: 'top', x: xr, y: `${topCol}3:${topCol}${last}`});
  }
  plan.metadata = {version: 1, variant: o.variant, segments: m, count: n, seriesCount: plan.seriesCount, rows: (plan.last || n + 2)};
  plan.writes.unshift({address: 'B1', formulas: [[JSON.stringify(plan.metadata)]], values: true});
  return plan;
}

export function chartType(variant) { return {total: 'ColumnStacked', complete: 'BarStacked', connected: 'AreaStacked'}[variant]; }
export function seriesTypes(plan) { return plan.series.map((s, i) => plan.variant === 'total' && i === plan.m ? 'Line' : chartType(plan.variant)); }
export function seriesColor(plan, s) {
  const pal = PALETTES[plan.variant], c = Number.isInteger(s.name) ? pal[s.name % pal.length] : s.segment !== undefined ? pal[s.segment % pal.length] : null;
  if (plan.variant === 'complete') return s.role === 'value' ? c : s.role === 'track' ? TRACK_COLOR : null;
  if (plan.variant === 'connected') return s.role === 'bar' ? c : s.role === 'bridge' ? lighten(c, 0.6) : null;
  return Number.isInteger(s.name) ? c : null;
}
export function layoutFor(plan) {
  if (plan.variant === 'complete') return {width: 560, height: Math.min(1600, Math.max(300, 90 + plan.n * 30))};
  if (plan.variant === 'connected') return {width: 560, height: 420};
  return {width: 560, height: 380};
}

async function writeHelper({context, helper, source, options, previous, chart}) {
  const p = buildPlan(source, options);
  if (previous) {
    const meta = helper.getRange('B1'); meta.load('values');
    chart.series.load('items/chartType'); await context.sync();
    let old; try { old = JSON.parse(meta.values[0][0]); } catch { throw new UserError('堆积增强图原实例记录损坏，未重绑。'); }
    if (old.variant !== p.variant || old.segments !== p.m) throw new UserError('分段列数或变体已改变；按选区更新只支持增删行，请另建图。');
    const types = seriesTypes(p);
    if (chart.series.items.length !== p.seriesCount || chart.series.items.some((s, i) => s.chartType !== types[i])) throw new UserError('原图系列数量或类型已被手动改动，未重绑。');
    p.previous = old;
    helper.getRange(`A2:${columnLetter(52 + 2 + 2 * p.m)}${Math.max(old.rows, p.last || 0, old.count + 2, p.n + 2)}`).clear('Contents');
  }
  for (const w of p.writes) { const r = helper.getRange(w.address); if (w.values) r.values = w.formulas; else r.formulas = w.formulas; }
  return p;
}

function bindSeries(chart, helper, p) {
  p.series.forEach((s, i) => { const series = chart.series.getItemAt(i); series.setXAxisValues(helper.getRange(s.x)); series.setValues(helper.getRange(s.y)); });
}
// 逐点标签计划（三个变体都只用点级 dataLabel；Mac 拒绝系列级 ChartDataLabels.position）。
// cell 有值＝链接辅助表文本；否则显示数值＋numberFormat。首次创建或新出现的点才写默认格式与位置。
export function labelPlan(p, {initial, previousCount = 0}) {
  const out = [], add = (series, point, cell, position, color, isNew, extra = {}) => out.push({series, point, cell, position, color, fresh: initial || isNew, ...extra});
  if (p.variant === 'total') {
    if (p.o.showValues) for (let j = 0; j < p.m; j++) for (let i = 0; i < p.n; i++) add(j, i, null, 'Center', j === 2 ? '#404040' : '#FFFFFF', i >= previousCount);
    if (p.o.showTotals) for (let i = 0; i < p.n; i++) add(p.m, i, null, 'Top', '#404040', i >= previousCount, {bold: true, border: '#BFBFBF'});
  } else if (p.variant === 'complete') {
    if (p.o.showValues) p.series.forEach((s, si) => { if (s.role === 'track') for (let i = 0; i < p.n; i++) add(si, i, `${columnLetter(p.textBase + s.segment)}$${i + 3}`, 'InsideBase', '#404040', i >= previousCount); });
  } else {
    // 面积图类型没有标签位置选项：不设 position
    p.series.forEach((s, si) => {
      if (s.role === 'bar' && p.o.showValues) p.midRows.forEach((row, k) => add(si, row, `${p.textCol(s.segment)}$${row + 3}`, null, '#FFFFFF', k >= previousCount));
      if (s.role === 'top') p.midRows.forEach((row, k) => add(si, row, `${p.topText}$${row + 3}`, null, '#595959', k >= previousCount));
    });
  }
  return out;
}
async function linkLabels(context, chart, helperName, p, opts) {
  const plan = labelPlan(p, opts), point = l => chart.series.getItemAt(l.series).points.getItemAt(l.point);
  for (const l of plan) {
    const pt = point(l); pt.hasDataLabel = true;
    if (l.cell) pt.dataLabel.formula = `=${quote(helperName)}!$${l.cell}`;
    else { pt.dataLabel.showValue = true; pt.dataLabel.showSeriesName = false; pt.dataLabel.showCategoryName = false; if (l.fresh) pt.dataLabel.numberFormat = p.o.valueFormat; }
    if (l.fresh) { const f = pt.dataLabel.format.font; f.name = FONT; f.size = 10; f.color = l.color; if (l.bold) f.bold = true; }
  }
  await context.sync();
  // 位置与边框各自单独提交：宿主若拒绝，保留默认位置/无边框并记入说明，不让整张图创建失败。
  const issues = [];
  for (const key of ['position', 'border']) {
    const todo = plan.filter(l => l.fresh && l[key]);
    if (!todo.length) continue;
    for (const l of todo) {
      if (key === 'position') point(l).dataLabel.position = l.position;
      else { const b = point(l).dataLabel.format.border; b.color = l.border; b.weight = 0.75; }
    }
    try { await context.sync(); } catch (e) { issues.push(`标签${key === 'position' ? '位置' : '边框'}未被宿主接受，保留默认：${e?.message || e}`); }
  }
  return issues;
}

async function createChart({context, sourceSheet, helper, instance, binding: p}) {
  const first = p.series[0];
  const chart = sourceSheet.charts.add(chartType(p.variant), helper.getRange(first.y), 'Columns');
  chart.name = instance.chartName; chart.plotVisibleOnly = false;
  chart.series.load('items'); await context.sync();
  if (chart.series.items.length !== 1) throw new UserError('堆积增强图初始系列数量异常。');
  for (let i = 1; i < p.seriesCount; i++) chart.series.add(`系列${i + 1}`);
  await context.sync();
  const types = seriesTypes(p);
  p.series.forEach((s, i) => {
    const series = chart.series.getItemAt(i);
    if (types[i] !== chartType(p.variant)) series.chartType = types[i];
    series.name = Number.isInteger(s.name) ? `=${quote(instance.helperName)}!$${p.variant === 'total' ? columnLetter(2 + s.name) : p.variant === 'complete' ? columnLetter(1 + 3 * s.name) : columnLetter(3 + s.name)}$2` : s.name;
  });
  bindSeries(chart, helper, p);
  const size = layoutFor(p); chart.width = size.width; chart.height = size.height;
  await context.sync();
  return chart;
}

function hideAxis(axis) { axis.tickLabelPosition = 'None'; axis.majorTickMark = 'None'; axis.minorTickMark = 'None'; axis.majorGridlines.visible = false; axis.format.line.lineStyle = 'None'; }

async function style({context, chart, instance, binding: p}) {
  const o = p.o;
  chart.format.fill.setSolidColor('#FFFFFF'); chart.format.font.name = FONT; chart.format.font.size = 10; chart.format.font.color = '#000000';
  chart.title.visible = Boolean(o.title); chart.title.text = o.title; chart.title.horizontalAlignment = 'Left';
  chart.title.format.font.name = FONT; chart.title.format.font.size = 14; chart.title.format.font.bold = false;
  const types = seriesTypes(p);
  p.series.forEach((s, i) => {
    const series = chart.series.getItemAt(i), color = seriesColor(p, s);
    if (types[i] === 'Line') { series.format.line.lineStyle = 'None'; series.markerStyle = 'None'; return; }
    if (color) series.format.fill.setSolidColor(color); else series.format.fill.clear();
    if (p.variant === 'connected') { series.format.line.lineStyle = s.role === 'bar' || s.role === 'bridge' ? 'Continuous' : 'None'; if (s.role === 'bar' || s.role === 'bridge') { series.format.line.color = '#FFFFFF'; series.format.line.weight = 0.75; } }
    else { series.format.line.lineStyle = 'None'; series.gapWidth = p.variant === 'complete' ? 60 : 80; series.overlap = 100; }
  });
  await context.sync();
  const x = chart.axes.categoryAxis, y = chart.axes.valueAxis;
  y.majorGridlines.visible = false;
  if (p.variant === 'total') {
    x.majorTickMark = 'None'; x.format.line.color = '#D9D9D9'; hideAxis(y); y.minimum = 0;
    chart.legend.visible = true; chart.legend.position = 'Top';
  } else if (p.variant === 'complete') {
    x.reversePlotOrder = true; x.majorTickMark = 'None'; x.format.line.color = '#D9D9D9';
    hideAxis(y); y.minimum = 0; y.maximum = p.axisMax;
    chart.legend.visible = true; chart.legend.position = 'Top';
  } else {
    x.categoryType = 'DateAxis'; x.baseTimeUnit = 'Days'; x.minimum = DATE_ORIGIN; x.maximum = DATE_ORIGIN + p.lastX; x.isBetweenCategories = false;
    hideAxis(x); hideAxis(y); y.minimum = 0; y.maximum = p.axisMax;
    chart.legend.visible = false;
  }
  await context.sync();
  if (chart.legend.visible) await hideAuxLegend(context, chart, p);
  const issues = await linkLabels(context, chart, instance.helperName, p, {initial: true});
  return {nativeValidation: 'not-run', variant: p.variant, notes: [...notesFor(p), ...issues]};
}
async function hideAuxLegend(context, chart, p) {
  // legendEntries 为 ExcelApi 1.13，与 clusterstack/barline 同样只做显隐。
  const entries = chart.legend.legendEntries; entries.load('items'); await context.sync();
  entries.items.forEach((e, i) => { const s = p.series[i]; e.visible = Boolean(s) && (p.variant === 'total' ? i < p.m : s.role === 'value'); });
  await context.sync();
}
function notesFor(p) {
  return {
    total: ['顶部合计＝无线无标记折线系列的“上方”标签；原图圆角标注框以浅灰直角边框代替（ExcelApi 1.9 无标签形状类型）。'],
    complete: ['每系列＝数值＋补全（浅灰轨道）＋间隔（无填充）；值归一化到补全上限，轴范围固定为 m+(m-1)×间距。', '数值标签挂在补全段内侧起点，单元格公式驱动。'],
    connected: ['堆积面积＋日期轴：每柱3点、柱间2点，重复X形成竖直边；底部无填充偏移系列实现居中。', '分段名显示在第一根柱的标签内（原图在柱左侧另列）；面积图标签位置需实机确认。'],
  }[p.variant];
}
async function rebind({context, chart, helper, instance, previous, binding: p}) {
  bindSeries(chart, helper, p);
  if (p.variant === 'connected') { chart.axes.categoryAxis.maximum = DATE_ORIGIN + p.lastX; }
  await context.sync();
  const prevCount = previous?.count ?? 0;
  const issues = await linkLabels(context, chart, instance.helperName, p, {initial: false, previousCount: prevCount});
  return {notes: ['只更新公式、系列引用、标签链接；保留已有颜色与格式。', ...issues]};
}
// core 的统一字体可能把标签刷成黑色；按标签计划逐点恢复白字/深灰字。仅创建/刷新格式时调用。
async function finalizePresentation({context, chart, binding: p}) {
  for (const l of labelPlan(p, {initial: true})) chart.series.getItemAt(l.series).points.getItemAt(l.point).dataLabel.format.font.color = l.color;
  await context.sync();
}

export const inputSpec = {
  category: {id: 'composition', title: '构成'}, minVersion: '1.9', columns: 5, minRows: 1, maxRows: MAX_ROWS,
  description: '第一列为分类（连体柱为期间，如年份），其后1–10列为堆积分段数值（非负，可留空）。变体：堆积柱＋顶部合计、补全堆积条形、连体面积柱形；可按选区增删行。',
  options: [
    {id: 'title', label: '标题', type: 'text', default: '带合计标签的堆积柱形图', maxLength: 160},
    {id: 'variant', label: '变体', type: 'select', default: 'total', choices: [
      {value: 'total', label: '堆积柱＋顶部合计'}, {value: 'complete', label: '补全堆积条形'}, {value: 'connected', label: '连体面积柱形'}]},
    {id: 'showValues', label: '显示分段数值', type: 'checkbox', default: true},
    {id: 'showTotals', label: '显示合计（连体柱附在期间名后）', type: 'checkbox', default: true},
    {id: 'showNames', label: '连体柱：首柱显示分段名', type: 'checkbox', default: true},
    {id: 'valueFormat', label: '数值格式', type: 'select', default: '0', choices: FORMATS.map(value => ({value, label: value}))},
    {id: 'capacity', label: '补全上限（0＝取数据最大值）', type: 'number', default: 0, min: 0, max: 1e12, step: 1},
    {id: 'gapPercent', label: '补全条间距', type: 'number', default: 10, min: 0, max: 100, step: 1, scale: 0.01, suffix: '%'},
    {id: 'bridgePercent', label: '连体间距（柱宽）', type: 'number', default: 80, min: 20, max: 300, step: 10, scale: 0.01, suffix: '%'},
    {id: 'align', label: '连体柱对齐', type: 'select', default: 'center', choices: [{value: 'center', label: '垂直居中'}, {value: 'bottom', label: '底部对齐'}]},
  ],
  storage: {marker: 'ChartCraft.StackedPlus.v1', helperPrefix: 'CC_堆积增强_', chartPrefix: 'CC_StackedPlus_'},
};
export default {id: 'stackedplus', title: '堆积增强图', inputSpec, validate, writeHelper, createChart, style, rebind, finalizePresentation};
