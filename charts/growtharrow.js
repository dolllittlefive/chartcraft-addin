// 景歌图族“涨跌箭头/增幅柱形”+“同环比/多年对比”。Native Excel series only:
// clustered columns + XY polylines / masked XY label carriers / secondary columns.
// No error bars, shapes, pictures or VBA. Every geometry value is a helper formula,
// so in-place source edits recalculate without the plugin.
import {UserError} from '../core/errors.js';

export const FONT = 'Microsoft YaHei';
export const MAX_ROWS = 60;
export const LAYOUT = {width:480, height:360};
// Bars and XY overlays share one normalised scale: value / max(all period values).
export const GEOMETRY = {
  arrow:{gap:80, top:1.3},
  outline:{gap:60, top:1.3, bracket:1.1, notch:0.06},
  delta:{gap:80, top:1.75, secondaryMin:-6.2, secondaryMax:1.3}
};
export const COLORS = {
  arrow:{bars:['#1170AA','#33BFD6'], line:'#595959', up:'#404040', down:'#E00000'},
  outline2:{bars:['#0A1C2C','#40A1ED'], line:'#A6A6A6', up:'#0E9FD8', down:'#E00000', barLabels:['#FFFFFF','#FFFFFF']},
  outline3:{bars:['#AEE6F7','#14B4E8','#0A4A63'], line:'#8497B0', up:'#0E9FD8', down:'#E00000', barLabels:['#262626','#FFFFFF','#FFFFFF']},
  delta:{bars:['#1B75A8','#3CC0D4','#8ADDE6'], up:'#76C8C0', down:'#D97373', changes:['#7DCB6C','#F39200'], zero:'#BFBFBF'},
  text:'#404040'
};
const ERROR_TEXT = /^#(REF!|VALUE!|N\/A|DIV\/0!|NAME\?|NUM!|SPILL!|CALC!|NULL!)/;
const quote = s => `'${String(s).replace(/'/g,"''")}'`;
const choose = (v, values, label) => { if (!values.includes(v)) throw new UserError(`${label}无效。`); return v; };

export function columnLetter(index) {
  if (!Number.isInteger(index) || index < 0 || index > 16383) throw new UserError('列索引无效。');
  let s = '', n = index + 1; while (n) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s;
}

export function resolveOptions(input = {}) {
  const o = {title:'同比增幅对比', variant:'arrow', metric:'auto', showLabels:true, ...input};
  if (typeof o.title !== 'string' || o.title.length > 80) throw new UserError('标题应为不超过 80 个字符的文本。');
  choose(o.variant, ['arrow','outline','delta'], '图表变体');
  choose(o.metric, ['auto','growth','difference'], '变化指标');
  if (typeof o.showLabels !== 'boolean') throw new UserError('显示标签必须是布尔值。');
  o.resolvedMetric = o.metric === 'auto' ? (o.variant === 'delta' ? 'difference' : 'growth') : o.metric;
  return o;
}

export function validate(values) {
  if (!Array.isArray(values) || values.length < 2 || values.length > MAX_ROWS + 1) throw new UserError(`选择含表头的 3 或 4 列，包含 1–${MAX_ROWS} 行数据。`);
  const width = values[0]?.length;
  if (![3,4].includes(width) || values.some(r => !Array.isArray(r) || r.length !== width)) throw new UserError('数据须是规则的 3 列（分类、基期、本期）或 4 列（分类、三个年度）。');
  if (values[0].some(v => typeof v !== 'string' || !v.trim() || v.length > 255 || ERROR_TEXT.test(v))) throw new UserError('每列必须有非空文本表头（最多 255 字符）。');
  values.slice(1).forEach((r, i) => {
    if (!['string','number'].includes(typeof r[0]) || String(r[0]).trim() === '' || (typeof r[0] === 'number' && !Number.isFinite(r[0])) || ERROR_TEXT.test(String(r[0]))) throw new UserError(`第 ${i+2} 行分类为空或无效。`);
    r.slice(1).forEach((v, j) => {
      if (typeof v !== 'number' || !Number.isFinite(v)) throw new UserError(`第 ${i+2} 行第 ${j+2} 列必须是真实数值；不接受文本数字或空白。`);
      if (v < 0) throw new UserError(`第 ${i+2} 行第 ${j+2} 列为负数；涨跌柱按“值/最大值”同比例缩放，要求各期数值非负。`);
    });
  });
  return {count:values.length - 1, periods:width - 1, categoryHeader:values[0][0], periodHeaders:values[0].slice(1)};
}

function validateSource(source, o) {
  const info = validate(source.values);
  if (info.count !== source.count) throw new UserError('选区行数与校验结果不一致。');
  if (!Number.isInteger(source.rowIndex) || source.rowIndex < 0 || source.rowIndex + source.count + 1 > 1048576 || !Number.isInteger(source.columnIndex) || source.columnIndex < 0 || source.columnIndex + source.values[0].length > 16384 || typeof source.sheet !== 'string' || !source.sheet) throw new UserError('源区域地址无效。');
  if (o.variant === 'arrow' && info.periods !== 2) throw new UserError('涨跌箭头柱只比较两期：请选 3 列（分类、基期、本期）。三年数据请用“轮廓线增幅”或“增减量柱”。');
  if (o.resolvedMetric === 'growth' && source.values.slice(1).some(r => r.slice(1, -1).some(v => v === 0))) throw new UserError('计算增长率时，前一期不能为 0；请改用“差值”指标。');
  if (source.values.slice(1).every(r => r.slice(1).every(v => v === 0))) throw new UserError('各期数值全为 0，无法确定比例。');
  return info;
}

const allIntegers = rows => rows.every(r => r.slice(1).every(v => Number.isInteger(v)));
/** Formats are chosen from the source data at plan time and reapplied on rebind. */
export function formats(o, values) {
  const integer = allIntegers(values.slice(1));
  const value = integer ? '0' : '0.0';
  if (o.resolvedMetric === 'growth') return {value, abs:'0.0%', signed:'+0.0%;-0.0%;0.0%'};
  return {value, abs:integer ? '0' : '0.0', signed:integer ? '+0;-0;0' : '+0.0;-0.0;0.0'};
}

function layoutColumns(o, k) {
  const m = k - 1, names = ['cat'];
  for (let j = 0; j < k; j++) names.push(`raw${j}`);
  for (let j = 0; j < k; j++) names.push(`norm${j}`);
  for (let j = 0; j < m; j++) names.push(`chg${j}`);
  for (let j = 0; j < m; j++) names.push(`lab${j}`);
  for (let j = 0; j < k; j++) names.push(`x${j}`);
  if (o.variant === 'arrow') names.push('hi','lo','xLow','xHigh','mid','headUp','headDown');
  if (o.variant === 'delta') for (let j = 0; j < (k === 2 ? 2 : m); j++) names.push(`bar2_${j}`);
  if (o.variant !== 'delta') names.push('lineX','lineY','tagX','tagUp','tagDown','tagText');
  return Object.fromEntries(names.map((name, i) => [name, columnLetter(i)]));
}
const POINTS_PER_CATEGORY = {arrow:() => 3, outline:k => (k === 2 ? 4 : 8), delta:() => 0};

/** Pure formula plan. A1, I1:M1 and R1 remain core-owned; B1:H1 and N1 are ours. */
export function buildHelperPlan(source, input = {}) {
  const o = resolveOptions(input), info = validateSource(source, o), n = source.count, k = info.periods, m = k - 1, last = n + 2;
  const C = layoutColumns(o, k), g = GEOMETRY[o.variant], f = formats(o, source.values);
  const src = (offset, row) => `${quote(source.sheet)}!$${columnLetter(source.columnIndex + offset)}$${row}`;
  const rawRange = `$${C.raw0}$3:$${C[`raw${k-1}`]}$${last}`, chgRange = `$${C.chg0}$3:$${C[`chg${m-1}`]}$${last}`;
  const scalars = {
    C1:String(n),
    D1:`=IF(MAX(${rawRange})>0,MAX(${rawRange}),1)`,
    E1:`=1/(${k}+${g.gap}/100)/$C$1`,
    F1:`=IF(MAX(MAX(${chgRange}),-MIN(${chgRange}))>0,MAX(MAX(${chgRange}),-MIN(${chgRange})),1)`,
    G1:String(g.bracket ?? g.top),
    H1:`=$E$1*${g.notch ?? 0}`,
    N1:String(k)
  };
  const cells = {}, put = (col, row, formula) => { cells[`${col}${row}`] = formula; };
  const ppc = POINTS_PER_CATEGORY[o.variant](k);
  const tagCount = o.variant === 'delta' ? 0 : n * m;
  for (let i = 0; i < n; i++) {
    const r = i + 3, sr = source.rowIndex + i + 2, ref = name => `${C[name]}${r}`;
    put(C.cat, r, `=IFERROR(${src(0, sr)}&"","")`);
    for (let j = 0; j < k; j++) {
      const s = src(j + 1, sr);
      put(C[`raw${j}`], r, `=IF(ISNUMBER(${s}),${s},"")`);
      put(C[`norm${j}`], r, `=IF(ISNUMBER(${ref(`raw${j}`)}),${ref(`raw${j}`)}/$D$1,NA())`);
      put(C[`x${j}`], r, `=(${i}+0.5)/$C$1+(${j - (k - 1) / 2})*$E$1`);
    }
    for (let j = 0; j < m; j++) {
      const a = ref(`raw${j}`), b = ref(`raw${j+1}`), chg = ref(`chg${j}`);
      put(C[`chg${j}`], r, o.resolvedMetric === 'growth'
        ? `=IF(AND(ISNUMBER(${a}),ISNUMBER(${b})),IF(${a}<>0,${b}/${a}-1,""),"")`
        : `=IF(AND(ISNUMBER(${a}),ISNUMBER(${b})),${b}-${a},"")`);
      const text = o.variant === 'delta' ? `TEXT(${chg},"${f.signed}")`
        : o.variant === 'arrow' ? `IF(${chg}>0,"增长 ",IF(${chg}<0,"下降 ","持平 "))&TEXT(ABS(${chg}),"${f.abs}")`
        : `IF(${chg}>0,"↗ ",IF(${chg}<0,"↘ -","→ "))&TEXT(ABS(${chg}),"${f.abs}")`;
      put(C[`lab${j}`], r, `=IF(ISNUMBER(${chg}),${text},"")`);
    }
    if (o.variant === 'arrow') {
      const n0 = ref('norm0'), n1 = ref('norm1'), valid = `AND(ISNUMBER(${n0}),ISNUMBER(${n1}))`;
      put(C.hi, r, `=IF(${valid},MAX(${n0},${n1}),NA())`);
      put(C.lo, r, `=IF(${valid},MIN(${n0},${n1}),NA())`);
      // The arrow stands on the lower bar; the cap runs from the higher bar's centre.
      put(C.xLow, r, `=IF(${valid},IF(${n0}<=${n1},${ref('x0')},${ref('x1')}),NA())`);
      put(C.xHigh, r, `=IF(${valid},IF(${n0}<=${n1},${ref('x1')},${ref('x0')}),NA())`);
      put(C.mid, r, `=(${ref('x0')}+${ref('x1')})/2`);
      put(C.headUp, r, `=IF(${valid},IF(${n1}>${n0},${ref('hi')},NA()),NA())`);
      put(C.headDown, r, `=IF(${valid},IF(${n1}<${n0},${ref('lo')},NA()),NA())`);
      const p = 3 + 3 * i;
      [[`=${ref('xHigh')}`, `=${ref('hi')}`], [`=${ref('xLow')}`, `=${ref('hi')}`], [`=${ref('xLow')}`, `=${ref('lo')}`]]
        .forEach(([x, y], d) => { put(C.lineX, p + d, x); put(C.lineY, p + d, y); });
    }
    if (o.variant === 'outline') {
      const p = 3 + ppc * i, x = j => ref(`x${j}`), y = j => ref(`norm${j}`);
      const pts = k === 2
        ? [[x(0), y(0)], [x(0), '$G$1'], [x(1), '$G$1'], [x(1), y(1)]]
        : [[x(0), y(0)], [x(0), '$G$1'], [`${x(1)}-$H$1`, '$G$1'], [`${x(1)}-$H$1`, y(1)], [`${x(1)}+$H$1`, y(1)], [`${x(1)}+$H$1`, '$G$1'], [x(2), '$G$1'], [x(2), y(2)]];
      pts.forEach(([px, py], d) => { put(C.lineX, p + d, `=${px}`); put(C.lineY, p + d, `=IF(ISNUMBER(${py}),${py},NA())`); });
    }
    if (o.variant === 'delta') {
      if (k === 2) {
        const chg = ref('chg0');
        put(C.bar2_0, r, `=IF(ISNUMBER(${chg}),IF(${chg}>=0,${chg}/$F$1,NA()),NA())`);
        put(C.bar2_1, r, `=IF(ISNUMBER(${chg}),IF(${chg}<0,${chg}/$F$1,NA()),NA())`);
      } else for (let j = 0; j < m; j++) { const chg = ref(`chg${j}`); put(C[`bar2_${j}`], r, `=IF(ISNUMBER(${chg}),${chg}/$F$1,NA())`); }
    }
    if (o.variant !== 'delta') for (let j = 0; j < m; j++) {
      const t = 3 + i * m + j, chg = ref(`chg${j}`);
      const tx = o.variant === 'arrow' ? ref('mid') : k === 2 ? `(${ref('x0')}+${ref('x1')})/2` : j === 0 ? `(${ref('x0')}+${ref('x1')}-$H$1)/2` : `(${ref('x1')}+$H$1+${ref('x2')})/2`;
      const ty = o.variant === 'arrow' ? ref('hi') : '$G$1';
      put(C.tagX, t, `=${tx}`);
      put(C.tagUp, t, `=IF(ISNUMBER(${chg}),IF(${chg}>=0,${ty},NA()),NA())`);
      put(C.tagDown, t, `=IF(ISNUMBER(${chg}),IF(${chg}<0,${ty},NA()),NA())`);
      put(C.tagText, t, `=${ref(`lab${j}`)}`);
    }
  }
  const fixed = fixedSeriesCount(o.variant, k);
  const seriesCount = fixed + (ppc ? n : 0);
  const bottom = Math.max(last, 2 + n * ppc, 2 + tagCount);
  const lastCol = Object.values(C).at(-1);
  return {options:o, info, n, k, m, last, columns:C, scalars, cells, formats:f, fixed, seriesCount, pointsPerCategory:ppc, tagCount, bottom, lastCol, geometry:g};
}

export function fixedSeriesCount(variant, k) {
  if (variant === 'arrow') return k + 4;   // bars, head▲, head▼, tag up, tag down
  if (variant === 'outline') return k + 2; // bars, tag up, tag down
  return k + (k === 2 ? 2 : k - 1);        // bars, change columns
}

export function seriesName(p, i) {
  const {k, info, options:o} = p;
  if (i < k) return info.periodHeaders[i];
  const r = i - k;
  if (o.variant === 'arrow' && r < 4) return ['增长箭头','下降箭头','增长标签','下降标签'][r];
  if (o.variant === 'outline' && r < 2) return ['增长标签','下降标签'][r];
  if (o.variant === 'delta') return k === 2 ? ['增加','减少'][r] : `${info.periodHeaders[r+1]} vs ${info.periodHeaders[r]}`;
  return `轮廓线 ${i - p.fixed + 1}`;
}

/** Contiguous helper writes grouped per column; rows are 3..bottom. */
export function helperWrites(p) {
  const writes = [];
  for (const [address, formula] of Object.entries(p.scalars)) writes.push({address, formulas:[[formula]]});
  const cols = Object.values(p.columns);
  const grid = Array.from({length:p.bottom - 2}, (_, r) => cols.map(c => p.cells[`${c}${r+3}`] ?? ''));
  writes.push({address:`A3:${p.lastCol}${p.bottom}`, formulas:grid});
  const headers = cols.map(c => headerName(p, c));
  writes.push({address:`A2:${p.lastCol}2`, formulas:[headers]});
  return writes;
}
function headerName(p, col) {
  const name = Object.entries(p.columns).find(([, c]) => c === col)[0];
  const h = p.info.periodHeaders;
  const map = {cat:'分类', hi:'高值', lo:'低值', xLow:'箭头X', xHigh:'横线起点X', mid:'标签X', headUp:'▲Y', headDown:'▼Y', lineX:'线X', lineY:'线Y', tagX:'标签X', tagUp:'增长标签Y', tagDown:'下降标签Y', tagText:'标签文本'};
  if (map[name]) return map[name];
  const [, kind, j] = /^([a-z]+2?_?)(\d+)$/.exec(name) || [];
  if (kind === 'raw') return h[j]; if (kind === 'norm') return `${h[j]}/最大值`; if (kind === 'x') return `${h[j]} X`;
  if (kind === 'chg') return `${h[+j+1]} vs ${h[j]}`; if (kind === 'lab') return `${h[+j+1]} vs ${h[j]} 文本`;
  return name;
}

async function writeHelper({context, helper, source, options, previous, chart}) {
  const plan = buildHelperPlan(source, options);
  if (previous) {
    const stored = helper.getRange('B1:N1'); stored.load('values');
    chart.series.load('items/name,items/hasDataLabels'); await context.sync();
    if (Number(stored.values[0][12]) !== plan.k) throw new UserError('原图与当前选区的期数（列数）不同，未重绑；请另建图表。');
    let names; try { names = JSON.parse(stored.values[0][0]); } catch { throw new UserError('原图系列身份记录无效，未重绑。'); }
    const expected = plan.fixed + (plan.pointsPerCategory ? previous.count : 0);
    if (!Array.isArray(names) || names.length !== expected || chart.series.items.length !== names.length || chart.series.items.some((s, i) => s.name !== names[i])) throw new UserError('原图系列数量、名称或顺序已变化，未重绑；请保留原实例并另建图表。');
    plan.previousLabelVisibility = chart.series.items.map(s => s.hasDataLabels === true);
    plan.identityNames = names.slice(0, plan.seriesCount);
    while (plan.identityNames.length < plan.seriesCount) plan.identityNames.push(seriesName(plan, plan.identityNames.length));
    const oldBottom = Math.max(previous.count + 2, 2 + previous.count * plan.pointsPerCategory, 2 + previous.count * plan.m);
    helper.getRange(`A2:${plan.lastCol}${Math.max(oldBottom, plan.bottom)}`).clear('Contents');
  } else plan.identityNames = Array.from({length:plan.seriesCount}, (_, i) => seriesName(plan, i));
  helper.getRange('B1').values = [[JSON.stringify(plan.identityNames)]];
  for (const w of helperWrites(plan)) helper.getRange(w.address).formulas = w.formulas;
  const C = plan.columns;
  for (let j = 0; j < plan.k; j++) helper.getRange(`${C[`raw${j}`]}3:${C[`raw${j}`]}${plan.last}`).numberFormat = Array.from({length:plan.n}, () => [plan.formats.value]);
  return plan;
}

function bind(chart, helper, p) {
  const C = p.columns, s = i => chart.series.getItemAt(i), r = a => helper.getRange(a), col = (c, from = 3, to = p.last) => r(`${c}${from}:${c}${to}`);
  for (let j = 0; j < p.k; j++) { s(j).setXAxisValues(col(C.cat)); s(j).setValues(col(C[`norm${j}`])); }
  const k = p.k, tagLast = 2 + p.tagCount;
  if (p.options.variant === 'arrow') {
    s(k).setXAxisValues(col(C.xLow)); s(k).setValues(col(C.headUp));
    s(k+1).setXAxisValues(col(C.xLow)); s(k+1).setValues(col(C.headDown));
  }
  if (p.options.variant !== 'delta') {
    const t = p.options.variant === 'arrow' ? k + 2 : k;
    s(t).setXAxisValues(col(C.tagX, 3, tagLast)); s(t).setValues(col(C.tagUp, 3, tagLast));
    s(t+1).setXAxisValues(col(C.tagX, 3, tagLast)); s(t+1).setValues(col(C.tagDown, 3, tagLast));
    for (let i = 0; i < p.n; i++) {
      const from = 3 + i * p.pointsPerCategory, to = from + p.pointsPerCategory - 1;
      s(p.fixed + i).setXAxisValues(col(C.lineX, from, to)); s(p.fixed + i).setValues(col(C.lineY, from, to));
    }
  } else for (let j = 0; j < p.fixed - k; j++) { s(k + j).setXAxisValues(col(C.cat)); s(k + j).setValues(col(C[`bar2_${j}`])); }
}

function seriesKind(p, i) {
  if (i < p.k) return 'bar';
  if (p.options.variant === 'delta') return 'change';
  if (i >= p.fixed) return 'line';
  if (p.options.variant === 'arrow') return ['headUp','headDown','tagUp','tagDown'][i - p.k];
  return ['tagUp','tagDown'][i - p.k];
}
function configureSeriesType(s, kind) {
  if (kind === 'bar') { s.chartType = 'ColumnClustered'; s.axisGroup = 'Primary'; return; }
  if (kind === 'change') { s.chartType = 'ColumnClustered'; s.axisGroup = 'Secondary'; return; }
  // Label/arrow-head carriers are invisible points. Mac reports a marker-less XYScatter as
  // XYScatterLinesNoMarkers, so create them as that type with the line switched off: same type on both hosts.
  s.chartType = 'XYScatterLinesNoMarkers'; s.axisGroup = 'Secondary';
}

async function createChart({context, helper, sourceSheet, source, instance, binding:p}) {
  const chart = sourceSheet.charts.add('ColumnClustered', helper.getRange(`${p.columns.norm0}3:${p.columns.norm0}${p.last}`), 'Columns');
  chart.name = instance.chartName; chart.plotVisibleOnly = false; chart.displayBlanksAs = 'NotPlotted';
  chart.series.load('items');
  const anchor = sourceSheet.getRangeByIndexes(source.rowIndex, source.columnIndex + source.columnCount, 1, 1); anchor.load('left,top');
  await context.sync();
  if (chart.series.items.length !== 1) throw new UserError('新图未生成预期的单个初始系列。');
  chart.series.getItemAt(0).name = p.identityNames[0];
  for (let i = 1; i < p.seriesCount; i++) chart.series.add(p.identityNames[i]);
  await context.sync();
  for (let i = 0; i < p.seriesCount; i++) configureSeriesType(chart.series.getItemAt(i), seriesKind(p, i));
  bind(chart, helper, p);
  chart.left = anchor.left + 12; chart.top = anchor.top; chart.width = LAYOUT.width; chart.height = LAYOUT.height;
  await context.sync(); return chart;
}

function font(format, size = 9, color = COLORS.text, bold = false) { format.font.name = FONT; format.font.size = size; format.font.color = color; format.font.bold = bold; }
async function nativeStage(context, name, write) {
  try { await write(); await context.sync(); }
  catch (error) { const failure = new Error(`[涨跌柱/${name}] ${error.message || String(error)}`); failure.name = error.name || 'Error'; failure.code = error.code; failure.debugInfo = error.debugInfo; failure.cause = error; throw failure; }
}
function palette(p) { return p.options.variant === 'arrow' ? COLORS.arrow : p.options.variant === 'outline' ? (p.k === 2 ? COLORS.outline2 : COLORS.outline3) : COLORS.delta; }

/** Which series carry which per-point labels; shared by style() and rebind(). */
export function labelTargets(p) {
  const C = p.columns, o = p.options, k = p.k, out = [];
  if (o.showLabels) for (let j = 0; j < k; j++) out.push({series:j, column:C[`raw${j}`], count:p.n, position:o.variant === 'delta' ? 'OutsideEnd' : 'InsideEnd', kind:'bar'});
  if (o.variant === 'arrow') {
    out.push({series:k, glyph:'▲', count:p.n, position:'Bottom', kind:'head'}, {series:k+1, glyph:'▼', count:p.n, position:'Top', kind:'head'});
    out.push({series:k+2, column:C.tagText, count:p.tagCount, position:'Top', kind:'tag'}, {series:k+3, column:C.tagText, count:p.tagCount, position:'Top', kind:'tag'});
  }
  if (o.variant === 'outline') out.push({series:k, column:C.tagText, count:p.tagCount, position:'Top', kind:'tag'}, {series:k+1, column:C.tagText, count:p.tagCount, position:'Top', kind:'tag'});
  if (o.variant === 'delta') for (let j = 0; j < p.fixed - k; j++) out.push({series:k + j, column:k === 2 ? C.lab0 : C[`lab${j}`], count:p.n, position:'OutsideEnd', kind:'change'});
  return out;
}
function labelColor(p, t) {
  const c = palette(p), o = p.options;
  if (t.kind === 'bar') return o.variant === 'outline' ? c.barLabels[t.series] : COLORS.text;
  if (t.kind === 'head') return t.glyph === '▲' ? c.line : c.down;
  if (t.kind === 'tag') return (t.series - (o.variant === 'arrow' ? p.k + 2 : p.k)) === 0 ? c.up : c.down;
  return p.k === 2 && t.series === p.k + 1 ? c.down : COLORS.text;
}

async function applyLabels(context, chart, helperName, p, start = 0, visibility = null) {
  const targets = labelTargets(p).filter(t => !visibility || visibility[t.series] === true);
  if (start === 0) await nativeStage(context, '标签/启用', () => {
    for (const t of targets) {
      const s = chart.series.getItemAt(t.series); s.hasDataLabels = true;
      s.dataLabels.showValue = Boolean(t.glyph); s.dataLabels.showCategoryName = false; s.dataLabels.showSeriesName = false;
      if (t.glyph) { s.dataLabels.linkNumberFormat = false; s.dataLabels.numberFormat = `"${t.glyph}";"${t.glyph}";"${t.glyph}"`; }
      font(s.dataLabels.format, t.kind === 'head' ? 10 : 9, labelColor(p, t), t.kind === 'tag');
    }
  });
  // Mac rejects series-level positions in mixed column/XY charts: position each point.
  await nativeStage(context, '标签/逐点链接与位置', () => {
    for (const t of targets) {
      const s = chart.series.getItemAt(t.series), from = t.kind === 'tag' ? start * p.m : start;
      for (let i = from; i < t.count; i++) {
        const point = s.points.getItemAt(i); point.hasDataLabel = true;
        if (t.column) point.dataLabel.formula = `=${quote(helperName)}!$${t.column}$${i + 3}`;
        point.dataLabel.position = t.position;
      }
    }
  });
}

async function hideAuxLegend(context, chart, p, start = 0) {
  const entries = chart.legend.legendEntries; entries.load('items'); await context.sync();
  const visibleUpTo = p.options.variant === 'delta' && p.k === 3 ? p.fixed : p.k;
  for (let i = Math.max(start, visibleUpTo); i < entries.items.length; i++) entries.items[i].visible = false;
  await context.sync();
}
function styleLine(s, color) { s.markerStyle = 'None'; s.format.line.lineStyle = 'Continuous'; s.format.line.color = color; s.format.line.weight = 0.75; s.hasDataLabels = false; }

async function style({context, chart, instance, binding:p}) {
  const o = p.options, c = palette(p), g = p.geometry, k = p.k;
  await nativeStage(context, '图表区域/标题图例', () => {
    font(chart.format); chart.format.fill.setSolidColor('#FFFFFF'); chart.format.border.color = '#D9D9D9'; chart.format.border.weight = 0.5;
    chart.title.visible = Boolean(o.title); chart.title.text = o.title; font(chart.title.format, 14, '#262626', true);
    chart.legend.visible = true; chart.legend.position = 'Top'; font(chart.legend.format, 9);
  });
  await nativeStage(context, '柱系列/填充', () => {
    for (let j = 0; j < k; j++) { const s = chart.series.getItemAt(j); s.gapWidth = g.gap; s.overlap = 0; s.format.fill.setSolidColor(c.bars[j]); s.format.line.lineStyle = 'None'; s.hasDataLabels = false; }
  });
  if (o.variant === 'delta') await nativeStage(context, '增减量柱/填充', () => {
    for (let j = k; j < p.fixed; j++) {
      const s = chart.series.getItemAt(j); s.gapWidth = k === 2 ? 150 : 80; s.overlap = k === 2 ? 100 : 0;
      s.format.fill.setSolidColor(k === 2 ? (j === k ? c.up : c.down) : c.changes[j - k]); s.format.line.lineStyle = 'None'; s.hasDataLabels = false;
    }
  });
  else await nativeStage(context, '散点/标记与线', () => {
    for (let j = k; j < p.fixed; j++) { const s = chart.series.getItemAt(j); s.markerStyle = 'None'; s.format.line.lineStyle = 'None'; s.hasDataLabels = false; }
    for (let i = p.fixed; i < p.seriesCount; i++) styleLine(chart.series.getItemAt(i), c.line);
  });
  await applyLabels(context, chart, instance.helperName, p);
  const category = chart.axes.getItem('Category', 'Primary'), primary = chart.axes.getItem('Value', 'Primary'), secondary = chart.axes.getItem('Value', 'Secondary'), sx = chart.axes.getItem('Category', 'Secondary');
  await nativeStage(context, '主分类轴/样式', () => {
    category.visible = true; category.categoryType = 'TextAxis'; category.tickLabelPosition = 'Low'; category.majorTickMark = 'None'; category.minorTickMark = 'None';
    category.format.line.color = '#BFBFBF'; category.format.line.weight = 0.75; font(category.format, 9);
  });
  for (const [axis, name] of [[primary, '主数值轴'], [secondary, '次数值轴']]) {
    await nativeStage(context, name + '/启用', () => { axis.visible = true; });
    await nativeStage(context, name + '/隐藏刻度与网格', () => {
      axis.majorGridlines.visible = false; axis.minorGridlines.visible = false; axis.majorTickMark = 'None'; axis.minorTickMark = 'None';
      axis.format.line.lineStyle = 'None'; axis.tickLabelPosition = 'None'; axis.title.visible = false;
    });
  }
  await nativeStage(context, '数值轴/同比例范围', () => {
    primary.minimum = 0; primary.maximum = g.top;
    if (o.variant === 'delta') { secondary.minimum = g.secondaryMin; secondary.maximum = g.secondaryMax; }
    else { secondary.minimum = 0; secondary.maximum = g.top; }
  });
  const notes = ['柱高=各期值/全部值最大值，坐标轴刻度隐藏；柱上数字链接辅助表原值。', '同一区域改数由公式重算；增删行请按选区重绑。'];
  if (o.variant !== 'delta') {
    await nativeStage(context, '次散点X轴/范围', () => { sx.visible = true; sx.minimum = 0; sx.maximum = 1; sx.tickLabelPosition = 'None'; sx.majorTickMark = 'None'; sx.minorTickMark = 'None'; sx.majorGridlines.visible = false; });
    await nativeStage(context, '次散点X轴/隐藏', () => { sx.visible = false; });
  } else {
    // Zero line of the change band = the secondary category axis crossing at 0. Non-critical on hosts without that format object.
    try { await nativeStage(context, '增减零线', () => { sx.visible = true; sx.tickLabelPosition = 'None'; sx.majorTickMark = 'None'; sx.format.line.color = c.zero; sx.format.line.weight = 0.75; }); }
    catch (e) { notes.push(`增减零线未设置：${e.message}`); }
  }
  await hideAuxLegend(context, chart, p);
  return {notes, variant:o.variant, metric:o.resolvedMetric, nativeVerified:false};
}

async function rebind({context, chart, helper, instance, binding:p, previous}) {
  if (p.pointsPerCategory) {
    const c = palette(p);
    if (p.n < previous.count) for (let i = p.fixed + previous.count - 1; i >= p.fixed + p.n; i--) chart.series.getItemAt(i).delete();
    if (p.n > previous.count) for (let i = previous.count; i < p.n; i++) { const s = chart.series.add(p.identityNames[p.fixed + i]); configureSeriesType(s, 'line'); styleLine(s, c.line); }
    await context.sync();
  }
  bind(chart, helper, p); await context.sync();
  if (p.n > previous.count) { await applyLabels(context, chart, instance.helperName, p, previous.count, p.previousLabelVisibility); await hideAuxLegend(context, chart, p, p.fixed); }
  return {variant:p.options.variant, preservedFormatting:true};
}

const inputSpec = {
  category:{id:'comparison', title:'比较'},
  description:'首行为表头。两期：分类、基期、本期（3 列）；三年：分类、第1年、第2年、第3年（4 列）。各期数值须非负。',
  minVersion:'1.9', columns:3, minRows:1, maxRows:MAX_ROWS,
  options:[
    {id:'title', label:'标题', type:'text', default:'同比增幅对比', maxLength:80},
    {id:'variant', label:'变体', type:'select', default:'arrow', choices:[
      {value:'arrow', label:'涨跌箭头柱（两期）'}, {value:'outline', label:'轮廓线增幅柱（两期/三年）'}, {value:'delta', label:'增减量柱（顶部涨跌值）'}]},
    {id:'metric', label:'变化指标', type:'select', default:'auto', choices:[
      {value:'auto', label:'自动（箭头/轮廓=增长率，增减量=差值）'}, {value:'growth', label:'增长率 本期/上期−1'}, {value:'difference', label:'差值 本期−上期'}]},
    {id:'showLabels', label:'显示柱数值', type:'checkbox', default:true}
  ],
  storage:{marker:'ChartCraft.GrowthArrow.v1', helperPrefix:'CC_涨跌柱_', chartPrefix:'CC_GrowthArrow_'}
};
export default {id:'growtharrow', title:'涨跌箭头与同比柱', inputSpec, validate, writeHelper, createChart, style, rebind};
