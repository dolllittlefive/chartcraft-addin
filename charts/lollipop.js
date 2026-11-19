// Two-column native scatter lollipops and three/four-column native bubble
// lollipops share the instance envelope, but retain separate chart topologies.
import {UserError} from '../core/errors.js';
import * as bubble from './lollipop/bubble.js';

const MARK = 'ChartCraft.Lollipop.v1';
const HELPER_PREFIX = 'CC_棒棒糖_';
const CHART_PREFIX = 'CC_Lollipop_';
const MAX_ROWS = 60;
const COLORS = {head: '#1A6092', stem: '#A3A3A3', text: '#595959', grid: '#EEEEEE'};
const quote = name => `'${String(name).replace(/'/g, "''")}'`;
function columnLetter(index) {
  let result = '', n = index + 1;
  while (n > 0) { const m = (n - 1) % 26; result = String.fromCharCode(65 + m) + result; n = Math.floor((n - 1) / 26); }
  return result;
}

function validate(values) {
  if ([3, 4].includes(values?.[0]?.length)) return bubble.validate(values);
  if (!Array.isArray(values) || values.length < 3) throw new UserError('请选择含表头的两列数据（分类、数值），至少 2 行数据。');
  if (!Array.isArray(values[0]) || values[0].length !== 2) throw new UserError('请选择两列（分类、数值），或三至四列（分类、增长率小数、规模、可选分组）。');
  if (values.length - 1 > MAX_ROWS) throw new UserError(`最多支持 ${MAX_ROWS} 行数据。`);
  values.slice(1).forEach((row, i) => {
    if (!Array.isArray(row) || row.length !== 2) throw new UserError(`第 ${i + 2} 行需要两列。`);
    if (row[0] == null || String(row[0]).trim() === '') throw new UserError(`第 ${i + 2} 行分类为空。`);
    if (typeof row[1] !== 'number' || !Number.isFinite(row[1])) throw new UserError(`第 ${i + 2} 行数值必须是有限数字。`);
  });
  return {count: values.length - 1, inputKind: 'equal', categoryHeader: String(values[0][0]), valueHeader: String(values[0][1])};
}

function resolvedVariant(source, options = {}) {
  const requested = options.variant ?? 'auto';
  if (!['auto', 'vertical', 'horizontal', 'bubble'].includes(requested)) throw new UserError('请选择有效的棒棒糖图类型。');
  const isBubble = source.inputKind === 'bubble' || [3, 4].includes(source.columnCount);
  if (requested === 'bubble' && !isBubble) throw new UserError('气泡棒棒糖需要三至四列：分类、增长率小数、规模、可选分组。');
  if (isBubble && requested === 'horizontal') throw new UserError('气泡棒棒糖使用纵向增长率与规模，请选“自动识别”或“气泡规模”。');
  return isBubble ? 'bubble' : requested === 'horizontal' ? 'horizontal' : 'vertical';
}

// The stem path returns to zero before moving to the next category:
// (i, 0) -> (i, value) -> (i, 0) -> (i+1, 0).
// This gives a genuine zero baseline and isolated stems without relying on
// host-specific #N/A gap settings or per-point error-bar length APIs.
function helperRecipe(source, options = {}) {
  const variant = resolvedVariant(source, options);
  if (options.negativeShade && variant !== 'bubble') throw new UserError('负值区浅灰底适用于气泡棒棒糖。请使用含分类、增长率、规模的三至四列数据。');
  if (variant === 'bubble') return bubble.helperRecipe(source, options);
  const horizontal = variant === 'horizontal', last = source.count + 2;
  const sheet = quote(source.sheet), category = columnLetter(source.columnIndex), value = columnLetter(source.columnIndex + 1);
  const rows = [], stems = [];
  for (let i = 0; i < source.count; i++) {
    const r = i + 3, sr = source.rowIndex + i + 2;
    rows.push([
      `=${sheet}!$${category}$${sr}&""`,
      `=IF(ISNUMBER(${sheet}!$${value}$${sr}),${sheet}!$${value}$${sr},NA())`,
      '=ROW()-2', `=IFERROR(B${r}&"","")`,
      horizontal ? `=B${r}` : `=C${r}`, horizontal ? `=C${r}` : `=B${r}`,
      horizontal ? '=$B$1' : `=C${r}`, horizontal ? `=C${r}` : '=$B$1'
    ]);
    for (const y of ['=0', `=B${r}`, '=0']) stems.push(horizontal ? [y, `=C${r}`] : [`=C${r}`, y]);
  }
  const values = `$B$3:$B$${last}`;
  return {
    horizontal, last, stemLast: 3 * source.count + 2, rows, stems,
    labelFloor: `=MIN(0,MIN(${values}))-MAX(1,MAX(${values})-MIN(${values}),ABS(MAX(${values})),ABS(MIN(${values})))*0.12`
  };
}

function writeHelper(args) {
  const {helper, source, options, previous} = args;
  // Validate variant before touching an existing helper.
  const binding = helperRecipe(source, options);
  if (previous && (previous.options?.variant === 'bubble') !== (binding.variant === 'bubble')) {
    throw new UserError('现有等大圆点图与气泡规模图的原生结构不同。请保留原图，使用新选区另建对应类型。');
  }
  if (binding.variant === 'bubble') return bubble.writeHelper({...args, binding});
  if (previous && previous.count > source.count) {
    helper.getRange(`A${source.count + 3}:H${previous.count + 2}`).clear(Excel.ClearApplyTo.contents);
    helper.getRange(`J${source.count * 3 + 3}:K${previous.count * 3 + 2}`).clear(Excel.ClearApplyTo.contents);
  }
  helper.getRange('B1').formulas = [[binding.labelFloor]];
  helper.getRange('A2:H2').values = [['分类', '数值', '序号', '数值标签', '圆点X', '圆点Y', '类别标签X', '类别标签Y']];
  helper.getRange('J2:K2').values = [['茎X', '茎Y']];
  helper.getRange(`A3:H${binding.last}`).formulas = binding.rows;
  helper.getRange(`J3:K${binding.stemLast}`).formulas = binding.stems;
  return binding;
}

function bindSeries(chart, helper, binding) {
  // Stems first; heads then cover their endpoints rather than being crossed
  // by a stem drawn over the centre of each marker.
  const ranges = [['J3', `J${binding.stemLast}`, 'K3', `K${binding.stemLast}`], ['E3', `E${binding.last}`, 'F3', `F${binding.last}`], ['G3', `G${binding.last}`, 'H3', `H${binding.last}`]];
  ranges.forEach(([x1, x2, y1, y2], i) => {
    const series = chart.series.getItemAt(i);
    series.setXAxisValues(helper.getRange(`${x1}:${x2}`));
    series.setValues(helper.getRange(`${y1}:${y2}`));
  });
}

// Preserve the native exception/code/debugInfo while making the failing batch
// visible in the pane and acceptance receipt. Never continue after a failure.
async function syncStage(context, stage) {
  try { await context.sync(); }
  catch (error) {
    error.message = `棒棒糖 [${stage}]：${error.message}`;
    throw error;
  }
}

async function createChart(args) {
  const {context, helper, sourceSheet, source, instance, binding} = args;
  if (binding.variant === 'bubble') return bubble.createChart(args);
  const chart = sourceSheet.charts.add(Excel.ChartType.xyscatter, helper.getRange(`K3:K${binding.stemLast}`), Excel.ChartSeriesBy.columns);
  chart.name = instance.chartName;
  chart.series.getItemAt(0).name = '零线与茎';
  chart.series.add(source.valueHeader); chart.series.add('分类标签');
  const anchor = sourceSheet.getRangeByIndexes(source.rowIndex, source.columnIndex + 2, 1, 1);
  anchor.load('left,top'); await syncStage(context, 'create-chart-and-anchor');
  chart.left = anchor.left + 12; chart.top = anchor.top;
  chart.width = binding.horizontal ? 620 : 760; chart.height = binding.horizontal ? 460 : 420;
  chart.plotVisibleOnly = false;
  bindSeries(chart, helper, binding);
  return chart;
}

function linkLabels({chart, instance, source, options, binding, previous}) {
  const oldCount = previous?.count ?? 0;
  for (let i = 0; i < source.count; i++) {
    const categoryPoint = chart.series.getItemAt(2).points.getItemAt(i);
    const valuePoint = chart.series.getItemAt(1).points.getItemAt(i);
    // On rebind, only new points get default formatting. Existing user label
    // positions, fonts, colours and visibility are not assigned again.
    if (i >= oldCount) {
      categoryPoint.hasDataLabel = true;
      categoryPoint.dataLabel.position = binding.horizontal ? Excel.ChartDataLabelPosition.left : Excel.ChartDataLabelPosition.bottom;
      categoryPoint.dataLabel.format.font.size = 10;
      categoryPoint.dataLabel.format.font.color = COLORS.text;
      valuePoint.hasDataLabel = options.showValues !== false;
      if (options.showValues !== false) {
        valuePoint.dataLabel.position = binding.horizontal ? Excel.ChartDataLabelPosition.right : Excel.ChartDataLabelPosition.top;
        valuePoint.dataLabel.format.font.size = 10;
        valuePoint.dataLabel.format.font.color = COLORS.text;
      }
    }
    categoryPoint.dataLabel.formula = `=${quote(instance.helperName)}!$A$${i + 3}`;
    if (options.showValues !== false) valuePoint.dataLabel.formula = `=${quote(instance.helperName)}!$D$${i + 3}`;
  }
}

async function style(args) {
  const {context, chart, options, binding} = args;
  if (binding.variant === 'bubble') return bubble.style(args);
  await syncStage(context, 'bind-series-data');
  chart.format.font.name = '微软雅黑'; chart.format.font.size = 10; chart.format.font.color = COLORS.text;
  // This fill belongs to the chart area, never to a scatter series or point.
  chart.format.fill.setSolidColor('#FFFFFF'); chart.format.border.lineStyle = 'None';
  chart.title.visible = true; chart.title.text = options.title;
  // Establish the requested typography before Excel measures its automatic
  // plot area. Resizing a 16pt title afterwards makes the old plot impossible
  // to restore identically in the Windows host.
  chart.title.format.font.size = 14; chart.title.format.font.bold = false;
  chart.title.horizontalAlignment = Excel.ChartTextHorizontalAlignment.left;
  chart.legend.visible = false;
  await syncStage(context, 'chart-area-title-legend');
  const stems = chart.series.getItemAt(0);
  stems.chartType = Excel.ChartType.xyscatterLines; stems.axisGroup = Excel.ChartAxisGroup.primary;
  stems.markerStyle = Excel.ChartMarkerStyle.none;
  stems.format.line.lineStyle = 'Continuous'; stems.format.line.weight = 0.75; stems.format.line.color = COLORS.stem;
  await syncStage(context, 'series-0-stem-line');
  const heads = chart.series.getItemAt(1);
  heads.chartType = Excel.ChartType.xyscatter; heads.axisGroup = Excel.ChartAxisGroup.primary;
  heads.markerStyle = Excel.ChartMarkerStyle.circle; heads.markerSize = 9;
  heads.markerBackgroundColor = COLORS.head; heads.markerForegroundColor = COLORS.head;
  // Scatter heads use marker colours. The Mac creation batch failed at
  // ChartFill.setSolidColor; markerBackgroundColor already supplies this fill.
  heads.format.line.lineStyle = 'None';
  await syncStage(context, 'series-1-head-markers');
  const names = chart.series.getItemAt(2);
  names.chartType = Excel.ChartType.xyscatter; names.axisGroup = Excel.ChartAxisGroup.primary;
  names.markerStyle = Excel.ChartMarkerStyle.none; names.format.line.lineStyle = 'None';
  await syncStage(context, 'series-2-category-anchors');
  for (const axis of [chart.axes.categoryAxis, chart.axes.valueAxis]) {
    axis.majorTickMark = 'None'; axis.minorTickMark = 'None'; axis.format.line.lineStyle = 'None';
    axis.majorGridlines.visible = false;
  }
  const indexAxis = binding.horizontal ? chart.axes.valueAxis : chart.axes.categoryAxis;
  const valueAxis = binding.horizontal ? chart.axes.categoryAxis : chart.axes.valueAxis;
  indexAxis.tickLabelPosition = 'None';
  valueAxis.tickLabelPosition = 'Low'; valueAxis.majorGridlines.visible = true;
  valueAxis.majorGridlines.format.line.color = COLORS.grid; valueAxis.majorGridlines.format.line.weight = 0.5;
  // Autoscale remains enabled so future positive/negative values remain visible
  // when Excel recalculates with the add-in closed.
  await syncStage(context, 'axes');
  linkLabels(args); await syncStage(context, 'series-1-and-2-linked-labels');
  return {variant: binding.horizontal ? 'horizontal' : 'vertical', notes: ['等大圆点原生图。']};
}

async function rebind(args) {
  if (args.binding.variant === 'bubble') return bubble.rebind(args);
  bindSeries(args.chart, args.helper, args.binding);
  await syncStage(args.context, 'rebind-series-data');
  linkLabels(args); await syncStage(args.context, 'rebind-linked-labels');
}

async function finalizePresentation(args) {
  if (args.binding.variant === 'bubble') return bubble.finalizePresentation(args);
}

const inputSpec = {
  category: {id: 'comparison', title: '比较'},
  description: '含表头：两列（分类、数值）生成等大圆点；三至四列（分类、增长率小数、规模、可选分组）生成原生气泡棒棒糖，支持2–15项。增长率按27%=0.27输入；修改绑定数值自动重算。',
  minVersion: '1.9', columns: 2, minRows: 2, maxRows: MAX_ROWS,
  options: [
    {id: 'title', label: '标题', type: 'text', default: '棒棒糖图', maxLength: 80},
    {id: 'variant', label: '类型', type: 'select', default: 'auto', choices: [{value: 'auto', label: '自动识别'}, {value: 'vertical', label: '等大圆点·纵向'}, {value: 'horizontal', label: '等大圆点·横向'}, {value: 'bubble', label: '气泡规模'}]},
    {id: 'showValues', label: '显示数值', type: 'checkbox', default: true},
    {id: 'negativeShade', label: '负值区浅灰底（气泡）', type: 'checkbox', default: false}
  ],
  storage: {marker: MARK, helperPrefix: HELPER_PREFIX, chartPrefix: CHART_PREFIX}
};
export {validate, helperRecipe, resolvedVariant};
export default {id: 'lollipop', title: '棒棒糖图', inputSpec, validate, writeHelper, createChart, style, rebind, finalizePresentation};
