// Extracted from the verified pareto/web/pareto.js; formula/style recipe retained.
import {UserError} from '../core/errors.js';
const MARK = 'ChartCraft.Pareto.v1';
const HELPER_PREFIX = 'CC_帕累托_';
const CHART_PREFIX = 'CC_Pareto_';
const MAX_ROWS = 200;
const COLORS = { bar: '#0A1C2C', line: '#40A1ED', cross: '#323FCA', grid: '#F0F0F0', frame: '#D9D9D9', text: '#595959', value: '#404040', threshold: '#000000', guide: '#595959' };
const FONT = '微软雅黑';
// 原图 752×835 px，按坐标轴数字推算约 1.68 px/pt。
// 绘图区按原图实测 x=73..666、y=140..699 px 换算（第一轮用 0.11 的 top=79，实测整体偏上约 6 px）。
const LAYOUT = { width: 448, height: 497, plot: { left: 44, top: 83.3, width: 353, height: 332.7 }, gapWidth: 3 };
// 辅助表第 1 行参数列。
const P = { mark: 'A1', threshold: 'B1', total: 'C1', thresholdValue: 'D1', crossIndex: 'E1', prevCum: 'F1', crossCum: 'G1', crossX: 'H1', source: 'I1', chartSheet: 'J1', count: 'K1', chartName: 'L1', title: 'M1' };

const quote = name => `'${String(name).replace(/'/g, "''")}'`;
// 写进单元格的说明文字不能被 Excel 当成公式或数字。
const literal = text => "'" + String(text);
function columnLetter(index) { let s = '', n = index + 1; while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; }

/** 校验选区：首行为表头，两列（分类、非负数值），至少 2 行数据。 */
function validate(values) {
  if (!values || values.length < 3) throw new UserError('请选择含表头的两列数据，至少 2 行数据。');
  if (values[0].length !== 2) throw new UserError(`需要正好两列（分类、数值），当前选区有 ${values[0].length} 列。`);
  const rows = values.slice(1);
  if (rows.length > MAX_ROWS) throw new UserError(`最多支持 ${MAX_ROWS} 行数据。`);
  let total = 0;
  rows.forEach((row, i) => {
    if (row[0] === '' || row[0] === null) throw new UserError(`第 ${i + 2} 行分类为空。`);
    if (typeof row[1] !== 'number' || !isFinite(row[1])) throw new UserError(`第 ${i + 2} 行数值不是数字：${row[1]}`);
    if (row[1] < 0) throw new UserError(`第 ${i + 2} 行数值为负数；帕累托图只接受非负数值。`);
    total += row[1];
  });
  if (total <= 0) throw new UserError('数值合计必须大于 0。');
  return { count: rows.length, categoryHeader: String(values[0][0]), valueHeader: String(values[0][1]), total };
}

/** 写入（或重写）辅助表公式。src: {sheet,rowIndex,columnIndex,count}；返回数据末行号。 */
function writeHelper(helper, src, threshold, meta) {
  const n = src.count, last = n + 2;
  const s = quote(src.sheet), catCol = columnLetter(src.columnIndex), valCol = columnLetter(src.columnIndex + 1);
  const rows = [];
  for (let i = 0; i < n; i++) {
    const r = i + 3, srcRow = src.rowIndex + 2 + i;
    rows.push([
      `=${s}!$${catCol}$${srcRow}&""`,
      `=N(${s}!$${valCol}$${srcRow})`,
      `=RANK(B${r},$B$3:$B$${last},0)+COUNTIF($B$3:B${r},B${r})-1`,
      `=INDEX($A$3:$A$${last},MATCH(ROW()-2,$C$3:$C$${last},0))`,
      `=INDEX($B$3:$B$${last},MATCH(ROW()-2,$C$3:$C$${last},0))`,
      `=SUM($E$3:E${r})`,
      `=IF($C$1>0,F${r}/$C$1,0)`,
      String(i + 1.5),
    ]);
  }
  helper.getRange('A1:M1').formulas = [[
    MARK, String(threshold), `=SUM($B$3:$B$${last})`, '=B1*C1',
    `=MIN(K1,COUNTIF($F$3:$F$${last},"<"&D1)+1)`,
    `=IF(E1=1,0,INDEX($F$3:$F$${last},E1-1))`,
    `=INDEX($F$3:$F$${last},E1)`,
    '=IF(G1=F1,E1,E1-1+(D1-F1)/(G1-F1))',
    literal(meta.source), literal(meta.chartSheet), String(n), literal(meta.chartName), literal(meta.title),
  ]];
  helper.getRange('A2:H2').formulas = [['', '', '', '', '', '0', '', '0.5']];
  helper.getRange(`A3:H${last}`).formulas = rows;
  helper.getRange(`G3:G${last}`).numberFormat = rows.map(() => ['0.0%']);
  // 阈值线、交点竖线、首个达标点、阈值标签文字。
  helper.getRange('J2:P3').formulas = [
    ['=H1+0.5', '=D1', '=E1+0.5', '0', '=E1+0.5', '=G1', '=UNICHAR(9664)&UNICHAR(65038)&" "&ROUND(B1*100,0)&"%"'],
    [String(n + 0.5), '=D1', '=E1+0.5', '=G1', '', '', ''],
  ];
  return last;
}

function seriesRanges(helper, last) {
  return {
    barsX: helper.getRange(`D3:D${last}`), barsY: helper.getRange(`E3:E${last}`),
    lineX: helper.getRange(`H2:H${last}`), lineY: helper.getRange(`F2:F${last}`),
    thrX: helper.getRange('J2:J3'), thrY: helper.getRange('K2:K3'),
    guideX: helper.getRange('L2:L3'), guideY: helper.getRange('M2:M3'),
    crossX: helper.getRange('N2:N2'), crossY: helper.getRange('O2:O2'),
  };
}

function bindSeries(chart, ranges) {
  const s = i => chart.series.getItemAt(i);
  s(0).setXAxisValues(ranges.barsX); s(0).setValues(ranges.barsY);
  s(1).setXAxisValues(ranges.lineX); s(1).setValues(ranges.lineY);
  s(2).setXAxisValues(ranges.thrX); s(2).setValues(ranges.thrY);
  s(3).setXAxisValues(ranges.guideX); s(3).setValues(ranges.guideY);
  s(4).setXAxisValues(ranges.crossX); s(4).setValues(ranges.crossY);
}

/** 每个累计点的百分比标签链接辅助表 G 列；第 0 点是零点，不显示标签。 */
function linkCumulativeLabels(chart, helperName, n, preserveCount = 0) {
  const line = chart.series.getItemAt(1);
  for (let i = 1; i <= n; i++) {
    // Existing formulas already address the same helper row. Do not touch their
    // visibility or formatting when the source range changes.
    if (i <= preserveCount) continue;
    const point = line.points.getItemAt(i);
    point.hasDataLabel = true;
    point.dataLabel.formula = `=${quote(helperName)}!$G$${i + 2}`;
    point.dataLabel.position = Excel.ChartDataLabelPosition.top;
    point.dataLabel.format.font.size = 10; point.dataLabel.format.font.color = '#000000'; point.dataLabel.format.font.name = FONT;
  }
}

function styleChart(chart, opts) {
  const s = i => chart.series.getItemAt(i);
  chart.format.font.name = FONT; chart.format.font.color = COLORS.text; chart.format.font.size = 10;
  chart.format.border.color = COLORS.frame; chart.format.border.weight = 0.75; chart.format.border.lineStyle = 'Continuous';

  chart.title.visible = true; chart.title.text = opts.title;
  chart.title.format.font.size = 16; chart.title.format.font.bold = false; chart.title.format.font.color = COLORS.text;
  chart.title.horizontalAlignment = Excel.ChartTextHorizontalAlignment.left;

  chart.legend.visible = true; chart.legend.position = Excel.ChartLegendPosition.top;
  chart.legend.format.font.size = 11; chart.legend.format.font.color = COLORS.text;

  const bars = s(0);
  bars.gapWidth = LAYOUT.gapWidth; bars.format.fill.setSolidColor(COLORS.bar); bars.format.line.lineStyle = 'None';
  bars.hasDataLabels = true; bars.dataLabels.showValue = true; bars.dataLabels.position = Excel.ChartDataLabelPosition.outsideEnd;
  bars.dataLabels.format.font.size = 10; bars.dataLabels.format.font.color = COLORS.value;

  const line = s(1);
  line.chartType = Excel.ChartType.xyscatterLines; line.axisGroup = Excel.ChartAxisGroup.primary;
  line.format.line.color = COLORS.line; line.format.line.weight = 1.25;
  line.markerStyle = Excel.ChartMarkerStyle.circle; line.markerSize = 8; line.markerBackgroundColor = '#FFFFFF'; line.markerForegroundColor = COLORS.line;
  line.points.getItemAt(0).markerStyle = Excel.ChartMarkerStyle.none;

  const thr = s(2);
  thr.chartType = Excel.ChartType.xyscatterLines; thr.axisGroup = Excel.ChartAxisGroup.primary;
  // 原图两条参考线都是短而密的虚线（RoundDot 太碎、Dash 太长，取 1pt 的 Dash 与原图更接近，见 comparison/）。
  thr.format.line.color = COLORS.threshold; thr.format.line.weight = 1; thr.format.line.lineStyle = Excel.ChartLineStyle.dash;
  thr.markerStyle = Excel.ChartMarkerStyle.none;
  // 原图交点是黑色细“+”；只设前景色，背景填充会让 Mac 把它画成实心方块。
  const cross = thr.points.getItemAt(0); cross.markerStyle = Excel.ChartMarkerStyle.plus; cross.markerSize = 7; cross.markerForegroundColor = COLORS.threshold;
  const end = thr.points.getItemAt(1); end.hasDataLabel = true;

  const guide = s(3);
  guide.chartType = Excel.ChartType.xyscatterLines; guide.axisGroup = Excel.ChartAxisGroup.primary;
  guide.format.line.color = COLORS.guide; guide.format.line.weight = 1; guide.format.line.lineStyle = Excel.ChartLineStyle.dash;
  guide.markerStyle = Excel.ChartMarkerStyle.none;

  const first = s(4);
  first.chartType = Excel.ChartType.xyscatter; first.axisGroup = Excel.ChartAxisGroup.primary;
  first.markerStyle = Excel.ChartMarkerStyle.square; first.markerSize = 8; first.markerBackgroundColor = COLORS.cross; first.markerForegroundColor = COLORS.cross;

  const value = chart.axes.valueAxis;
  value.minimum = 0; value.majorTickMark = 'None'; value.format.font.size = 10; value.format.font.color = COLORS.text;
  value.format.line.lineStyle = 'None';
  value.majorGridlines.visible = true; value.majorGridlines.format.line.color = COLORS.grid; value.majorGridlines.format.line.weight = 0.75;

  const category = chart.axes.categoryAxis;
  category.majorTickMark = 'None'; category.format.font.size = 11; category.format.font.color = COLORS.text;
  category.format.line.lineStyle = 'None';
}

/** 分类标签 45° 斜排、标题左对齐。单独提交，个别宿主不支持时记录而不中断生成。 */
async function styleTextPlacement(context, chart) {
  const notes = [];
  try { chart.axes.categoryAxis.textOrientation = 45; await context.sync(); } catch (e) { notes.push('分类标签斜排未生效：' + e.message); }
  try { chart.title.left = 8; chart.title.top = 6; await context.sync(); } catch (e) { notes.push('标题位置未生效：' + e.message); }
  return notes;
}

/** 阈值标签：链接辅助表 P2，放在阈值线右端外侧。 */
function linkThresholdLabel(chart, helperName) {
  const label = chart.series.getItemAt(2).points.getItemAt(1).dataLabel;
  label.formula = `=${quote(helperName)}!$P$2`;
  label.position = Excel.ChartDataLabelPosition.right;
  label.format.font.size = 10; label.format.font.color = COLORS.value; label.format.font.name = FONT;
}

/** 一些设置需要在系列和字体就位后重复一次，Excel 才不会自动改写（审核探针 P20 的现象）。 */
async function placePlotArea(context, chart) {
  const plot = chart.plotArea, want = LAYOUT.plot;
  for (let k = 0; k < 2; k++) {
    plot.insideWidth = want.width; plot.insideHeight = want.height; plot.insideLeft = want.left; plot.insideTop = want.top;
    await context.sync();
  }
  plot.load('insideLeft,insideTop,insideWidth,insideHeight'); await context.sync();
  return { requested: want, actual: { left: plot.insideLeft, top: plot.insideTop, width: plot.insideWidth, height: plot.insideHeight } };
}

async function hideHelperLegend(context, chart) {
  const entries = chart.legend.legendEntries;
  entries.load('items'); await context.sync();
  // 系列顺序：柱、累计线、阈值线、交点竖线、首个达标点；后两个不进图例。
  if (entries.items.length === 5) { entries.items[3].visible = false; entries.items[4].visible = false; }
  await context.sync();
  return entries.items.length;
}


async function writeHelperModule({context, helper, source, options, instance, previous}) {
  if (previous) {
    const used = helper.getUsedRange(); used.load('rowCount'); await context.sync();
    helper.getRange(`A3:H${Math.max(used.rowCount, previous.count + 2)}`).clear(Excel.ClearApplyTo.contents);
  }
  return {last: writeHelper(helper, source, options.threshold, instance)};
}

async function createChart({context, sourceSheet, helper, source, options, instance, binding}) {
  const ranges = seriesRanges(helper, binding.last);
  const chart = sourceSheet.charts.add(Excel.ChartType.columnClustered, ranges.barsY, Excel.ChartSeriesBy.columns);
  chart.name = instance.chartName;
  const anchor = sourceSheet.getRangeByIndexes(source.rowIndex, source.columnIndex + 2, 1, 1); anchor.load('left,top');
  await context.sync();
  chart.left = anchor.left + 12; chart.top = anchor.top; chart.width = LAYOUT.width; chart.height = LAYOUT.height;
  chart.series.getItemAt(0).name = source.valueHeader;
  for (const name of ['累计百分比', `${Math.round(options.threshold * 100)}%`, '交点', '首个达标']) chart.series.add(name);
  await context.sync();
  bindSeries(chart, ranges);
  return chart;
}

async function style({context, chart, options, instance, source}) {
  styleChart(chart, options); await context.sync();
  linkCumulativeLabels(chart, instance.helperName, source.count);
  linkThresholdLabel(chart, instance.helperName); await context.sync();
  const legendCount = await hideHelperLegend(context, chart);
  const notes = await styleTextPlacement(context, chart);
  const plot = await placePlotArea(context, chart);
  return {legendCount, notes, plot};
}

async function rebind({context, chart, helper, instance, source, binding, previous}) {
  bindSeries(chart, seriesRanges(helper, binding.last)); await context.sync();
  linkCumulativeLabels(chart, instance.helperName, source.count, previous?.count ?? 0);
  // The existing threshold label remains linked to P2 and keeps user styling.
  await context.sync();
}

const inputSpec = {
  crossSheetRebind: true,
  category: {id: 'composition', title: '构成与累计'},
  description: '选中两列数据（首行为表头：分类、数值），生成原生 Excel 图表。数值修改后图表自动更新，无需本插件。',
  minVersion: '1.9', columns: 2, minRows: 2, maxRows: MAX_ROWS,
  options: [
    {id: 'title', label: '标题', type: 'text', default: '图表标题', maxLength: 80},
    {id: 'threshold', label: '阈值线', type: 'number', default: 80, min: 1, max: 99, step: 1, scale: 0.01, suffix: '%'}
  ],
  storage: {marker: MARK, helperPrefix: HELPER_PREFIX, chartPrefix: CHART_PREFIX, legacyOptions: {threshold: 1}}
};
export {MARK, HELPER_PREFIX, CHART_PREFIX, COLORS, LAYOUT, FONT, validate, writeHelper as writeParetoHelper};
export default {id: 'pareto', title: '帕累托图', inputSpec, validate, writeHelper: writeHelperModule, createChart, style, rebind};
