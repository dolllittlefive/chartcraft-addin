import {UserError} from '../../core/errors.js';
import {LOLLIPOP_TEMPLATE_BASE64, LOLLIPOP_TEMPLATE_META as TEMPLATE} from './native-template.js';
import {LOLLIPOP_NEGATIVE_TEMPLATE_BASE64} from './negative-template.js';

// The imported worksheet is a private formula carrier. Its custom error-bar
// references must survive: never setData(), rebuild a series, or delete it.
export const CAPACITY = 15;
const TEMPLATE_PREFIX = 'CC_LP_T_';
const quote = name => `'${String(name).replace(/'/g, "''")}'`;
const formulaText = text => `"${String(text).replace(/"/g, '""')}"`;
const GROUP_COLORS = Object.freeze({
  '欧洲': ['#2251FF', '#FFFFFF'], '其他亚太': ['#20E9DA', '#000000'],
  '东盟': ['#14B8AB', '#000000'], '加拿大与墨西哥': ['#00A9F4', '#FFFFFF'],
  '拉丁美洲': ['#99E6FF', '#000000'], '中国（大陆）与香港': ['#9B187A', '#FFFFFF'],
  '中国大陆': ['#9B187A', '#FFFFFF']
});
const COUNTRY_GROUPS = Object.freeze({
  '爱尔兰': '欧洲', '法国': '欧洲', '意大利': '欧洲', '英国': '欧洲', '荷兰': '欧洲', '德国': '欧洲',
  '中国台湾': '其他亚太', '印度': '其他亚太', '日本': '其他亚太', '韩国': '其他亚太',
  '越南': '东盟', '墨西哥': '加拿大与墨西哥', '加拿大': '加拿大与墨西哥',
  '巴西': '拉丁美洲', '中国大陆': '中国（大陆）与香港'
});
const ORIGINAL_COUNTRIES = TEMPLATE.series.filter(s => s.role === 'trade').map(s => s.country);
const FALLBACK_COLORS = Object.values(GROUP_COLORS).slice(0, 6);

function columnLetter(index) {
  let result = '', n = index + 1;
  while (n > 0) { const m = (n - 1) % 26; result = String.fromCharCode(65 + m) + result; n = Math.floor((n - 1) / 26); }
  return result;
}

export function validate(values) {
  const width = values?.[0]?.length;
  if (!Array.isArray(values) || values.length < 3 || ![3, 4].includes(width)) {
    throw new UserError('气泡棒棒糖需要含表头的三至四列：分类、增长率小数、规模、可选分组，至少2行数据。');
  }
  if (values.length - 1 > CAPACITY) throw new UserError(`气泡棒棒糖最多支持${CAPACITY}行；不会截断或忽略其余数据。`);
  values.slice(1).forEach((row, i) => {
    if (!Array.isArray(row) || row.length !== width) throw new UserError(`第${i + 2}行需要${width}列。`);
    if (row[0] == null || String(row[0]).trim() === '') throw new UserError(`第${i + 2}行分类为空。`);
    if (typeof row[1] !== 'number' || !Number.isFinite(row[1])) throw new UserError(`第${i + 2}行增长率必须是有限数字（27%=0.27）。`);
    if (typeof row[2] !== 'number' || !Number.isFinite(row[2]) || row[2] <= 0) throw new UserError(`第${i + 2}行规模必须是大于0的有限数字。`);
  });
  return {count: values.length - 1, inputKind: 'bubble', categoryHeader: String(values[0][0]),
    valueHeader: String(values[0][1]), sizeHeader: String(values[0][2]), hasGroup: width === 4};
}

export function helperRecipe(source) {
  if (!Number.isInteger(source.count) || source.count < 2 || source.count > CAPACITY) throw new UserError('气泡棒棒糖只支持2–15行数据。');
  const sheet = quote(source.sheet), category = columnLetter(source.columnIndex);
  const growth = columnLetter(source.columnIndex + 1), size = columnLetter(source.columnIndex + 2);
  const hasGroup = source.hasGroup || source.columnCount === 4;
  const groupColumn = columnLetter(source.columnIndex + 3), rows = [];
  const sourceRows = source.values?.slice(1) || [];
  const groupKeys = sourceRows.map(row => String((hasGroup && row[3]) || COUNTRY_GROUPS[String(row[0]).trim()] || '其他'));
  const unknownGroups = [...new Set(groupKeys.filter(key => !GROUP_COLORS[key]))];
  const palette = groupKeys.map(key => GROUP_COLORS[key] || FALLBACK_COLORS[unknownGroups.indexOf(key) % FALLBACK_COLORS.length]);
  const originalTrade = source.count === CAPACITY && sourceRows.length === CAPACITY
    && sourceRows.every((row, i) => String(row[0]).trim() === ORIGINAL_COUNTRIES[i]);
  for (let i = 0; i < CAPACITY; i++) {
    const r = i + 3, sr = source.rowIndex + i + 2;
    if (i >= source.count) {
      // Safe numeric placeholders retain all 23 native series. Zero size, no
      // fill, no error bars and no labels prevent a removed row from lingering.
      rows.push(['=""', '=0', '=0', '=0', '=0', '=0', '=""', '=0', '=0', '=""']);
      continue;
    }
    const cell = col => `${sheet}!$${col}$${sr}`;
    rows.push([
      `=${cell(category)}&""`, `=${i + 1}`,
      `=IF(ISNUMBER(${cell(growth)}),${cell(growth)},NA())`,
      `=IF(AND(ISNUMBER(${cell(size)}),${cell(size)}>0),${cell(size)},NA())`,
      `=IFERROR(MAX(-C${r},0),0)`, `=IFERROR(MAX(C${r},0),0)`,
      `=IFERROR(TEXT(D${r},"0"),"")`, `=B${r}`,
      `=IFERROR(MAX(-0.39,MIN(0.79,C${r}+IF(C${r}<0,-1,1)*(0.055+0.095*SQRT(D${r}/MAX($D$3:$D$17))))),0)`,
      hasGroup ? `=${cell(groupColumn)}&""` : `=${formulaText(COUNTRY_GROUPS[String(sourceRows[i]?.[0]).trim()] || '其他')}`
    ]);
  }
  return {variant: 'bubble', rows, count: source.count, last: CAPACITY + 2, palette, groupKeys, originalTrade};
}

async function sync(context, stage) {
  const mark = state => {
    try { globalThis.CCGenerateDiagnostics?.mark?.('bubble:sync:' + state, {stage}); }
    catch { /* Optional acceptance diagnostics must not affect host errors. */ }
  };
  mark('start');
  try { await context.sync(); }
  catch (error) { mark('rejected'); error.message = `气泡棒棒糖 [${stage}]：${error.message}`; throw error; }
  mark('resolved');
}

async function templateFor({context, instance, options}) {
  const name = options.nativeTemplateSheet;
  if (!name || !String(name).startsWith(TEMPLATE_PREFIX)) throw new UserError('此气泡图缺少模板辅助表标识，未重写原图。');
  const sheet = context.workbook.worksheets.getItemOrNullObject(name);
  await sync(context, 'read-template-carrier');
  if (sheet.isNullObject) throw new UserError(`原生图表引用的辅助表“${name}”不存在，请保留原图并重新创建。`);
  const owner = sheet.getRange('T1:U1'); owner.load('values');
  await sync(context, 'verify-template-owner');
  if (owner.values[0][0] !== instance.helperName || owner.values[0][1] !== instance.chartName) {
    throw new UserError('模板辅助表与当前图表身份不匹配，未修改其公式。');
  }
  return sheet;
}

function linkTemplate(template, helperName) {
  template.getRange('A3:J17').formulas = Array.from({length: CAPACITY}, (_, i) =>
    Array.from({length: 10}, (_, c) => `=${quote(helperName)}!$${columnLetter(c)}$${i + 3}`));
  // K and L:R carry transparent label anchors, annotation geometry, and the
  // custom X-error-bar endpoints. They belong to the native template.
}

function setCurveCarrierVisibility(template, binding) {
  // A zero-sized transparent anchor hides its native label without deleting
  // the custom geometry, so a later rebind can restore the same editable curve.
  if (TEMPLATE.curveAnnotations?.length) template.getRange('O7:O9').values = Array.from({length:3}, () => [binding.originalTrade ? 1e-7 : 0]);
}

export async function writeHelper(args) {
  const {context, helper, source, options, previous} = args;
  const binding = args.binding || helperRecipe(source);
  // Validate carrier existence/ownership before altering an existing instance.
  if (previous) {
    const template = await templateFor(args);
    if (options.nativeTemplateVersion === TEMPLATE.version) setCurveCarrierVisibility(template, binding);
  }
  helper.getRange('A2:J2').values = [['分类', '序号', '增长率', '规模', '正误差', '负误差', '规模标签', '分类标签X', '分类标签Y', '分组']];
  helper.getRange('A3:J17').formulas = binding.rows;
  helper.getRange('C3:C17').numberFormat = Array.from({length: CAPACITY}, () => ['0%']);
  options.variant = 'bubble';
  if (!previous && binding.originalTrade && (!options.title || options.title === '棒棒糖图')) {
    options.title = TEMPLATE.mainTitle;
    args.instance.title = options.title;
  }
  return binding;
}

async function assertNative(context, chart) {
  chart.load('chartType'); chart.series.load('items/name');
  await sync(context, 'verify-native-bubble');
  if (chart.chartType !== 'Bubble' || chart.series.items.length !== TEMPLATE.seriesCount) {
    throw new UserError(`模板复制未保留${TEMPLATE.seriesCount}个原生Bubble系列，未认定创建成功。`);
  }
  const errors = Array.from({length: CAPACITY}, (_, i) => chart.series.getItemAt(i).yErrorBars);
  errors.forEach(error => error.load('type'));
  const xError = chart.series.getItemAt(16).xErrorBars;
  xError.load('type');
  await sync(context, 'verify-custom-error-bars');
  if (errors.some(error => error.type !== 'Custom') || xError.type !== 'Custom') {
    throw new UserError('原生自定义误差线未完整保留，未认定创建成功。');
  }
}

export async function createChart(args) {
  const {context, sourceSheet, instance, options} = args;
  if (typeof context.workbook.insertWorksheetsFromBase64 !== 'function'
    || (typeof Office !== 'undefined' && Office.context?.requirements
      && !Office.context.requirements.isSetSupported('ExcelApi', '1.19'))) {
    throw new UserError('气泡棒棒糖需要支持ExcelApi 1.19的桌面版Excel。');
  }
  const sheets = context.workbook.worksheets;
  sheets.load('items/name'); await sync(context, 'read-template-sheet-names');
  const names = new Set(sheets.items.map(sheet => sheet.name));
  let suffix = instance.id || 1;
  while (names.has(TEMPLATE_PREFIX + suffix)) suffix++;
  const templateBytes = options.negativeShade === true ? LOLLIPOP_NEGATIVE_TEMPLATE_BASE64 : LOLLIPOP_TEMPLATE_BASE64;
  const imported = context.workbook.insertWorksheetsFromBase64(templateBytes,
    {sheetNamesToInsert: [TEMPLATE.sheet], positionType: 'End'});
  await sync(context, 'import-native-template');
  // The API returns IDs, not necessarily the requested name. This also avoids
  // touching a user's pre-existing LP_Template worksheet.
  if (!Array.isArray(imported.value) || imported.value.length !== 1) throw new UserError('模板导入未返回唯一新工作表，已保留现场。');
  const template = sheets.getItem(imported.value[0]);
  template.name = TEMPLATE_PREFIX + suffix;
  template.getRange('T1:U1').values = [[instance.helperName, instance.chartName]];
  linkTemplate(template, instance.helperName);
  setCurveCarrierVisibility(template, args.binding);
  await sync(context, 'link-native-template');
  options.nativeTemplateSheet = TEMPLATE_PREFIX + suffix;
  options.nativeTemplateVersion = TEMPLATE.version;
  instance.nativeTemplateSheet = options.nativeTemplateSheet;
  const templateChart = template.charts.getItem(TEMPLATE.chartName);
  const pasted = template.shapes.getItem(TEMPLATE.chartName).copyTo(sourceSheet);
  pasted.load('name'); await sync(context, 'copy-native-chart-to-source');
  const chart = sourceSheet.charts.getItem(pasted.name);
  chart.name = instance.chartName;
  chart.width = TEMPLATE.chartDimensionsPt[0]; chart.height = TEMPLATE.chartDimensionsPt[1];
  chart.plotVisibleOnly = false;
  await assertNative(context, chart);
  // Delete only the newly imported duplicate chart; the hidden formula carrier
  // is essential to the copied chart's custom error bars and linked labels.
  templateChart.delete();
  sourceSheet.activate(); template.visibility = Excel.SheetVisibility.hidden;
  await sync(context, 'retain-hidden-formula-carrier');
  return chart;
}

function updateVisibilityAndColors({chart, source, binding, options, previous}, resetStyle) {
  const oldCount = previous?.count ?? 0, oldGroups = previous?.options?.bubbleGroups || [];
  for (let i = 0; i < CAPACITY; i++) {
    const series = chart.series.getItemAt(i), point = series.points.getItemAt(0);
    const category = chart.series.getItemAt(15).points.getItemAt(i), active = i < source.count;
    if (!active) {
      series.format.fill.clear(); series.format.line.lineStyle = 'None';
      point.format.fill.clear(); point.format.border.lineStyle = 'None';
      series.yErrorBars.visible = false;
      point.hasDataLabel = false; category.hasDataLabel = false;
      continue;
    }
    if (resetStyle || i >= oldCount || oldGroups[i] !== binding.groupKeys[i]) {
      const [color, text] = binding.palette[i] || FALLBACK_COLORS[0];
      series.format.fill.setSolidColor(color); series.format.line.lineStyle = 'None';
      point.format.fill.setSolidColor(color); point.format.border.lineStyle = 'None';
      series.yErrorBars.visible = true;
      point.hasDataLabel = options.showValues !== false; category.hasDataLabel = true;
      // hasDataLabel=false can remove the old label object. Restore the formula
      // explicitly when a capacity slot becomes active again.
      category.dataLabel.formula = `=${quote(options.nativeTemplateSheet)}!$A$${i + 3}`;
      if (options.showValues !== false) {
        point.dataLabel.formula = `=${quote(options.nativeTemplateSheet)}!$G$${i + 3}`;
        point.dataLabel.format.font.color = text;
      }
    }
  }
  // Trade-specific prose/10% reference cannot describe an arbitrary new table.
  // Preserve the carrier cells and only toggle their native series/label display.
  const tradeVisible = binding.originalTrade;
  if (resetStyle || options.bubbleOriginalTrade !== tradeVisible) {
    for (let i = 16; i < TEMPLATE.seriesCount; i++) {
      const point = chart.series.getItemAt(i).points.getItemAt(0);
      // The original 001 chart omits the worksheet's simulation footnote.
      // Keep its native role/cell for provenance, without adding it to the plot.
      const visible = tradeVisible && i !== 18 && i !== 19;
      const customCurve = i >= 20 && TEMPLATE.curveAnnotations?.length && options.nativeTemplateVersion === TEMPLATE.version;
      point.hasDataLabel = customCurve ? true : visible;
      if (visible || customCurve) point.dataLabel.formula = `=${quote(options.nativeTemplateSheet)}!$P$${i - 13}`;
    }
    chart.series.getItemAt(16).xErrorBars.visible = tradeVisible;
  }
  options.bubbleGroups = binding.groupKeys;
  options.bubbleOriginalTrade = tradeVisible;
}

export async function style(args) {
  const {context, chart, options, binding} = args;
  if (args.previous && options.nativeTemplateVersion !== TEMPLATE.version) throw new UserError('当前图表使用旧版注释样式。请选择源数据，新建带曲线注释的棒棒糖；原图将保留。');
  chart.title.visible = true; chart.title.text = args.binding.originalTrade ? `${options.title}\n${TEMPLATE.subtitle}` : options.title; chart.title.format.font.bold = false;
  chart.plotVisibleOnly = false;
  // Calibrated against the supplied 866x730 original and actual host circles.
  chart.series.getItemAt(0).bubbleScale = TEMPLATE.bubbleScale;
  if (binding.originalTrade) {
    const title = chart.axes.valueAxis.title;
    title.visible = true; title.text = '贸易变化 %';
    title.format.font.bold = false; title.format.font.name = '微软雅黑'; title.format.font.size = 10; title.format.font.color = '#000000';
  }
  updateVisibilityAndColors(args, true);
  await sync(context, 'native-bubble-style');
  return {variant: 'bubble', nativeTemplateSheet: options.nativeTemplateSheet, nativeSeriesCount: TEMPLATE.seriesCount,
    notes: ['原生气泡面积映射规模，自定义误差线连接零线。']};
}

export async function rebind(args) {
  const template = await templateFor(args);
  linkTemplate(template, args.instance.helperName);
  updateVisibilityAndColors(args, false);
  await sync(args.context, 'rebind-native-bubble-formulas');
  return {variant: 'bubble', nativeTemplateSheet: args.options.nativeTemplateSheet};
}

// Run only after the common 14/10pt typography on create/restyle. Moving native
// labels retains their worksheet formulas; the template fixes the editable box size.
// Rebind deliberately does not call this, preserving user label positions.
export async function finalizePresentation(args) {
  const {context, chart, binding} = args;
  if (!binding.originalTrade) return null;
  const targets = TEMPLATE.curveAnnotations;
  if (!Array.isArray(targets) || targets.length !== 3) throw new UserError('曲线注释模板不完整，请重新安装当前版本。');
  const template = await templateFor(args), anchors = template.getRange('M7:N9');
  const title = chart.title, plot = chart.plotArea, xAxis = chart.axes.categoryAxis, yAxis = chart.axes.valueAxis;
  const labels = [17, ...targets.map(t => t.series)].map(index => ({index, label: chart.series.getItemAt(index).points.getItemAt(0).dataLabel}));
  const byIndex = new Map(labels.map(item => [item.index, item.label]));
  chart.load('width,height');
  plot.load('insideLeft,insideTop,insideWidth,insideHeight');
  await sync(context, 'read-curve-canvas');
  const sx = chart.width / TEMPLATE.chartDimensionsPt[0], sy = chart.height / TEMPLATE.chartDimensionsPt[1];
  const plotBefore = {height: plot.insideHeight};
  const calibratedLeft = TEMPLATE.originalTradePlotLeftPt * sx;
  const calibratedWidth = TEMPLATE.originalTradePlotWidthPt * sx;
  const calibratedTop = TEMPLATE.originalTradePlotTopPt * sy;
  if (![calibratedLeft, calibratedWidth, calibratedTop].every(Number.isFinite) || calibratedWidth <= 0)
    throw new UserError('原例绘图区校准参数缺失，未认定排版完成。');
  // Absolute targets keep source 001/002 aligned without accumulating offsets
  // on restyle. Height stays unchanged; circle geometry is checked in Excel.
  plot.insideLeft = calibratedLeft; plot.insideTop = calibratedTop;
  plot.insideWidth = calibratedWidth; plot.insideHeight = plotBefore.height;
  for (const target of targets) {
    const label = byIndex.get(target.series);
    label.setWidth(target.widthPt * sx); label.setHeight(target.heightPt * sy);
    chart.series.getItemAt(target.series).showLeaderLines = false;
  }
  await sync(context, 'size-native-bezier-labels');
  chart.getImage(); await sync(context, 'flush-native-curve-layout');
  title.load('left,top,height'); plot.load('insideLeft,insideTop,insideWidth,insideHeight');
  xAxis.load('minimum,maximum'); yAxis.load('minimum,maximum'); anchors.load('values');
  labels.forEach(({label}) => label.load('left,top,width,height'));
  await sync(context, 'measure-native-curve-anchors');
  if (Math.abs(plot.insideTop - calibratedTop) > 0.25
      || Math.abs(plot.insideLeft - calibratedLeft) > 0.25
      || Math.abs(plot.insideWidth - calibratedWidth) > 0.25
      || Math.abs(plot.insideHeight - plotBefore.height) > 0.25) {
    throw new UserError('宿主未保持原例绘图区位置或尺寸，未认定排版完成。');
  }
  if (![plot.insideWidth, plot.insideHeight, xAxis.minimum, xAxis.maximum, yAxis.minimum, yAxis.maximum].every(Number.isFinite)
      || plot.insideWidth <= 0 || plot.insideHeight <= 0 || xAxis.maximum <= xAxis.minimum || yAxis.maximum <= yAxis.minimum) {
    throw new UserError('未取得有效的曲线注释坐标，未认定排版完成。');
  }
  title.left = 12; title.top = 8;
  const legend = byIndex.get(17);
  legend.horizontalAlignment = 'Left'; legend.left = chart.width * .51;
  legend.top = Math.max(8 + title.height + 8, plot.insideTop - legend.height - 8);
  targets.forEach((target, i) => {
    const label = byIndex.get(target.series), [x, y] = anchors.values[i];
    const px = plot.insideLeft + (x - xAxis.minimum) / (xAxis.maximum - xAxis.minimum) * plot.insideWidth;
    const py = plot.insideTop + (yAxis.maximum - y) / (yAxis.maximum - yAxis.minimum) * plot.insideHeight;
    label.left = px - target.headOffsetPt[0] * sx;
    label.top = py - target.headOffsetPt[1] * sy;
  });
  await sync(context, 'place-native-bezier-heads');
  labels.forEach(({label}) => label.load('left,top,width,height'));
  await sync(context, 'read-native-bezier-placement');
  return {nativeLabels: labels.map(({index, label}) => ({series:index,left:label.left,top:label.top,width:label.width,height:label.height})),
    nativeCurveGeometry:true, visualVerified:false};
}
