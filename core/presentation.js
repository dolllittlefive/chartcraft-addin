import {UserError} from './errors.js';

export const PRESENTATION_FONT = Object.freeze({name: '微软雅黑', size: 10, color: '#000000'});
export const PRESENTATION_TITLE_FONT = Object.freeze({...PRESENTATION_FONT, size: 14});
export const PRESENTATION_SUBTITLE_FONT = PRESENTATION_FONT;
export const PRESENTATION_BORDER = Object.freeze({lineStyle: 'Continuous', color: '#D9D9D9', weight: 0.75});
export const PREVIEW_LAYOUT = Object.freeze({leftGutterPt: 28, gapColumnPt: 24, topRowPt: 16, columnPaddingPt: 6});
const plotKeys = ['insideLeft', 'insideTop', 'insideWidth', 'insideHeight'];
const fontKeys = 'name,size,color';
const readPlot = plot => Object.fromEntries(plotKeys.map(key => [key, plot[key]]));
const samePlot = (a, b) => plotKeys.every(key => Math.abs(a[key] - b[key]) <= 0.25);
const normalizeColor = color => String(color).replace(/^#/, '').toLowerCase().replace(/^black$/, '000000').replace(/^white$/, 'ffffff').replace(/^red$/, 'ff0000');
const fontMatches = (font, expected = PRESENTATION_FONT) => ['微软雅黑', 'microsoftyahei'].includes(String(font.name).toLowerCase().replace(/\s/g, ''))
  && Number.isFinite(font.size) && Math.abs(font.size - expected.size) < 0.01
  && normalizeColor(font.color) === normalizeColor(expected.color);

function setFont(font, expected = PRESENTATION_FONT) {
  font.name = expected.name;
  font.size = expected.size;
  font.color = expected.color;
}

// Call from a recipe BEFORE its title/label bbox measurement, then context.sync()
// (and its existing getImage/sync layout flush) before reading fresh geometry.
// text is supplied explicitly, so this helper never reads an unloaded proxy or
// rewrites a title/formula. A multiline main title can supply its exact length.
export function titleTypographyRuns(text, mainTitleLength) {
  if (typeof text !== 'string') throw new UserError('标题字体需要明确的标题文本。');
  const firstBreak = text.search(/[\r\n]/);
  const mainLength = mainTitleLength ?? (firstBreak < 0 ? text.length : firstBreak);
  if (!Number.isInteger(mainLength) || mainLength < 0 || mainLength > text.length
      || (mainLength > 0 && mainLength < text.length && !/[\r\n]/.test(text[mainLength]))) {
    throw new UserError('主标题长度必须完整覆盖主标题，并在副标题换行之前结束。');
  }
  let subtitleStart = mainLength;
  while (subtitleStart < text.length && /[\r\n]/.test(text[subtitleStart])) subtitleStart++;
  return [
    ...(mainLength ? [{role: 'main', start: 0, length: mainLength, font: PRESENTATION_TITLE_FONT}] : []),
    ...(subtitleStart < text.length ? [{role: 'subtitle', start: subtitleStart, length: text.length - subtitleStart, font: PRESENTATION_SUBTITLE_FONT}] : [])
  ];
}

export function setChartTitleTypography(chart, {text, mainTitleLength} = {}) {
  const runs = titleTypographyRuns(text, mainTitleLength);
  setFont(chart.title.format.font, PRESENTATION_TITLE_FONT);
  for (const run of runs) setFont(chart.title.getSubstring(run.start, run.length).font, run.font);
  return runs;
}

function preservedLabelColor(text, color, explicitGraphic = false) {
  const hex = normalizeColor(color);
  if (!/^[0-9a-f]{6}$/.test(hex)) return undefined;
  if (explicitGraphic || ['■', '│', '▲', '▼'].includes(text)) return {color: '#' + hex, role: 'graphic'};
  // White is an explicit recipe contrast choice on a dark mark. Ordinary grey,
  // blue or pink prose becomes black; a negative value keeps an existing red.
  if (hex === 'ffffff') return {color: '#FFFFFF', role: 'contrast'};
  const negative = /^\s*[-−﹣－]\s*\d/.test(text) || /^\s*\(\s*\d[^)]*\)\s*%?\s*$/.test(text);
  const [red, green, blue] = hex.match(/../g).map(value => parseInt(value, 16));
  if (negative && red >= 64 && red > green * 1.5 && red > blue * 1.5) return {color: '#' + hex, role: 'negative'};
  return undefined;
}

async function sync(context, stage) {
  const mark = state => {
    try { globalThis.CCGenerateDiagnostics?.mark?.('presentation:sync:' + state, {stage}); }
    catch { /* Optional acceptance diagnostics must not affect host errors. */ }
  };
  mark('start');
  try { await context.sync(); }
  catch (cause) {
    mark('rejected');
    const error = new UserError(`统一格式未验证（${stage}）：${cause?.message || String(cause)}${cause?.debugInfo?.errorLocation ? ' [' + cause.debugInfo.errorLocation + ']' : ''}`);
    error.cause = cause; error.code = cause?.code; error.debugInfo = cause?.debugInfo;
    error.stage = stage;
    throw error;
  }
  mark('resolved');
}

function labelState(series) {
  return series.map(s => ({series: s.hasDataLabels, points: s.points.items.map(p => p.hasDataLabel)}));
}

// Doughnut and Pie series have no axes. A mixed group still needs axes if
// any other series does. Unknown types retain the established behavior;
// this narrow rule is not an API-error fallback or a generic chart taxonomy.
export function axisGroupsForSeries(series) {
  const axislessTypes = new Set(['Doughnut', 'DoughnutExploded', 'Pie', 'PieExploded']);
  return ['Primary', 'Secondary'].filter(group => series.some(s => s.axisGroup === group && !axislessTypes.has(s.chartType)));
}

// Call after the recipe's create/style, inside its existing Excel.run.
// No show*, hasDataLabel*, text, formula, position, width or height setters are used.
// Point.hasDataLabel is read before obtaining point.dataLabel: hidden spacer
// points must not acquire labels just because typography is being standardized.
export async function applyChartPresentation({context, chart, mainTitleLength, graphicLabels = []}) {
  chart.series.load('items/name,items/chartType,items/axisGroup,items/hasDataLabels');
  chart.title.load('visible,text');
  chart.legend.load('visible');
  chart.plotArea.load(plotKeys.join(','));
  await sync(context, '读取原图字体范围');
  const beforePlot = readPlot(chart.plotArea);
  if (plotKeys.some(key => !Number.isFinite(beforePlot[key])) || beforePlot.insideWidth <= 0 || beforePlot.insideHeight <= 0) {
    throw new UserError('图表未返回有效plotInside，未统一字体。');
  }
  const series = chart.series.items;
  const groups = axisGroupsForSeries(series);
  const axes = [];
  for (const group of groups) {
    for (const type of ['Category', 'Value']) {
      const axis = chart.axes.getItem(type, group);
      axis.load('visible'); axis.title.load('visible');
      axes.push({axis, name: `${group}/${type}`});
    }
  }
  for (const s of series) s.points.load('items/hasDataLabel');
  await sync(context, '读取已有标签开关');
  const beforeLabels = labelState(series);
  // Recipes can identify native graphical labels even when their linked text
  // is currently empty (e.g. a zero-growth arrow that later becomes positive).
  const explicitGraphics = new Set();
  if (!Array.isArray(graphicLabels)) throw new UserError('图形标签引用必须是数组。');
  for (const ref of graphicLabels) {
    if (!Number.isInteger(ref?.series) || !Number.isInteger(ref?.point) || ref.series < 0 || ref.point < 0
        || !series[ref.series]?.points.items[ref.point]?.hasDataLabel) throw new UserError('图形标签引用无效或未启用。');
    explicitGraphics.add(series[ref.series].points.items[ref.point]);
  }
  // Capture explicit semantic/contrast colors before chart/series font setters
  // can overwrite inherited point colors. Never infer a new color from values.
  const preservedColors = new Map();
  for (const s of series) for (const point of s.points.items) if (point.hasDataLabel === true) {
    point.dataLabel.load('text'); point.dataLabel.format.font.load('color');
  }
  await sync(context, '读取标签语义色与图例符号颜色');
  for (const s of series) for (const point of s.points.items) if (point.hasDataLabel === true) {
    const label = point.dataLabel;
    const preserve = preservedLabelColor(label.text, label.format.font.color, explicitGraphics.has(point));
    if (preserve) preservedColors.set(point, preserve);
  }
  const fonts = [];
  const useFont = (font, scope, expected = PRESENTATION_FONT) => { setFont(font, expected); fonts.push({font, scope, expected}); };
  useFont(chart.format.font, 'chart');
  Object.assign(chart.format.border, PRESENTATION_BORDER);
  let titleRuns = [];
  if (chart.title.visible) {
    titleRuns = setChartTitleTypography(chart, {text: chart.title.text, mainTitleLength});
    // A mixed 14/10 title may not expose one aggregate font size. Read each
    // intended run; do not incorrectly assert a uniform whole-title font.
    for (const run of titleRuns) fonts.push({font: chart.title.getSubstring(run.start, run.length).font,
      scope: `title/${run.role}`, expected: run.font});
    if (!titleRuns.length) fonts.push({font: chart.title.format.font, scope: 'title', expected: PRESENTATION_TITLE_FONT});
  }
  if (chart.legend.visible) useFont(chart.legend.format.font, 'legend');
  for (const {axis, name} of axes) {
    // Excel rejects ChartFont setters for deleted/hidden secondary axes.
    // Keep their visibility intact; only displayed axis text needs formatting.
    if (!axis.visible) continue;
    useFont(axis.format.font, `axis/${name}`);
    if (axis.title.visible) useFont(axis.title.format.font, `axis-title/${name}`);
  }
  let enabledPoints = 0;
  series.forEach((s, i) => {
    if (s.hasDataLabels === true) useFont(s.dataLabels.format.font, `series/${i}`);
    s.points.items.forEach((point, j) => {
      if (point.hasDataLabel !== true) return;
      useFont(point.dataLabel.format.font, `point/${i}/${j}`, {...PRESENTATION_FONT,
        color: preservedColors.get(point)?.color || PRESENTATION_FONT.color});
      enabledPoints++;
    });
  });
  chart.plotArea.load(plotKeys.join(','));
  chart.format.border.load('lineStyle,color,weight');
  for (const {font} of fonts) font.load(fontKeys);
  chart.series.load('items/hasDataLabels');
  for (const s of series) s.points.load('items/hasDataLabel');
  await sync(context, '回读字体与标签开关');
  const invalidFonts = fonts.filter(({font, expected}) => !fontMatches(font, expected))
    .map(({font, scope, expected}) => ({scope, name: font.name, size: font.size, color: font.color, expected}));
  if (invalidFonts.length) {
    const error = new UserError(`宿主未保持指定字体：${invalidFonts[0].scope}。未认定统一格式通过。`);
    error.details = {invalidFonts};
    throw error;
  }
  if (JSON.stringify(beforeLabels) !== JSON.stringify(labelState(series))) {
    throw new UserError('字体设置后标签显隐状态发生变化，未认定统一格式通过。');
  }
  const border = chart.format.border;
  if (border.lineStyle !== PRESENTATION_BORDER.lineStyle || normalizeColor(border.color) !== normalizeColor(PRESENTATION_BORDER.color)
      || !Number.isFinite(border.weight) || Math.abs(border.weight - PRESENTATION_BORDER.weight) > 0.01) {
    const error = new UserError('宿主未保持全图浅灰外边框，未认定统一格式通过。');
    error.details = {border: {lineStyle: border.lineStyle, color: border.color, weight: border.weight}, expected: PRESENTATION_BORDER};
    throw error;
  }
  // Font metrics can make Excel recompute the plot area. Preserve the recipe's
  // measured geometry, but never restore old label width/height after resizing text.
  const fontChangedPlot = !samePlot(beforePlot, readPlot(chart.plotArea));
  if (fontChangedPlot) {
    for (const key of plotKeys) chart.plotArea[key] = beforePlot[key];
    await sync(context, '保留模块plotInside');
    chart.plotArea.load(plotKeys.join(','));
    await sync(context, '回读模块plotInside');
  }
  const afterPlot = readPlot(chart.plotArea);
  if (!samePlot(beforePlot, afterPlot)) {
    const error = new UserError('字体设置后plotInside未能保持，需实图核对；未认定版式通过。');
    error.details = {beforePlot, afterPlot};
    throw error;
  }
  const roles = [...preservedColors.values()].map(value => value.role);
  return {font: PRESENTATION_FONT, titleFont: PRESENTATION_TITLE_FONT, subtitleFont: PRESENTATION_SUBTITLE_FONT, titleRuns,
    border: PRESENTATION_BORDER, enabledPointLabels: enabledPoints, fontScopes: fonts.length,
    preservedGraphicSwatches: roles.filter(role => role === 'graphic').length,
    preservedNegativeLabels: roles.filter(role => role === 'negative').length,
    preservedContrastLabels: roles.filter(role => role === 'contrast').length,
    labelVisibilityPreserved: true, plotRestored: fontChangedPlot, beforePlot, afterPlot,
    verification: 'current-host-property-readback', visualVerified: false, crossHostVerified: false};
}

// Dedicated preview sheets only. The caller supplies the existing B2 range.
// Run before chart creation (or after Absolute placement), since column sizing
// otherwise resizes an existing TwoCell chart. Never relocates or rewrites data.
export async function formatPreviewSource({context, sourceSheet, range}) {
  range.load('address,rowIndex,columnIndex,rowCount,columnCount');
  await sync(context, '读取专用预览范围');
  if (range.rowIndex !== 1 || range.columnIndex !== 1 || range.rowCount < 1 || range.columnCount < 1) {
    throw new UserError('专用预览格式仅接受B2起始的已有数据区，不自动搬迁业务数据。');
  }
  const {rowIndex, columnIndex, rowCount, columnCount} = range;
  const gapColumn = columnIndex + columnCount;
  const header = sourceSheet.getRangeByIndexes(rowIndex, columnIndex, 1, columnCount);
  const columns = Array.from({length: columnCount}, (_, i) => sourceSheet.getRangeByIndexes(rowIndex, columnIndex + i, rowCount, 1));
  setFont(range.format.font);
  range.format.font.bold = false;
  range.format.font.italic = false;
  range.format.horizontalAlignment = 'General';
  range.format.verticalAlignment = 'Center';
  range.format.wrapText = false;
  range.format.shrinkToFit = false;
  range.format.fill.clear();
  header.format.font.bold = true;
  header.format.horizontalAlignment = 'Center';
  header.format.fill.color = '#D9D9D9';
  const edges = ['EdgeLeft', 'EdgeRight', 'EdgeTop', 'EdgeBottom'];
  if (rowCount > 1) edges.push('InsideHorizontal');
  if (columnCount > 1) edges.push('InsideVertical');
  for (const edge of edges) {
    const border = range.format.borders.getItem(edge);
    border.style = 'Continuous'; border.weight = 'Thin'; border.color = '#000000';
  }
  sourceSheet.showGridlines = false;
  await sync(context, '专用预览字体表头边框');
  range.format.autofitColumns();
  range.format.autofitRows();
  await sync(context, '按10号字完整适配数据');
  for (const column of columns) column.format.load('columnWidth');
  await sync(context, '读取数据列point宽');
  const columnWidthsPt = columns.map(column => column.format.columnWidth);
  if (columnWidthsPt.some(width => !Number.isFinite(width) || width <= 0)) {
    throw new UserError('自动列宽未返回有效point值，未确认专用预览格式。');
  }
  // Freeze explicit point dimensions, with a small font-metric safety margin.
  // This is not the OOXML character-width number (such as 36.83203125).
  columns.forEach((column, i) => {
    columnWidthsPt[i] = Math.ceil(columnWidthsPt[i] * 4) / 4 + PREVIEW_LAYOUT.columnPaddingPt;
    column.format.columnWidth = columnWidthsPt[i];
  });
  sourceSheet.getRangeByIndexes(0, 0, 1, 1).format.columnWidth = PREVIEW_LAYOUT.leftGutterPt;
  sourceSheet.getRangeByIndexes(0, gapColumn, 1, 1).format.columnWidth = PREVIEW_LAYOUT.gapColumnPt;
  sourceSheet.getRangeByIndexes(0, columnIndex, 1, columnCount).format.rowHeight = PREVIEW_LAYOUT.topRowPt;
  await sync(context, '固定预览锚点前置尺寸');
  for (const column of columns) column.format.load('columnWidth');
  sourceSheet.load('showGridlines');
  range.format.font.load(fontKeys);
  header.format.font.load('bold');
  header.format.fill.load('color');
  await sync(context, '回读专用预览格式');
  if (!fontMatches(range.format.font) || !header.format.font.bold || sourceSheet.showGridlines !== false
      || String(header.format.fill.color).toUpperCase() !== '#D9D9D9') {
    throw new UserError('宿主未保持专用预览字体/表头/网格线设置，未认定格式通过。');
  }
  return {address: range.address, font: PRESENTATION_FONT, headerFill: '#D9D9D9', border: 'Continuous/Thin/#000000',
    columnWidthsPt: columns.map(column => column.format.columnWidth), ...PREVIEW_LAYOUT,
    chartAnchor: {rowIndex, columnIndex: gapColumn + 1}, sourceMoved: false,
    verification: 'current-host-property-readback', visualVerified: false, crossHostVerified: false};
}
