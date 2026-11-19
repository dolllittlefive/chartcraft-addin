import {UserError} from './errors.js';

// Microsoft uses charts.add + shapes.getItem(chart.name) in shape-get-active.yaml.
// Shape.placement requires ExcelApi 1.10 (the Placement enum itself is 1.9).
// https://learn.microsoft.com/javascript/api/excel/excel.shape#placement
// https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/excel/44-shape/shape-get-active.yaml
export const PLACEMENT_MIN_EXCEL_API = '1.10';
const geometryKeys = ['left', 'top', 'width', 'height'];
const chartProperties = 'id,name,left,top,width,height';
const shapeProperties = chartProperties + ',type,placement';
const read = (object, keys) => Object.fromEntries(keys.map(key => [key, object[key]]));
const sameGeometry = (a, b) => geometryKeys.every(key => Number.isFinite(a[key]) && Number.isFinite(b[key]) && Math.abs(a[key] - b[key]) <= 0.01);

export class ChartPlacementError extends UserError {
  constructor(code, stage, message, details = {}, cause) {
    super(message);
    this.name = 'ChartPlacementError';
    this.code = code;
    this.stage = stage;
    this.details = details;
    if (cause) this.cause = cause;
  }
}

function requireApi() {
  const requirements = globalThis.Office?.context?.requirements;
  if (typeof requirements?.isSetSupported !== 'function' || !requirements.isSetSupported('ExcelApi', PLACEMENT_MIN_EXCEL_API)) {
    throw new ChartPlacementError('PLACEMENT_API_UNAVAILABLE', 'api', '固定图表尺寸需要 ExcelApi 1.10，当前宿主未确认支持；未设置锚定。');
  }
}

async function host(stage, action, details = {}) {
  try { return await action(); }
  catch (cause) {
    if (cause instanceof ChartPlacementError) throw cause;
    throw new ChartPlacementError('PLACEMENT_HOST_ERROR', stage,
      `图表锚定未验证（${stage}）：${cause?.message || String(cause)}`,
      {...details, hostCode: cause?.code, errorLocation: cause?.debugInfo?.errorLocation}, cause);
  }
}

function snapshot(chart, shape) {
  return {chart: read(chart, chartProperties.split(',')), shape: read(shape, shapeProperties.split(','))};
}

function requireMatch(evidence, stage) {
  const {chart, shape} = evidence;
  // ShapeType has no Chart member. Keep its actual value as evidence, including Unsupported.
  if (!chart.id || !shape.id || !chart.name || chart.name !== shape.name || !sameGeometry(chart, shape) || chart.width <= 0 || chart.height <= 0) {
    throw new ChartPlacementError('PLACEMENT_SHAPE_MISMATCH', stage, '原生图表与同名 Shape 的身份或几何不一致，未确认固定尺寸。', evidence);
  }
  if (!['Absolute', 'OneCell', 'TwoCell'].includes(shape.placement)) {
    throw new ChartPlacementError('PLACEMENT_VALUE_UNAVAILABLE', stage, '宿主未返回可识别的 Shape.placement，未确认固定尺寸。', evidence);
  }
}

async function resolve({context, sourceSheet, chart}) {
  requireApi();
  await host('read-chart', async () => { chart.load(chartProperties); await context.sync(); });
  // Resolve through both collections so a drawing never replaces the native Chart.
  let nativeChart, shape;
  await host('resolve-shape', async () => {
    nativeChart = sourceSheet.charts.getItem(chart.name);
    shape = sourceSheet.shapes.getItem(chart.name);
    nativeChart.load(chartProperties); shape.load(shapeProperties);
    await context.sync();
  }, {chartName: chart.name});
  const before = snapshot(nativeChart, shape);
  requireMatch(before, 'resolve-shape');
  if (before.chart.id !== chart.id) throw new ChartPlacementError('PLACEMENT_CHART_MISMATCH', 'resolve-shape', '指定工作表中的图表与传入原生 Chart 不一致，未设置锚定。', before);
  return {shape, before};
}

// Read-only probe. Call inside an existing Excel.run; do not nest Excel.run here.
export async function inspectChartPlacement(args) {
  return (await resolve(args)).before;
}

// Call after the recipe has set the final geometry. Only placement is written;
// no shapes/charts are added, grouped, converted, resized, deleted, or saved.
export async function setChartPlacement(args, placement = 'Absolute') {
  if (!['Absolute', 'OneCell'].includes(placement)) {
    throw new ChartPlacementError('PLACEMENT_INVALID_ARGUMENT', 'argument', '固定尺寸锚定仅支持 Absolute 或 OneCell。');
  }
  const {context, sourceSheet} = args;
  const {shape, before} = await resolve(args);
  await host('set-placement', async () => { shape.placement = placement; await context.sync(); }, {before, requestedPlacement: placement, writeAttempted: true});
  let after;
  await host('read-back', async () => {
    const chartAgain = sourceSheet.charts.getItem(before.chart.name);
    const shapeAgain = sourceSheet.shapes.getItem(before.chart.name);
    chartAgain.load(chartProperties); shapeAgain.load(shapeProperties);
    await context.sync();
    after = snapshot(chartAgain, shapeAgain);
  }, {before, requestedPlacement: placement, writeAttempted: true});
  requireMatch(after, 'read-back');
  if (after.shape.placement !== placement || after.chart.id !== before.chart.id || after.shape.id !== before.shape.id || !sameGeometry(before.chart, after.chart)) {
    throw new ChartPlacementError('PLACEMENT_NOT_VERIFIED', 'read-back', '宿主读回的锚定、原生图表身份或尺寸与预期不符；可能已有局部修改，未确认固定尺寸。', {before, after, requestedPlacement: placement, writeAttempted: true});
  }
  return {requestedPlacement: placement, before, after, verification: 'current-host-readback', persistenceVerified: false, crossHostVerified: false};
}
