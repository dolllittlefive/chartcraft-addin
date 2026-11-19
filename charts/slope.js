import {UserError} from '../core/errors.js';

const MARK = 'ChartCraft.Slope.v1';
const HELPER_PREFIX = 'CC_斜率_';
const CHART_PREFIX = 'CC_Slope_';
const MAX_ROWS = 200, MAX_PERIODS = 12, FIXED_SERIES = 9;
// R6 user typography takes precedence over the PDF's original typography.
// Apply the final font before measuring native label boxes for spacing.
const FONT = '微软雅黑';
const TYPOGRAPHY = {title: {name: FONT, size: 14}, endpoints: {name: FONT, size: 10}, changes: {name: FONT, size: 10}, periods: {name: FONT, size: 10}};
// PDF 018 original image samples; order follows the original five source rows.
const COLORS = ['#323FCA', '#0A1C2C', '#4E95B4', '#40A1ED', '#A2E5FB'];
// Preserve image proportions (722 × 434). 0.75 pt/px is a working scale, not an original workbook measurement.
const LAYOUT = {width: 541.5, height: 325.5, plot: {left: 47.25, top: 69.75, width: 457.5, height: 237.3}};
const quote = name => `'${String(name).replace(/'/g, "''")}'`;
function columnLetter(index) { let s = '', n = index + 1; while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; }

function validate(values) {
  if (!Array.isArray(values) || values.length < 3 || !Array.isArray(values[0])) throw new UserError('请选择含表头的数据：第一列分类，其后至少两期数值，至少 2 行数据。');
  const columns = values[0].length, periods = columns - 1;
  if (periods < 2 || periods > MAX_PERIODS) throw new UserError(`需要分类列及 2–${MAX_PERIODS} 期数值，共 3–${MAX_PERIODS + 1} 列。`);
  if (values.length - 1 > MAX_ROWS) throw new UserError(`最多支持 ${MAX_ROWS} 行数据。`);
  const text = value => (typeof value === 'string' || typeof value === 'number') && String(value).trim() !== '';
  if (!values[0].every(text)) throw new UserError('每列都需要非空表头。');
  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    if (!Array.isArray(row) || row.length !== columns) throw new UserError(`第 ${i + 1} 行列数不一致。`);
    if (!text(row[0])) throw new UserError(`第 ${i + 1} 行分类为空或无效。`);
    for (let j = 1; j < columns; j++) if (typeof row[j] !== 'number' || !Number.isFinite(row[j])) throw new UserError(`第 ${i + 1} 行第 ${j + 1} 列必须是数字；缺失值不能代填为 0。`);
  }
  return {count: values.length - 1, periods, categoryHeader: String(values[0][0]), periodHeaders: values[0].slice(1).map(String)};
}

const inputSpec = {
  category: {id: 'comparison', title: '比较'},
  description: '首列分类，其后 2–12 期数值，首行为表头。百分比输入用 0.46 表示 46%。原生线条、端点和变化标签随单元格自动重算；多期变化比较首末两期。',
  minVersion: '1.9', columns: 3, minRows: 2, maxRows: MAX_ROWS,
  options: [
    {id: 'title', label: '标题', type: 'text', default: '不同时期的数值变化', maxLength: 80},
    {id: 'variant', label: '变化标签位置', type: 'select', default: 'right-column', choices: [{value: 'right-column', label: '右侧变化列'}, {value: 'on-line', label: '线上（水平文字）'}, {value: 'endpoints', label: '仅端点'}]},
    {id: 'changeMode', label: '变化口径', type: 'select', default: 'relative', choices: [{value: 'relative', label: '相对变化 (末−首)/ABS(首)'}, {value: 'difference', label: '绝对差 末−首'}, {value: 'percentage-points', label: '百分点 (末−首)×100'}]},
    {id: 'valueFormat', label: '端点数值格式', type: 'select', default: 'number', choices: [{value: 'number', label: '数值'}, {value: 'percent', label: '百分比'}]},
    {id: 'valueDecimals', label: '端点小数位', type: 'number', default: 0, min: 0, max: 3, step: 1},
    {id: 'changeDecimals', label: '变化小数位', type: 'number', default: 1, min: 0, max: 3, step: 1},
    {id: 'categoryLabels', label: '分类名称', type: 'select', default: 'both', choices: [{value: 'both', label: '两端'}, {value: 'left', label: '仅左端'}, {value: 'right', label: '仅右端'}]},
    {id: 'showMarkers', label: '显示节点圆点', type: 'checkbox', default: false},
    {id: 'showGuides', label: '显示两端竖导线', type: 'checkbox', default: true},
    {id: 'showReference', label: '显示水平参考线', type: 'checkbox', default: false},
    {id: 'referenceValue', label: '参考值（与源数据同单位）', type: 'number', default: 50, min: -1000000000, max: 1000000000, step: 1}
  ],
  storage: {marker: MARK, helperPrefix: HELPER_PREFIX, chartPrefix: CHART_PREFIX}
};

function normalizeOptions(options = {}) {
  const out = Object.fromEntries(inputSpec.options.map(option => [option.id, options[option.id] ?? option.default]));
  for (const option of inputSpec.options) {
    const value = out[option.id];
    if (option.type === 'select' && !option.choices.some(c => c.value === value)) throw new UserError(`无效选项：${option.label}`);
    if (option.type === 'number' && (typeof value !== 'number' || !Number.isFinite(value) || value < option.min || value > option.max)) throw new UserError(`无效数值：${option.label}`);
    if (option.type === 'checkbox' && typeof value !== 'boolean') throw new UserError(`无效开关：${option.label}`);
  }
  for (const id of ['valueDecimals', 'changeDecimals']) if (!Number.isInteger(out[id])) throw new UserError('小数位数必须为整数。');
  if (typeof out.title !== 'string' || out.title.length > 80) throw new UserError('标题最多 80 个字符。');
  return out;
}

/** Pure formula plan. All business values and labels refer to source cells, never cached JS values. */
function makeRecipe(source, options) {
  const o = normalizeOptions(options), n = source.count, p = source.periods ?? source.columnCount - 1, last = n + 2;
  const sheet = quote(source.sheet), rawEnd = columnLetter(16 + p - 1), normEnd = columnLetter(29 + p - 1);
  const rawRange = `$Q$3:$${rawEnd}$${last}`, ref = (column, row) => `${sheet}!$${columnLetter(column)}$${row}`;
  const decimals = d => '0' + (d ? '.' + '0'.repeat(d) : '');
  const valueMask = decimals(o.valueDecimals) + (o.valueFormat === 'percent' ? '%' : '');
  const suffix = o.changeMode === 'relative' ? '%' : o.changeMode === 'percentage-points' ? '" pp"' : '';
  // TEXT's literal suffix quotes need doubling inside the worksheet formula string.
  const changeMask = ['+', '-', ''].map(sign => sign + decimals(o.changeDecimals) + suffix).join(';').replace(/"/g, '""');
  const delta = r => o.changeMode === 'relative' ? `(C${r}-B${r})/ABS(B${r})` : o.changeMode === 'percentage-points' ? `(C${r}-B${r})*100` : `C${r}-B${r}`;
  const blocks = [];
  blocks.push({address: 'B1:H1', formulas: [[
    '=INT(MIN(0,IFERROR(AGGREGATE(15,6,' + rawRange + ',1),0))/$G$1)*$G$1',
    '=MAX(B1+G1,-INT(-MAX(0,IFERROR(AGGREGATE(14,6,' + rawRange + ',1),0))/G1)*G1)',
    '=C1-B1', `=IFERROR(${ref(source.columnIndex + 1, source.rowIndex + 1)}&"","")`,
    `=IFERROR(${ref(source.columnIndex + p, source.rowIndex + 1)}&"","")`,
    '=IFERROR(10^INT(LOG10(MAX(ABS(AGGREGATE(15,6,' + rawRange + ',1)),ABS(AGGREGATE(14,6,' + rawRange + ',1))))),1)',
    String(p)
  ]]});
  blocks.push({address: 'A2:O2', formulas: [['分类', '首期', '末期', '端点有效', '左标签', '右标签', '变化', '左Y', '右Y', '左X', '右X', '变化X', '变化Y', '非负标签', '负标签']]});
  const rows = [], raw = [], normalized = [];
  for (let i = 0; i < n; i++) {
    const r = i + 3, sr = source.rowIndex + i + 2;
    raw.push(Array.from({length: p}, (_, j) => {
      const cell = ref(source.columnIndex + j + 1, sr);
      return `=IFERROR(IF(ISNUMBER(${cell}),${cell},NA()),NA())`;
    }));
    normalized.push(Array.from({length: p}, (_, j) => `=IFERROR((${columnLetter(16 + j)}${r}-$B$1)/$D$1,NA())`));
    const left = `TEXT(B${r},"${valueMask}")`, right = `TEXT(C${r},"${valueMask}")`;
    const change = `TEXT(${delta(r)},"${changeMask}")`;
    const mid = p % 2 ? `${columnLetter(29 + (p - 1) / 2)}${r}` : `AVERAGE(${columnLetter(29 + p / 2 - 1)}${r}:${columnLetter(29 + p / 2)}${r})`;
    rows.push([
      `=IFERROR(${ref(source.columnIndex, sr)}&"","")`, '=Q' + r, `=${rawEnd}${r}`, `=AND(ISNUMBER(B${r}),ISNUMBER(C${r}))`,
      `=IFERROR(IF(D${r},${o.categoryLabels === 'right' ? left : `A${r}&"  "&${left}`},""),"")`,
      `=IFERROR(IF(D${r},${o.categoryLabels === 'left' ? right : `${right}&"  "&A${r}`},""),"")`,
      `=IFERROR(IF(D${r},${o.changeMode === 'relative' ? `IF(B${r}=0,"—",${change})` : change},""),"")`,
      // Extra anchor spacing approximates the original 18 px endpoint-to-label gap;
      // native label padding still needs host measurement. Changes use a common right edge.
      `=IFERROR(AD${r},0)`, `=IFERROR(${normEnd}${r},0)`, o.variant === 'right-column' ? '-0.045' : '-0.035', o.variant === 'right-column' ? '1.045' : '1.035', o.variant === 'right-column' ? '1.84' : '0.5',
      `=IFERROR(${o.variant === 'on-line' ? mid : `I${r}`},0)`,
      o.variant === 'endpoints' ? '=""' : `=IFERROR(IF(AND(C${r}>=B${r},${o.variant === 'on-line' ? `ISNUMBER(${mid})` : `D${r}`}),G${r},""),"")`,
      o.variant === 'endpoints' ? '=""' : `=IFERROR(IF(AND(C${r}<B${r},${o.variant === 'on-line' ? `ISNUMBER(${mid})` : `D${r}`}),G${r},""),"")`
    ]);
  }
  blocks.push({address: `A3:O${last}`, formulas: rows}, {address: `Q3:${rawEnd}${last}`, formulas: raw},
    {address: `AD2:${normEnd}2`, formulas: [Array.from({length: p}, (_, j) => String(j / (p - 1)))]},
    {address: `AD3:${normEnd}${last}`, formulas: normalized});
  blocks.push({address: 'AQ3:AX4', formulas: [
    ['0', o.showGuides ? '0' : '=NA()', '1', o.showGuides ? '0' : '=NA()', '-0.5', '0', '0', o.showReference ? `=(${o.referenceValue}-$B$1)/$D$1` : '=NA()'],
    ['0', o.showGuides ? '1' : '=NA()', '1', o.showGuides ? '1' : '=NA()', o.variant === 'right-column' ? '1.5' : '1.497', '0', '1', o.showReference ? `=(${o.referenceValue}-$B$1)/$D$1` : '=NA()']
  ]});
  blocks.push({address: `AZ3:BB${p + 2}`, formulas: Array.from({length: p}, (_, j) => [String(j / (p - 1)), '-0.065', `=IFERROR(${ref(source.columnIndex + j + 1, source.rowIndex + 1)}&"","")`])});
  return {blocks, last, periods: p, normEnd, options: o};
}

async function writeHelper({context, helper, source, options, previous, chart}) {
  const binding = makeRecipe(source, options);
  let oldPeriods = 0;
  if (previous) {
    // Protect user-modified topology before touching helper cells. Individual formatting is allowed.
    chart.series.load('items'); await context.sync();
    if (chart.series.items.length !== FIXED_SERIES + previous.count) throw new UserError('图表系列数量已被手工修改；为保留现有内容，本次未重绑。');
    const headers = chart.series.getItemAt(4).points; headers.load('items'); await context.sync(); oldPeriods = headers.items.length;
    helper.getRange(`A2:BB${Math.max(previous.count, source.count, MAX_PERIODS) + 2}`).clear(Excel.ClearApplyTo.contents);
  }
  for (const block of binding.blocks) helper.getRange(block.address).formulas = block.formulas;
  return {...binding, oldPeriods};
}

const fixedNames = ['左端标签', '右端标签', '非负变化', '负变化', '期间', '左导线', '右导线', '基线', '参考线'];
function bindSeries(chart, helper, binding, count) {
  const ranges = [['J', 'H'], ['K', 'I'], ['L', 'M'], ['L', 'M']].map(([x, y]) => [`${x}3:${x}${binding.last}`, `${y}3:${y}${binding.last}`]);
  ranges.push([`AZ3:AZ${binding.periods + 2}`, `BA3:BA${binding.periods + 2}`], ['AQ3:AQ4', 'AR3:AR4'], ['AS3:AS4', 'AT3:AT4'], ['AU3:AU4', 'AV3:AV4'], ['AW3:AW4', 'AX3:AX4']);
  for (let i = 0; i < count; i++) ranges.push([`AD2:${binding.normEnd}2`, `AD${i + 3}:${binding.normEnd}${i + 3}`]);
  ranges.forEach(([x, y], i) => { const series = chart.series.getItemAt(i); series.setXAxisValues(helper.getRange(x)); series.setValues(helper.getRange(y)); });
}

function styleEntity(series, index, options) {
  series.format.line.color = COLORS[index % COLORS.length]; series.format.line.weight = 1.25; series.format.line.lineStyle = 'Continuous';
  series.markerStyle = options.showMarkers ? Excel.ChartMarkerStyle.circle : Excel.ChartMarkerStyle.none;
  if (options.showMarkers) { series.markerSize = 5; series.markerForegroundColor = COLORS[index % COLORS.length]; series.markerBackgroundColor = '#FFFFFF'; }
}

// Least-squares packing in the existing vertical order. Only colliding clusters
// move; true series coordinates and formula-linked text never change. Bounds
// prevent a dense user dataset from being pushed outside the existing chart.
function spaceLabelBoxes(boxes, topLimit, bottomLimit, gap = 2) {
  const ordered = boxes.map((box, index) => ({...box, index})).sort((a, b) => a.top - b.top || a.index - b.index);
  let offset = 0;
  const blocks = [];
  for (let i = 0; i < ordered.length; i++) {
    const box = ordered[i]; box.offset = offset;
    blocks.push({first: i, last: i, sum: box.top - offset, count: 1});
    offset += box.height + gap;
    while (blocks.length > 1) {
      const b = blocks.at(-1), a = blocks.at(-2);
      if (a.sum / a.count <= b.sum / b.count) break;
      blocks.splice(-2, 2, {first: a.first, last: b.last, sum: a.sum + b.sum, count: a.count + b.count});
    }
  }
  const requiredHeight = ordered.length ? offset - gap : 0;
  if (requiredHeight > bottomLimit - topLimit) return null;
  const tops = Array(boxes.length);
  for (const block of blocks) {
    const start = Math.max(topLimit, Math.min(bottomLimit - requiredHeight, block.sum / block.count));
    for (let i = block.first; i <= block.last; i++) tops[ordered[i].index] = start + ordered[i].offset;
  }
  return tops;
}

async function spaceEndpointLabels({context, chart, count, options, plot}) {
  const endpointGroups = [0, 1].map(si => Array.from({length: count}, (_, i) => chart.series.getItemAt(si).points.getItemAt(i).dataLabel));
  const changeGroups = options.variant === 'right-column'
    ? [2, 3].map(si => Array.from({length: count}, (_, i) => chart.series.getItemAt(si).points.getItemAt(i).dataLabel)) : [];
  for (const labels of [...endpointGroups, ...changeGroups]) for (const label of labels) label.load('text,top,height');
  await context.sync();
  const result = {gapPt: 2, movedLabels: 0, groups: []};
  endpointGroups.forEach((labels, side) => {
    const boxes = labels.map((label, index) => ({label, index, top: label.top, height: label.height}))
      .filter(box => typeof box.label.text === 'string' && box.label.text.trim() !== '' && Number.isFinite(box.top) && Number.isFinite(box.height) && box.height > 0);
    if (!boxes.length) { result.groups.push({side, measured: 0, status: 'no-measurable-labels'}); return; }
    const tops = spaceLabelBoxes(boxes, plot.top - Math.max(...boxes.map(box => box.height)), plot.top + plot.height, result.gapPt);
    if (!tops) { result.groups.push({side, measured: boxes.length, status: 'insufficient-height'}); return; }
    boxes.forEach((box, i) => {
      const shift = tops[i] - box.top;
      if (Math.abs(shift) <= 0.01) return;
      box.label.top = tops[i]; result.movedLabels++;
      // Keep each right-column change on the same row as its endpoint label.
      // Blank positive/negative counterparts stay blank and are not positioned.
      if (side === 1) for (const changes of changeGroups) {
        const change = changes[box.index];
        if (typeof change.text === 'string' && change.text.trim() !== '' && Number.isFinite(change.top)) change.top += shift;
      }
    });
    result.groups.push({side, measured: boxes.length, status: 'positioned'});
  });
  // Excel enables bent leader lines as a side effect of label.top writes.
  // These tiny spacing adjustments must not add brackets absent in the sample.
  for (const index of [0,1,2,3]) chart.series.getItemAt(index).showLeaderLines = false;
  await context.sync();
  return result;
}

// Rebinding changes only formulas. Initial/new points receive defaults; existing custom font/position stays intact.
function linkLabels(chart, helperName, binding, count, previousCount = 0, previousPeriods = 0) {
  const o = binding.options, changePosition = o.variant === 'right-column' ? 'Left' : 'Center';
  const groups = [{col: 'E', position: 'Left'}, {col: 'F', position: 'Right'},
    {col: 'N', position: changePosition}, {col: 'O', position: changePosition}, {col: 'BB', position: 'Center'}];
  groups.forEach((group, si) => {
    const size = si === 4 ? binding.periods : count, oldSize = si === 4 ? previousPeriods : previousCount;
    for (let i = 0; i < size; i++) {
      const point = chart.series.getItemAt(si).points.getItemAt(i), label = point.dataLabel;
      if (i >= oldSize) {
        point.hasDataLabel = true; label.position = group.position;
        const font = si < 2 ? TYPOGRAPHY.endpoints : si === 4 ? TYPOGRAPHY.periods : TYPOGRAPHY.changes;
        label.format.font.name = font.name; label.format.font.size = font.size; label.format.font.color = si === 3 ? '#FF0000' : '#000000';
        label.format.font.bold = false; label.format.font.italic = false;
        if ((si === 2 || si === 3) && o.variant === 'on-line') label.format.fill.setSolidColor('#FFFFFF');
      }
      label.formula = `=${quote(helperName)}!$${group.col}$${i + 3}`;
    }
  });
}

async function createChart({context, sourceSheet, helper, source, instance, binding}) {
  const chart = sourceSheet.charts.add(Excel.ChartType.xyscatterLinesNoMarkers, helper.getRange(`H3:H${binding.last}`), Excel.ChartSeriesBy.columns);
  chart.name = instance.chartName; chart.plotVisibleOnly = false; chart.displayBlanksAs = 'NotPlotted';
  const anchor = sourceSheet.getRangeByIndexes(source.rowIndex, source.columnIndex + source.columnCount, 1, 1); anchor.load('left,top');
  await context.sync();
  chart.left = anchor.left + 12; chart.top = anchor.top; chart.width = LAYOUT.width; chart.height = LAYOUT.height;
  chart.series.getItemAt(0).name = fixedNames[0];
  for (const name of fixedNames.slice(1)) chart.series.add(name);
  for (let i = 0; i < source.count; i++) chart.series.add(`分类 ${i + 1}`);
  await context.sync(); bindSeries(chart, helper, binding, source.count); await context.sync();
  return chart;
}

async function style({context, chart, instance, source, binding}) {
  const o = binding.options;
  // Changing chartType can restore Excel's default marker. Commit all types first,
  // then apply marker/line formats; all helper carriers explicitly use NoMarkers.
  for (let i = 0; i < FIXED_SERIES + source.count; i++) {
    const series = chart.series.getItemAt(i);
    series.chartType = i >= FIXED_SERIES && o.showMarkers ? Excel.ChartType.xyscatterLines : Excel.ChartType.xyscatterLinesNoMarkers;
    series.axisGroup = Excel.ChartAxisGroup.primary;
  }
  await context.sync();
  chart.format.font.name = FONT; chart.format.font.color = '#000000'; chart.format.font.size = TYPOGRAPHY.endpoints.size;
  chart.format.font.bold = false; chart.format.font.italic = false;
  chart.format.fill.setSolidColor('#FFFFFF'); chart.format.border.color = '#D9D9D9'; chart.format.border.weight = 0.75;
  chart.title.visible = o.title !== ''; chart.title.text = o.title; chart.title.format.font.name = TYPOGRAPHY.title.name; chart.title.format.font.size = TYPOGRAPHY.title.size; chart.title.format.font.bold = false; chart.title.format.font.italic = false; chart.title.format.font.color = '#000000'; chart.title.horizontalAlignment = 'Left';
  chart.legend.visible = false;
  for (let i = 0; i < FIXED_SERIES; i++) {
    const series = chart.series.getItemAt(i);
    series.markerStyle = Excel.ChartMarkerStyle.none;
    if (i < 5) series.format.line.lineStyle = 'None';
    // Image274 guide/baseline ~2 px at 722 px width => 1.5 pt at this chart scale.
    else { series.format.line.lineStyle = i === 8 ? 'Dash' : 'Continuous'; series.format.line.color = i === 8 ? '#C00000' : i === 7 ? '#D9D9D9' : '#A6A6A6'; series.format.line.weight = i === 8 ? 0.75 : 1.5; }
  }
  for (let i = 0; i < source.count; i++) styleEntity(chart.series.getItemAt(FIXED_SERIES + i), i, o);
  const x = chart.axes.categoryAxis, y = chart.axes.valueAxis;
  x.minimum = o.variant === 'right-column' ? -0.498 : -0.5; x.maximum = o.variant === 'right-column' ? 1.9035 : 1.5; x.majorUnit = 1;
  y.minimum = -0.12; y.maximum = 1; y.majorUnit = 1;
  for (const axis of [x, y]) { axis.tickLabelPosition = 'None'; axis.majorTickMark = 'None'; axis.minorTickMark = 'None'; axis.format.line.lineStyle = 'None'; axis.majorGridlines.visible = false; axis.minorGridlines.visible = false; }
  await context.sync(); linkLabels(chart, instance.helperName, binding, source.count); await context.sync();
  chart.title.left = 9; chart.title.top = 7;
  const want = o.variant === 'right-column' ? LAYOUT.plot : {...LAYOUT.plot, left: 47.25, width: 457.5};
  for (let i = 0; i < 2; i++) { chart.plotArea.insideLeft = want.left; chart.plotArea.insideTop = want.top; chart.plotArea.insideWidth = want.width; chart.plotArea.insideHeight = want.height; await context.sync(); }
  chart.plotArea.load('insideLeft,insideTop,insideWidth,insideHeight'); await context.sync();
  const labelSpacing = await spaceEndpointLabels({context, chart, count: source.count, options: o, plot: want});
  return {notes: ['按用户要求采用微软雅黑10pt黑字；端点标签按原生文字框保留至少2pt纵向间距，右侧变化率随所属标签移位，数据端点不动。实际字形与保存重开后的间距仍需实图核验。'], labelSpacing, plot: {requested: want, actual: {left: chart.plotArea.insideLeft, top: chart.plotArea.insideTop, width: chart.plotArea.insideWidth, height: chart.plotArea.insideHeight}}};
}

// Re-measure after the common font cascade: Excel may move a custom top label
// across its point. Endpoint X anchors are fixed; only their text boxes move.
async function finalizePresentation({context, chart, source, binding}) {
  const o=binding.options, plot=LAYOUT.plot;
  chart.getImage();await context.sync();
  await spaceEndpointLabels({context,chart,count:source.count,options:o,plot});
  const min=o.variant==='right-column'?-.498:-.5,max=o.variant==='right-column'?1.9035:1.5;
  const anchors=[o.variant==='right-column'?-.045:-.035,o.variant==='right-column'?1.045:1.035];
  const labels=[0,1].flatMap(side=>Array.from({length:source.count},(_,i)=>({side,label:chart.series.getItemAt(side).points.getItemAt(i).dataLabel})));
  for(const {label} of labels)label.load('text,width');await context.sync();
  for(const {side,label} of labels)if(label.text?.trim()&&Number.isFinite(label.width)){
    const x=plot.left+(anchors[side]-min)/(max-min)*plot.width;
    label.left=side===0?x-label.width:x;
  }
  await context.sync();
}

async function rebind({context, chart, helper, instance, source, previous, binding}) {
  // Stable helper prefix + one series per source row: surviving series objects retain all user formatting.
  if (source.count < previous.count) for (let i = previous.count - 1; i >= source.count; i--) chart.series.getItemAt(FIXED_SERIES + i).delete();
  if (source.count > previous.count) for (let i = previous.count; i < source.count; i++) {
    const series = chart.series.add(`分类 ${i + 1}`);
    series.chartType = binding.options.showMarkers ? Excel.ChartType.xyscatterLines : Excel.ChartType.xyscatterLinesNoMarkers;
    series.axisGroup = Excel.ChartAxisGroup.primary;
  }
  await context.sync(); bindSeries(chart, helper, binding, source.count); await context.sync();
  for (let i = previous.count; i < source.count; i++) styleEntity(chart.series.getItemAt(FIXED_SERIES + i), i, binding.options);
  linkLabels(chart, instance.helperName, binding, source.count, Math.min(previous.count, source.count), binding.oldPeriods);
  await context.sync();
  return {notes: ['已更新引用并保留存续系列与标签格式；格式按源行位置保留。']};
}

export {MARK, HELPER_PREFIX, CHART_PREFIX, COLORS, FONT, TYPOGRAPHY, LAYOUT, FIXED_SERIES, validate, normalizeOptions, makeRecipe, spaceLabelBoxes};
export default {id: 'slope', title: '斜率图', inputSpec, validate, writeHelper, createChart, style, rebind, finalizePresentation};
