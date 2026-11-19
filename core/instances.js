import {UserError} from './errors.js';
import {setChartPlacement} from './placement.js';
import {applyChartPresentation} from './presentation.js';

// A1 identifies the recipe. I1:M1 and R1 are the shared instance envelope.
// B1:H1, N1:Q1 and rows 2+ belong to the recipe.
const literal = value => "'" + String(value);
const address = src => `${src.sheet}!${src.address.split('!').pop()}`;

// Older envelopes store an unquoted sheet name; Excel addresses may quote it.
// Resolve only this instance's local rectangle, never a named/external range.
function sourceReference(instance, module) {
  const source = String(instance.source || ''), bang = source.lastIndexOf('!');
  const rawSheet = source.slice(0, bang), range = source.slice(bang + 1);
  const sheet = rawSheet.startsWith("'") && rawSheet.endsWith("'")
    ? rawSheet.slice(1, -1).replace(/''/g, "'") : rawSheet;
  if (bang < 1 || !sheet || /[\[\]]/.test(sheet) ||
      (!module.inputSpec.crossSheetRebind && sheet !== instance.chartSheet) ||
      !/^\$?[A-Z]{1,3}\$?[1-9]\d*(?::\$?[A-Z]{1,3}\$?[1-9]\d*)?$/i.test(range)) {
    throw new UserError('此图保存的源区域与工作表不匹配。未选择或修改数据，请核对源表名称。');
  }
  return {sheet, range};
}

export function createCore(modules) {
  const byId = new Map(modules.map(m => [m.id, m]));
  const moduleFor = id => {
    const module = byId.get(id);
    if (!module) throw new UserError(`未注册图表：${id}`);
    return module;
  };
  function decode(helperName, values) {
    const module = modules.find(m => values[0] === m.inputSpec.storage.marker && helperName.startsWith(m.inputSpec.storage.helperPrefix));
    if (!module) return null;
    let options = {};
    if (values[17]) {
      try { options = JSON.parse(values[17]); }
      catch { throw new UserError(`辅助表“${helperName}”的选项损坏，未修改。`); }
    } else {
      for (const [key, column] of Object.entries(module.inputSpec.storage.legacyOptions || {})) options[key] = values[column];
    }
    options.title = values[12];
    // Recipe options must not replace the shared envelope's identity/range.
    return {...options, moduleId: module.id, helperName, source: values[8], chartSheet: values[9], count: Number(values[10]), chartName: values[11], title: values[12], options};
  }
  function persist(helper, module, source, options, instance) {
    helper.getRange('A1').values = [[module.inputSpec.storage.marker]];
    helper.getRange('I1:M1').formulas = [[literal(address(source)), literal(instance.chartSheet), String(source.count), literal(instance.chartName), literal(options.title)]];
    helper.getRange('R1').values = [[JSON.stringify(options)]];
  }
  async function readSelection(moduleId) {
    const module = moduleFor(moduleId);
    const trace = stage => {
      try { globalThis.CCGenerateDiagnostics?.mark?.(`core:read-selection:${stage}`); }
      catch { /* Acceptance observation must not affect selection or validation. */ }
    };
    return Excel.run(async context => {
      const chart = context.workbook.getActiveChartOrNullObject();
      trace('active-chart:start'); await context.sync(); trace('active-chart:resolved');
      if (!chart.isNullObject) throw new UserError('请先选择含表头的完整数据区域。修改现有图表可先点“编辑数据”。');
      // ExcelApi 1.1 getSelectedRange is supported; there is no OrNullObject variant.
      let range;
      try {
        range = context.workbook.getSelectedRange();
        range.load('address,values,rowIndex,columnIndex,rowCount,columnCount'); range.worksheet.load('name');
        trace('selected-range:start'); await context.sync(); trace('selected-range:resolved');
      } catch (cause) {
        if (['InvalidSelection', 'ItemNotFound'].includes(cause?.code)) {
          throw new UserError('请选择一块连续的完整数据区域（含表头），不要同时选择多个区域。');
        }
        throw cause;
      }
      // Only recipes explicitly verified against the original opt into a
      // nonempty single cell's current region. Other selections stay exact.
      let selectionOrigin = {};
      if (module.inputSpec.expandSingleCell && range.rowCount === 1 && range.columnCount === 1) {
        const selectedValue = range.values[0][0];
        if (selectedValue === '' || selectedValue === null) throw new UserError('所选单元格为空，未扩展或生成。');
        selectionOrigin = {fromSingleCell: true, selectedValue};
        range = range.getSurroundingRegion();
        range.load('address,values,rowIndex,columnIndex,rowCount,columnCount'); range.worksheet.load('name');
        await context.sync();
      }
      trace('validate:start');
      const info = module.validate(range.values);
      trace('validate:resolved');
      return {...info, ...selectionOrigin, values: range.values, sheet: range.worksheet.name, address: range.address, rowIndex: range.rowIndex, columnIndex: range.columnIndex, rowCount: range.rowCount, columnCount: range.columnCount};
    });
  }
  async function list(moduleId) {
    if (moduleId) moduleFor(moduleId);
    return Excel.run(async context => {
      const sheets = context.workbook.worksheets; sheets.load('items/name'); await context.sync();
      const candidates = sheets.items.filter(s => modules.some(m => s.name.startsWith(m.inputSpec.storage.helperPrefix)));
      const heads = candidates.map(sheet => { const head = sheet.getRange('A1:R1'); head.load('values'); return {sheet, head}; });
      await context.sync();
      return heads.map(({sheet, head}) => decode(sheet.name, head.values[0])).filter(m => m && (!moduleId || m.moduleId === moduleId));
    });
  }
  async function loadInstance(context, helperName) {
    const helper = context.workbook.worksheets.getItem(helperName);
    const head = helper.getRange('A1:R1'); head.load('values'); await context.sync();
    const instance = decode(helperName, head.values[0]);
    if (!instance) throw new UserError('这不是已注册的 ChartCraft 辅助表，未修改。');
    return {helper, instance, module: moduleFor(instance.moduleId)};
  }
  async function selectSource(helperName) {
    return Excel.run(async context => {
      const {instance, module} = await loadInstance(context, helperName);
      const reference = sourceReference(instance, module);
      const sheet = context.workbook.worksheets.getItemOrNullObject(reference.sheet); await context.sync();
      if (sheet.isNullObject) throw new UserError(`源工作表“${reference.sheet}”不存在，未修改数据。`);
      sheet.load('visibility');
      const range = sheet.getRange(reference.range); range.load('address'); await context.sync();
      if (sheet.visibility !== Excel.SheetVisibility.visible) throw new UserError(`源工作表“${reference.sheet}”已隐藏，请先在 Excel 中显示该表。`);
      // Navigation only: do not write values/formulas, unhide helpers or rebind.
      sheet.activate(); range.select(); await context.sync();
      return {helperName, chartName: instance.chartName, source: range.address, count: instance.count};
    });
  }
  async function activeInstance(items) {
    const instances = items || await list();
    return Excel.run(async context => {
      const chart = context.workbook.getActiveChartOrNullObject(); await context.sync();
      if (chart.isNullObject) return null;
      chart.load('name'); chart.worksheet.load('name'); await context.sync();
      const matches = instances.filter(item => item.chartSheet === chart.worksheet.name && (moduleFor(item.moduleId).memberChartNames?.(item) || [item.chartName]).includes(chart.name));
      return {chartName: chart.name, chartSheet: chart.worksheet.name, instance: matches.length === 1 ? matches[0] : null, ambiguous: matches.length > 1};
    });
  }
  async function create(moduleId, options) {
    const trace = stage => globalThis.CCGenerateDiagnostics?.mark?.(`core:create:${stage}`);
    trace('read-selection:start');
    const module = moduleFor(moduleId), source = await readSelection(moduleId);
    trace('read-selection:resolved');
    return Excel.run(async context => {
      trace('read-sheets:start');
      const sheets = context.workbook.worksheets; sheets.load('items/name'); await context.sync();
      trace('read-sheets:resolved');
      trace('read-charts:start');
      const sourceSheet = sheets.getItem(source.sheet); sourceSheet.charts.load('items/name'); await context.sync();
      trace('read-charts:resolved');
      const names = new Set(sheets.items.map(s => s.name)), charts = new Set(sourceSheet.charts.items.map(c => c.name));
      const storage = module.inputSpec.storage;
      let id = 1;
      while (names.has(storage.helperPrefix + id) || (module.memberChartNames
        ? module.memberChartNames({chartName:storage.chartPrefix + id,count:source.count}).some(name=>charts.has(name)||names.has('_'+name))
        : charts.has(storage.chartPrefix + id))) id++;
      const instance = {id, moduleId, helperName: storage.helperPrefix + id, chartName: storage.chartPrefix + id, chartSheet: source.sheet, title: options.title, source: address(source), count: source.count};
      const helper = sheets.add(instance.helperName);
      const args = {context, helper, sourceSheet, source, options, instance, previous: null};
      trace('write-helper:start');
      const binding = await module.writeHelper(args); trace('write-helper:resolved');
      trace('write-helper-sync:start'); await context.sync(); trace('write-helper-sync:resolved');
      trace('create-chart:start');
      const chart = await module.createChart({...args, binding});
      trace('create-chart:resolved');
      trace('style:start');
      const detail = await module.style({...args, binding, chart});
      trace('style:resolved');
      trace('presentation:start');
      if (module.inputSpec.preserveOriginalPresentation !== true) await applyChartPresentation({context, chart, graphicLabels: detail?.graphicLabels || []});
      trace('presentation:resolved');
      if (module.finalizePresentation) {
        trace('finalize:start'); await module.finalizePresentation({...args, binding, chart}); trace('finalize:resolved');
      }
      // Keep one blank column between the selected source and its native chart.
      trace('anchor:start');
      const chartAnchor = module.resolveAnchor
        ? module.resolveAnchor({source, options, binding})
        : {rowIndex: source.rowIndex, columnIndex: source.columnIndex + source.columnCount + 1};
      const anchor = sourceSheet.getRangeByIndexes(chartAnchor.rowIndex, chartAnchor.columnIndex, 1, 1);
      anchor.load('left,top'); await context.sync(); trace('anchor:read');
      chart.left = anchor.left;
      if (!(moduleId === 'gantt' && options.placement === 'table-aligned')) chart.top = anchor.top;
      await context.sync(); trace('anchor:placed');
      // OneCell follows F2 (or the source-dependent anchor) across host column
      // metrics while retaining point width/height. Absolute kept dimensions
      // but drifted away from F2 on the verified Mac-to-Windows round trip.
      trace('placement:start');
      await setChartPlacement({context, sourceSheet, chart}, 'OneCell');
      trace('placement:resolved');
      trace('persist:start');
      persist(helper, module, source, options, instance);
      helper.visibility = Excel.SheetVisibility.hidden; sourceSheet.activate(); await context.sync();
      trace('persist-sync:resolved');
      return {...instance, sheet: source.sheet, ...detail};
    });
  }
  async function rebind(helperName) {
    // Selection validation belongs to this instance's module, even if another gallery card is active.
    const before = (await list()).find(m => m.helperName === helperName);
    if (!before) throw new UserError('未找到可更新的 ChartCraft 实例。');
    const source = await readSelection(before.moduleId);
    return Excel.run(async context => {
      const {helper, instance, module} = await loadInstance(context, helperName);
      if (!module.inputSpec.crossSheetRebind && source.sheet !== instance.chartSheet) throw new UserError(`请在原数据所在工作表“${instance.chartSheet}”上选择数据。`);
      const sourceSheet = context.workbook.worksheets.getItem(instance.chartSheet);
      const chart = sourceSheet.charts.getItemOrNullObject(instance.chartName); await context.sync();
      if (chart.isNullObject) throw new UserError(`未找到原图“${instance.chartName}”，未更新数据引用。`);
      const args = {context, helper, sourceSheet, chart, source, options: instance.options, instance: {...instance, source: address(source), count: source.count}, previous: instance};
      const binding = await module.writeHelper(args); await context.sync();
      const detail = await module.rebind({...args, binding});
      persist(helper, module, source, instance.options, args.instance); await context.sync();
      return {helperName, chartName: instance.chartName, source: address(source), oldCount: instance.count, count: source.count, ...detail};
    });
  }
  async function refreshStyle(helperName) {
    return Excel.run(async context => {
      const {helper, instance, module} = await loadInstance(context, helperName);
      const sourceSheet = context.workbook.worksheets.getItemOrNullObject(instance.chartSheet); await context.sync();
      if (sourceSheet.isNullObject) throw new UserError(`源工作表“${instance.chartSheet}”不存在，未应用新版格式。`);
      const reference = sourceReference(instance, module);
      const dataSheet = context.workbook.worksheets.getItem(reference.sheet);
      const range = dataSheet.getRange(reference.range);
      range.load('address,values,rowIndex,columnIndex,rowCount,columnCount'); await context.sync();
      const source = {...module.validate(range.values), values: range.values, sheet: reference.sheet, address: range.address,
        rowIndex: range.rowIndex, columnIndex: range.columnIndex, rowCount: range.rowCount, columnCount: range.columnCount};
      if (source.count !== instance.count) throw new UserError('源数据行数已变化。请先选择完整数据区域并“按选区更新”，再应用新版格式。');
      const chart = sourceSheet.charts.getItemOrNullObject(instance.chartName); await context.sync();
      if (chart.isNullObject) throw new UserError(`未找到原图“${instance.chartName}”，未应用新版格式。`);
      chart.load('id'); chart.series.load('items'); await context.sync();
      const chartId = chart.id, seriesCount = chart.series.items.length;
      const options = {...instance.options};
      const legacyPreviewTitles = {'斜率图自带右侧变化标签': '2022年第三季度至2023年第二季度的变化', 'tusimpleBI 偏差图': '双期数值及差异对比', '棒棒糖合成工程验收': '分类数值对比（合成示例）'};
      if (legacyPreviewTitles[options.title]) options.title = legacyPreviewTitles[options.title];
      if (options.title === '图表标题') {
        const title = module.inputSpec.options.find(option => option.id === 'title')?.default;
        options.title = title && title !== '图表标题' ? title : module.title;
      }
      // Bounded guards for the existing recipes; this action never migrates an
      // old chart by adding/removing series. Row-count-changing rebind is separate.
      const fixedCount = {pareto: 5, bullet: 12, gantt: 9, deviation: 18, lollipop: options.variant === 'bubble' ? 23 : 3}[module.id];
      let expected = fixedCount;
      if (module.id === 'matrixbubble') {
        const kind = options.variant && options.variant !== 'auto' ? options.variant : source.inputKind;
        const bands = options.colorMode === 'bands25_30';
        expected = 4 + (kind === 'matrix' && !bands && !options.referenceRings ? 1 : 0) + (options.referenceRings ? 2 : 0) + (bands ? 4 : 0);
      }
      if (expected !== undefined && seriesCount !== expected) throw new UserError(`此图结构不兼容新版格式（现有 ${seriesCount} 个系列，需 ${expected} 个）。请保留原图并用源数据新建，未自动增删系列。`);
      const updated = {...instance, title: options.title, options, source: address(source), count: source.count};
      const args = {context, helper, sourceSheet, chart, source, options, instance: updated, previous: instance};
      const binding = await module.writeHelper(args); await context.sync();
      // These two recipes normally append missing capacity. Style refresh must
      // instead stop before rebind if their existing topology cannot be reused.
      const needed = module.id === 'variablewidthbar' ? binding.firstBody + binding.bodyColumns
        : module.id === 'waterfall' ? binding.body.length + 6 + binding.info.components + binding.n - 1 : undefined;
      if (needed !== undefined && seriesCount < needed) throw new UserError(`原图系列不足（现有 ${seriesCount}，需 ${needed}），未自动补系列。请保留原图并新建实例；辅助公式可能已刷新。`);
      await module.rebind({...args, binding});
      const detail = await module.style({...args, binding});
      if (module.inputSpec.preserveOriginalPresentation !== true) await applyChartPresentation({context, chart, graphicLabels: detail?.graphicLabels || []});
      if (module.finalizePresentation) await module.finalizePresentation({...args, binding, chart});
      await setChartPlacement({context, sourceSheet, chart}, 'OneCell');
      chart.load('id'); chart.series.load('items'); await context.sync();
      if (chart.id !== chartId || chart.series.items.length !== seriesCount) throw new UserError('格式更新后的图表身份或系列数变化，未认定更新成功；请保留当前现场。');
      persist(helper, module, source, options, updated); await context.sync();
      return {helperName, chartName: instance.chartName, source: updated.source, title: updated.title, count: source.count, seriesCount, ...detail};
    });
  }
  async function remove(helperName) {
    return Excel.run(async context => {
      const {helper, instance, module} = await loadInstance(context, helperName);
      if(module.removeMembers) await module.removeMembers({context,instance});
      else {
        const chart = context.workbook.worksheets.getItem(instance.chartSheet).charts.getItemOrNullObject(instance.chartName); await context.sync();
        if (!chart.isNullObject) chart.delete();
      }
      // Custom error-bar ranges remain linked to this instance's private sheet.
      const templateName = instance.options?.nativeTemplateSheet;
      if (instance.moduleId === 'lollipop' && templateName && /^CC_LP_T_/.test(templateName)) {
        const template = context.workbook.worksheets.getItemOrNullObject(templateName);
        await context.sync();
        if (!template.isNullObject) { template.visibility = Excel.SheetVisibility.visible; template.delete(); }
      }
      helper.visibility = Excel.SheetVisibility.visible; helper.delete(); await context.sync();
      return {helperName, chartName: instance.chartName};
    });
  }
  return {modules, moduleFor, readSelection, create, list, selectSource, activeInstance, rebind, refreshStyle, remove};
}
