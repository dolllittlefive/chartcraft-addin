import {modules, build} from './registry.generated.js';
import {createCore} from './core/instances.js';
import {UserError} from './core/errors.js';
import * as pareto from './charts/pareto.js';
import {addAverageLine,listAverageLines,selectAverageSource,rebindAverageLine} from './tools/average-line.js';

const core = createCore(modules);
const $ = id => document.getElementById(id);
const requested = new URLSearchParams(location.search).get('chart');
const state = {moduleId: modules.some(m => m.id === requested) ? requested : modules[0].id, selection: null, busy: false, confirmDelete: null, instances: [], focusHelper: null, contextPending: false};
// Diagnostic candidate only. No Office calls, network writes or source values.
const diagEvents = [];
let diagEnabled = false, diagSeq = 0, diagClickCount = 0, diagGeneration = 0, diagRender = 0;
function diag(event, detail = {}) {
  if (!diagEnabled) return;
  try {
    diagEvents.push({seq: ++diagSeq, at: new Date().toISOString(), monotonicMs: performance.now(), event,
      busy: state.busy, clickCount: diagClickCount, generateCalls: diagGeneration,
      status: $('status')?.textContent ?? null, statusClass: $('status')?.className ?? null,
      generateDisabled: $('generate')?.disabled ?? null, useSelectionDisabled: $('use-selection')?.disabled ?? null,
      ...detail});
    if (diagEvents.length > 256) diagEvents.shift();
    window.CCGenerateDiagnostics?.onMark?.(diagEvents[diagEvents.length - 1]);
  } catch { /* Diagnostics must not change the production result. */ }
}
const diagnosticSnapshot = () => ({version: 1, seq: diagSeq,
  clickCount: diagClickCount, generateCalls: diagGeneration, busy: state.busy,
  status: $('status')?.textContent ?? null, statusClass: $('status')?.className ?? null,
  generateDisabled: $('generate')?.disabled ?? null, useSelectionDisabled: $('use-selection')?.disabled ?? null,
  events: diagEvents.map(event => ({...event}))});
const watchedSheets = new Set();
let selectionWatched = false, contextTimer, contextRevision = 0;
const current = () => core.moduleFor(state.moduleId);
function setStatus(text, kind) { $('status').textContent = text; $('status').className = 'status' + (kind ? ' ' + kind : ''); }
function describe(e) { return e instanceof UserError ? e.message : `操作失败：${e.message}${e.debugInfo?.errorLocation ? '（' + e.debugInfo.errorLocation + '）' : ''}`; }
function refreshButtons() {
  $('generate').disabled = state.busy;
  $('use-selection').disabled = state.busy;
  $('refresh-instances').disabled = state.busy;
  $('add-average-line').disabled = state.busy;
  $('legend-narrow').disabled = state.busy || Office.context.platform !== Office.PlatformType.PC;
  $('reverse-category').disabled = state.busy || Office.context.platform !== Office.PlatformType.PC;
  $('unify-size').disabled = state.busy || Office.context.platform !== Office.PlatformType.PC;
  document.querySelectorAll('#gallery button, #instances button, #options input, #options select').forEach(b => { b.disabled = state.busy; });
}
async function busy(task) {
  if (state.busy) { diag('busy:rejected'); throw new UserError('请等待当前操作完成。'); }
  state.busy = true; refreshButtons(); diag('busy:entered');
  try { return await task(); } finally {
    state.busy = false; refreshButtons(); diag('busy:released');
    if (state.contextPending) { state.contextPending = false; scheduleContextRefresh(); }
  }
}
function activate(id) {
  state.moduleId = id; state.selection = null;
  const module = current();
  $('chart-title').textContent = module.title; $('chart-description').textContent = module.inputSpec.description;
  $('generate').textContent = `生成${module.title}`; $('selection-info').textContent = '尚未选择数据';
  $('options').replaceChildren();
  for (const option of module.inputSpec.options) {
    const label = document.createElement('label'); label.textContent = option.label + ' ';
    const input = document.createElement(option.type === 'select' ? 'select' : 'input'); input.id = option.id;
    if (option.type === 'select') {
      for (const choice of option.choices) { const item = document.createElement('option'); item.value = choice.value; item.textContent = choice.label; input.append(item); }
      input.value = option.default;
    } else if (option.type === 'checkbox') { input.type = 'checkbox'; input.checked = option.default; }
    else { input.type = option.type; input.value = option.default; }
    for (const key of ['min', 'max', 'step', 'maxLength']) if (option[key] !== undefined) input[key] = option[key];
    label.append(input); if (option.suffix) label.append(' ' + option.suffix); $('options').append(label);
  }
  document.querySelectorAll('#gallery button').forEach(b => b.setAttribute('aria-pressed', String(b.dataset.chart === id)));
  setStatus(''); refreshButtons();
}
function gallery() {
  const categories = new Map();
  for (const module of modules) {
    const category = module.inputSpec.category;
    if (!categories.has(category.id)) {
      const group = document.createElement('div'); const label = document.createElement('p'); label.className = 'hint'; label.textContent = category.title;
      group.append(label); $('gallery').append(group); categories.set(category.id, group);
    }
    const button = document.createElement('button'); button.type = 'button'; button.dataset.chart = module.id; button.textContent = module.title;
    button.onclick = () => activate(module.id); categories.get(category.id).append(button);
  }
  activate(state.moduleId);
}
function readOptions() {
  const options = {};
  for (const spec of current().inputSpec.options) {
    const input = $(spec.id);
    if (spec.type === 'checkbox') options[spec.id] = input.checked;
    else if (spec.type === 'select') {
      if (!spec.choices.some(choice => choice.value === input.value)) throw new UserError(`${spec.label}请选择有效选项。`);
      options[spec.id] = input.value;
    } else if (spec.type === 'number') {
      const number = Number(input.value);
      if (!input.value.trim() || !Number.isFinite(number) || !input.checkValidity()) throw new UserError(`${spec.label}请输入 ${spec.min}–${spec.max} 范围内的有效数值。`);
      options[spec.id] = number * (spec.scale ?? 1);
    } else options[spec.id] = input.value.trim() || String(spec.default);
  }
  return options;
}
async function useSelection() {
  return busy(async () => {
    try {
      state.selection = await core.readSelection(state.moduleId);
      const s = state.selection;
      $('selection-info').textContent = `${s.sheet}!${s.address.split('!').pop()}：${s.count} 行${s.total !== undefined ? `，合计 ${s.total}` : ''}`; setStatus('');
    } catch (e) { state.selection = null; $('selection-info').textContent = '尚未选择数据'; setStatus(describe(e), 'err'); }
    return state.selection;
  });
}
async function generate(event) {
  const generation = diagEnabled ? ++diagGeneration : null; diag('generate:called', {generation});
  return busy(async () => {
    try {
      const rawOptions = readOptions();
      const options = current().creationOptions ? current().creationOptions(rawOptions,event) : rawOptions; setStatus('正在生成…');
      diag('generate:core:start', {generation});
      const result = await core.create(state.moduleId, options);
      if (Office.context.requirements.isSetSupported('SharedRuntime', '1.1')) {
        await Office.addin.setStartupBehavior(Office.StartupBehavior.load);
      }
      diag('generate:core:resolved', {generation});
      state.focusHelper = result.helperName;
      setStatus(`已生成 ${result.chartName}（${result.count} 项）。${result.notes?.join('；') || ''}`, 'ok');
      diag('generate:render:start', {generation});
      await renderInstances(); diag('generate:render:resolved', {generation}); return result;
    } catch (e) { setStatus(describe(e), 'err'); diag('generate:rejected', {generation, code: e.code, message: e.message, location: e.debugInfo?.errorLocation}); throw e; }
  });
}
async function addAverageLineToSelection() {
  return busy(async () => {
    try {
      setStatus('正在为所选图表添加平均线…'); diag('tool:averageline:start');
      const result = await Excel.run(context => addAverageLine({context}));
      setStatus(`已为“${result.chartName}”添加平均线。修改原数据即可自动更新。`, 'ok');
      diag('tool:averageline:resolved'); await renderInstances(); return result;
    } catch (e) { setStatus(describe(e), 'err'); diag('tool:averageline:rejected', {code:e.code, message:e.message, location:e.debugInfo?.errorLocation}); throw e; }
  });
}
async function unifySelectedSize() {
  return busy(async () => {
    try {
      setStatus('正在按第一个选中对象统一大小…');
      const request=await Excel.run(async context=>{
        const wb=context.workbook,sheet=wb.worksheets.getActiveWorksheet();
        wb.load('name');sheet.load('name');await context.sync();return {workbook:wb.name,worksheet:sheet.name};
      });
      const response=await fetch('/native-selected-size',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify(request)});
      const result=await response.json();if(!response.ok||!result.ok)throw Error(result.error||'统一大小失败');
      setStatus(result.skipped ? '当前未多选绘图对象，未修改。' : `已按“${result.reference}”统一 ${result.applied.length} 个对象的大小。`, 'ok');return result;
    }catch(e){setStatus(describe(e),'err');throw e;}
  });
}
async function reverseSelectedCategory() {
  return busy(async()=>{
    try{
      setStatus('正在按原版规则翻转类别…');
      const request=await Excel.run(async context=>{
        const wb=context.workbook,sheet=wb.worksheets.getActiveWorksheet();wb.load('name');sheet.load('name');await context.sync();return {workbook:wb.name,worksheet:sheet.name};
      });
      const response=await fetch('/native-reverse-category',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify(request)});
      const result=await response.json();if(!response.ok||!result.ok)throw Error(result.error||'翻转类别失败');
      setStatus(`已翻转 ${result.applied.length} 张条形图；跳过 ${result.skipped.length} 张不适用图表。`,'ok');return result;
    }catch(e){setStatus(describe(e),'err');throw e;}
  });
}
async function narrowSelectedLegend() {
  return busy(async()=>{
    try{
      setStatus('正在按原版规则窄化图例…');
      const request=await Excel.run(async context=>{
        const wb=context.workbook,sheet=wb.worksheets.getActiveWorksheet();wb.load('name');sheet.load('name');await context.sync();return {workbook:wb.name,worksheet:sheet.name};
      });
      const response=await fetch('/native-legend-narrow',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify(request)});
      const result=await response.json();if(!response.ok||!result.ok)throw Error(result.error||'图例窄化失败');
      setStatus(result.skipped==='no-active-chart'?'请先选择一个嵌入式图表。':`已复制为“${result.chartName}”，窄化 ${result.applied.length} 个系列的图例；原图保留。`,'ok');return result;
    }catch(e){setStatus(describe(e),'err');throw e;}
  });
}
async function rebind(helperName) {
  return busy(async () => {
    try {
      const target = state.instances.find(item => item.helperName === helperName);
      setStatus(`正在更新“${target?.title || helperName}”的数据引用…`); const result = target?.moduleId==='average-line' ? await Excel.run(context=>rebindAverageLine(context,helperName)) : await core.rebind(helperName);
      state.focusHelper = helperName;
      setStatus(`已更新 ${result.chartName}：${result.oldCount} 项 → ${result.count} 项。源区域 ${result.source}，源单元格内容未改写。`, 'ok');
      await renderInstances(); return result;
    } catch (e) { setStatus(describe(e), 'err'); throw e; }
  });
}
async function editSource(helperName) {
  return busy(async () => {
    try {
      const target=state.instances.find(item=>item.helperName===helperName);
      const result = target?.moduleId==='average-line' ? await Excel.run(context=>selectAverageSource(context,helperName)) : await core.selectSource(helperName);
      state.focusHelper = helperName; paintInstanceContext(null);
      setStatus(`已选中 ${result.source}。直接编辑其中的值，图表会自动重算；增删数据行后，选中含表头的完整区域，再点此图的“按选区更新”。`, 'ok');
      return result;
    } catch (e) { setStatus(describe(e), 'err'); throw e; }
  });
}
async function refreshStyle(helperName) {
  return busy(async () => {
    try {
      setStatus('正在应用新版格式，将重设此图样式…');
      const result = await core.refreshStyle(helperName);
      state.focusHelper = helperName;
      setStatus(`“${result.title}”已应用新版格式。保留原数据区域与图表，源单元格内容未改写。`, 'ok');
      await renderInstances(); return result;
    } catch (e) { setStatus(`新版格式未完成，可能已部分更新样式或辅助公式。${describe(e)}`, 'err'); throw e; }
  });
}
function paintInstanceContext(selected) {
  if (selected) state.focusHelper = selected.instance?.helperName || null;
  const target = state.instances.find(item => item.helperName === state.focusHelper);
  $('instance-context').textContent = selected && !target
    ? (selected.ambiguous ? '当前图表对应多份实例记录，未自动选择。请按下方源区域核对。' : `当前图表“${selected.chartName}”没有可识别的 ChartCraft 实例记录。`)
    : target ? `${selected ? '已选图表' : '编辑目标'}：${(target.moduleId==='average-line'?'平均线':core.moduleFor(target.moduleId).title)} · ${target.title}`
      : '选中工作表中的图表可识别对应实例，也可直接在下方选择“编辑数据”。';
  document.querySelectorAll('#instances [data-helper]').forEach(li => {
    const active = li.dataset.helper === target?.helperName;
    li.classList.toggle('is-current', active);
    if (active) li.setAttribute('aria-current', 'true'); else li.removeAttribute('aria-current');
  });
}
async function refreshInstanceContext(allowBusy = false) {
  if (state.busy && !allowBusy) { state.contextPending = true; return; }
  const revision = ++contextRevision;
  try {
    diag('context:active-instance:start', {revision, allowBusy});
    const selected = await core.activeInstance(state.instances);
    diag('context:active-instance:resolved', {revision, allowBusy});
    if (revision !== contextRevision) return;
    if (state.busy && !allowBusy) { state.contextPending = true; return; }
    paintInstanceContext(selected);
  } catch (e) { diag('context:active-instance:rejected', {revision, message: e.message}); $('instance-context').textContent = `暂时无法识别所选图表，可按源区域使用下方入口。${describe(e)}`; }
}
function scheduleContextRefresh() {
  clearTimeout(contextTimer);
  contextTimer = setTimeout(() => refreshInstanceContext(), 150);
}
async function watchInstanceContext() {
  // In-memory listeners only; no polling, workbook settings or data writes.
  await Excel.run(async context => {
    const sheets = context.workbook.worksheets; sheets.load('items/id');
    diag('watch:sheets-sync:start'); await context.sync(); diag('watch:sheets-sync:resolved');
    const pending = sheets.items.filter(sheet => !watchedSheets.has(sheet.id));
    for (const sheet of pending) sheet.charts.onActivated.add(scheduleContextRefresh);
    if (!selectionWatched) {
      context.workbook.onSelectionChanged.add(scheduleContextRefresh);
      sheets.onActivated.add(scheduleContextRefresh);
    }
    diag('watch:subscribe-sync:start', {newSheets: pending.length, selectionWatched});
    await context.sync(); diag('watch:subscribe-sync:resolved');
    for (const sheet of pending) watchedSheets.add(sheet.id);
    selectionWatched = true;
  });
}
async function remove(helperName, button) {
  if (state.confirmDelete !== helperName) {
    state.confirmDelete = helperName; button.textContent = '再点一次确认删除';
    setTimeout(() => { if (state.confirmDelete === helperName) { state.confirmDelete = null; button.textContent = '删除'; } }, 4000); return;
  }
  state.confirmDelete = null;
  return busy(async () => {
    try { const result = await core.remove(helperName); setStatus(`已删除 ${result.chartName} 及其辅助表。`, 'ok'); await renderInstances(); return result; }
    catch (e) { setStatus(describe(e), 'err'); throw e; }
  });
}
async function renderInstances() {
  const render = diagEnabled ? ++diagRender : null; diag('render:list:start', {render});
  const items = [...await core.list(),...await Excel.run(context=>listAverageLines(context))]; diag('render:list:resolved', {render, count: items.length}); state.instances = items;
  const list = $('instances'); list.replaceChildren();
  if (!items.length) { const li = document.createElement('li'); li.className = 'empty'; li.textContent = '未找到 ChartCraft 图表。生成后的图表会显示在这里。'; list.append(li); }
  for (const item of items) {
    const li = document.createElement('li'); li.dataset.helper = item.helperName; li.dataset.chart = item.moduleId;
    const title = document.createElement('div'); title.textContent = `${(item.moduleId==='average-line'?'平均线':core.moduleFor(item.moduleId).title)} · ${item.title}`;
    const meta = document.createElement('div'); meta.className = 'meta'; meta.textContent = `源数据 ${item.source}（${item.count} 项）`;
    const actions = document.createElement('div'); actions.className = 'actions';
    const edit = document.createElement('button'); edit.type = 'button'; edit.textContent = '编辑数据'; edit.dataset.action = 'edit-source'; edit.title = `定位并选中 ${item.source}，不改写数据`; edit.onclick = () => editSource(item.helperName).catch(() => {});
    const update = document.createElement('button'); update.type = 'button'; update.textContent = '按选区更新'; update.dataset.action = 'rebind'; update.title = `用当前完整选区更新“${item.title}”`; update.onclick = () => rebind(item.helperName).catch(() => {});
    const style = document.createElement('button'); style.type = 'button'; style.textContent = '应用新版格式'; style.dataset.action = 'refresh-style'; style.title = '重设此图的字体、颜色、标签及版式；保留原数据和实例选项'; style.onclick = () => refreshStyle(item.helperName).catch(() => {});
    const del = document.createElement('button'); del.type = 'button'; del.textContent = '删除'; del.dataset.action = 'remove'; del.onclick = () => remove(item.helperName, del).catch(() => {});
    const styleHint = document.createElement('div'); styleHint.className = 'meta'; styleHint.textContent = '应用新版格式会重设此图样式，覆盖手工调整的格式。';
    if(item.moduleId==='average-line'){actions.append(edit,update);styleHint.textContent='选择含表头的两列数据后更新；保留已有图表格式。';}else actions.append(edit, update, style, del); li.append(title, meta, actions, styleHint); list.append(li);
  }
  diag('render:context:start', {render}); await refreshInstanceContext(true); diag('render:context:resolved', {render});
  try { diag('render:watch:start', {render}); await watchInstanceContext(); diag('render:watch:resolved', {render}); }
  catch (e) { diag('render:watch:rejected', {render, message: e.message}); $('instance-context').textContent += ' 自动识别暂不可用，可点击“识别所选图表 / 刷新”。'; }
  refreshButtons(); diag('render:resolved', {render}); return items;
}
// Legacy API retained for the existing independent acceptance/inspection tools.
window.CCPareto = {...pareto, UserError, create: options => core.create('pareto', options), list: () => core.list('pareto'), readSelection: () => core.readSelection('pareto'), rebind: core.rebind, remove: core.remove};
window.CCPane = {useSelection, generate, rebind, editSource, refreshStyle, renderInstances, readOptions};
window.CCCore = core; window.CCBuild = build;
Office.onReady(async info => {
  if (info.host !== Office.HostType.Excel) { setStatus('请在 Excel 中打开本插件。', 'err'); return; }
  if (!Office.context.requirements.isSetSupported('ExcelApi', build.minExcelApi)) { setStatus(`当前 Excel 版本过旧（需要 ExcelApi ${build.minExcelApi} 及以上）。`, 'err'); return; }
  if (build.acceptance === true && Office.context.document.settings.get('ccAcceptance') === true) {
    diagEnabled = true; window.CCGenerateDiagnostics = {snapshot: diagnosticSnapshot, mark: diag};
  }
  gallery();
  window.CCResolveRibbon?.(async route => {
    try {
      if (state.busy) throw new UserError('请等待当前操作完成。');
      // The build supplies one route per native button. Every button activates
      // its own recipe before opening options or generating from the selection.
      core.moduleFor(route.moduleId);
      activate(route.moduleId);
      if (route.direction) $('direction').value = route.direction;
      if (route.mode === 'create') return await generate();
      await Office.addin.showAsTaskpane();
    } catch (error) {
      setStatus(describe(error), 'err');
      await Office.addin.showAsTaskpane();
      throw error;
    }
  });
  $('use-selection').onclick = () => useSelection().catch(e => setStatus(describe(e), 'err'));
  $('generate').onclick = event => {
    if (diagEnabled) ++diagClickCount; diag('generate:dom-click');
    const operation = generate(event);
    if (diagEnabled) operation.then(
      () => window.CCGenerateDiagnostics?.onSettled?.('resolved'),
      () => window.CCGenerateDiagnostics?.onSettled?.('rejected')
    ).catch(() => {});
    return operation.catch(() => {});
  };
  $('refresh-instances').onclick = () => busy(renderInstances).catch(e => setStatus(describe(e), 'err'));
  $('add-average-line').onclick = () => addAverageLineToSelection().catch(() => {});
  $('unify-size').onclick = () => unifySelectedSize().catch(() => {});
  $('reverse-category').onclick = () => reverseSelectedCategory().catch(() => {});
  $('legend-narrow').onclick = () => narrowSelectedLegend().catch(() => {});
  try {
    await renderInstances();
    if (build.acceptance === true && Office.context.document.settings.get('ccAcceptance') === true
        && Office.context.document.settings.get('ccAcceptanceCase') === 'averageline-tool') {
      const {start} = await import('./acceptance/averageline-tool.js'); await start(build); return;
    }
    
  } catch (e) { setStatus(describe(e), 'err'); }
});
