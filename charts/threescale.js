// 景歌“三数量级组合图”图族（21 篇）。三个不同数量级的指标放进一张原生图：
//   bubble  柱形 + 折线 + 气泡行（2024-09-18 / 2026-06-09 / 2026-10-07 方法二）
//   layered 上下分层：底部柱 + 中间折线 + 顶部柱（2026-03-18）
//   log     对数刻度双柱 + 折线（2026-10-07 方法一）
// 只用 Chart API。Excel 气泡图不能与柱/线组合，所以气泡用“带标记折线、线条隐藏”的
// 系列画出：按 15 档固定标记大小拆成 15 个系列，每个点由公式落到其所属档位，改数后
// 公式重算即换档，不需要插件刷新（大小量化误差 ≤ 最大气泡/30）。
import {UserError} from '../core/errors.js';

export const FONT = '微软雅黑';
export const MAX_ROWS = 100;
export const SIZE_CLASSES = 15;
export const COLORS = Object.freeze({bottom:'#2FBFD3', line:'#E2B464', top:'#D56B6B', layeredTop:'#1BA1E2', layeredBottom:'#76C1BA', logA:'#1E6FA8', logB:'#9FD3F0', text:'#404040', axis:'#D9D9D9'});
export const LAYOUT = Object.freeze({width:560, height:420});
// 版带（绘图区高度的比例，0=底，1=顶）。bubble/layered 主轴固定 0..1，公式把各指标压进自己的版带。
export const BANDS = Object.freeze({
  bubble:{bottom:[0,0.48], line:[0.54,0.20], bubbleY:0.86, labelY:0.96},
  layered:{bottom:[0,0.36], line:[0.44,0.16], top:[0.66,0.28]},
  log:{line:[0.70,0.22], columnShare:0.62}
});
export const VARIANTS = ['bubble','layered','log'];
const quote = s => `'${String(s).replace(/'/g,"''")}'`;
const invalidText = v => /^#(?:REF!|VALUE!|N\/A|DIV\/0!|NAME\?|NUM!|SPILL!|CALC!|NULL!)/.test(String(v));
const round6 = v => Math.round(v*1e6)/1e6;

export function columnLetter(index) {
  if (!Number.isInteger(index) || index < 0 || index > 16383) throw new UserError('源列地址无效。');
  let s='', n=index+1; while(n){const m=(n-1)%26;s=String.fromCharCode(65+m)+s;n=Math.floor((n-1)/26);} return s;
}
export function resolveOptions(input={}) {
  const o={title:'三指标组合图',variant:'bubble',sizeRule:'diameter',maxBubble:30,showLabels:true,...input};
  if(typeof o.title!=='string'||o.title.length>80)throw new UserError('标题最多 80 个字符。');
  if(!VARIANTS.includes(o.variant))throw new UserError('图表变体无效。');
  if(!['diameter','area'].includes(o.sizeRule))throw new UserError('气泡大小规则无效。');
  const size=Number(o.maxBubble);
  if(!Number.isFinite(size)||size<10||size>60)throw new UserError('最大气泡应为 10–60 磅。');
  o.maxBubble=Math.round(size);
  if(typeof o.showLabels!=='boolean')throw new UserError('显示标签必须是布尔值。');
  return o;
}
export function validate(values) {
  if(!Array.isArray(values)||values.length<3||values.length>MAX_ROWS+1||values.some(r=>!Array.isArray(r)||r.length!==4))
    throw new UserError(`选择含表头的 4 列：分类、底部柱指标、折线指标、顶部指标；2–${MAX_ROWS} 行数据。`);
  if(values[0].some(v=>typeof v!=='string'||!v.trim()||v.length>255||invalidText(v)))throw new UserError('4 列均须有非空文本表头。');
  values.slice(1).forEach((r,i)=>{
    if(!['string','number'].includes(typeof r[0])||!String(r[0]).trim()||invalidText(r[0])||(typeof r[0]==='number'&&!Number.isFinite(r[0])))throw new UserError(`第 ${i+2} 行分类为空或无效。`);
    r.slice(1).forEach((v,j)=>{if(typeof v!=='number'||!Number.isFinite(v))throw new UserError(`第 ${i+2} 行第 ${j+2} 列必须是真实数值（不接受文本数字或空白）。`);});
    if(r[1]<0||r[3]<0)throw new UserError(`第 ${i+2} 行：柱形/气泡指标不能为负。`);
  });
  return {count:values.length-1,categoryHeader:values[0][0],headers:values[0].slice(1).map(String)};
}
function validateSource(source,o) {
  const info=validate(source.values);
  if(info.count!==source.count||!Number.isInteger(source.rowIndex)||source.rowIndex<0||source.rowIndex+info.count+1>1048576||!Number.isInteger(source.columnIndex)||source.columnIndex<0||source.columnIndex+4>16384||typeof source.sheet!=='string'||!source.sheet)throw new UserError('源区域地址或行数无效。');
  const rows=source.values.slice(1);
  if(o.variant==='log'&&rows.some(r=>r[1]<=0||r[3]<=0))throw new UserError('对数刻度要求两个柱形指标都大于 0。');
  if(o.variant!=='log'&&(Math.max(...rows.map(r=>r[1]))<=0||Math.max(...rows.map(r=>r[3]))<=0))throw new UserError('柱形/气泡指标至少要有一个大于 0 的值。');
  return info;
}

/** 二维值的数学模型（与辅助表公式一一对应，供单测和验收期望使用）。 */
export function computeModel(values,input={}) {
  const o=resolveOptions(input),info=validate(values),rows=values.slice(1);
  const m1=rows.map(r=>r[1]),m2=rows.map(r=>r[2]),m3=rows.map(r=>r[3]);
  const max1=Math.max(...m1),max3=Math.max(...m3),lo2=Math.min(0,...m2),hi2=Math.max(...m2);
  const norm2=v=>hi2>lo2?(v-lo2)/(hi2-lo2):0.5, B=BANDS[o.variant];
  const model={options:o,info,categories:rows.map(r=>String(r[0]))};
  if(o.variant==='log') {
    model.y0=m1.slice();model.y1=m2.map(v=>round6(norm2(v)));model.y2=m3.slice();
    model.secondary=bandAxis(B.line);model.primary=logBounds([...m1,...m3]);
  } else {
    model.y0=m1.map(v=>round6(v/max1*B.bottom[1]));
    model.y1=m2.map(v=>round6(B.line[0]+norm2(v)*B.line[1]));
    if(o.variant==='layered'){model.y2=m3.map(v=>round6(v/max3));model.secondary=bandAxis(B.top);}
    else {
      model.ratio=m3.map(v=>round6(o.sizeRule==='area'?Math.sqrt(v/max3):v/max3));
      // Excel ROUND rounds half away from zero; JS Math.round rounds half up. Values are ≥0 here, identical.
      model.classes=m3.map((v,i)=>v>0?Math.max(1,Math.round(model.ratio[i]*SIZE_CLASSES)):null);
      model.sizes=model.classes.map(k=>k===null?null:classSize(k,o.maxBubble));
    }
    model.primary={minimum:0,maximum:1};
  }
  return model;
}
export const classSize=(k,maxBubble)=>Math.max(2,Math.min(72,Math.round(maxBubble*k/SIZE_CLASSES)));
/** 次轴范围：把 0..1 的归一值放进 [lo, lo+h] 版带。 */
export function bandAxis([lo,h]) {return {minimum:round6(-lo/h),maximum:round6((1-lo)/h)};}
/** 对数主轴：最小值比最小柱低一个数量级，最大柱落在绘图区约 62% 高度处。 */
export function logBounds(values) {
  const positives=values.filter(v=>v>0);if(!positives.length)throw new UserError('对数刻度需要正值。');
  const logMin=Math.floor(Math.log10(Math.min(...positives)))-1,top=Math.log10(Math.max(...positives));
  const logMax=logMin+Math.max(1,top-logMin)/BANDS.log.columnShare;
  return {minimum:10**logMin,maximum:Number((10**logMax).toPrecision(3))};
}
/** 标签公式：沿用源单元格数字格式；常规格式整数不带小数、非整数保留 1–2 位。 */
export function labelText(ref,format) {
  const f=String(format??'').trim();
  if(!f||/^(General|G\/通用格式|@)$/i.test(f))return `=IFERROR(IF(ROUND(${ref},0)=${ref},TEXT(${ref},"0"),TEXT(${ref},"0.0#")),"")`;
  return `=IFERROR(TEXT(${ref},"${f.replace(/"/g,'""')}"),"")`;
}
export const seriesCount=o=>o.variant==='bubble'?3+SIZE_CLASSES:3;
const CLASS_FIRST=12; // M 列
const classColumn=k=>columnLetter(CLASS_FIRST+k-1); // M..AA
const CARRIER='AB';

/** 纯公式计划。A1、I1:M1、R1 归 core；本模块只用 B1:F1 和第 2 行起。 */
export function buildHelperPlan(source,input={},formats=[]) {
  const o=resolveOptions(input),info=validateSource(source,o),n=info.count,last=n+2,B=BANDS[o.variant];
  const ref=(offset,row)=>`${quote(source.sheet)}!$${columnLetter(source.columnIndex+offset)}$${row}`;
  const raw=r=>`=IF(ISNUMBER(${r}),${r},NA())`, header=source.rowIndex+1;
  const norm2=r=>`IF($D$1>$C$1,(C${r}-$C$1)/($D$1-$C$1),0.5)`;
  const scalars=[[`=MAX(B3:B${last})`,`=MIN(0,MIN(C3:C${last}))`,`=MAX(C3:C${last})`,`=MAX(D3:D${last})`,`=${n}`]];
  const rows=[],bubble=[];
  for(let i=0;i<n;i++) {
    const r=i+3,sr=source.rowIndex+i+2;
    let y0,y1,y2;
    if(o.variant==='log'){y0=`=IFERROR(IF(B${r}>0,B${r},NA()),NA())`;y1=`=IFERROR(${norm2(r)},NA())`;y2=`=IFERROR(IF(D${r}>0,D${r},NA()),NA())`;}
    else {
      y0=`=IFERROR(B${r}/$B$1*${B.bottom[1]},NA())`;y1=`=IFERROR(${B.line[0]}+${norm2(r)}*${B.line[1]},NA())`;
      y2=o.variant==='layered'?`=IFERROR(D${r}/$E$1,NA())`:'=NA()';
    }
    rows.push([`=IFERROR(${ref(0,sr)}&"","")`,raw(ref(1,sr)),raw(ref(2,sr)),raw(ref(3,sr)),y0,y1,y2,
      labelText(`B${r}`,formats[0]),labelText(`C${r}`,formats[1]),labelText(`D${r}`,formats[2])]);
    if(o.variant==='bubble') {
      const ratio=o.sizeRule==='area'?`SQRT(MAX(0,D${r})/$E$1)`:`MAX(0,D${r})/$E$1`;
      const line=[`=IFERROR(${ratio},NA())`,`=IFERROR(IF(D${r}>0,MAX(1,ROUND(K${r}*${SIZE_CLASSES},0)),NA()),NA())`];
      for(let k=1;k<=SIZE_CLASSES;k++)line.push(`=IFERROR(IF(L${r}=${k},${B.bubbleY},NA()),NA())`);
      line.push(`=IF(ISNUMBER(D${r}),${B.labelY},NA())`);
      bubble.push(line);
    }
  }
  const headers=[`=${ref(0,header)}&""`,`=${ref(1,header)}&""`,`=${ref(2,header)}&""`,`=${ref(3,header)}&""`,'Y·底部','Y·折线','Y·顶部','标签1','标签2','标签3'];
  const writes=[{address:'B1:F1',formulas:scalars},{address:'A2:J2',formulas:[headers]},{address:`A3:J${last}`,formulas:rows}];
  if(o.variant==='bubble'){
    writes.push({address:`K2:${CARRIER}2`,formulas:[['比例','档位',...Array.from({length:SIZE_CLASSES},(_,k)=>`档${k+1}`),'标签行']]});
    writes.push({address:`K3:${CARRIER}${last}`,formulas:bubble});
  }
  return {options:o,info,n,last,seriesCount:seriesCount(o),writes,
    logAxis:o.variant==='log'?logBounds(source.values.slice(1).flatMap(r=>[r[1],r[3]])):null};
}
async function stage(context,name,write) {
  try{await write();await context.sync();}
  catch(cause){const e=new UserError(`三数量级图（${name}）：${cause.message||String(cause)}`);e.code=cause.code;e.debugInfo=cause.debugInfo;e.cause=cause;throw e;}
}
async function writeHelper({context,helper,sourceSheet,source,options,previous,chart}) {
  const fmt=sourceSheet.getRangeByIndexes(source.rowIndex+1,source.columnIndex+1,1,3);fmt.load('numberFormat');await context.sync();
  const plan=buildHelperPlan(source,options,fmt.numberFormat[0]);
  if(previous) {
    if((previous.options?.variant||'bubble')!==plan.options.variant)throw new UserError('重绑不能改变变体，请另建图表。');
    chart.series.load('items/chartType,items/axisGroup,items/hasDataLabels');await context.sync();
    if(chart.series.items.length!==plan.seriesCount)throw new UserError('原图系列数量已变化，未重绑；请保留原实例并另建图表。');
    plan.labelVisibility=chart.series.items.map(s=>s.hasDataLabels===true);
    helper.getRange(`A3:${CARRIER}${Math.max(previous.count,plan.n)+2}`).clear('Contents');
  }
  for(const w of plan.writes)helper.getRange(w.address).formulas=w.formulas;
  return plan;
}
function seriesName(plan,i) {
  const h=plan.info.headers;
  if(i<3)return h[i];
  return i===plan.seriesCount-1?`${h[2]}·标签`:`${h[2]}·档${i-1}`;
}
// bubble: 0 柱 | 1 折线 | 2..16 气泡档1..15（档1兼作图例）| 17 标签载体
function seriesRole(plan,i) {
  const v=plan.options.variant;
  if(i===1)return {type:'LineMarkers',axis:v==='log'?'Secondary':'Primary',y:'F'};
  if(i===0)return {type:'ColumnClustered',axis:'Primary',y:'E'};
  if(v==='layered')return {type:'ColumnClustered',axis:'Secondary',y:'G'};
  if(v==='log')return {type:'ColumnClustered',axis:'Primary',y:'G'};
  if(i===plan.seriesCount-1)return {type:'Line',axis:'Primary',y:CARRIER};
  return {type:'LineMarkers',axis:'Primary',y:classColumn(i-1),size:classSize(i-1,plan.options.maxBubble)};
}
function bind(chart,helper,plan) {
  const r=a=>helper.getRange(a),last=plan.last;
  for(let i=0;i<plan.seriesCount;i++){const s=chart.series.getItemAt(i),role=seriesRole(plan,i);s.setXAxisValues(r(`A3:A${last}`));s.setValues(r(`${role.y}3:${role.y}${last}`));}
}
async function createChart({context,helper,sourceSheet,source,instance,binding:plan}) {
  const chart=sourceSheet.charts.add('ColumnClustered',helper.getRange(`E3:E${plan.last}`),'Columns');
  chart.name=instance.chartName;chart.plotVisibleOnly=false;chart.displayBlanksAs='NotPlotted';chart.series.load('items');
  const anchor=sourceSheet.getRangeByIndexes(source.rowIndex,source.columnIndex+source.columnCount+1,1,1);anchor.load('left,top');await context.sync();
  if(chart.series.items.length!==1)throw new UserError('Excel 未生成预期的单个初始系列。');
  chart.series.getItemAt(0).name=seriesName(plan,0);
  for(let i=1;i<plan.seriesCount;i++)chart.series.add(seriesName(plan,i));
  await context.sync();
  await stage(context,'系列类型与轴组',()=>{for(let i=1;i<plan.seriesCount;i++){const s=chart.series.getItemAt(i),role=seriesRole(plan,i);s.chartType=role.type;s.axisGroup=role.axis;}});
  bind(chart,helper,plan);
  chart.left=anchor.left;chart.top=anchor.top;chart.width=LAYOUT.width;chart.height=LAYOUT.height;await context.sync();
  return chart;
}
function font(f,size=9,color=COLORS.text){f.name=FONT;f.size=size;f.color=color;f.bold=false;}
function palette(v){return v==='layered'?[COLORS.layeredBottom,COLORS.line,COLORS.layeredTop]:v==='log'?[COLORS.logA,COLORS.line,COLORS.logB]:[COLORS.bottom,COLORS.line,COLORS.top];}
function labelTargets(plan) {
  const v=plan.options.variant;
  return [{series:0,column:'H',position:'OutsideEnd'},{series:1,column:'I',position:'Top'},
    v==='bubble'?{series:plan.seriesCount-1,column:'J',position:'Center'}:{series:2,column:'J',position:'OutsideEnd'}];
}
// 逐点标签链接辅助表文本单元格：改数即变；重绑只给新增点补标签，不改已有标签格式。
async function labels(context,chart,helper,plan,start=0,visibility=null) {
  if(!plan.options.showLabels)return;
  helper.load('name');await context.sync();
  const targets=labelTargets(plan).filter(t=>!visibility||visibility[t.series]);
  if(start===0)await stage(context,'启用标签',()=>{for(const t of targets){const s=chart.series.getItemAt(t.series);s.hasDataLabels=true;s.dataLabels.showValue=false;s.dataLabels.showCategoryName=false;s.dataLabels.showSeriesName=false;s.dataLabels.showLegendKey=false;font(s.dataLabels.format.font,9,palette(plan.options.variant)[Math.min(2,t.series)]);}});
  await stage(context,'链接标签单元格',()=>{for(const t of targets)for(let i=start;i<plan.n;i++){const p=chart.series.getItemAt(t.series).points.getItemAt(i);p.hasDataLabel=true;p.dataLabel.formula=`=${quote(helper.name)}!$${t.column}$${i+3}`;p.dataLabel.position=t.position;}});
}
async function axes(context,chart,plan) {
  const v=plan.options.variant,m=computeAxisPlan(plan);
  await stage(context,'分类轴',()=>{const c=chart.axes.getItem('Category','Primary');c.tickLabelPosition='Low';c.majorTickMark='None';c.minorTickMark='None';c.format.line.color=COLORS.axis;c.format.line.weight=0.75;font(c.format.font);});
  const p=chart.axes.getItem('Value','Primary');
  if(v==='log')await stage(context,'主数值轴/对数刻度',()=>{p.scaleType='Logarithmic';p.logBase=10;});
  await stage(context,'主数值轴/范围',()=>{p.minimum=m.primary.minimum;p.maximum=m.primary.maximum;p.majorGridlines.visible=false;p.minorGridlines.visible=false;p.majorTickMark='None';p.minorTickMark='None';p.tickLabelPosition='None';p.title.visible=false;});
  await stage(context,'主数值轴/隐藏',()=>{p.visible=false;});
  if(!m.secondary)return;
  const s=chart.axes.getItem('Value','Secondary');
  await stage(context,'次数值轴/启用',()=>{s.visible=true;});
  await stage(context,'次数值轴/范围',()=>{s.minimum=m.secondary.minimum;s.maximum=m.secondary.maximum;s.majorGridlines.visible=false;s.minorGridlines.visible=false;s.majorTickMark='None';s.minorTickMark='None';s.tickLabelPosition='None';s.title.visible=false;});
  await stage(context,'次数值轴/隐藏',()=>{s.visible=false;});
}
export function computeAxisPlan(plan) {
  const v=plan.options.variant;
  if(v==='log')return {primary:plan.logAxis,secondary:bandAxis(BANDS.log.line)};
  return {primary:{minimum:0,maximum:1},secondary:v==='layered'?bandAxis(BANDS.layered.top):null};
}
async function style({context,chart,helper,binding:plan}) {
  const o=plan.options,[c0,c1,c2]=palette(o.variant);
  await stage(context,'系列格式',()=>{
    for(let i=0;i<plan.seriesCount;i++)chart.series.getItemAt(i).hasDataLabels=false;
    const s0=chart.series.getItemAt(0);s0.gapWidth=o.variant==='log'?60:90;s0.format.fill.setSolidColor(c0);s0.format.line.lineStyle='None';
    const s1=chart.series.getItemAt(1);s1.format.line.color=c1;s1.format.line.weight=2;s1.markerStyle='Circle';s1.markerSize=5;s1.markerBackgroundColor=c1;s1.markerForegroundColor=c1;s1.smooth=false;
    if(o.variant!=='bubble'){const s2=chart.series.getItemAt(2);s2.gapWidth=o.variant==='log'?60:90;if(o.variant==='log')s2.overlap=-10;s2.format.fill.setSolidColor(c2);s2.format.line.lineStyle='None';return;}
    for(let i=2;i<plan.seriesCount-1;i++){const s=chart.series.getItemAt(i),role=seriesRole(plan,i);s.format.line.lineStyle='None';s.markerStyle='Circle';s.markerSize=role.size;s.markerBackgroundColor=c2;s.markerForegroundColor=c2;}
    const carrier=chart.series.getItemAt(plan.seriesCount-1);carrier.format.line.lineStyle='None';carrier.markerStyle='None';
  });
  await axes(context,chart,plan);
  await stage(context,'标题图例',()=>{
    chart.format.fill.setSolidColor('#FFFFFF');chart.format.border.color=COLORS.axis;chart.format.border.weight=0.5;font(chart.format.font);
    chart.title.visible=Boolean(o.title);chart.title.text=o.title;font(chart.title.format.font,14,'#202020');
    chart.legend.visible=true;chart.legend.position='Top';font(chart.legend.format.font,9);
  });
  await labels(context,chart,helper,plan);
  if(o.variant==='bubble')await hideAuxLegend(context,chart);
  return {notes:[
    o.variant==='bubble'?`气泡=带标记折线（线条隐藏），按 ${SIZE_CLASSES} 档固定标记大小分系列，改数由公式换档；大小量化误差 ≤ ${(o.maxBubble/SIZE_CLASSES/2).toFixed(1)} 磅。`:o.variant==='layered'?'底部柱（主轴）/折线（主轴）/顶部柱（次轴）分层，版带由公式归一化 + 固定轴范围实现。':'两柱共享对数主轴，折线归一化放到次轴顶部版带；对数轴范围按生成时数据计算，重绑时更新。',
    '标签逐点链接辅助表文本单元格，沿用源列数字格式（生成/重绑时读取）。'],variant:o.variant,nativeVerified:false};
}
async function hideAuxLegend(context,chart) {
  const entries=chart.legend.legendEntries;entries.load('items');await context.sync();
  for(let i=3;i<entries.items.length;i++)entries.items[i].visible=false;await context.sync();
}
async function rebind({context,chart,helper,previous,binding:plan}) {
  bind(chart,helper,plan);await context.sync();
  if(plan.options.variant==='log')await stage(context,'对数轴范围',()=>{const p=chart.axes.getItem('Value','Primary');p.minimum=plan.logAxis.minimum;p.maximum=plan.logAxis.maximum;});
  if(plan.n>previous.count)await labels(context,chart,helper,plan,previous.count,plan.labelVisibility);
  return {notes:['只更新辅助公式、系列引用和新增点标签；保留现有系列颜色、标签格式和尺寸。'],variant:plan.options.variant};
}
const inputSpec={category:{id:'comparison',title:'比较'},minVersion:'1.9',columns:4,minRows:2,maxRows:MAX_ROWS,
  description:'4 列含表头：分类、底部柱指标（≥0）、折线指标、顶部指标（气泡/顶部柱，≥0）。三个指标数量级不同也能同图呈现。',
  options:[
    {id:'title',label:'标题',type:'text',default:'三指标组合图',maxLength:80},
    {id:'variant',label:'变体',type:'select',default:'bubble',choices:[{value:'bubble',label:'柱形＋折线＋气泡'},{value:'layered',label:'上下分层：双柱＋折线'},{value:'log',label:'对数刻度：双柱＋折线'}]},
    {id:'sizeRule',label:'气泡大小',type:'select',default:'diameter',choices:[{value:'diameter',label:'直径与数值成正比（原文）'},{value:'area',label:'面积与数值成正比'}]},
    {id:'maxBubble',label:'最大气泡',type:'number',default:30,min:10,max:60,step:1,suffix:'磅'},
    {id:'showLabels',label:'显示数值标签',type:'checkbox',default:true}],
  storage:{marker:'ChartCraft.ThreeScale.v1',helperPrefix:'CC_三量级_',chartPrefix:'CC_ThreeScale_'}};
export default {id:'threescale',title:'三数量级组合图',inputSpec,validate,writeHelper,createChart,style,rebind};
