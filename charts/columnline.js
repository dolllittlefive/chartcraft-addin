// F086 / PDF221. Native Excel series only; source edits recalculate through formulas.
import {UserError} from '../core/errors.js';

export const FONT = 'SimSun'; // Serif candidate matched to the source image; original font remains unknown.
export const NUMBER_FONT = 'Times New Roman';
export const COLORS = {before:'#0A1C2C', after:'#40A1ED', indicator:'#323FCA', baseline:'#999999', text:'#404040', frame:'#D9D9D9'};
export const LAYOUT = {width:378, height:413.25, plot:{left:20.25, top:72.75, width:334.5, height:294.75}};
export const MAX_ROWS = 100;
const quote = s => `'${String(s).replace(/'/g,"''")}'`;
const choose = (v, values, label) => { if (!values.includes(v)) throw new UserError(`${label}无效。`); return v; };
export function columnLetter(index) {
  if (!Number.isInteger(index) || index < 0 || index > 16383) throw new UserError('源列索引无效。');
  let s='', n=index+1; while(n) {const m=(n-1)%26;s=String.fromCharCode(65+m)+s;n=Math.floor((n-1)/26);} return s;
}
export function resolveOptions(input={}) {
  const o={title:'出色的在线浏览体验',variant:'column-pin',indicator:'difference',indicatorUnit:'points',layout:'pdf221',showLabels:true,...input};
  if(typeof o.title!=='string'||o.title.length>80) throw new UserError('标题应为不超过 80 个字符的文本。');
  choose(o.variant,['column-pin','cluster-line','column-line'],'图表变体');
  choose(o.indicator,['difference','growth','source'],'上方指标');
  choose(o.indicatorUnit,['points','percent','number'],'第三指标单位');
  choose(o.layout,['pdf221','automatic'],'坐标范围');
  if(typeof o.showLabels!=='boolean') throw new UserError('显示标签必须是布尔值。');
  if(o.variant==='column-line'&&o.indicator!=='difference') throw new UserError('单柱＋线使用第二、三列本身；上方指标请保留“之后−之前”。');
  if(o.layout==='pdf221'&&(o.variant==='column-line'||o.indicator==='growth'||o.indicatorUnit==='percent')) throw new UserError('此变体/单位请使用“自动坐标范围”；PDF221 范围仅用于百分数柱和百分点变化。');
  return o;
}
export function validate(values) {
  if(!Array.isArray(values)||values.length<3||values.length>MAX_ROWS+1) throw new UserError(`选择含表头的 3 或 4 列，包含 2–${MAX_ROWS} 行数据。`);
  const width=values[0]?.length;
  if(![3,4].includes(width)||values.some(r=>!Array.isArray(r)||r.length!==width)) throw new UserError('数据必须是规则的 3 列或 4 列矩形。');
  if(values[0].some(v=>typeof v!=='string'||!v.trim()||v.length>255||/^#(REF!|VALUE!|N\/A|DIV\/0!|NAME\?|NUM!|SPILL!|CALC!|NULL!)/.test(v))) throw new UserError('每列必须有非空文本表头（最多 255 字符）。');
  values.slice(1).forEach((r,i)=>{
    if(!['string','number'].includes(typeof r[0])||String(r[0]).trim()===''||(typeof r[0]==='number'&&!Number.isFinite(r[0]))||/^#(REF!|VALUE!|N\/A|DIV\/0!|NAME\?|NUM!|SPILL!|CALC!|NULL!)/.test(String(r[0]))) throw new UserError(`第 ${i+2} 行分类为空或无效。`);
    r.slice(1).forEach((v,j)=>{if(typeof v!=='number'||!Number.isFinite(v)) throw new UserError(`第 ${i+2} 行第 ${j+2} 列必须是真实数值；不接受文本数字或空白。`);});
  });
  return {count:values.length-1,categoryHeader:values[0][0],beforeHeader:values[0][1],afterHeader:values[0][2],indicatorHeader:width===4?values[0][3]:null};
}
function validateSource(source,o) {
  const info=validate(source.values);
  if(info.count!==source.count) throw new UserError('选区行数与校验结果不一致。');
  if(!Number.isInteger(source.rowIndex)||source.rowIndex<0||source.rowIndex+source.count+1>1048576||!Number.isInteger(source.columnIndex)||source.columnIndex<0||source.columnIndex+source.values[0].length>16384||typeof source.sheet!=='string'||!source.sheet) throw new UserError('源区域地址无效。');
  const required=o.indicator==='source'&&o.variant!=='column-line'?4:3;
  if(source.values[0].length!==required) throw new UserError(required===4?'独立第三指标需要第 4 列。':'此模式只使用 3 列，请缩小选区；不静默忽略第 4 列。');
  if(o.indicator==='growth'&&source.values.slice(1).some(r=>r[1]===0)) throw new UserError('增长率的“之前”不能为 0。');
  if(o.layout==='pdf221') {
    if(source.values.slice(1).some(r=>r[1]<0||r[1]>100||r[2]<0||r[2]>100)) throw new UserError('PDF221 版式要求前后值在 0–100；其他数据请使用自动坐标。');
    if(source.values.slice(1).some(r=>{const d=o.indicator==='source'?r[3]:r[2]-r[1];return d< -69.5||d>30.5;})) throw new UserError('变化指标超出 PDF221 测量版式，请使用自动坐标。');
  }
  return info;
}
const indicatorFormat=o=>o.variant==='column-line'?(o.indicatorUnit==='percent'?'0.0%':'0.0'):o.indicator==='growth'||o.indicatorUnit==='percent'?'+0.0%;-0.0%;0.0%':'+0.0;-0.0;0.0';
const countSeries=(o,n)=>o.variant==='column-pin'?n+5:o.variant==='cluster-line'?3:2;
const safeNumber=ref=>`IF(ISNUMBER(${ref}),${ref},NA())`;
// R6 renders native category labels in YaHei 10pt. Budget each category slot
// minus a 6pt gutter; the old 7/7.5-glyph lines exceeded the 66.9pt five-row slots.
// Keep the break as a live formula, preserving all edited text and explicit newlines.
function categoryFormula(ref,value,o,n) {
  const text=`${ref}&""`;
  if(o.layout!=='pdf221')return `=IFERROR(${text},"")`;
  const capacity=Math.max(1,(LAYOUT.plot.width/n-6)/10);let units=0,k=0,overflow=false;
  for(const ch of String(value)){const width=/[\x20-\x7e]/.test(ch)?0.5:1;if(units+width>capacity){overflow=true;break;}units+=width;k+=ch.length;}
  if(!overflow)k=Math.max(1,Math.floor(capacity));
  k=Math.max(1,k);
  return `=IFERROR(IF(ISNUMBER(FIND(CHAR(10),${text})),${text},IF(LEN(${text})>${k},LEFT(${text},${k})&CHAR(10)&MID(${text},${k+1},32767),${text})),"")`;
}


/** Pure formula plan. A1 / I1:M1 / R1 remain exclusively owned by core. */
export function buildHelperPlan(source,input={}) {
  const o=resolveOptions(input),info=validateSource(source,o), n=source.count, last=n+2;
  const cell=(offset,row)=>`${quote(source.sheet)}!$${columnLetter(source.columnIndex+offset)}$${row}`;
  const header=source.rowIndex+1;
  const rows=[];
  for(let i=0;i<n;i++) {
    const r=i+3,sr=source.rowIndex+i+2,a=cell(0,sr),b=cell(1,sr),c=cell(2,sr);
    const d=o.variant==='column-line'?`C${r}`:o.indicator==='source'?safeNumber(cell(3,sr)):o.indicator==='growth'?`IF(B${r}<>0,C${r}/B${r}-1,NA())`:`C${r}-B${r}`;
    rows.push([categoryFormula(a,source.values[i+1][0],o,n),`=${safeNumber(b)}`,`=${safeNumber(c)}`,`=IFERROR(${d},NA())`,
      `=(${i}+0.5)/$G$1`,`=IFERROR(TEXT(D${r},"${indicatorFormat(o)}"),"")`,
      `=IFERROR(IF(AND(ISNUMBER(B${r}),ISNUMBER(C${r}),ISNUMBER(D${r})),"有效","输入无效"),"输入无效")`,
      `=IFERROR(IF(D${r}>=0,D${r},NA()),NA())`,`=IFERROR(IF(D${r}<0,D${r},NA()),NA())`]);
  }
  const stems=[];
  for(let i=0;i<n;i++) {const r=i+3;stems.push([`=E${r}`,'=0'],[`=E${r}`,`=D${r}`]);}
  return {options:o,info,n,last,seriesCount:countSeries(o,n),
    headers:[`=${cell(0,header)}&""`,`=${cell(1,header)}&""`,`=${cell(2,header)}&""`,o.indicator==='source'?`=${cell(3,header)}&""`:'变化','X','变化标签','输入状态','非负变化','负变化'],rows,stems,
    writes:[{address:'G1',formulas:[[String(n)]]},{address:'A2:I2',formulas:null},{address:`A3:I${last}`,formulas:rows},{address:`J3:K${2*n+2}`,formulas:stems},{address:'L3:M4',formulas:[['=0','=0'],['=1','=0']]}]};
}
async function writeHelper({context,helper,source,options,previous,chart}) {
  const plan=buildHelperPlan(source,options);
  if(previous) {
    // Stable series IDs are not exposed here. Reject edits to series identity/order before writes.
    const identities=helper.getRange('B1');identities.load('values');chart.series.load('items/name,items/hasDataLabels');await context.sync();
    let names;try{names=JSON.parse(identities.values[0][0]);}catch{throw new UserError('原图系列身份记录无效，未重绑。');}
    if(!Array.isArray(names)||names.length!==countSeries(plan.options,previous.count)||chart.series.items.length!==names.length||chart.series.items.some((s,i)=>s.name!==names[i])) throw new UserError('原图系列数量、名称或顺序已变化，未重绑；请保留原实例并另建图表。');
    plan.identityNames=names.slice(0,plan.seriesCount);
    plan.previousLabelVisibility=chart.series.items.map(s=>s.hasDataLabels===true);
    while(plan.identityNames.length<plan.seriesCount)plan.identityNames.push(seriesName(plan,plan.identityNames.length));
    helper.getRange(`A3:M${Math.max(2*previous.count+2,2*plan.n+2)}`).clear('Contents');
  } else plan.identityNames=Array.from({length:plan.seriesCount},(_,i)=>seriesName(plan,i));
  helper.getRange('B1').values=[[JSON.stringify(plan.identityNames)]];
  plan.writes[1].formulas=[plan.headers];
  for(const w of plan.writes) helper.getRange(w.address).formulas=w.formulas;
  helper.getRange(`B3:C${plan.last}`).numberFormat=plan.rows.map(()=>['0','0']);
  helper.getRange(`D3:D${plan.last}`).numberFormat=plan.rows.map(()=>[indicatorFormat(plan.options)]);
  return plan;
}
function seriesName(plan,i) {
  if(i===0)return plan.info.beforeHeader;
  if(plan.options.variant==='column-line')return plan.info.afterHeader;
  if(i===1)return plan.info.afterHeader;
  if(i===2)return plan.info.indicatorHeader||(plan.options.indicator==='growth'?'增长率':'变化（百分点）');
  if(i===3)return '负变化'; if(i===4)return '变化零线'; return `变化竖线 ${i-4}`;
}
function bind(chart,helper,p) {
  const s=i=>chart.series.getItemAt(i), r=address=>helper.getRange(address),last=p.last;
  s(0).setXAxisValues(r(`A3:A${last}`));s(0).setValues(r(`B3:B${last}`));
  if(p.options.variant==='column-line') {s(1).setXAxisValues(r(`E3:E${last}`));s(1).setValues(r(`C3:C${last}`));return;}
  s(1).setXAxisValues(r(`A3:A${last}`));s(1).setValues(r(`C3:C${last}`));
  s(2).setXAxisValues(r(`E3:E${last}`));s(2).setValues(r(`${p.options.variant==='column-pin'?'H':'D'}3:${p.options.variant==='column-pin'?'H':'D'}${last}`));
  if(p.options.variant!=='column-pin')return;
  s(3).setXAxisValues(r(`E3:E${last}`));s(3).setValues(r(`I3:I${last}`));
  s(4).setXAxisValues(r('L3:L4'));s(4).setValues(r('M3:M4'));
  for(let i=0;i<p.n;i++){const row=2*i+3;s(i+5).setXAxisValues(r(`J${row}:J${row+1}`));s(i+5).setValues(r(`K${row}:K${row+1}`));}
}
function configureSeriesType(s,i,p) {
  if(i===0||(i===1&&p.options.variant!=='column-line')){s.chartType='ColumnClustered';s.axisGroup='Primary';return;}
  s.chartType=p.options.variant==='column-pin'?(i===2||i===3?'XYScatter':'XYScatterLinesNoMarkers'):'XYScatterLines';s.axisGroup='Secondary';
}
async function createChart({context,helper,sourceSheet,source,instance,binding:p}) {
  const chart=sourceSheet.charts.add('ColumnClustered',helper.getRange(`B3:B${p.last}`),'Columns');
  chart.name=instance.chartName;chart.plotVisibleOnly=false;chart.displayBlanksAs='NotPlotted';
  chart.series.load('items');const anchor=sourceSheet.getRangeByIndexes(source.rowIndex,source.columnIndex+source.columnCount,1,1);anchor.load('left,top');await context.sync();
  if(chart.series.items.length!==1) throw new UserError('新图未生成预期的单个初始系列。');
  chart.series.getItemAt(0).name=seriesName(p,0);
  for(let i=1;i<p.seriesCount;i++)chart.series.add(seriesName(p,i));
  await context.sync();
  for(let i=0;i<p.seriesCount;i++)configureSeriesType(chart.series.getItemAt(i),i,p);
  bind(chart,helper,p);
  chart.left=anchor.left+12;chart.top=anchor.top;chart.width=LAYOUT.width;chart.height=LAYOUT.height;
  await context.sync();return chart;
}
function font(format,size=8,color=COLORS.text,family=FONT){format.font.name=family;format.font.size=size;format.font.color=color;format.font.bold=false;}
function styleStem(s){s.chartType='XYScatterLinesNoMarkers';s.axisGroup='Secondary';s.markerStyle='None';s.format.line.color=COLORS.indicator;s.format.line.weight=0.5;s.hasDataLabels=false;}
function labels(s,format,columnPosition=null){s.hasDataLabels=true;s.dataLabels.showValue=true;s.dataLabels.showCategoryName=false;s.dataLabels.showSeriesName=false;if(columnPosition)s.dataLabels.position=columnPosition;s.dataLabels.linkNumberFormat=false;s.dataLabels.numberFormat=format;font(s.dataLabels.format,8,COLORS.text,NUMBER_FONT);}
// A mixed column/XY chart rejected ChartDataLabels.position on Mac. Position the
// native ChartDataLabel on each point after labels have been enabled and synced.
// Both sign-mask series get positions for every source slot, including NA slots,
// so crossing zero later needs neither plugin code nor a label recreation.
function positionPointLabels(chart,p,start=0,visibility=null) {
  const bars=p.options.variant==='column-line'?1:2;
  const targets=[{series:bars,position:'Top'}];
  if(p.options.variant==='column-pin')targets.push({series:3,position:'Bottom'});
  for(const target of targets) {
    if(visibility&&visibility[target.series]!==true)continue;
    const series=chart.series.getItemAt(target.series);
    for(let i=start;i<p.n;i++){const point=series.points.getItemAt(i);point.hasDataLabel=true;point.dataLabel.position=target.position;}
  }
}
// Re-throw with a precise phase; no failed native operation is suppressed.
async function nativeStage(context,name,write) {
  try {await write();await context.sync();}
  catch(error){const failure=new Error(`[柱线/${name}] ${error.message||String(error)}`);failure.name=error.name||'Error';failure.code=error.code;failure.debugInfo=error.debugInfo;failure.cause=error;throw failure;}
}
// Numeric labels stay linked to the live series values.
async function styleSeries(context,chart,p) {
  const o=p.options,bars=o.variant==='column-line'?1:2;
  for(let i=0;i<bars;i++) await nativeStage(context,`柱系列${i}/样式与标签`,()=>{
    const s=chart.series.getItemAt(i);s.gapWidth=100;s.overlap=0;s.format.fill.setSolidColor(i===0?COLORS.before:COLORS.after);s.format.line.lineStyle='None';
    if(o.showLabels)labels(s,'0','OutsideEnd');else s.hasDataLabels=false;
  });
  await nativeStage(context,o.variant==='column-pin'?'正针头/标记与标签':'折线/线条标记与标签',()=>{
    const s=chart.series.getItemAt(bars);s.markerStyle='Circle';s.markerSize=6;s.markerBackgroundColor=COLORS.indicator;s.markerForegroundColor=COLORS.indicator;
    // XYScatter already means markers only. Its absent line need not be styled.
    if(o.variant!=='column-pin'){s.format.line.color=COLORS.indicator;s.format.line.weight=0.75;}
    if(o.showLabels)labels(s,indicatorFormat(o));else s.hasDataLabels=false;
  });
  if(o.variant!=='column-pin')return;
  await nativeStage(context,'负针头/标记与标签',()=>{
    const s=chart.series.getItemAt(3);s.markerStyle='Circle';s.markerSize=6;s.markerBackgroundColor=COLORS.indicator;s.markerForegroundColor=COLORS.indicator;
    if(o.showLabels)labels(s,indicatorFormat(o));else s.hasDataLabels=false;
  });
  await nativeStage(context,'变化零线/样式',()=>{const s=chart.series.getItemAt(4);s.markerStyle='None';s.format.line.color=COLORS.baseline;s.format.line.weight=0.5;s.hasDataLabels=false;});
  await nativeStage(context,'原生针杆/样式',()=>{for(let i=5;i<p.seriesCount;i++)styleStem(chart.series.getItemAt(i));});
}
async function style({context,chart,binding:p}) {
  const o=p.options;
  await nativeStage(context,'图表区域/标题图例',()=>{
    font(chart.format);chart.format.fill.setSolidColor('#FFFFFF');chart.format.border.color=COLORS.frame;chart.format.border.weight=0.5;
    chart.title.visible=Boolean(o.title);chart.title.text=o.title;font(chart.title.format,14,'#202020');chart.title.format.font.bold=false;chart.title.horizontalAlignment='Left';
    chart.legend.visible=true;chart.legend.position='Top';font(chart.legend.format,9);
  });
  await styleSeries(context,chart,p);
  if(o.showLabels)await nativeStage(context,'XY单点/标签位置',()=>positionPointLabels(chart,p));
  const category=chart.axes.getItem('Category','Primary'),primary=chart.axes.getItem('Value','Primary'),secondary=chart.axes.getItem('Value','Secondary'),sx=chart.axes.getItem('Category','Secondary');
  await nativeStage(context,'主分类X轴/启用',()=>{category.visible=true;});
  await nativeStage(context,'主分类X轴/样式',()=>{
    category.categoryType='TextAxis';category.textOrientation=0;category.tickLabelPosition='Low';category.tickLabelSpacing=1;category.tickMarkSpacing=1;category.alignment='Center';category.majorTickMark='None';category.minorTickMark='None';category.format.line.color='#000000';category.format.line.weight=0.5;font(category.format,9);category.isBetweenCategories=true;
  });
  for(const [axis,name] of [[primary,'主数值Y轴'],[secondary,'次数值Y轴']]) {
    await nativeStage(context,name+'/启用',()=>{axis.visible=true;});
    await nativeStage(context,name+'/线条刻度',()=>{
      axis.majorGridlines.visible=false;axis.minorGridlines.visible=false;axis.majorTickMark='None';axis.minorTickMark='None';axis.format.line.lineStyle='None';font(axis.format);axis.title.visible=false;
    });
  }
  // Hidden secondary X axes may have no line-format object on Mac. Set only
  // the numeric alignment domain and visibility, never lineStyle/font on it.
  await nativeStage(context,'次散点X轴/启用',()=>{sx.visible=true;});
  await nativeStage(context,'次散点X轴/坐标范围',()=>{
    sx.minimum=0;sx.maximum=1;sx.tickLabelPosition='None';sx.majorGridlines.visible=false;sx.minorGridlines.visible=false;sx.majorTickMark='None';sx.minorTickMark='None';sx.title.visible=false;
  });
  await nativeStage(context,'次散点X轴/隐藏',()=>{sx.visible=false;});
  await nativeStage(context,'双数值Y轴/坐标范围与单位',()=>{
    if(o.layout==='pdf221') {
      // Candidate bounds fit measured y=490 (bars) / y=217 (pins), not recovered workbook settings.
      primary.minimum=0;primary.maximum=100;primary.majorUnit=20;secondary.minimum=-69.5;secondary.maximum=30.5;secondary.majorUnit=20;
      primary.tickLabelPosition='None';secondary.tickLabelPosition='None';
    } else {
      primary.minimum='';primary.maximum='';secondary.minimum='';secondary.maximum='';primary.tickLabelPosition='Low';secondary.tickLabelPosition='High';
      primary.title.visible=true;primary.title.text=o.variant==='column-line'?p.info.beforeHeader:'前后值';
      secondary.title.visible=true;secondary.title.text=o.variant==='column-line'?p.info.afterHeader:o.indicator==='growth'?'增长率':o.indicatorUnit==='points'?'变化（百分点）':p.info.indicatorHeader||'变化';
    }
    secondary.numberFormat=indicatorFormat(o);
  });
  const entries=chart.legend.legendEntries;
  await nativeStage(context,'图例/读取条目',()=>entries.load('items'));
  await nativeStage(context,'标题图例/位置与可见项',()=>{
    if(o.variant==='column-pin')for(let i=2;i<entries.items.length;i++)entries.items[i].visible=false;
    // PDF221 r3 legend was 7–8px too far right at 504px width: shift 5.625pt.
    chart.title.left=6;chart.title.top=4;chart.legend.left=o.layout==='pdf221'?9.375:15;chart.legend.top=49;
  });
  if(o.layout==='pdf221')for(let pass=0;pass<2;pass++)await nativeStage(context,`绘图区/定位${pass+1}`,()=>{
    const area=chart.plotArea,l=LAYOUT.plot;area.insideWidth=l.width;area.insideHeight=l.height;area.insideLeft=l.left;area.insideTop=l.top;
  });
  return {notes:['中文使用宋体、数值使用 Times New Roman 候选；原版字体与轴参数未知，PDF221 为测量版式，尚未双端验收。','同一区域改数由 Excel 公式重算；改变数据行数请按选区重绑。'],variant:o.variant,nativeVerified:false};
}
async function rebind({context,chart,helper,binding:p,previous}) {
  // Existing series formatting, axes, legend, title and label positions are never reset.
  if(p.options.variant==='column-pin') {
    if(p.n<previous.count)for(let i=previous.count+4;i>=p.n+5;i--)chart.series.getItemAt(i).delete();
    if(p.n>previous.count)for(let i=previous.count;i<p.n;i++){const s=chart.series.add(seriesName(p,i+5));styleStem(s);}
    await context.sync();
  }
  bind(chart,helper,p);await context.sync();
  if(p.options.showLabels&&p.n>previous.count){positionPointLabels(chart,p,previous.count,p.previousLabelVisibility);await context.sync();}
  if(p.options.variant==='column-pin'&&p.n>previous.count){const entries=chart.legend.legendEntries;entries.load('items');await context.sync();for(let i=previous.count+5;i<entries.items.length;i++)entries.items[i].visible=false;await context.sync();}
  return {variant:p.options.variant,preservedFormatting:true};
}
const inputSpec={category:{id:'comparison',title:'比较'},description:'首行为表头：分类、之前、之后。PDF221 版式的输入约定：43 代表 43%。独立第三指标模式另选第 4 列；单柱＋线为分类、柱值、线值。',minVersion:'1.9',columns:3,minRows:2,maxRows:MAX_ROWS,
  options:[{id:'title',label:'标题',type:'text',default:'出色的在线浏览体验',maxLength:80},
    {id:'variant',label:'变体',type:'select',default:'column-pin',choices:[{value:'column-pin',label:'簇状柱＋大头针（PDF221）'},{value:'cluster-line',label:'簇状柱＋折线'},{value:'column-line',label:'单柱＋折线'}]},
    {id:'indicator',label:'变化指标',type:'select',default:'difference',choices:[{value:'difference',label:'之后−之前'},{value:'growth',label:'之后/之前−1'},{value:'source',label:'第4列独立指标'}]},
    {id:'indicatorUnit',label:'指标单位',type:'select',default:'points',choices:[{value:'points',label:'百分点'},{value:'percent',label:'百分比（0.13=13%）'},{value:'number',label:'普通数值'}]},
    {id:'layout',label:'坐标范围',type:'select',default:'pdf221',choices:[{value:'pdf221',label:'PDF221 测量版式'},{value:'automatic',label:'自动坐标（显示双轴）'}]},
    {id:'showLabels',label:'显示数值标签',type:'checkbox',default:true}],
  storage:{marker:'ChartCraft.ColumnLine.v1',helperPrefix:'CC_柱线_',chartPrefix:'CC_ColumnLine_'}};
export default {id:'columnline',title:'柱线与柱钉图',inputSpec,validate,writeHelper,createChart,style,rebind};
