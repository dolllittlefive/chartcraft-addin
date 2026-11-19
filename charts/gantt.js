import {UserError} from '../core/errors.js';

// F114: a native stacked-bar/XY combination chart. Every plotted value and
// displayed label refers to an Excel formula, including the date-axis labels.
const MARK = 'ChartCraft.Gantt.v1';
const HELPER_PREFIX = 'CC_甘特_';
const CHART_PREFIX = 'CC_Gantt_';
const MAX_ROWS = 200;
const MAX_DATE = 2958465;
const FONT = '微软雅黑'; // Original chart font is not available in the raster PDF.
const COLORS = {bar:'#43A2ED', completed:'#0A1C2C', line:'#171717', grid:'#EEEEEE', text:'#777777', today:'#8E2632', marker:'#485444', frame:'#D9D9D9'};
// PDF342 Image4744 is 800×287 px, placed at 600×215.25 pt. Plot is measured,
// not an original Excel property: x≈78..751, y≈76..267 in the embedded image.
const LAYOUT = {width:600, height:215.25, plot:{left:58.5, top:57, width:504.75, height:143.25}, gapWidth:90, tickIntervals:8};
const quote = value => `'${String(value).replace(/'/g,"''")}'`;
const header = value => String(value ?? '').trim().toLowerCase().replace(/[\s（）()_%％]/g,'');
const columnLetter = index => {let out='', n=index+1; while(n>0){const m=(n-1)%26;out=String.fromCharCode(65+m)+out;n=Math.floor((n-1)/26);}return out;};
const fail = message => {throw new UserError(message);};
// Keep the original Office error/code/debugInfo and stop at the failing stage.
async function syncStage(context,stage){
  try{await context.sync();}
  catch(error){error.message=`甘特图[${stage}]：${error.message}`;throw error;}
}
const isDate = value => typeof value==='number' && Number.isInteger(value) && value>=1 && value<=MAX_DATE;
const isDuration = value => typeof value==='number' && Number.isInteger(value) && value>=0 && value<=MAX_DATE;
const progressHeaders = new Set(['完成率','完成比例','进度','进度比例','完成百分比','progress','completion']);
const ownerHeaders = new Set(['责任人','负责人','所有者','owner','assignee']);
const endHeaders = new Set(['结束日期','结束时间','结束','end','enddate','finish','finishdate']);
const durationHeaders = new Set(['持续天数','工期','工期天','持续时间','天数','duration','durationdays']);

function validate(values) {
  if (!Array.isArray(values) || values.length<2 || !Array.isArray(values[0])) fail('请选择含表头及至少一个任务的数据。');
  const width=values[0].length, rows=values.slice(1);
  if (![3,4,5].includes(width) || values.some(row=>!Array.isArray(row)||row.length!==width)) fail('甘特图需要3至5列：任务、开始日期、结束日期或持续天数，可加完成率和责任人。');
  if (rows.length>MAX_ROWS) fail(`最多支持${MAX_ROWS}个任务。`);
  if (values[0].some(v=>typeof v!=='string'||!v.trim())) fail('每列必须有非空文字表头。');
  const mode=endHeaders.has(header(values[0][2]))?'end':durationHeaders.has(header(values[0][2]))?'duration':null;
  if (!mode) fail('第三列表头请使用“结束日期”或“持续天数”，以明确结束日期/工期口径。');
  let progressColumn=null, ownerColumn=null;
  for(let c=3;c<width;c++){
    const h=header(values[0][c]);
    if(progressHeaders.has(h)&&progressColumn===null) progressColumn=c;
    else if(ownerHeaders.has(h)&&ownerColumn===null) ownerColumn=c;
    else fail('附加列表头须为不重复的“完成率”或“责任人”。');
  }
  rows.forEach((row,i)=>{
    const line=i+2;
    if ((typeof row[0]!=='string'&&typeof row[0]!=='number') || !String(row[0]).trim() || (typeof row[0]==='number'&&!Number.isFinite(row[0]))) fail(`第${line}行任务名称为空或无效。`);
    if (!isDate(row[1])) fail(`第${line}行开始日期必须是Excel日期数值（整天），不接受日期文本或时间小数。`);
    if (mode==='end' ? !isDate(row[2]) || row[2]<row[1] : !isDuration(row[2]) || row[1]+row[2]>MAX_DATE) fail(`第${line}行结束日期/工期无效：结束不得早于开始，工期须为非负整数，结束不得超出Excel日期范围。`);
    if (progressColumn!==null && (typeof row[progressColumn]!=='number' || !Number.isFinite(row[progressColumn]) || row[progressColumn]<0 || row[progressColumn]>1)) fail(`第${line}行完成率须为0至1的数字，例如50%存储为0.5。`);
    if (ownerColumn!==null && (row[ownerColumn]===null || row[ownerColumn]===undefined || !String(row[ownerColumn]).trim() || !['string','number'].includes(typeof row[ownerColumn]) || (typeof row[ownerColumn]==='number'&&!Number.isFinite(row[ownerColumn])))) fail(`第${line}行责任人为空或无效。`);
  });
  return {count:rows.length, inputMode:mode, progressColumn, ownerColumn, categoryHeader:String(values[0][0]), valueHeader:mode==='duration'?'持续天数':'结束日期'};
}

function resolveOptions(options, source) {
  const out={title:'图表标题', variant:'auto', placement:'standalone', tickDays:20, referenceDate:'', connections:true, today:true, milestone:true, endLine:false, ...options};
  if(typeof out.title!=='string'||out.title.length>80) fail('标题须为不超过80字的文字。');
  if(!['auto','basic','progress','owner','complete'].includes(out.variant)) fail('未知甘特图变体。');
  if(!['standalone','table-aligned'].includes(out.placement)) fail('未知甘特图放置模式。');
  if(!Number.isInteger(out.tickDays)||out.tickDays<1||out.tickDays>3660) fail('最小日期刻度须为1至3660整天。');
  for(const key of ['connections','today','milestone','endLine']) if(typeof out[key]!=='boolean') fail(`${key}须为布尔值。`);
  const hasProgress=source.progressColumn!==null, hasOwner=source.ownerColumn!==null;
  const variant=out.variant==='auto' ? hasProgress?(hasOwner?'complete':'progress'):(hasOwner?'owner':'basic') : out.variant;
  if(['progress','complete'].includes(variant)&&!hasProgress) fail('此变体需要“完成率”列。');
  if(['owner','complete'].includes(variant)&&!hasOwner) fail('此变体需要“责任人”列。');
  let todayFormula='=TODAY()';
  if(out.referenceDate!==''){
    if(typeof out.referenceDate!=='string'||!/^\d{4}-\d{2}-\d{2}$/.test(out.referenceDate)) fail('基准日请留空（今天）或填写YYYY-MM-DD。');
    const [y,m,d]=out.referenceDate.split('-').map(Number), date=new Date(Date.UTC(y,m-1,d));
    if(y<1900||y>9999||date.getUTCFullYear()!==y||date.getUTCMonth()+1!==m||date.getUTCDate()!==d) fail('基准日不是有效公历日期。');
    todayFormula=`=DATE(${y},${m},${d})`;
  }
  return {...out, variant, todayFormula, variantCode:['basic','progress','owner','complete'].indexOf(variant)};
}

// Pure formula recipe. No source values are copied as chart series constants.
function helperRecipe(source, options={}) {
  const verified=validate(source.values), src={...source,...verified}, opts=resolveOptions(options,src);
  for(const k of ['rowIndex','columnIndex']) if(!Number.isInteger(src[k])||src[k]<0) fail('源区域坐标无效。');
  if(typeof src.sheet!=='string'||!src.sheet) fail('源工作表名称无效。');
  const n=src.count,last=n+2, range=col=>`$${col}$3:$${col}$${last}`;
  const ref=(c,r)=>`${quote(src.sheet)}!$${columnLetter(src.columnIndex+c)}$${src.rowIndex+r+2}`;
  const data=[];
  for(let i=0;i<n;i++){
    const r=i+3, task=ref(0,i), start=ref(1,i), third=ref(2,i), progress=src.progressColumn===null?'0':ref(src.progressColumn,i), owner=src.ownerColumn===null?'""':ref(src.ownerColumn,i);
    const valid=[`LEN(TRIM(${task}&""))>0`,`ISNUMBER(${start})`,`${start}>=1`,`${start}<=${MAX_DATE}`,`MOD(${start},1)=0`,`ISNUMBER(${third})`,`MOD(${third},1)=0`,src.inputMode==='end'?`${third}>=${start}`:`${third}>=0`,src.inputMode==='end'?`${third}<=${MAX_DATE}`:`${start}+${third}<=${MAX_DATE}`];
    if(src.progressColumn!==null) valid.push(`ISNUMBER(${progress})`,`${progress}>=0`,`${progress}<=1`);
    if(src.ownerColumn!==null) valid.push(`LEN(TRIM(${owner}&""))>0`);
    const duration=src.inputMode==='end'?`${third}-${start}`:third;
    data.push([
      `=IFERROR(${task}&"","输入无效")`,
      `=IF(H${r}="OK",${start},0)`,
      `=IF(H${r}="OK",${duration},0)`,
      `=B${r}+C${r}`,
      `=IF(H${r}="OK",${progress},0)`,
      `=IFERROR(${owner}&"","")`,
      `=IF(H${r}<>"OK",0,IF(C${r}=0,0,NETWORKDAYS(B${r},D${r}-1)))`,
      `=IF(IFERROR(AND(${valid.join(',')}),FALSE),"OK","输入无效")`,
      `=IF(H${r}="OK",(B${r}-$B$1)/$E$1,0)`,
      `=IF(AND(H${r}="OK",OR($G$1=1,$G$1=3)),C${r}*E${r}/$E$1,0)`,
      `=IF(H${r}="OK",C${r}/$E$1-J${r},0)`,
      `=A${r}&IF(H${r}="OK",""," [输入无效]")`,
      `=IF(H${r}<>"OK","输入无效",IF($G$1=0,TEXT(C${r},"0")&"("&TEXT(G${r},"0")&")",IF($G$1=2,F${r},TEXT(E${r},"0%")&IF(E${r}=1," √",IF($F$1>=D${r}," ×",""))&IF($G$1=3,"  "&F${r},""))))`,
      `=1-(${i+1}-0.5)/${n}`,
      `=IF(H${r}="OK",(D${r}-$B$1)/$E$1,0)`,
      `=IF(H${r}="OK",TEXT(E${r},"0%"),"")`,
      `=IF(H${r}="OK",1,0)`
    ]);
  }
  const tickRows=Array.from({length:9},(_,i)=>[String(i/8),'1',`=TEXT($B$1+${i}*$E$1/8,"yyyy/m/d")`]);
  const connections=[], arrows=[];
  for(let i=0;i<Math.max(1,n-1);i++){
    const a=i+3,b=a+1, enabled=i<n-1?`AND($N$1=1,H${a}="OK",H${b}="OK")`:'FALSE';
    connections.push([`=IF(${enabled},O${a},NA())`,`=IF(${enabled},N${a}-0.26/${n},NA())`],[`=IF(${enabled},I${b},NA())`,`=IF(${enabled},N${b}+0.26/${n},NA())`],['=NA()','=NA()']);
    arrows.push([`=IF(${enabled},I${b},NA())`,`=IF(${enabled},N${b}+0.26/${n},NA())`,`=IF(${enabled},"▼","")`]);
  }
  const todayVisible=`AND($O$1=1,COUNTIF(${range('H')},"OK")>0,$F$1>=$B$1,$F$1<=MAX(${range('D')}))`;
  const endVisible=`AND($P$1=1,COUNTIF(${range('H')},"OK")>0)`;
  return {source:src,options:opts,last,blocks:[
    {address:'B1:H1',formulas:[[
      `=MIN(2958457,IF(COUNTIF(${range('H')},"OK")=0,TODAY(),MINIFS(${range('B')},${range('H')},"OK")))`,
      `=MAX(${opts.tickDays},CEILING(MAX(0,MAX(${range('D')})-B1)/8,${opts.tickDays}))`,
      '=MIN(2958465,B1+8*C1)', '=MAX(1,D1-B1)', opts.todayFormula, String(opts.variantCode),String(n)
    ]]},
    {address:'N1:Q1',formulas:[[String(Number(opts.connections)),String(Number(opts.today)),String(Number(opts.milestone)),String(Number(opts.endLine))]]},
    {address:'A2:Q2',formulas:[['任务','开始日期','持续天数','结束日期（不含）','完成率','责任人','工作日','校验','开始偏移比例','完成比例长度','剩余比例长度','分类标签','右侧标签','行位置','终点位置','完成率标签','有效行']]},
    {address:`A3:Q${last}`,formulas:data},
    {address:'S3:U11',formulas:tickRows},
    {address:`W3:X${connections.length+2}`,formulas:connections},
    {address:`Z3:AB${arrows.length+2}`,formulas:arrows},
    {address:'AD3:AF4',formulas:[[ `=IF(${todayVisible},($F$1-$B$1)/$E$1,NA())`,`=IF(${todayVisible},0,NA())`,`=IF(${todayVisible},"▲","")`],[ '=AD3',`=IF(${todayVisible},1,NA())`,'=""']]},
    {address:'AH3:AJ4',formulas:[[ `=IF(${endVisible},(MAX(${range('D')})-$B$1)/$E$1,NA())`,`=IF(${endVisible},0,NA())`,`=IF(${endVisible},"▲","")`],[ '=AH3',`=IF(AND(${endVisible},$Q$1=1),1,NA())`,'=""']]}
  ],connectionLast:connections.length+2,arrowLast:arrows.length+2};
}

async function writeHelper({helper,source,options,previous}) {
  const recipe=helperRecipe(source,options); // Validate before queueing recipe mutations.
  if(previous){
    const old=Math.max(Number(previous.count)||0,recipe.source.count);
    helper.getRange(`A3:Q${old+2}`).clear(Excel.ClearApplyTo.contents);
    helper.getRange(`W3:AB${Math.max(3,3*(old-1))+2}`).clear(Excel.ClearApplyTo.contents);
  }
  for(const block of recipe.blocks) helper.getRange(block.address).formulas=block.formulas;
  helper.getRange(`B3:B${recipe.last}`).numberFormat=Array.from({length:recipe.source.count},()=>['yyyy/m/d']);
  helper.getRange(`D3:D${recipe.last}`).numberFormat=Array.from({length:recipe.source.count},()=>['yyyy/m/d']);
  return {last:recipe.last,connectionLast:recipe.connectionLast,arrowLast:recipe.arrowLast,count:recipe.source.count};
}

function bindSeries(chart,helper,binding){
  const {last,connectionLast,arrowLast}=binding,s=i=>chart.series.getItemAt(i);
  for(let i=0;i<3;i++){s(i).setXAxisValues(helper.getRange(`L3:L${last}`));s(i).setValues(helper.getRange(`${['I','J','K'][i]}3:${['I','J','K'][i]}${last}`));}
  const specs=[[3,'S3:S11','T3:T11'],[4,`W3:W${connectionLast}`,`X3:X${connectionLast}`],[5,`Z3:Z${arrowLast}`,`AA3:AA${arrowLast}`],[6,'AD3:AD4','AE3:AE4'],[7,'AH3:AH4','AI3:AI4'],[8,`O3:O${last}`,`N3:N${last}`]];
  for(const [index,x,y] of specs){s(index).setXAxisValues(helper.getRange(x));s(index).setValues(helper.getRange(y));}
}

function labelGroups(binding){
  return [
    {series:8,name:'任务标签',count:binding.count,col:'M',position:'Right',size:7,color:COLORS.text,allPoints:true},
    {series:3,name:'日期刻度',count:9,col:'U',position:'Top',size:7,color:COLORS.text,allPoints:true},
    {series:5,name:'连接箭头',count:binding.arrowLast-2,col:'AB',position:'Center',size:5.5,color:COLORS.line,allPoints:true},
    {series:6,name:'今天里程碑',count:1,col:'AF',position:'Bottom',size:8,color:COLORS.marker},
    {series:7,name:'项目终点',count:1,col:'AJ',position:'Bottom',size:8,color:COLORS.line}
  ];
}

function linkLabels(chart,helperName,binding){
  for(const g of labelGroups(binding))for(let i=0;i<g.count;i++)chart.series.getItemAt(g.series).points.getItemAt(i).dataLabel.formula=`=${quote(helperName)}!$${g.col}$${i+3}`;
}

async function initializeLabels(context,chart,helperName,binding){
  for(const g of labelGroups(binding)){
    const s=chart.series.getItemAt(g.series),stage=`${g.name}[系列${g.series}]`;
    // r4 Mac rejected the collection's position, despite valid XY types.
    // Materialize linked point labels first; never clear showValue before the
    // labels exist, and never write ChartDataLabels.position on this combo.
    if(g.allPoints){s.hasDataLabels=true;s.dataLabels.showValue=true;}
    for(let i=0;i<g.count;i++){
      const p=s.points.getItemAt(i);p.hasDataLabel=true;
      p.dataLabel.formula=`=${quote(helperName)}!$${g.col}$${i+3}`;
    }
    await syncStage(context,stage+'/启用与公式');
    // Keep a default font for subsequently extended series. Rebind itself
    // remains references-only and does not reset user-formatted point labels.
    if(g.allPoints){s.dataLabels.format.font.name=FONT;s.dataLabels.format.font.size=g.size;s.dataLabels.format.font.color=g.color;}
    for(let i=0;i<g.count;i++){
      const label=s.points.getItemAt(i).dataLabel;
      label.position=g.position;label.format.font.name=FONT;label.format.font.size=g.size;label.format.font.color=g.color;
    }
    await syncStage(context,stage+`/单点位置${g.position}与字体`);
  }
}

async function resolveLayout({context,sourceSheet,source,options,binding}){
  const opts=resolveOptions(options,validate(source.values));
  const anchor=sourceSheet.getRangeByIndexes(source.rowIndex,source.columnIndex+source.columnCount,1,1);anchor.load('left,top');
  let rows;
  if(opts.placement==='table-aligned'){
    rows=sourceSheet.getRangeByIndexes(source.rowIndex+1,source.columnIndex,source.count,1);rows.load('top,height');rows.format.load('rowHeight');
  }
  await syncStage(context,'读取源表位置');
  const extra=Math.max(0,source.count-9)*(LAYOUT.plot.height/9);
  const layout={...LAYOUT,height:LAYOUT.height+extra,plot:{...LAYOUT.plot,height:LAYOUT.plot.height+extra}};
  let top=anchor.top;
  if(rows){
    const height=rows.height,rowHeight=rows.format.rowHeight;
    if(!Number.isFinite(rowHeight)||rowHeight<=0||!Number.isFinite(height)||Math.abs(height-rowHeight*source.count)>0.1) fail('嵌表对齐要求所选任务行等高且可见；请统一行高或选独立图表。');
    top=Math.max(0,anchor.top-24);layout.plot={left:8,top:rows.top-top,width:578,height};layout.height=layout.plot.top+height+14;
  }
  binding.layout=layout;binding.placement=opts.placement;
  return {anchor,top,layout};
}

async function createChart({context,sourceSheet,helper,source,options,instance,binding}){
  const {anchor,top,layout}=await resolveLayout({context,sourceSheet,source,options,binding});
  const chart=sourceSheet.charts.add(Excel.ChartType.barStacked,helper.getRange(`I3:I${binding.last}`),Excel.ChartSeriesBy.columns);
  chart.name=instance.chartName;chart.plotVisibleOnly=false;
  chart.series.getItemAt(0).name='开始偏移';
  for(const name of ['已完成','任务','日期刻度','相邻任务连接','连接箭头','今天','项目终点','任务标签'])chart.series.add(name);
  await syncStage(context,'创建图表与系列');
  for(let i=0;i<3;i++){const s=chart.series.getItemAt(i);s.chartType=Excel.ChartType.barStacked;s.axisGroup=Excel.ChartAxisGroup.primary;}
  for(let i=3;i<9;i++){const s=chart.series.getItemAt(i);s.chartType=[4,6,7].includes(i)?Excel.ChartType.xyscatterLinesNoMarkers:Excel.ChartType.xyscatter;s.axisGroup=Excel.ChartAxisGroup.secondary;}
  bindSeries(chart,helper,binding);
  chart.left=anchor.left+12;chart.top=top;chart.width=layout.width;chart.height=layout.height;
  await syncStage(context,'系列类型、轴组与引用');
  return chart;
}

async function style({context,sourceSheet,chart,source,options,instance,binding}){
  if(!binding.layout) await resolveLayout({context,sourceSheet,source,options,binding});
  const opts=resolveOptions(options,validate(source.values)),s=i=>chart.series.getItemAt(i);
  chart.format.font.name=FONT;chart.format.font.size=7;chart.format.font.color=COLORS.text;
  chart.format.fill.setSolidColor('#FFFFFF');chart.format.border.color=COLORS.frame;chart.format.border.weight=0.5;chart.format.border.lineStyle='Continuous';
  chart.title.visible=opts.placement==='standalone'&&opts.title!=='';chart.title.text=opts.title;chart.title.format.font.size=14;chart.title.format.font.bold=false;chart.title.format.font.color='#101010';chart.title.horizontalAlignment='Left';
  chart.legend.visible=false;
  s(0).format.fill.clear();s(0).format.line.lineStyle='None';
  for(let i=0;i<3;i++){s(i).gapWidth=LAYOUT.gapWidth;s(i).overlap=100;}
  s(1).format.fill.setSolidColor(COLORS.completed);s(1).format.line.lineStyle='None';
  s(2).format.fill.setSolidColor(COLORS.bar);s(2).format.line.lineStyle='Continuous';s(2).format.line.color=COLORS.line;s(2).format.line.weight=0.5;
  for(let i=3;i<9;i++){s(i).markerStyle=Excel.ChartMarkerStyle.none;s(i).format.line.lineStyle=[4,6,7].includes(i)?'Continuous':'None';}
  s(4).format.line.color=COLORS.line;s(4).format.line.weight=0.5;
  s(6).format.line.color=COLORS.today;s(6).format.line.weight=0.75;
  s(7).format.line.color=COLORS.line;s(7).format.line.weight=0.5;
  await syncStage(context,'图表与系列格式');
  const cat=chart.axes.categoryAxis,val=chart.axes.valueAxis;
  cat.reversePlotOrder=true;cat.tickLabelPosition=opts.placement==='table-aligned'?'None':'Low';cat.majorTickMark='None';cat.format.font.name=FONT;cat.format.font.size=8;cat.format.font.color=COLORS.text;cat.format.line.lineStyle='None';
  cat.majorGridlines.visible=true;cat.majorGridlines.format.line.color=COLORS.grid;cat.majorGridlines.format.line.weight=0.4;
  await syncStage(context,'主分类轴');
  val.minimum=0;val.maximum=1;val.majorUnit=1/8;val.tickLabelPosition='None';val.majorTickMark='None';val.format.line.lineStyle='None';val.majorGridlines.visible=true;val.majorGridlines.format.line.color=COLORS.grid;val.majorGridlines.format.line.weight=0.4;
  await syncStage(context,'主数值轴');
  const secondaryAxes=[];
  for(const [type,stage] of [[Excel.ChartAxisType.category,'次散点X轴'],[Excel.ChartAxisType.value,'次散点Y轴']]){
    const a=chart.axes.getItem(type,Excel.ChartAxisGroup.secondary);
    a.minimum=0;a.maximum=1;
    await syncStage(context,stage+'范围');
    // Mac can reject the line-format object of an already hidden secondary
    // scatter axis. Hide the native axis itself; do not access its format or
    // gridline objects. Its 0–1 scale still locates the XY series and labels.
    a.visible=false;
    await syncStage(context,stage+'隐藏');
    a.load('minimum,maximum,visible');
    await syncStage(context,stage+'回读');
    if(a.visible!==false||!Number.isFinite(a.minimum)||!Number.isFinite(a.maximum)||Math.abs(a.minimum)>1e-9||Math.abs(a.maximum-1)>1e-9) fail(`甘特图[${stage}回读]：隐藏次轴后未保留0–1范围，图表创建未通过。`);
    secondaryAxes.push({axis:type,minimum:a.minimum,maximum:a.maximum,visible:a.visible});
  }
  await initializeLabels(context,chart,instance.helperName,binding);
  if(opts.placement==='standalone'){chart.title.left=7.5;chart.title.top=4.5;}
  await syncStage(context,'标题位置');
  const plot=binding.layout.plot;
  for(let i=0;i<2;i++){const p=chart.plotArea;p.insideWidth=plot.width;p.insideHeight=plot.height;p.insideLeft=plot.left;p.insideTop=plot.top;await syncStage(context,`绘图区第${i+1}次定位`);}
  return {notes:['PDF342前8项工期可读；原始起止日期、培训数据与字体参数未认证。','横轴为公式日期标签；真实BarStacked+XY双轴对齐、箭头及标签位置须双端实机验收。','结束日期不含当日：结束=开始+工期；工作日仅扣周末，不含节假日表。'],variant:opts.variant,secondaryAxes};
}

async function rebind({context,chart,helper,instance,binding}){
  // Deliberately no style, chart recreation, axis reset, title assignment,
  // color/marker writes or series add/remove: all nine series are preallocated.
  bindSeries(chart,helper,binding);await syncStage(context,'重绑系列引用');
  linkLabels(chart,instance.helperName,binding);await syncStage(context,'重绑标签引用');
}

const inputSpec={
  category:{id:'project',title:'项目与进度'},
  description:'含表头3–5列：任务、开始日期、结束日期或持续天数，可加完成率(0–1)和责任人。日期须为Excel日期数值；结束不含当日，0天保留零长度。源值由公式自动重算。',
  minVersion:'1.9',columns:3,minRows:1,maxRows:MAX_ROWS,
  options:[
    {id:'title',label:'标题',type:'text',default:'项目任务进度计划',maxLength:80},
    {id:'variant',label:'标签与进度',type:'select',default:'auto',choices:[{value:'auto',label:'按输入列自动选择'},{value:'basic',label:'工期（工作日）'},{value:'progress',label:'完成率与进度条'},{value:'owner',label:'责任人'},{value:'complete',label:'完成率与责任人'}]},
    {id:'placement',label:'放置模式',type:'select',default:'standalone',choices:[{value:'standalone',label:'独立图表（标题与任务标签）'},{value:'table-aligned',label:'对齐源表任务行（等高行）'}]},
    {id:'tickDays',label:'最小日期刻度（天）',type:'number',default:20,min:1,max:3660,step:1},
    {id:'referenceDate',label:'基准日（留空为今天，YYYY-MM-DD）',type:'text',default:'',maxLength:10},
    {id:'connections',label:'连接相邻任务（按行序）',type:'checkbox',default:true},
    {id:'today',label:'显示周期内今天线',type:'checkbox',default:true},
    {id:'milestone',label:'显示项目终点',type:'checkbox',default:true},
    {id:'endLine',label:'显示项目终点竖线',type:'checkbox',default:false}
  ],
  storage:{marker:MARK,helperPrefix:HELPER_PREFIX,chartPrefix:CHART_PREFIX}
};
export {MARK,HELPER_PREFIX,CHART_PREFIX,FONT,COLORS,LAYOUT,validate,resolveOptions,helperRecipe};
export default {id:'gantt',title:'甘特图',inputSpec,validate,writeHelper,createChart,style,rebind};
