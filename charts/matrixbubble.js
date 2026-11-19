import {setChartTitleTypography} from '../core/presentation.js';
import {UserError} from '../core/errors.js';

const MARK = 'ChartCraft.MatrixBubble.v1';
const HELPER_PREFIX = 'CC_矩阵气泡_';
const CHART_PREFIX = 'CC_MatrixBubble_';
const FONT = '微软雅黑';
const COLORS = {bubble:'#40A1ED',text:'#404040',grid:'#E8E8E8',frame:'#D9D9D9'};
// PDF281 / Mac 16.114: Scale100 produced a 51.96pt maximum circle versus
// the 33.75pt original. 64 is the measured next-pass calibration candidate;
// native cross-host rendering still needs acceptance (see visual-fix evidence).
const DEFAULT_BUBBLE_SCALE = 64;
const MAX_ROWS = 100, MAX_COLUMNS = 24, MAX_POINTS = 600;
const quote = s => `'${String(s).replaceAll("'", "''")}'`;
const literal = s => "'" + String(s);
function columnLetter(i) { let s=''; for(let n=i+1;n;n=Math.floor((n-1)/26)) s=String.fromCharCode(65+(n-1)%26)+s; return s; }
function present(v) { return (typeof v==='string'&&v.trim()!=='') || (typeof v==='number'&&Number.isFinite(v)); }

/** Wide matrix: row categories + >=2 numeric columns. Two columns: named tiles.
 * Source values are inspected only for validation; chart values always use refs. */
function validate(values) {
  if(!Array.isArray(values)||values.length<3||!Array.isArray(values[0])) throw new UserError('请选择含表头、至少2行数据的矩阵或两列“名称、数值”。');
  const width=values[0].length, count=values.length-1;
  if(width<2||width>MAX_COLUMNS+1||values.some(r=>!Array.isArray(r)||r.length!==width)) throw new UserError(`选区须为完整矩形，含名称列及1–${MAX_COLUMNS}个数值列。`);
  if(count>MAX_ROWS||count*(width-1)>MAX_POINTS) throw new UserError(`最多${MAX_ROWS}行、${MAX_POINTS}个气泡。`);
  values[0].forEach((v,c)=>{if(!present(v))throw new UserError(`第${c+1}列表头为空。`);});
  let positive=0;
  values.slice(1).forEach((row,r)=>{
    if(!present(row[0]))throw new UserError(`第${r+2}行名称为空。`);
    row.slice(1).forEach((v,c)=>{
      if(typeof v!=='number'||!Number.isFinite(v)||v<0) throw new UserError(`第${r+2}行、第${c+2}列须为有限非负数；空白和文本不能当作准确原值。`);
      if(v>0)positive++;
    });
  });
  return {count,matrixColumns:width-1,inputKind:width===2?'tiles':'matrix',categoryHeader:String(values[0][0]),valueHeader:width===2?String(values[0][1]):'矩阵数值',positiveCount:positive};
}
function normalizeOptions(options={}) {
  const defaults={title:'矩阵气泡图',subtitle:'',footnote:'',variant:'auto',tileColumns:4,transpose:false,showValues:true,decimals:0,bubbleScale:DEFAULT_BUBBLE_SCALE,referenceRings:false,colorMode:'single'};
  if(Object.keys(options).some(k=>!Object.hasOwn(defaults,k)))throw new UserError('存在未知矩阵气泡选项。');
  const o={...defaults,...options};
  if(typeof o.title!=='string'||o.title.length>160)throw new UserError('标题须为不超过160字的文本。');
  for(const key of ['subtitle','footnote'])if(typeof o[key]!=='string'||o[key].length>200)throw new UserError('副标题和脚注须为不超过200字的文本。');
  if(!['auto','matrix','tiles'].includes(o.variant))throw new UserError('未知矩阵气泡变体。');
  if(!['single','bands25_30'].includes(o.colorMode))throw new UserError('未知颜色分档。');
  for(const [key,min,max] of [['tileColumns',2,8],['decimals',0,3],['bubbleScale',1,300]])if(!Number.isInteger(o[key])||o[key]<min||o[key]>max)throw new UserError(`${key}须为${min}–${max}的整数。`);
  if(['transpose','showValues','referenceRings'].some(k=>typeof o[k]!=='boolean'))throw new UserError('标签、转置和参考环选项须为布尔值。');
  return o;
}
const rangeRef=(s,r,c)=>`${quote(s.sheet)}!$${columnLetter(s.columnIndex+c)}$${s.rowIndex+r+1}`;
const textRef=ref=>`=IFERROR(${ref}&"","")`;
// Display-only line breaks for the long PDF281 category; source text remains
// unchanged and the label remains formula-linked when the source is edited.
const matrixRowLabel=ref=>`=SUBSTITUTE(${textRef(ref).slice(1)},"2019冠状病毒病大流行","2019冠状"&CHAR(10)&"病毒病大流行")`;
const plainMatrix=p=>p.kind==='matrix'&&!p.bands&&!p.options.referenceRings;
const safeSize=ref=>`=IFERROR(IF(ISNUMBER(${ref}),MAX(0,${ref}),0),0)`;
const invalidValue=ref=>`=IFERROR(IF(ISNUMBER(${ref}),IF(${ref}<0,1,0),IF(${ref}="",0,1)),1)`;
function valueLabel(ref,format,show,percent=false) {
  return show?`=IFERROR(IF(ISNUMBER(${ref}),IF(${ref}<0,"负值",IF(${ref}=0,"",TEXT(${ref},"${format}")${percent?'&"%"':''})),IF(${ref}="","","非数值")),"源错误")`:'=""';
}

/** Pure, auditable formula plan. Shared metadata A1/I1:M1/R1 is never written. */
function formulaPlan(source, options={}) {
  const data=validate(source.values),o=normalizeOptions(options),kind=o.variant==='auto'?data.inputKind:o.variant;
  if(kind==='tiles'&&data.matrixColumns!==1)throw new UserError('城市/项目网格使用两列“名称、数值”；多列表请选择行列矩阵。');
  if(kind==='matrix'&&data.matrixColumns<2)throw new UserError('行列矩阵至少需要名称列和2个数值列。');
  if(kind==='tiles'&&o.transpose)throw new UserError('网格变体不使用转置；请调整每行气泡数。');
  if(o.referenceRings&&(kind!=='matrix'||source.values.slice(1).some(r=>r.slice(1).some(v=>v>100))))throw new UserError('100%参考环适用于0–100的行列矩阵。');
  if(o.colorMode!=='single'&&(kind!=='matrix'||o.referenceRings))throw new UserError('颜色分档使用行列矩阵，不能同时显示100%参考环。');
  const bands=o.colorMode==='bands25_30';
  const rows=kind==='tiles'?Math.ceil(data.count/o.tileColumns):(o.transpose?data.matrixColumns:data.count);
  const columns=kind==='tiles'?o.tileColumns:(o.transpose?data.count:data.matrixColumns);
  const count=kind==='tiles'?data.count:data.count*data.matrixColumns;
  const format='0'+(o.decimals?'.'+'0'.repeat(o.decimals):'');
  const body=[],columnsLabels=[],rowLabels=[],tileLabels=[];
  for(let i=0;i<count;i++){
    const sr=kind==='tiles'?i:Math.floor(i/data.matrixColumns),sc=kind==='tiles'?0:i%data.matrixColumns;
    const r=kind==='tiles'?Math.floor(i/columns):(o.transpose?sc:sr),c=kind==='tiles'?i%columns:(o.transpose?sr:sc);
    const ref=rangeRef(source,sr+1,sc+1),rowRef=rangeRef(source,sr+1,0),colRef=rangeRef(source,0,sc+1);
    body.push([textRef(rowRef),textRef(colRef),`=(${c}+0.5)/${columns}`,`=1-(${r}+0.5)/${rows}`,safeSize(ref),valueLabel(ref,format,o.showValues,o.referenceRings),kind==='tiles'?textRef(rowRef):'=""',invalidValue(ref)]);
    tileLabels.push(kind==='tiles'?[`=C${i+3}`,`=1-(${r}+0.91)/${rows}`,'=0.0000001',`=G${i+3}`]:['=0','=0','=0','=""']);
  }
  for(let c=0;c<columns;c++)columnsLabels.push(kind==='matrix'?[`=(${c}+0.5)/${columns}`,'=0','=0.0000001',textRef(o.transpose?rangeRef(source,c+1,0):rangeRef(source,0,c+1))]:['=0','=0','=0','=""']);
  for(let r=0;r<rows;r++)rowLabels.push(kind==='matrix'?['=0',`=1-(${r}+0.5)/${rows}`,'=0.0000001',(kind==='matrix'&&!bands&&!o.referenceRings?matrixRowLabel:textRef)(o.transpose?rangeRef(source,0,r+1):rangeRef(source,r+1,0))]:['=0','=0','=0','=""']);
  // The unused first transparent tile-name point carries the optional plain
  // matrix footnote. Real C:D:E coordinates/sizes and all series stay intact.
  if(kind==='matrix'&&!bands&&!o.referenceRings&&o.footnote)tileLabels[0]=['=0','=0','=0.0000001',`="${o.footnote.replaceAll('"','""')}"`];
  const blocks=[
    {address:'B1:H1',formulas:[[String(data.matrixColumns),String(rows),String(columns),String(count),`=SUM(H3:H${count+2})`,literal(kind),o.transpose?'1':'0']]},
    {address:'A2:H2',formulas:[['行类别','列类别','X位置','Y位置','面积数值','数值标签','网格名称','无效数值'].map(literal)]},
    {address:`A3:H${count+2}`,formulas:body},
    {address:'J2:M2',formulas:[['X','Y','隐藏气泡','横轴类别'].map(literal)]},
    {address:`J3:M${columns+2}`,formulas:columnsLabels},
    {address:'O2:R2',formulas:[['X','Y','隐藏气泡','纵轴类别'].map(literal)]},
    {address:`O3:R${rows+2}`,formulas:rowLabels},
    {address:'T2:W2',formulas:[['X','Y','隐藏气泡','网格项目名'].map(literal)]},
    {address:`T3:W${count+2}`,formulas:tileLabels}
  ];
  if(kind==='matrix'&&!bands&&!o.referenceRings)blocks.push({address:`AP3:AR${rows+4}`,formulas:[...rowLabels.map((_,i)=>['=0.5',`=P${i+3}`,'=0.0000001']),['=-MAX(10,C1+2)','=-10','=0.0000001'],['=MAX(10,C1+2)','=-10','=0.0000001']]});
  if(o.referenceRings)blocks.push({address:'Y2:AA2',formulas:[['X','Y','100%参考面积'].map(literal)]},{address:`Y3:AA${count+2}`,formulas:body.map((_,i)=>[`=C${i+3}`,`=D${i+3}`,'=100'])});
  if(o.referenceRings||bands)blocks.push({address:'AC2:AF2',formulas:[['X','Y','隐藏气泡','顶部列名'].map(literal)]},{address:`AC3:AF${columns+2}`,formulas:columnsLabels.map((r,i)=>[r[0],'=1','=0.0000001',`=M${i+3}`])});
  if(bands)blocks.push({address:'AH2:AK2',formulas:[['低于25','25至不足30','30及以上','隐藏标签气泡'].map(literal)]},{address:`AH3:AK${count+2}`,formulas:body.map((_,i)=>{const r=i+3;return [`=IF(E${r}<25,E${r},0)`,`=IF(AND(E${r}>=25,E${r}<30),E${r},0)`,`=IF(E${r}>=30,E${r},0)`,'=0.0000001'];})});
  return {kind,rows,columns,pointCount:count,sourceColumns:data.matrixColumns,options:o,bands,blocks,last:count+2};
}
async function writeHelper({context,helper,source,options,previous}) {
  const plan=formulaPlan(source,options);
  if(previous){
    const old=helper.getRange('B1:H1');old.load('values');await context.sync();
    const v=old.values[0],oldCount=Number(v[3]);
    // Existing point identities are row-major. Changing width would silently
    // attach user-edited point styles to different cells; require a new chart.
    if(Number(v[0])!==plan.sourceColumns||String(v[5])!==plan.kind||Number(v[6])!==(plan.options.transpose?1:0))throw new UserError('重绑须保留原数值列数和布局；改变列结构请创建新图，现有手改格式未动。');
    plan.previousPointCount=oldCount;plan.previousRows=Number(v[1]);plan.previousColumns=Number(v[2]);
    const max=Math.max(oldCount,plan.pointCount,Number(v[1]),Number(v[2]),plan.rows,plan.columns)+2;
    for(const columns of [['A','H'],['J','M'],['O','R'],['T','W'],...(plan.options.referenceRings?[['Y','AA']]:[]),...(plan.options.referenceRings||plan.bands?[['AC','AF']]:[]),...(plan.bands?[['AH','AK']]:[]),...(plainMatrix(plan)?[['AP','AR']]:[])])helper.getRange(`${columns[0]}3:${columns[1]}${max}`).clear(Excel.ClearApplyTo.contents);
  }
  for(const block of plan.blocks)helper.getRange(block.address).formulas=block.formulas;
  return plan;
}
function bindSeries(chart,helper,p) {
  const defs=[['C','D',p.bands?'AK':'E',p.pointCount],['J','K','L',p.columns],['O','P','Q',p.rows],['T','U','V',p.pointCount]];
  if(plainMatrix(p))defs.push(['AP','AQ','AR',p.rows+2]);
  if(p.options.referenceRings)defs.push(['Y','Z','AA',p.pointCount],['AC','AD','AE',p.columns]);
  if(p.bands)defs.push(...['AH','AI','AJ'].map(z=>['C','D',z,p.pointCount]),['AC','AD','AE',p.columns]);
  defs.forEach(([x,y,z,n],i)=>{const s=chart.series.getItemAt(i);s.setXAxisValues(helper.getRange(`${x}3:${x}${n+2}`));s.setValues(helper.getRange(`${y}3:${y}${n+2}`));s.setBubbleSizes(helper.getRange(`${z}3:${z}${n+2}`));});
}
function linkLabels(chart,helperName,p,{initial=false}={}) {
  const defs=[['F',p.pointCount,p.options.referenceRings?'Right':'Center'],['M',p.columns,'Bottom'],['R',p.rows,'Left'],['W',p.pointCount,'Center']];
  if(p.options.referenceRings)defs.push(['',0,'Center'],['AF',p.columns,'Top']);
  if(p.bands)defs.push(['',0,'Center'],['',0,'Center'],['',0,'Center'],['AF',p.columns,'Top']);
  defs.forEach(([column,count,position],i)=>{
    const series=chart.series.getItemAt(i);
    for(let j=0;j<count;j++){
      const point=series.points.getItemAt(j);
      const previousCount=i>=4?p.previousColumns:[p.previousPointCount,p.previousColumns,p.previousRows,p.previousPointCount][i];
      if(initial||j>=previousCount)point.hasDataLabel=true;
      point.dataLabel.formula=`=${quote(helperName)}!$${column}$${j+3}`;
      // Rebind touches links only; it must not reset user-edited fonts/positions.
      if(initial){point.dataLabel.position=position;point.dataLabel.format.font.name=FONT;point.dataLabel.format.font.size=plainMatrix(p)?10:(i===0||p.bands?8:9);point.dataLabel.format.font.color=plainMatrix(p)?'#000000':(p.bands&&i===0?'#FFFFFF':COLORS.text);}
    }
  });
}
async function createChart({context,sourceSheet,helper,source,instance,binding}) {
  const chart=sourceSheet.charts.add(Excel.ChartType.bubble,helper.getRange(`C3:E${binding.last}`),Excel.ChartSeriesBy.columns);
  chart.name=instance.chartName;chart.plotVisibleOnly=false;
  chart.series.load('items');await context.sync();
  // Clear only auto-generated series on a newly-created chart, never rebind.
  for(const s of chart.series.items)s.delete();
  for(const name of ['矩阵数值','横轴类别','纵轴类别','网格名称'])chart.series.add(name);
  if(plainMatrix(binding))chart.series.add('横向网格');
  if(binding.options.referenceRings){chart.series.add('100%参考环');chart.series.add('顶部列名');}
  if(binding.bands){for(const name of ['低于25','25至不足30','30及以上','顶部列名'])chart.series.add(name);}
  await context.sync();bindSeries(chart,helper,binding);
  const anchor=sourceSheet.getRangeByIndexes(source.rowIndex,source.columnIndex+source.columnCount,1,1);anchor.load('left,top');await context.sync();
  chart.left=anchor.left+12;chart.top=anchor.top;
  // ChartTitle.width is read-only. Give the PDF281 14pt main title room to
  // remain on one line; keep the data plot's absolute size/spacing unchanged.
  chart.width=binding.bands?356.77:(binding.options.referenceRings?480:(plainMatrix(binding)?600:528));chart.height=binding.bands?396.12:(binding.options.referenceRings?286:(binding.kind==='tiles'?357:294));
  return chart;
}
function chartGeometry(p) {
  const dataPlot=p.bands?{left:80,top:64.6,width:261.4,height:299}:(p.options.referenceRings?{left:100.875,top:72.375,width:366,height:181.5}:(p.kind==='tiles'?{left:32,top:57,width:481,height:265}:{left:70.5,top:63.5,width:444,height:193}));
  const padding=p.kind==='matrix'&&!p.bands&&!p.options.referenceRings?4:0;
  // Expand clipping bounds, not point spacing or the exported image. With
  // H'=H+2p and Y=[-p/H,1+p/H], every point keeps its original chart position.
  return {dataPlot,plot:{...dataPlot,top:dataPlot.top-padding,height:dataPlot.height+2*padding},yMin:-padding/dataPlot.height,yMax:1+padding/dataPlot.height,padding};
}
async function positionMatrixRowLabels({context,chart,binding,dataPlot}) {
  if(!plainMatrix(binding))return null;
  const labels=Array.from({length:binding.rows},(_,i)=>chart.series.getItemAt(2).points.getItemAt(i).dataLabel);
  for(const label of labels)label.load('text,width,height');
  await context.sync();
  const rightEdge=dataPlot.left-4;
  let positioned=0;
  labels.forEach((label,i)=>{
    if(typeof label.text!=='string'||!label.text.trim()||!Number.isFinite(label.width)||!Number.isFinite(label.height))return;
    // Width/height are read-only in ExcelApi 1.9. Explicit formula line breaks
    // produce the natural box; only its position is changed, never its font size.
    label.left=rightEdge-label.width;
    label.top=dataPlot.top+(i+0.5)*dataPlot.height/binding.rows-label.height/2;
    positioned++;
  });
  await context.sync();
  return {rightEdge,positioned,requestedGapPt:4,visualVerified:false};
}
async function positionMatrixAnnotations({context,chart,binding}) {
  if(!plainMatrix(binding))return null;
  const titleVisible=Boolean(binding.options.title||binding.options.subtitle),hasFootnote=Boolean(binding.options.footnote);
  if(titleVisible){chart.title.left=6;chart.title.top=2;}
  if(hasFootnote){
    chart.load('height');await context.sync();
    if(!Number.isFinite(chart.height)||chart.height<=17)throw new UserError('矩阵气泡图[脚注定位]：宿主未返回有效图表高度。');
    const series=chart.series.getItemAt(3),label=series.points.getItemAt(0).dataLabel;
    label.left=6;label.top=chart.height-17;series.showLeaderLines=false;
  }
  await context.sync();
  return {title:titleVisible?{left:6,top:2}:null,footnote:hasFootnote?{series:3,point:0,left:6,top:chart.height-17}:null,visualVerified:false};
}
async function styleMatrixGrid(context,chart){
  const grid=chart.series.getItemAt(4);
  grid.hasDataLabels=false;grid.format.fill.clear();grid.format.line.lineStyle='None';
  grid.xErrorBars.visible=true;await context.sync();
  const errors=chart.series.getItemAt(4).xErrorBars;
  errors.type='StError';errors.include='Both';errors.endStyleCap=false;
  chart.series.getItemAt(4).yErrorBars.visible=false;await context.sync();
  const line=chart.series.getItemAt(4).xErrorBars.format.line;
  line.lineStyle='Continuous';line.color='#F2F2F2';line.weight=.5;
  await context.sync();
}
async function style({context,chart,instance,binding}) {
  const p=binding,o=p.options,geometry=chartGeometry(p);
  chart.format.font.name=FONT;chart.format.font.size=9;chart.format.font.color=COLORS.text;
  chart.format.fill.setSolidColor('#FFFFFF');chart.format.border.lineStyle='Continuous';chart.format.border.color=COLORS.frame;chart.format.border.weight=0.5;
  const titleText=plainMatrix(p)?[o.title,o.subtitle].filter(Boolean).join('\n'):o.title;
  chart.title.visible=titleText!=='';chart.title.text=titleText;chart.title.format.font.name=FONT;chart.title.format.font.size=14;chart.title.format.font.bold=false;chart.title.horizontalAlignment='Left';
  if(plainMatrix(p))chart.title.format.font.color='#000000';
  setChartTitleTypography(chart,{text:titleText});
  chart.legend.visible=false;
  for(let i=0;i<(p.bands?8:(o.referenceRings?6:(plainMatrix(p)?5:4)));i++){
    const s=chart.series.getItemAt(i);s.chartType=Excel.ChartType.bubble;s.bubbleScale=o.bubbleScale;s.hasDataLabels=!(plainMatrix(p)&&i===4);s.dataLabels.showValue=false;s.dataLabels.showBubbleSize=false;s.dataLabels.showCategoryName=false;
    s.format.line.lineStyle='None';
    if(i===0&&!p.bands)s.format.fill.setSolidColor(COLORS.bubble);else s.format.fill.clear();
    if(i===4&&o.referenceRings){s.hasDataLabels=false;s.format.line.lineStyle='Continuous';s.format.line.color='#000000';s.format.line.weight=0.5;}
    if(p.bands&&i>=4&&i<=6){s.hasDataLabels=false;s.format.fill.setSolidColor(['#00A9F4','#2251FF','#061F79'][i-4]);}
  }
  for(const [axis,n] of [[chart.axes.categoryAxis,p.columns],[chart.axes.valueAxis,p.rows]]){
    axis.minimum=0;axis.maximum=1;axis.majorTickMark='None';axis.minorTickMark='None';axis.tickLabelPosition='None';axis.format.line.lineStyle='None';axis.minorGridlines.visible=false;axis.majorGridlines.visible=false;
    axis.majorUnit=1/n;axis.minorUnit=1/(2*n);
    if(!o.referenceRings){const grid=p.kind==='tiles'?axis.majorGridlines:axis.minorGridlines;grid.visible=true;grid.format.line.color=p.kind==='tiles'?'#CCCCCC':'#F2F2F2';grid.format.line.weight=0.5;}
  }
  chart.axes.valueAxis.minimum=geometry.yMin;chart.axes.valueAxis.maximum=geometry.yMax;
  if(plainMatrix(p)){
    chart.axes.categoryAxis.majorGridlines.visible=true;
    chart.axes.categoryAxis.majorGridlines.format.line.lineStyle='None';
    chart.axes.valueAxis.majorGridlines.visible=false;chart.axes.valueAxis.minorGridlines.visible=false;
  }
  // Native gridlines only: tile boundaries use major ticks; matrix centers use
  // half-interval minor ticks. Rebind deliberately preserves user axis settings.
  await context.sync();
  if(plainMatrix(p))await styleMatrixGrid(context,chart);
  linkLabels(chart,instance.helperName,p,{initial:true});await context.sync();
  const layout=geometry.plot;
  for(let k=0;k<2;k++){chart.plotArea.insideLeft=layout.left;chart.plotArea.insideTop=layout.top;chart.plotArea.insideWidth=layout.width;chart.plotArea.insideHeight=layout.height;await context.sync();}
  if(plainMatrix(p)){chart.getImage();await context.sync();}
  const annotationLayout=await positionMatrixAnnotations({context,chart,binding:p});
  const rowLabelLayout=await positionMatrixRowLabels({context,chart,binding:p,dataPlot:geometry.dataPlot});
  return {nativeChart:true,layoutKind:p.kind,pointCount:p.pointCount,notes:['基础矩阵标签先设微软雅黑10pt黑字；长分类按词组换行，原生标签右沿置于plot左侧4pt，需实图核验。','基础矩阵使用原生次网格竖线与透明Bubble的标准误差线横线，保持原始气泡坐标及面积值。','仅原生Bubble系列及公式辅助表；面积模式、微型透明标签点和参考环须双端实机核验。','重绑保留坐标轴手改格式；增减矩阵行后网格间距保留原设置，需用户按新行数调整。'],rowLabelLayout,annotationLayout,requestedPlot:layout,requestedValueAxis:[geometry.yMin,geometry.yMax]};
}
async function rebind({context,chart,helper,instance,binding}) {
  let addedGrid=false;
  if(plainMatrix(binding)){chart.series.load('items');await context.sync();if(chart.series.items.length===4){chart.series.add('横向网格');await context.sync();addedGrid=true;}}
  bindSeries(chart,helper,binding);await context.sync();if(addedGrid)await styleMatrixGrid(context,chart);linkLabels(chart,instance.helperName,binding);await context.sync();
  return {pointCount:binding.pointCount,formatPolicy:'references-and-label-links-only'};
}
// Run after the common final 14/10pt presentation, only on create/restyle.
// Rebind deliberately leaves manual title, row label, and footnote positions.
async function finalizePresentation({context,chart,binding}) {
  if(!plainMatrix(binding))return null;
  const geometry=chartGeometry(binding),layout=geometry.plot;
  chart.getImage();await context.sync();
  for(let k=0;k<2;k++){chart.plotArea.insideLeft=layout.left;chart.plotArea.insideTop=layout.top;chart.plotArea.insideWidth=layout.width;chart.plotArea.insideHeight=layout.height;await context.sync();}
  const annotationLayout=await positionMatrixAnnotations({context,chart,binding});
  const rowLabelLayout=await positionMatrixRowLabels({context,chart,binding,dataPlot:geometry.dataPlot});
  chart.title.load('visible,left,top,width,height');chart.load('width,height');await context.sync();
  return {annotationLayout,rowLabelLayout,requestedPlot:layout,
    chartSize:{width:chart.width,height:chart.height},
    titleRect:chart.title.visible?{left:chart.title.left,top:chart.title.top,width:chart.title.width,height:chart.title.height}:null,
    visualVerified:false};
}
const inputSpec={
  category:{id:'comparison',title:'比较与矩阵'},
  description:'行列矩阵：首行列名、首列行名，其余非负数；两列“名称、数值”可用城市/项目网格。改数与标签由Excel公式更新。',
  minVersion:'1.9',columns:2,minRows:2,maxRows:MAX_ROWS,
  options:[
    {id:'title',label:'标题',type:'text',default:'矩阵气泡图',maxLength:160},
    {id:'subtitle',label:'基础矩阵副标题',type:'text',default:'',maxLength:200},
    {id:'footnote',label:'基础矩阵脚注',type:'text',default:'',maxLength:200},
    {id:'variant',label:'布局',type:'select',default:'auto',choices:[{value:'auto',label:'按选区自动识别'},{value:'matrix',label:'横纵类别矩阵'},{value:'tiles',label:'城市/项目气泡网格'}]},
    {id:'tileColumns',label:'网格每行气泡数',type:'number',default:4,min:2,max:8,step:1},
    {id:'transpose',label:'矩阵行列转置',type:'checkbox',default:false},
    {id:'showValues',label:'显示数值',type:'checkbox',default:true},
    {id:'decimals',label:'小数位',type:'number',default:0,min:0,max:3,step:1},
    {id:'bubbleScale',label:'原生气泡缩放（%）',type:'number',default:DEFAULT_BUBBLE_SCALE,min:1,max:300,step:1},
    {id:'referenceRings',label:'显示100%参考环（数值0–100）',type:'checkbox',default:false},
    {id:'colorMode',label:'面积气泡颜色',type:'select',default:'single',choices:[{value:'single',label:'统一蓝色'},{value:'bands25_30',label:'数值分档（25 / 30）'}]}
  ],
  storage:{marker:MARK,helperPrefix:HELPER_PREFIX,chartPrefix:CHART_PREFIX}
};
export {MARK,HELPER_PREFIX,CHART_PREFIX,FONT,COLORS,DEFAULT_BUBBLE_SCALE,validate,normalizeOptions,formulaPlan,chartGeometry};
export default {id:'matrixbubble',title:'矩阵气泡图',inputSpec,validate,writeHelper,createChart,style,finalizePresentation,rebind};
