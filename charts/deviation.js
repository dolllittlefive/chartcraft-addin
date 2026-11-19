// F092 / PDF 245: one editable native Chart, three formula-driven panels.
// The source is a raster image: native source object structure and font are unknown.
import {UserError} from '../core/errors.js';
export const MARK = 'ChartCraft.Deviation.v1';
export const HELPER_PREFIX = 'CC_偏差_';
export const CHART_PREFIX = 'CC_Deviation_';
export const FONT = '微软雅黑';
export const COLORS = {before:'#C8C8C8', after:'#40A1ED', positive:'#71B37C', negative:'#E14D57', text:'#595959', baseline:'#595959', frame:'#D9D9D9'};
export const LAYOUT = {width:720,height:360,plot:{left:47.25,top:45,width:611.25,height:240},slots:12};
const MAX_ROWS = 200, SLOTS = LAYOUT.slots, SERIES_COUNT = 18;
const quote = name => `'${String(name).replace(/'/g,"''")}'`;
export function columnLetter(index) { let n=index+1,s=''; while(n>0){const m=(n-1)%26;s=String.fromCharCode(65+m)+s;n=Math.floor((n-1)/26);}return s; }
export function validate(values) {
  if(!Array.isArray(values)||values.length<2||!Array.isArray(values[0])||values[0].length!==3) throw new UserError('请选择含表头的三列：分类、基期（或目标）、本期（或实际），至少 1 行数据。');
  if(values.length-1>MAX_ROWS) throw new UserError(`最多支持 ${MAX_ROWS} 行。`);
  values.slice(1).forEach((row,i)=>{
    if(!Array.isArray(row)||row.length!==3||row[0]===null||row[0]===undefined||String(row[0]).trim()==='') throw new UserError(`第 ${i+2} 行分类为空或列数错误。`);
    for(const c of [1,2]) if(typeof row[c]!=='number'||!Number.isFinite(row[c])||row[c]<0) throw new UserError(`第 ${i+2} 行第 ${c+1} 列须为非负数值；差异额可以为负。`);
  });
  return {count:values.length-1,categoryHeader:String(values[0][0]??'分类'),beforeHeader:String(values[0][1]??'基期'),afterHeader:String(values[0][2]??'本期')};
}
export function normalizeOptions(options={}) {
  const o={title:'偏差图',view:'three-in-one',valueDecimals:0,rateDecimals:1,showLabels:true,...options};
  if(typeof o.title!=='string'||o.title.length>80) throw new UserError('标题须为不超过80字的文字。');
  if(!['three-in-one','absolute','relative','comparison'].includes(o.view)) throw new UserError('未知偏差图视图。');
  for(const k of ['valueDecimals','rateDecimals']) if(!Number.isInteger(o[k])||o[k]<0||o[k]>4) throw new UserError('标签小数位须为0至4的整数。');
  if(typeof o.showLabels!=='boolean') throw new UserError('显示标签须为布尔值。');
  return o;
}
function viewConfig(view) {
  if(view==='absolute')return {values:false,delta:true,rate:false,valueHeight:0,deltaZero:500,rateZero:0,deltaSpan:400,rateSpan:0};
  if(view==='relative')return {values:false,delta:false,rate:true,valueHeight:0,deltaZero:0,rateZero:500,deltaSpan:0,rateSpan:400};
  if(view==='comparison')return {values:true,delta:false,rate:false,valueHeight:900,deltaZero:0,rateZero:0,deltaSpan:0,rateSpan:0};
  return {values:true,delta:true,rate:true,valueHeight:400,deltaZero:543.75,rateZero:818.75,deltaSpan:100,rateSpan:140};
}
/** Returns worksheet formula blocks; no chart series receives a JS value array.
 * A:X raw data/labels; AA:AG native stacked columns; AJ:AL stems;
 * AM:AY dots/linked labels; BA:BF baselines/panel labels.
 * Only recipe cells are cleared: core's A1/I1:M1/R1 remain untouched. */
export function formulaBlocks(source, rawOptions={}) {
  const options=normalizeOptions(rawOptions),cfg=viewConfig(options.view),n=source.count,last=n+2,microLast=n*SLOTS+2,stemLast=n*3+2;
  if(!Number.isInteger(n)||n<1||n>MAX_ROWS)throw new UserError('数据行数超出范围。');
  const sheet=quote(source.sheet),cat=columnLetter(source.columnIndex),b=columnLetter(source.columnIndex+1),c=columnLetter(source.columnIndex+2);
  const vf='0'+(options.valueDecimals?'.'+'0'.repeat(options.valueDecimals):''),rf='0'+(options.rateDecimals?'.'+'0'.repeat(options.rateDecimals):'')+'%';
  const label=(expr,format)=>`TEXT(${expr},"${format}")`;
  const rows=[],columns=[],stems=[],dots=[];
  for(let i=0;i<n;i++){
    const r=i+3,sr=source.rowIndex+i+2,br=`${sheet}!$${b}$${sr}`,cr=`${sheet}!$${c}$${sr}`;
    const x=i*SLOTS+6,xb=i*SLOTS+5,xc=i*SLOTS+7;
    rows.push([
      `=IFERROR(${sheet}!$${cat}$${sr}&"","无效分类")`,
      `=IF(F${r},${br},"")`,`=IF(F${r},${cr},"")`,`=IF(F${r},C${r}-B${r},"")`,
      `=IF(G${r},D${r}/B${r},"")`,
      `=IFERROR(AND(ISNUMBER(${br}),ISNUMBER(${cr}),${br}>=0,${cr}>=0),FALSE)`,
      `=IF(F${r},B${r}>0,FALSE)`,
      `=IF(F${r},${label('B'+r,vf)},"无效")`,`=IF(F${r},${label('C'+r,vf)},"无效")`,
      `=IF(F${r},${label('D'+r,vf)},"无效")`,
      `=IF(F${r},IF(G${r},IF(E${r}>0,"+","")&${label('E'+r,rf)},"—"),"无效")`,
      `=${x}`,`=IF(F${r},B${r}*$F$1,0)`,`=IF(F${r},C${r}*$F$1,0)`,
      `=IF(F${r},$N$1+D${r}*$G$1,$N$1)`,`=IF(G${r},$O$1+E${r}*$H$1,$O$1)`,
      `=O${r}+IF(IF(F${r},D${r},0)>=0,18,-18)`,`=P${r}+IF(IF(G${r},E${r},0)>=0,26,-26)`,
      `=${xb}`,`=${xc}`,'=-16',`=IF(G${r},ABS(E${r}),0)`,`=IF(F${r},ABS(D${r}),0)`,`=IF(F${r},MAX(B${r},C${r}),0)`
    ]);
    for(let slot=0;slot<SLOTS;slot++){
      const t=i*SLOTS+slot+3,oldOnly=slot>=2&&slot<4,overlap=slot>=4&&slot<=6,newBar=slot>=4&&slot<=8,deltaBar=slot>=2&&slot<=8;
      columns.push(['=""',`=${cfg.values&&oldOnly?'M'+r:'0'}`,`=${cfg.values&&newBar?'N'+r:'0'}`,`=${cfg.values&&overlap?`MAX(0,M${r}-N${r})`:'0'}`,
        `=${cfg.delta&&deltaBar?`IF(F${r},MAX(0,$N$1+MIN(0,D${r}*$G$1)-SUM(AB${t}:AD${t})),0)`:'0'}`,
        `=${cfg.delta&&deltaBar?`IF(F${r},MAX(0,D${r}*$G$1),0)`:'0'}`,`=${cfg.delta&&deltaBar?`IF(F${r},MAX(0,-D${r}*$G$1),0)`:'0'}`]);
    }
    for(let k=0;k<3;k++)stems.push([`=L${r}`,`=${cfg.rate&&k<2?`IF(AND(G${r},E${r}>=0),${k===0?'$O$1':'P'+r},NA())`:'NA()'}`,`=${cfg.rate&&k<2?`IF(AND(G${r},E${r}<0),${k===0?'$O$1':'P'+r},NA())`:'NA()'}`]);
    dots.push([`=L${r}`,`=${cfg.rate?`IF(AND(G${r},E${r}>=0),P${r},NA())`:'NA()'}`,`=${cfg.rate?`IF(AND(G${r},E${r}<0),P${r},NA())`:'NA()'}`,
      `=S${r}`,`=${cfg.values?'M'+r:'NA()'}`,`=T${r}`,`=${cfg.values?'N'+r:'NA()'}`,`=L${r}`,`=${cfg.delta?'Q'+r:'NA()'}`,`=L${r}`,`=${cfg.rate?'R'+r:'NA()'}`,`=L${r}`,'=0']);
  }
  const scale=`${cfg.valueHeight}/MAX(1,C1)`;
  const deltaScale=cfg.delta?(cfg.values?`MIN(${scale},${cfg.deltaSpan}/MAX(1,D1))`:`${cfg.deltaSpan}/MAX(1,D1)`):'0';
  const rateScale=cfg.rate?(options.view==='three-in-one'?`MIN(625/3,${cfg.rateSpan}/MAX(0.000001,E1))`:`${cfg.rateSpan}/MAX(0.000001,E1)`):'0';
  const blocks=[
    {range:'B1:H1',formulas:[[String(n),`=MAX(X3:X${last})`,`=MAX(W3:W${last})`,`=MAX(V3:V${last})`,`=${cfg.values?scale:'0'}`,`=${deltaScale}`,`=${rateScale}`]]},
    {range:'N1:Q1',formulas:[[String(cfg.deltaZero),String(cfg.rateZero),String(options.valueDecimals),String(options.rateDecimals)]]},
    {range:`A3:X${last}`,formulas:rows},{range:`AA3:AG${microLast}`,formulas:columns},{range:`AJ3:AL${stemLast}`,formulas:stems},{range:`AM3:AY${last}`,formulas:dots},
    {range:'BA3:BD4',formulas:[['0.5',cfg.values?'0':'=NA()',cfg.delta?'=$N$1':'=NA()',cfg.rate?'=$O$1':'=NA()'],[String(n*SLOTS+0.5),cfg.values?'0':'=NA()',cfg.delta?'=$N$1':'=NA()',cfg.rate?'=$O$1':'=NA()']]},
    {range:'BE3:BF3',formulas:[['="差异额(Δ)"','="差异率(Δ%)"']]}
  ];
  return {blocks,last,microLast,stemLast,count:n,options,cfg};
}
async function writeHelper({context,helper,source,options,previous}) {
  const binding=formulaBlocks(source,options);
  if(previous)helper.getRange(`A3:BF${Math.max(previous.count,source.count)*SLOTS+2}`).clear(Excel.ClearApplyTo.contents);
  for(const b of binding.blocks)helper.getRange(b.range).formulas=b.formulas;
  helper.getRange(`E3:E${binding.last}`).numberFormat=Array.from({length:source.count},()=>['0.000000%']);
  await context.sync();return binding;
}
export function seriesBindings(binding) {
  const n=binding.last,m=binding.microLast,s=binding.stemLast;
  return [['AA','AB',m],['AA','AC',m],['AA','AD',m],['AA','AE',m],['AA','AF',m],['AA','AG',m],
    ['AJ','AK',s],['AJ','AL',s],['AM','AN',n],['AM','AO',n],['AP','AQ',n],['AR','AS',n],['AT','AU',n],['AV','AW',n],['AX','AY',n],['BA','BB',4],['BA','BC',4],['BA','BD',4]];
}
function bindSeries(chart,helper,binding) {
  seriesBindings(binding).forEach(([x,y,last],i)=>{const s=chart.series.getItemAt(i);s.setXAxisValues(helper.getRange(`${x}3:${x}${last}`));s.setValues(helper.getRange(`${y}3:${y}${last}`));});
}
function labelSpecs(binding){const a=[];if(binding.cfg.values)a.push([10,'H','Top'],[11,'I','Top']);if(binding.cfg.delta)a.push([12,'J','Center']);if(binding.cfg.rate)a.push([13,'K','Center']);a.push([14,'A','Bottom']);return a;}
function linkLabels(chart,helperName,binding,{format=false,previousCount=0}={}) {
  const show=binding.options.showLabels;
  for(const [series,col,position] of labelSpecs(binding)){
    if(series!==14&&!show)continue;
    const s=chart.series.getItemAt(series);
    for(let i=0;i<binding.count;i++){
      const p=s.points.getItemAt(i);p.hasDataLabel=true;p.dataLabel.formula=`=${quote(helperName)}!$${col}$${i+3}`;
      // Existing labels retain their user-edited position, font and number format on rebind.
      if(format||i>=previousCount){p.dataLabel.position=position;p.dataLabel.format.font.name=FONT;p.dataLabel.format.font.size=10;p.dataLabel.format.font.color='#000000';if(series===14)p.dataLabel.textOrientation=45;}
    }
  }
  if(format){for(const [series,cell,enabled] of [[16,'BE3',binding.cfg.delta],[17,'BF3',binding.cfg.rate]])if(enabled){const p=chart.series.getItemAt(series).points.getItemAt(1);p.hasDataLabel=true;p.dataLabel.formula=`=${quote(helperName)}!$${cell}`;p.dataLabel.position='Right';p.dataLabel.format.font.name=FONT;p.dataLabel.format.font.size=10;p.dataLabel.format.font.color='#000000';}}
}
async function createChart({context,sourceSheet,helper,source,instance,binding}) {
  const chart=sourceSheet.charts.add(Excel.ChartType.columnStacked,helper.getRange(`AB3:AB${binding.microLast}`),Excel.ChartSeriesBy.columns);
  chart.name=instance.chartName;
  const names=[source.beforeHeader||'基期',source.afterHeader||'本期','基期重叠区','差异悬浮基座','正差异','负差异','正率引线','负率引线','正差异率','负差异率','基期标签','本期标签','差异额标签','差异率标签','分类','数值零线','差异额零线','差异率零线'];
  chart.series.getItemAt(0).name=names[0];for(let i=1;i<SERIES_COUNT;i++)chart.series.add(names[i]);
  const anchor=sourceSheet.getRangeByIndexes(source.rowIndex,source.columnIndex+3,1,1);anchor.load('left,top');await context.sync();
  chart.left=anchor.left+12;chart.top=anchor.top;chart.width=Math.max(LAYOUT.width,source.count*24+168);chart.height=LAYOUT.height;
  for(let i=6;i<SERIES_COUNT;i++){const s=chart.series.getItemAt(i);s.chartType=[6,7,15,16,17].includes(i)?Excel.ChartType.xyscatterLinesNoMarkers:Excel.ChartType.xyscatter;s.axisGroup=Excel.ChartAxisGroup.secondary;}
  bindSeries(chart,helper,binding);await context.sync();return chart;
}
async function syncStage(context,stage){
  try{await context.sync();}
  catch(error){error.message=`偏差图[${stage}]：${error.message}`;throw error;}
}
async function hideSecondaryAxis(context,axis,minimum,maximum,stage){
  axis.minimum=minimum;axis.maximum=maximum;
  await syncStage(context,stage+'范围');
  // Hidden secondary axes may reject line/gridline formatting on Mac.
  // Hide the native axis itself, preserving its scatter coordinate mapping.
  axis.visible=false;
  await syncStage(context,stage+'隐藏');
  axis.load('minimum,maximum,visible');
  await syncStage(context,stage+'回读');
  if(axis.visible!==false||!Number.isFinite(axis.minimum)||!Number.isFinite(axis.maximum)||Math.abs(axis.minimum-minimum)>1e-9||Math.abs(axis.maximum-maximum)>1e-9)
    throw new UserError(`偏差图[${stage}回读]：隐藏次轴后未保留所需范围 ${minimum}–${maximum}，原生图未验收。`);
  return {axis:stage,minimum:axis.minimum,maximum:axis.maximum,visible:axis.visible};
}
// Only native label boxes move. Forbidden top intervals come from labels whose
// X extents overlap; choosing the nearest free edge avoids a one-way staircase.
export function planLabelStagger(boxes,plot,cfg,bars=[]) {
  const planned=boxes.map(b=>({...b,originTop:b.top})),gap=1,epsilon=1e-7;
  const intersectsX=(a,b,gutter=gap)=>a.left<b.left+b.width+gutter-epsilon&&a.left+a.width+gutter>b.left+epsilon;
  const yAt=value=>plot.top+plot.height*(1-value/1000);
  const fixed=planned.filter(b=>![10,11,13].includes(b.series));
  function available(box,obstacles,hints=[]) {
    const peers=obstacles.filter(other=>intersectsX(box,other));
    const intervals=peers.map(other=>[other.top-box.height-gap,other.top+other.height+gap]).sort((a,b)=>a[0]-b[0]);
    const forbidden=[];
    for(const interval of intervals){
      const last=forbidden[forbidden.length-1];
      if(last&&interval[0]<last[1]-epsilon)last[1]=Math.max(last[1],interval[1]);
      else forbidden.push([...interval]);
    }
    // Future neighbours' preferred positions/bounds supply useful edge choices;
    // they are not treated as occupied until that neighbour is actually placed.
    const edges=hints.filter(other=>intersectsX(box,other)).flatMap(other=>[other.originTop,other.minimum,other.maximum].flatMap(top=>[top-box.height-gap,top+other.height+gap]));
    const candidates=[Math.max(box.minimum,Math.min(box.maximum,box.originTop)),box.minimum,box.maximum,...forbidden.flat(),...edges];
    const positions=[...new Set(candidates)].filter(top=>Number.isFinite(top)&&top>=box.minimum-epsilon&&top<=box.maximum+epsilon&&!forbidden.some(([a,b])=>top>a+epsilon&&top<b-epsilon));
    positions.sort((a,b)=>Math.abs(a-box.originTop)-Math.abs(b-box.originTop)||a-b);
    return {positions,forbidden,conflicts:peers.length};
  }
  function solve(group,obstacles,order) {
    const components=[];
    for(const box of [...group].sort((a,b)=>a.left-b.left)){
      const last=components[components.length-1];
      if(last&&box.left<last.right+gap-epsilon){last.items.push(box);last.right=Math.max(last.right,box.left+box.width);}
      else components.push({right:box.left+box.width,items:[box]});
    }
    for(const component of components){
      const queue=component.items.sort(order),placed=[];
      let attempts=0,blocked,exhausted=false;
      function search(at){
        if(at===queue.length)return true;
        if(++attempts>20000){exhausted=true;return false;}
        const box=queue[at],future=queue.slice(at+1),free=available(box,[...obstacles,...placed],future);
        blocked={box,...free};
        for(const top of free.positions){
          box.top=top;placed.push(box);
          // A nearest choice is only accepted if every remaining X-neighbour
          // still has a free position; backtrack if a later combination fails.
          const blockedFuture=future.find(other=>intersectsX(box,other)&&!available(other,[...obstacles,...placed],future.filter(item=>item!==other)).positions.length);
          if(!blockedFuture&&search(at+1))return true;
          if(blockedFuture)blocked={box:blockedFuture,...available(blockedFuture,[...obstacles,...placed],future.filter(item=>item!==blockedFuture))};
          placed.pop();
          if(exhausted)return false;
        }
        return false;
      }
      if(!search(0)){
        const {box,forbidden,conflicts}=blocked;
        const family=box.series===13?'差异率':box.series===10?'基期值':'本期值';
        const error=new UserError(`偏差图[标签留白]：${family}第${box.index+1}项${exhausted?'在限定布局搜索内未找到解':'没有可用位置'}；原top=${box.originTop.toFixed(2)}，允许top=${box.minimum.toFixed(2)}–${box.maximum.toFixed(2)}，横向冲突=${conflicts}，禁占Y=${forbidden.slice(0,6).map(pair=>pair.map(v=>v.toFixed(2)).join('–')).join('/') || '无'}。`);
        error.details={series:box.series,index:box.index,originTop:box.originTop,minimum:box.minimum,maximum:box.maximum,forbidden,searchExhausted:exhausted};
        throw error;
      }
    }
  }
  const rates=planned.filter(b=>b.series===13);
  for(const rate of rates){
    const negative=/^[-−]/.test(rate.text),zero=yAt(cfg.rateZero);
    rate.minimum=negative?zero+gap:plot.top;
    rate.maximum=(negative?(cfg.delta?yAt(cfg.deltaZero)-gap:plot.top+plot.height):zero-gap)-rate.height;
  }
  solve(rates,fixed,(a,b)=>a.left-b.left||a.index-b.index);
  const lower=planned.filter(b=>b.series===10||b.series===11);
  for(const value of lower){
    value.minimum=cfg.delta?yAt(cfg.deltaZero)+gap:plot.top;
    // The source places some black values inside their columns. Keep labels
    // near the original column top, allowing the nearest free position below it.
    value.maximum=Math.min(plot.top+plot.height-value.height,value.originTop+40);
  }
  // Work across categories so local conflicts are resolved before distant rows.
  // Within each pair the shorter column goes first; its neighbour takes the
  // nearest free location around the column top, including inside the column.
  solve(lower,[...fixed,...rates],(a,b)=>a.index-b.index||b.originTop-a.originTop||a.left-b.left||a.series-b.series);
  return planned.filter(b=>Math.abs(b.top-b.originTop)>0.01).map(b=>({series:b.series,index:b.index,top:b.top}));
}
async function staggerNativeLabels(context,chart,helper,binding,plot) {
  if(!binding.options.showLabels)return {moved:0,visualVerified:false};
  const targets=[...(binding.cfg.values?[10,11]:[]),...(binding.cfg.rate?[13]:[])],labels=[];
  if(!targets.length)return {moved:0,visualVerified:false};
  const specs=[...targets.map(series=>[series,binding.count,0]),...(binding.cfg.delta?[[12,binding.count,0],[16,1,1]]:[]),...(binding.cfg.rate?[[17,1,1]]:[])];
  for(const [series,count,start] of specs)for(let index=start;index<start+count;index++){
    const label=chart.series.getItemAt(series).points.getItemAt(index).dataLabel;
    label.format.font.name=FONT;label.format.font.size=10;label.format.font.color='#000000';
    labels.push({series,index,label});
  }
  for(const series of targets)chart.series.getItemAt(series).showLeaderLines=false;
  await syncStage(context,'标签10pt排版');
  // Mac can expose pre-render automatic label positions after font/plot sync.
  // Request a native render before acquiring fresh proxies for measurement;
  // this image is neither exported nor used as visual acceptance evidence.
  chart.getImage();
  await syncStage(context,'原生图像布局刷新');
  for(const item of labels){
    item.label=chart.series.getItemAt(item.series).points.getItemAt(item.index).dataLabel;
    item.label.load('left,top,width,height,text');
  }
  chart.plotArea.load('insideLeft,insideTop,insideWidth,insideHeight');
  const columnValues=binding.cfg.values?helper.getRange(`M3:N${binding.last}`):null;
  if(columnValues)columnValues.load('values');
  await syncStage(context,'标签边界读取');
  const boxes=labels.map(({series,index,label})=>({series,index,left:label.left,top:label.top,width:label.width,height:label.height,text:label.text}));
  if(boxes.some(b=>!['left','top','width','height'].every(k=>Number.isFinite(b[k]))||b.width<=0||b.height<=0))
    throw new UserError('偏差图[标签留白]：宿主未返回可用原生标签边界，未认定布局通过。');
  const bars=[];
  if(columnValues)for(let index=0;index<binding.count;index++)for(let period=0;period<2;period++){
    const value=columnValues.values[index]?.[period];
    if(!Number.isFinite(value)||value<0)throw new UserError(`偏差图[标签留白]：第${index+1}组柱顶坐标无效。`);
    bars.push({series:period?11:10,index,top:plot.top+plot.height*(1-value/1000)});
  }
  // YaHei 10pt has a conservative 13.2pt line height; native outer boxes also
  // contain padding. XML confirms 3pt horizontal insets. These are collision
  // bounds only: never resize or crop the actual native label.
  const inkBoxes=boxes.map(box=>{
    const height=Math.min(box.height,13.2),insetTop=(box.height-height)/2;
    if(box.width<=6)throw new UserError('偏差图[标签留白]：原生标签宽度不足以扣除已核实的水平内边距。');
    return {...box,left:box.left+3,top:box.top+insetTop,width:box.width-6,height,insetTop};
  });
  let moves;
  try{moves=planLabelStagger(inkBoxes,plot,binding.cfg,bars).map(move=>({
    ...move,top:move.top-inkBoxes.find(box=>box.series===move.series&&box.index===move.index).insetTop
  }));}
  catch(error){
    const round=value=>Number.isFinite(value)?Number(value.toFixed(3)):value;
    const diagnostic={requestedPlot:plot,actualPlot:Object.fromEntries(['insideLeft','insideTop','insideWidth','insideHeight'].map(key=>[key,chart.plotArea[key]])),
      boxColumns:['series','index','left','top','width','height','text'],boxes:boxes.map(box=>[box.series,box.index,round(box.left),round(box.top),round(box.width),round(box.height),box.text]),
      inkBounds:{lineHeight:13.2,horizontalInset:3},inkBoxes:inkBoxes.map(box=>[box.series,box.index,round(box.left),round(box.top),round(box.width),round(box.height),box.text]),
      barColumns:['series','index','top'],bars:bars.map(bar=>[bar.series,bar.index,round(bar.top)])};
    error.details={...error.details,measurement:diagnostic};
    error.message+=` BBoxDiagnostic=${JSON.stringify(diagnostic)}`;
    throw error;
  }
  for(const move of moves){
    const label=labels.find(b=>b.series===move.series&&b.index===move.index).label;
    label.left=boxes.find(b=>b.series===move.series&&b.index===move.index).left;
    label.top=move.top;
  }
  await syncStage(context,'标签最近空位定位');
  for(const move of moves)labels.find(b=>b.series===move.series&&b.index===move.index).label.load('left,top');
  const leaderStates=targets.map(index=>chart.series.getItemAt(index));
  for(const series of leaderStates)series.load('showLeaderLines');
  await syncStage(context,'标签位置与引导线回读');
  if(leaderStates.some(series=>series.showLeaderLines!==false)||moves.some(move=>{
    const label=labels.find(b=>b.series===move.series&&b.index===move.index).label,original=boxes.find(b=>b.series===move.series&&b.index===move.index);
    return !Number.isFinite(label.top)||!Number.isFinite(label.left)||Math.abs(label.top-move.top)>0.25||Math.abs(label.left-original.left)>0.25;
  }))throw new UserError('偏差图[标签留白]：宿主未保留原生标签位置或关闭引导线，未认定布局通过。');
  return {moved:moves.length,moves,visualVerified:false};
}
function hideAxis(axis){axis.majorTickMark='None';axis.minorTickMark='None';axis.tickLabelPosition='None';axis.format.line.lineStyle='None';axis.majorGridlines.visible=false;axis.minorGridlines.visible=false;axis.title.visible=false;}
async function style({context,chart,helper,binding,instance}) {
  chart.plotVisibleOnly=false;chart.displayBlanksAs='NotPlotted';
  chart.format.font.name=FONT;chart.format.font.size=10;chart.format.font.color='#000000';chart.format.fill.setSolidColor('#FFFFFF');
  chart.format.border.color=COLORS.frame;chart.format.border.weight=0.5;chart.format.border.lineStyle='Continuous';
  chart.title.visible=Boolean(binding.options.title);chart.title.text=binding.options.title;chart.title.format.font.name=FONT;chart.title.format.font.size=14;chart.title.format.font.color='#000000';chart.title.format.font.bold=false;chart.title.horizontalAlignment='Left';
  chart.legend.visible=binding.cfg.values;
  // Mac rejects assigning Custom even though it appears in the Office enum.
  if(binding.cfg.values){chart.legend.overlay=true;chart.legend.position='Top';chart.legend.format.font.name=FONT;chart.legend.format.font.size=10;chart.legend.format.font.color='#000000';}
  await syncStage(context,'图表与图例基本样式');
  const fills=[COLORS.before,COLORS.after,COLORS.before,null,COLORS.positive,COLORS.negative];
  for(let i=0;i<6;i++){const s=chart.series.getItemAt(i);s.gapWidth=0;s.overlap=100;s.format.line.lineStyle='None';if(fills[i])s.format.fill.setSolidColor(fills[i]);else s.format.fill.clear();}
  await syncStage(context,'主柱系列样式');
  for(let i=6;i<SERIES_COUNT;i++){const s=chart.series.getItemAt(i);s.markerStyle='None';s.format.line.lineStyle='None';}
  for(const [i,color] of [[6,COLORS.positive],[7,COLORS.negative]]){const s=chart.series.getItemAt(i);s.format.line.lineStyle='Continuous';s.format.line.color=color;s.format.line.weight=0.6;}
  for(const [i,color] of [[8,COLORS.positive],[9,COLORS.negative]]){const s=chart.series.getItemAt(i);s.markerStyle='Circle';s.markerSize=7;s.markerBackgroundColor=color;s.markerForegroundColor=color;}
  for(const i of [15,16,17]){const s=chart.series.getItemAt(i);s.format.line.lineStyle='Continuous';s.format.line.color=COLORS.baseline;s.format.line.weight=0.5;}
  await syncStage(context,'散点与线系列样式');
  const y=chart.axes.getItem('Value','Primary');hideAxis(y);y.minimum=0;y.maximum=1000;y.majorUnit=200;
  await syncStage(context,'主数值轴');
  hideAxis(chart.axes.getItem('Category','Primary'));
  await syncStage(context,'主分类轴');
  const secondaryAxes=[];
  secondaryAxes.push(await hideSecondaryAxis(context,chart.axes.getItem('Category','Secondary'),0.5,binding.count*SLOTS+0.5,'次散点X轴'));
  secondaryAxes.push(await hideSecondaryAxis(context,chart.axes.getItem('Value','Secondary'),0,1000,'次散点Y轴'));
  linkLabels(chart,instance.helperName,binding,{format:true});
  await syncStage(context,'链接标签');
  // Mac may reject even accessing legendEntries when the legend is hidden.
  if(binding.cfg.values){
    const entries=chart.legend.legendEntries;entries.load('items');await syncStage(context,'年份图例读取');
    if(entries.items.length<2)throw new UserError('年份图例项不足，无法完成图例配置。');
    entries.items.forEach((e,i)=>{e.visible=i<2;});
    await syncStage(context,'年份图例筛选');
    // Position the visible two-entry native legend after the host lays it out.
    chart.legend.width=140;chart.legend.height=16;chart.legend.left=21.5;chart.legend.top=47.5;
  }
  if(binding.options.title){chart.title.left=8;chart.title.top=6;}
  await syncStage(context,'图例与标题位置');
  const plot={...LAYOUT.plot,width:Math.max(LAYOUT.width,binding.count*24+168)-LAYOUT.width+LAYOUT.plot.width};
  for(let k=0;k<2;k++){chart.plotArea.insideLeft=plot.left;chart.plotArea.insideTop=plot.top;chart.plotArea.insideWidth=plot.width;chart.plotArea.insideHeight=plot.height;await syncStage(context,`绘图区定位${k+1}`);}
  const labelLayout=await staggerNativeLabels(context,chart,helper,binding,plot);
  return {notes:['三层原生组合图；采样堆积柱重建重叠柱，非原厂对象结构证明。','字体、精确间距及双端效果待实测；基期为0时差异率显示—。'],requestedPlot:plot,secondaryAxes,seriesCount:SERIES_COUNT,labelLayout};
}
async function rebind({context,chart,helper,binding,instance,previous}) {
  bindSeries(chart,helper,binding);await context.sync();
  // Keep axis limits stable on value edits. Only row-count changes require the x mapping to move.
  if(previous.count!==binding.count)chart.axes.getItem('Category','Secondary').maximum=binding.count*SLOTS+0.5;
  linkLabels(chart,instance.helperName,binding,{previousCount:previous.count});await context.sync();
}
const inputSpec={category:{id:'comparison',title:'比较与差异'},description:'三列含表头：分类、基期/目标、本期/实际。两列数值均须非负；差异额=本期−基期，差异率=差异额÷基期。基期0的率显示—。',minVersion:'1.9',columns:3,minRows:1,maxRows:MAX_ROWS,
  options:[{id:'title',label:'标题',type:'text',default:'偏差图',maxLength:80},{id:'view',label:'视图',type:'select',default:'three-in-one',choices:[{value:'three-in-one',label:'三合一（原例）'},{value:'absolute',label:'差异额'},{value:'relative',label:'差异率'},{value:'comparison',label:'双期对比'}]},{id:'valueDecimals',label:'数值小数位',type:'number',default:0,min:0,max:4,step:1},{id:'rateDecimals',label:'差异率小数位',type:'number',default:1,min:0,max:4,step:1},{id:'showLabels',label:'显示数值标签',type:'checkbox',default:true}],
  storage:{marker:MARK,helperPrefix:HELPER_PREFIX,chartPrefix:CHART_PREFIX}};
export default {id:'deviation',title:'偏差图',inputSpec,validate,writeHelper,createChart,style,rebind};
