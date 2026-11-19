import {UserError} from '../core/errors.js';
import {setChartTitleTypography} from '../core/presentation.js';

// Floating native stacked bars. Formula normalization keeps both native axes in [0,1]
// so horizontal/vertical scatter connectors continue to align after source edits.
export const MARK = 'ChartCraft.Waterfall.v1';
export const COLORS = {up:'#40A1ED',down:'#F79042',total:'#0A1C2C',line:'#D9D9D9',text:'#404040',frame:'#D9D9D9',stack:['#40A1ED','#107DAD','#0A1C2C','#89D1E5','#5572A8','#AAAAAA','#B5D5F5','#798D9B']};
export function colorsFor(palette='default') {
  if(palette==='pdf156')return {...COLORS,up:'#90D3E6',down:'#90D3E6',total:'#1C548B'};
  if(palette==='pdf165')return {...COLORS,stack:['#0A1C2C','#40A1ED','#323FCA','#A2E5FB','#4E95B4','#AFC4F4','#8C5AC8']};
  if(palette==='pdf168')return {...COLORS,stack:['#0A1C2C','#40A1ED',...COLORS.stack.slice(2)]};
  return COLORS;
}
const contrast=color=>{const rgb=color.slice(1).match(/../g).map(x=>parseInt(x,16));return (.2126*rgb[0]+.7152*rgb[1]+.0722*rgb[2])<140?'#FFFFFF':COLORS.text;};
export const PAGE58 = {width:700,height:396.35,plot:{left:20.49,top:84.89,width:640.88,height:279.89}};
const MAX_ROWS=200, MAX_COMPONENTS=8, RAW=100, AUX=110;
export const col = index => { let s='', n=index+1; while(n) { n--; s=String.fromCharCode(65+n%26)+s; n=Math.floor(n/26); } return s; };
const quote = name => `'${String(name).replace(/'/g,"''")}'`;
const literal = value => "'"+String(value);
const range = (helper,c,first,last=first) => helper.getRange(`${col(c)}${first}:${col(c)}${last}`);
const roleWords = new Map([['变化','change'],['增减','change'],['change','change'],['总计','total'],['合计','total'],['total','total'],['小计','subtotal'],['subtotal','subtotal']]);
const role = value => roleWords.get(String(value).trim().toLowerCase());

export function validate(values) {
  if (!Array.isArray(values) || values.length<3 || values.length>MAX_ROWS+1) throw new UserError(`请选择含表头的 2–${MAX_ROWS} 行数据。`);
  const width=values[0].length, hasRole=['类型','type'].includes(String(values[0][width-1]).trim().toLowerCase()), components=width-1-(hasRole?1:0);
  if(components<1 || components>MAX_COMPONENTS) throw new UserError('瀑布图需要分类列、1–8列数值，可选末列“类型”（变化/总计/小计）。');
  let magnitude=0;
  for(let i=1;i<values.length;i++) {
    const row=values[i];
    if(row.length!==width || row[0]===null || String(row[0]).trim()==='') throw new UserError(`第 ${i+1} 行分类为空或列数不一致。`);
    for(let j=1;j<=components;j++) if(typeof row[j]!=='number' || !Number.isFinite(row[j])) throw new UserError(`第 ${i+1} 行第 ${j+1} 列须为有限数字；空白请显式填 0。`);
    magnitude+=row.slice(1,components+1).reduce((sum,value)=>sum+Math.abs(value),0);
    if(!Number.isFinite(magnitude)) throw new UserError('数值合计超出可计算范围。');
    if(hasRole && !role(row[width-1])) throw new UserError(`第 ${i+1} 行类型只能是变化、总计、小计。`);
  }
  if(values[0].slice(1,components+1).some(v=>!String(v??'').trim())) throw new UserError('数值列必须有表头。');
  return {count:values.length-1, components, hasRole, componentHeaders:values[0].slice(1,components+1).map(String), valueHeader:String(values[0][1])};
}
export function optionsFor(options={}) {
  const o={title:'图表标题',variant:'page058',orientation:'vertical',totals:'append',firstTotal:true,totalLabel:'2020-24 年',connectors:true,decimals:1,linkedHeadline:false,coordinateMode:'scaled',palette:'default',...options};
  if(!['page058','standard','stacked'].includes(o.variant) || !['vertical','horizontal'].includes(o.orientation) || !['append','last','marked','none'].includes(o.totals)) throw new UserError('瀑布变体、方向或总计设置无效。');
  if(!['default','pdf156','pdf165','pdf168'].includes(o.palette))throw new UserError('原例配色无效。');
  if(!['scaled','values'].includes(o.coordinateMode)) throw new UserError('数值坐标模式无效。');
  if(o.coordinateMode==='values'&&o.orientation!=='vertical') throw new UserError('原始数值坐标当前支持垂直瀑布；水平原始数值坐标仍待验证。');
  if(!Number.isInteger(o.decimals) || o.decimals<0 || o.decimals>4) throw new UserError('标签小数位须为 0–4 的整数。');
  return o;
}
export function descriptors(components) {
  const out=[{name:'透明底座',kind:'base',component:-1,column:9}];
  for(const sign of ['negative','positive']) {
    const order=Array.from({length:components},(_,i)=>sign==='negative'?components-1-i:i);
    for(const component of order) for(const kind of ['change','total']) out.push({name:`${sign}-${kind}-${component}`,kind,sign,component,column:9+out.length});
  }
  return out;
}
// Raw-value counterpart: separate positive and negative stacks at zero. Each
// visible segment is intersected with each side of zero; mixed-sign components
// retain their common starting level. Values are never divided by a scale.
export function rawDescriptors(components) {
  const out=[{name:'正值透明底座',kind:'base',component:-1,column:160},{name:'负值透明底座',kind:'base',component:-1,column:161}];
  for(const part of ['positive','negative']) {
    const signs=part==='positive'?['negative','positive']:['positive','negative'];
    for(const sign of signs) {
      const reverse=part==='positive'?sign==='negative':sign==='positive';
      const order=Array.from({length:components},(_,i)=>reverse?components-1-i:i);
      for(const component of order)for(const kind of ['change','total'])out.push({name:`原值-${part}-${sign}-${kind}-${component}`,kind,sign,part,component,column:160+out.length});
    }
  }
  return out;
}
function rawGeometry(p) {
  const {info,o,n,rows,cells}=p; const body=rawDescriptors(info.components);
  if(body.length+6+info.components+n-1>255)throw new UserError('当前细分系列与行数超过原始数值组合图的255系列限制，请减少数据行。');
  for(let i=0;i<n;i++) {
    const r=i+3,v=rows[i];while(v.length<160+body.length)v.push('');
    v[160]=`=MAX(0,H${r})`;v[161]=`=MIN(0,I${r})`;
    for(const d of body.slice(2)) {
      const value=`${col(45+d.component)}${r}`;
      const before=Array.from({length:d.component},(_,j)=>`${d.sign==='positive'?'MAX':'MIN'}(0,${col(45+j)}${r})`);
      const start=`(D${r}+${before.length?'SUM('+before.join(',')+')':'0'})`;
      const low=d.sign==='positive'?start:`(${start}+${value})`,high=d.sign==='positive'?`(${start}+${value})`:start;
      const segment=d.part==='positive'?`MAX(0,${high})-MAX(0,${low})`:`MIN(0,${low})-MIN(0,${high})`;
      v[d.column]=`=IF(AND(B${r}${d.kind==='change'?'=0':'<>0'},${value}${d.sign==='positive'?'>=0':'<0'}),${segment},0)`;
    }
    v[AUX]=String(i+1);v[AUX+1]=`=I${r}`;v[AUX+3]=String(i+1);v[AUX+4]=o.variant==='page058'?`=H${r}`:'0';
    for(let j=0;j<info.components;j++)v[75+j]=v[75+j].replace(/-\$B\$1\)\/\$D\$1$/,')');
  }
  cells.E1='0';cells.F1='=E3';cells.H1=`=E${p.last}`;
  return {...p,body};
}
function sourceCell(source,r,c) { return `${quote(source.sheet)}!$${col(source.columnIndex+c)}$${source.rowIndex+r+1}`; }

export function helperPlan(source, options={}) {
  const info=validate(source.values), o=optionsFor(options), n=source.count+(o.totals==='append'?1:0), first=3,last=n+2;
  if(o.totals==='marked'&&!info.hasRole) throw new UserError('按类型列指定总计时，请选择包含末列“类型”的完整数据。');
  const body=descriptors(info.components), rows=[], decimals='0'+(o.decimals?'.'+'0'.repeat(o.decimals):'');
  for(let i=0;i<n;i++) {
    const r=i+first, appended=i===source.count, v=Array(AUX+14).fill('');
    const srcRole=info.hasRole?sourceCell(source,i+1,source.columnCount-1):null;
    v[0]=appended?literal(o.totalLabel):`=${sourceCell(source,i+1,0)}&""`;
    // Role flags: 0 delta; 1 absolute total (resets cumulative); 2 subtotal (shows prior cumulative).
    v[1]=appended?'2':srcRole?`=IF(OR(LOWER(${srcRole})="total",${srcRole}="总计",${srcRole}="合计"),1,IF(OR(LOWER(${srcRole})="subtotal",${srcRole}="小计"),2,0))`:(o.firstTotal&&i===0)||(o.totals==='last'&&i===source.count-1)?'1':'0';
    for(let j=0;j<info.components;j++) v[RAW+j]=appended?'0':`=IF(ISNUMBER(${sourceCell(source,i+1,j+1)}),${sourceCell(source,i+1,j+1)},0)`;
    // A subtotal is split into the current cumulative contribution of each component.
    // Each component baseline is reset at absolute totals, retained at subtotal rows.
    for(let j=0;j<info.components;j++) {
      const raw=`${col(RAW+j)}${r}`, value=`${col(45+j)}${r}`, prev=i?`${col(65+j)}${r-1}`:'0';
      v[45+j]=`=IF(B${r}=2,${prev},${raw})`;
      v[65+j]=`=IF(B${r}=0,${prev}+${value},${value})`;
    }
    const vals=Array.from({length:info.components},(_,j)=>`${col(45+j)}${r}`);
    v[2]=`=SUM(${vals.join(',')})`;
    v[3]=i?`=IF(B${r}=0,E${r-1},0)`:'0';
    v[4]=`=D${r}+C${r}`;
    v[5]=`=SUM(${vals.map(c=>`MIN(0,${c})`).join(',')})`;
    v[6]=`=SUM(${vals.map(c=>`MAX(0,${c})`).join(',')})`;
    v[7]=`=D${r}+F${r}`; v[8]=`=D${r}+G${r}`;
    v[9]=`=(H${r}-$B$1)/$D$1`;
    for(const d of body.slice(1)) {
      const vc=`${col(45+d.component)}${r}`, positive=d.sign==='positive';
      v[d.column]=`=IF(B${r}${d.kind==='change'?'=0':'<>0'},${positive?'MAX':'-MIN'}(0,${vc})/$D$1,0)`;
    }
    const center=(i+.5)/n;
    // Normalized coordinates: body/endpoint labels and categories can follow floating bars.
    const high=`(I${r}-$B$1)/$D$1`, low=`(H${r}-$B$1)/$D$1`;
    v[AUX]=o.orientation==='vertical'?String(center):`=${high}`;
    v[AUX+1]=o.orientation==='vertical'?`=${high}`:String(1-center);
    v[AUX+2]=`=TEXT(C${r},"${decimals}")`;
    v[AUX+3]=o.orientation==='vertical'?String(center):'0';
    v[AUX+4]=o.orientation==='vertical'?`=${o.variant==='page058'?low:'0'}`:String(1-center);
    v[AUX+5]=`=A${r}`;
    // Per-component label positions at the middle of the visible segment, not net delta.
    for(let j=0;j<info.components;j++) {
      const value=`${col(45+j)}${r}`;
      const earlier=Array.from({length:j},(_,k)=>`${col(45+k)}${r}`);
      const positiveBefore=earlier.length?`SUM(${earlier.map(c=>`MAX(0,${c})`).join(',')})`:'0';
      const negativeBefore=earlier.length?`SUM(${earlier.map(c=>`MIN(0,${c})`).join(',')})`:'0';
      v[75+j]=`=(D${r}+IF(${value}>=0,${positiveBefore},${negativeBefore})+${value}/2-$B$1)/$D$1`;
      v[85+j]=`=IF(${value}=0,"",TEXT(${value},"${decimals}"))`;
    }
    rows.push(v);
  }
  const cells={
    B1:`=MIN(0,MIN(H3:H${last}))`,C1:`=MAX(0,MAX(I3:I${last}))`,
    D1:'=MAX((C1-B1)*1.12,1)',E1:'=(0-B1)/D1',F1:'=(E3-B1)/D1',G1:`=E${last}-E3`,H1:`=(E${last}-B1)/D1`,
    N1:String(n),O1:String(info.components),P1:'=IF(G1>=0,"+","")&IF(ROUND(G1,4)=ROUND(G1,0),TEXT(G1,"0"),TEXT(G1,"0.####"))',Q1:'="所谓的“七巨头”在2020至2024年间为全球经济利润增加了 "&IF(ROUND(G1*10,4)=ROUND(G1*10,0),TEXT(G1*10,"0"),TEXT(G1*10,"0.####"))&" 亿美元。"'
  };
  const result={info,o,n,last,body,rows,cells};
  return o.coordinateMode==='values'?rawGeometry(result):result;
}

async function writeHelper({context,helper,source,options,previous}) {
  const p=helperPlan(source,options);
  if(previous) {
    const used=helper.getUsedRange(), oldComponents=helper.getRange('O1'); used.load('rowCount'); oldComponents.load('values'); await context.sync();
    if(Number(oldComponents.values[0][0])!==p.info.components) throw new UserError('请保持数值列数量不变；改变细分系列数量请新建图表，原图未修改。');
    helper.getRange(`A2:IZ${Math.max(used.rowCount,p.last+6)}`).clear(Excel.ClearApplyTo.contents);
  }
  for(const [cell,formula] of Object.entries(p.cells)) helper.getRange(cell).formulas=[[formula]];
  helper.getRange(`A3:${col(p.rows[0].length-1)}${p.last}`).formulas=p.rows;
  helper.getRange('P2').values=[['“七巨头”在2005-09年与2020-24年间的全球经济利润变化，单位：十亿美元（2023年价格）']];
  helper.getRange('EN15:EO15').formulas=[['0','1']];
  // Dedicated two-point ranges avoid NA/error caches and never connect across bars.
  const connectorStart=p.last+3, gap=(p.o.variant==='page058'?35/106:.35), edge=(1-gap)/2;
  for(let i=0;i<p.n-1;i++) {
    const r=connectorStart+i*2, level=p.o.coordinateMode==='values'?`E${i+3}`:`(E${i+3}-$B$1)/$D$1`;
    const a=p.o.coordinateMode==='values'?i+1+edge:(i+.5+edge)/p.n, b=p.o.coordinateMode==='values'?i+2-edge:(i+1.5-edge)/p.n;
    const coords=p.o.orientation==='vertical'?[[String(a),'='+level],[String(b),'='+level]]:[['='+level,String(1-a)],['='+level,String(1-b)]];
    helper.getRange(`EK${r}:EL${r+1}`).formulas=coords;
  }
  const extras={zero:[['0','=E1'],['1','=E1']],reference:[['0','=F1'],['1','=F1']],growth:[['0.989','=F1'],['0.989','=H1']],badge:[['0.989','=(F1+H1)/2']]};
  if(p.o.coordinateMode==='values') {extras.zero[0][0]='0.5';extras.zero[1][0]=String(p.n+.5);extras.reference[0][0]='0.5';extras.reference[1][0]=String(p.n+.5);extras.growth[0][0]=extras.growth[1][0]=extras.badge[0][0]=String(p.n+.4);}
  if(p.o.orientation==='horizontal') for(const data of Object.values(extras)) for(const row of data) {const a=row[0];row[0]=row[1];row[1]=`=1-(${a})`;}
  for(const [name,data] of Object.entries(extras)) {
    const offset={zero:0,reference:3,growth:6,badge:9}[name];
    helper.getRange(`EN${3+offset}:EO${2+offset+data.length}`).formulas=data;
  }
  if(hasSubtitle(p)) {
    // One formula-linked native data label supplies the filled head.
    // Both its glyph and normalized anchor follow growth sign without a pane.
    // Keep five native XY points / existing series count for bounded migration.
    const head=`IF(H1>=F1,1,-1)*MIN(7/${PAGE58.plot.height},ABS(H1-F1))`;
    const center=`=H1-(${head})/2`;
    helper.getRange('P3').formulas=[['=IF(H1>F1,"▲",IF(H1<F1,"▼",""))']];
    helper.getRange('EN18:EO22').formulas=[
      ['0.989','=F1'],['0.989',center],
      ['0.989',center],['0.989',center],['0.989',center]
    ];
  }
  return {...p,connectorStart};
}
const labelBase=p=>p.body.length;
const hasSubtitle=p=>p.o.variant==='page058'&&p.o.orientation==='vertical'&&p.o.coordinateMode==='scaled'&&p.info.components===1;
function setData(series,helper,xcol,ycol,first,last) {series.setXAxisValues(range(helper,xcol,first,last));series.setValues(range(helper,ycol,first,last));}
function bind(chart,helper,binding) {
  const p=binding;
  for(let i=0;i<p.body.length;i++) {const s=chart.series.getItemAt(i);s.setXAxisValues(helper.getRange(`A3:A${p.last}`));s.setValues(range(helper,p.body[i].column,3,p.last));}
  const at=labelBase(p);
  setData(chart.series.getItemAt(at),helper,AUX,AUX+1,3,p.last);
  setData(chart.series.getItemAt(at+1),helper,AUX+3,AUX+4,3,p.last);
  for(let j=0;j<p.info.components;j++) {
    const s=chart.series.getItemAt(at+6+j);
    if(hasSubtitle(p))setData(s,helper,143,144,15,15);
    else setData(s,helper,p.o.orientation==='vertical'?AUX:75+j,p.o.orientation==='vertical'?75+j:AUX+1,3,p.last);
  }
  for(const [offset,row,len] of [[2,3,2],[3,6,2],[4,9,2],[5,12,1]]) setData(chart.series.getItemAt(at+offset),helper,143,144,row,row+len-1); // EN/EO
  if(hasSubtitle(p))setData(chart.series.getItemAt(at+4),helper,143,144,18,22);
}
function scatter(series,visible=false,mode='scaled') {
  series.chartType=Excel.ChartType.xyscatterLines;series.axisGroup=mode==='values'?Excel.ChartAxisGroup.primary:Excel.ChartAxisGroup.secondary;
  series.markerStyle=Excel.ChartMarkerStyle.none;series.format.line.lineStyle=visible?'Continuous':'None';
  if(visible){series.format.line.color=COLORS.line;series.format.line.weight=.6;}
}
async function connectors({context,chart,helper,binding},initial=false) {
  const start=labelBase(binding)+6+binding.info.components;
  chart.series.load('items');await context.sync();
  for(let i=0;i<binding.n-1;i++) {
    let s;
    if(start+i>=chart.series.items.length) {s=chart.series.add('连接线'+(i+1));scatter(s,binding.o.connectors,binding.o.coordinateMode);}
    else s=chart.series.getItemAt(start+i);
    const r=binding.connectorStart+i*2;setData(s,helper,140,141,r,r+1); // EK/EL
  }
  // Keep removed connector objects (and custom formats), bind to a zero-length line.
  for(let i=start+binding.n-1;i<chart.series.items.length;i++) setData(chart.series.getItemAt(i),helper,143,144,3,3);
  await context.sync();
}
async function createChart(args) {
  const {context,sourceSheet,source,helper,instance,binding:p}=args;
  const chart=sourceSheet.charts.add(p.o.orientation==='vertical'?Excel.ChartType.columnStacked:Excel.ChartType.barStacked,range(helper,9,3,p.last),Excel.ChartSeriesBy.columns);
  chart.name=instance.chartName; chart.series.getItemAt(0).name=p.body[0].name;
  for(const d of p.body.slice(1)) chart.series.add(d.name);
  for(const name of ['值标签','分类标签','零线','起点参考线','增长线','增长标签',...p.info.componentHeaders.map(h=>'细分标签-'+h)]) chart.series.add(name);
  await context.sync();bind(chart,helper,p);
  const anchor=sourceSheet.getRangeByIndexes(source.rowIndex,source.columnIndex+source.columnCount,1,1);anchor.load('left,top');await context.sync();
  chart.left=anchor.left+12;chart.top=anchor.top;chart.width=p.o.variant==='page058'?PAGE58.width:700;chart.height=p.o.variant==='page058'&&p.o.orientation==='vertical'?PAGE58.height:420;
  return chart;
}
function labels(chart,helperName,p,styleLabels=false,oldN=0) {
  const at=labelBase(p), q=quote(helperName);
  function link(index,column,position,show=true,color=COLORS.text) {
    const s=chart.series.getItemAt(index);
    for(let i=0;i<p.n;i++) {
      const point=s.points.getItemAt(i);if(styleLabels || i>=oldN) point.hasDataLabel=show;
      if(!show) continue;
      point.dataLabel.formula=`=${q}!$${col(column)}$${i+3}`;
      if(styleLabels || i>=oldN){point.dataLabel.position=position;point.dataLabel.format.font.name='微软雅黑';point.dataLabel.format.font.size=10;point.dataLabel.format.font.color=color;}
    }
  }
  link(at,AUX+2,p.o.orientation==='vertical'?Excel.ChartDataLabelPosition.top:Excel.ChartDataLabelPosition.right);
  link(at+1,AUX+5,p.o.orientation==='vertical'?Excel.ChartDataLabelPosition.bottom:Excel.ChartDataLabelPosition.left);
  for(let j=0;j<p.info.components;j++) if(!hasSubtitle(p))link(at+6+j,85+j,Excel.ChartDataLabelPosition.center,p.info.components>1,contrast(colorsFor(p.o.palette).stack[j]));
  if(hasSubtitle(p)) {
    const head=chart.series.getItemAt(at+4).points.getItemAt(1);
    if(styleLabels)head.hasDataLabel=true;
    head.dataLabel.formula=`=${q}!$P$3`;
    if(styleLabels){head.dataLabel.position=Excel.ChartDataLabelPosition.center;head.dataLabel.format.font.name='微软雅黑';head.dataLabel.format.font.size=10;head.dataLabel.format.font.bold=false;head.dataLabel.format.font.color=COLORS.line;head.dataLabel.format.fill.clear();head.dataLabel.format.border.lineStyle='None';}
    const point=chart.series.getItemAt(at+6).points.getItemAt(0);
    if(styleLabels)point.hasDataLabel=true;point.dataLabel.formula=`=${q}!$P$2`;
    if(styleLabels){point.dataLabel.position=Excel.ChartDataLabelPosition.right;point.dataLabel.format.font.name='微软雅黑';point.dataLabel.format.font.size=10;point.dataLabel.format.font.color=COLORS.text;}
  }
  const badge=chart.series.getItemAt(at+5).points.getItemAt(0);if(styleLabels)badge.hasDataLabel=p.o.variant==='page058';
  if(p.o.variant==='page058'){badge.dataLabel.formula=`=${q}!$P$1`;if(styleLabels){badge.dataLabel.position=Excel.ChartDataLabelPosition.center;badge.dataLabel.format.font.name='微软雅黑';badge.dataLabel.format.font.size=10;badge.dataLabel.format.fill.setSolidColor('#FFFFFF');}}
}
function hideAxis(axis) {axis.tickLabelPosition=Excel.ChartAxisTickLabelPosition.none;axis.majorTickMark='None';axis.format.line.lineStyle='None';axis.majorGridlines.visible=false;}
async function style(args) {
  const {context,chart,binding:p,options,instance}=args;
  const colors=colorsFor(p.o.palette);
  chart.format.font.name='微软雅黑';chart.format.font.size=10;chart.format.font.color=COLORS.text;
  chart.format.border.color=COLORS.frame;chart.format.border.weight=.7;
  chart.title.visible=true;chart.title.text=options.title;chart.title.format.font.size=14;chart.title.format.font.bold=false;chart.title.horizontalAlignment=Excel.ChartTextHorizontalAlignment.left;
  setChartTitleTypography(chart,{text:options.title});
  if(p.o.variant==='page058'&&p.o.linkedHeadline) chart.title.setFormula(`=${quote(instance.helperName)}!$Q$1`);
  chart.legend.visible=false;
  for(let i=0;i<p.body.length;i++) {
    const s=chart.series.getItemAt(i),d=p.body[i];s.gapWidth=p.o.variant==='page058'?49.2958:55;s.overlap=100;s.format.line.lineStyle='None';
    if(d.kind==='base')s.format.fill.clear();
    else s.format.fill.setSolidColor(p.info.components>1?colors.stack[d.component]:d.kind==='total'?colors.total:d.sign==='positive'?colors.up:colors.down);
  }
  const at=labelBase(p);
  for(let i=at;i<at+6+p.info.components;i++) scatter(chart.series.getItemAt(i),i===at+2 || (p.o.variant==='page058'&&(i===at+3||i===at+4)),p.o.coordinateMode);
  chart.series.getItemAt(at+3).format.line.lineStyle=p.o.variant==='page058'?Excel.ChartLineStyle.dash:'None';
  await context.sync();
  chart.axes.valueAxis.minimum=p.o.coordinateMode==='values'?'':0;chart.axes.valueAxis.maximum=p.o.coordinateMode==='values'?'':1;hideAxis(chart.axes.valueAxis);hideAxis(chart.axes.categoryAxis);
  if(p.o.orientation==='horizontal')chart.axes.categoryAxis.reversePlotOrder=true;
  if(p.o.coordinateMode!=='values')for(const type of [Excel.ChartAxisType.category,Excel.ChartAxisType.value]) {const axis=chart.axes.getItem(type,Excel.ChartAxisGroup.secondary);axis.minimum=0;axis.maximum=1;hideAxis(axis);}
  labels(chart,instance.helperName,p,true);await context.sync();
  await connectors(args,true);
  chart.title.left=10;chart.title.top=8;await context.sync();
  const plot=p.o.variant==='page058'&&p.o.orientation==='vertical'?PAGE58.plot:{left:p.o.orientation==='horizontal'?140:45,top:65,width:p.o.orientation==='horizontal'?505:610,height:310};
  for(let k=0;k<2;k++){chart.plotArea.insideWidth=plot.width;chart.plotArea.insideHeight=plot.height;chart.plotArea.insideLeft=plot.left;chart.plotArea.insideTop=plot.top;await context.sync();}
  if(hasSubtitle(p)) {
    const series=chart.series.getItemAt(at+6), subtitle=series.points.getItemAt(0).dataLabel;
    // The one-point subtitle series supports horizontal text alignment. The
    // point's width property is read-only; ExcelApi 1.19 supplies setWidth().
    // Size after plot layout, then position after Excel has resized the label.
    // Leave height automatic so a host mismatch cannot silently clip the text.
    series.dataLabels.textOrientation=0;
    series.dataLabels.horizontalAlignment=Excel.ChartTextHorizontalAlignment.left;
    subtitle.setWidth(PAGE58.width-20);await context.sync();
    subtitle.left=10;subtitle.top=33;await context.sync();
    // Native text label geometry remains attached to XY point 1 as values move.
    const head=chart.series.getItemAt(at+4).points.getItemAt(1).dataLabel;
    head.geometricShapeType='Rectangle';head.textOrientation=0;
    head.setWidth(12);head.setHeight(12);head.position=Excel.ChartDataLabelPosition.center;
    await context.sync();
    // First-time native label geometry can restore the chart's inherited text
    // color. Set the graphic color after geometry, before core captures it.
    head.format.font.color=COLORS.line;await context.sync();
    // ExcelApi 1.19 exposes the native label shape itself; keep the linked P1
    // label at its data point, so changing growth also moves the ellipse.
    const badge=chart.series.getItemAt(at+5).points.getItemAt(0).dataLabel;
    badge.geometricShapeType='Ellipse';
    badge.format.fill.setSolidColor('#FFFFFF');
    badge.format.border.lineStyle='Continuous';badge.format.border.color=COLORS.line;badge.format.border.weight=.6;
    badge.setWidth(52);badge.setHeight(28);badge.position=Excel.ChartDataLabelPosition.center;
    await context.sync();
  }
  return {graphicLabels:hasSubtitle(p)?[{series:at+4,point:1}]:[],notes:[p.o.coordinateMode==='values'?'原始数值轴由Excel自动调整；跨零分段的悬停值为该段长度，原图视觉仍待验收。':'图册定位使用公式归一化坐标，标签显示源数；字体、双端布局及视觉仍待验收。']};
}
async function finalizePresentation({context,chart,binding}) {
  if(!hasSubtitle(binding))return;
  const subtitle=chart.series.getItemAt(labelBase(binding)+6).points.getItemAt(0).dataLabel;
  subtitle.setWidth(PAGE58.width-20);await context.sync();
  subtitle.left=10;subtitle.top=33;await context.sync();
  const head=chart.series.getItemAt(labelBase(binding)+4).points.getItemAt(1).dataLabel;
  // This hook is create/restyle-only; existing user colors remain untouched on
  // rebind. Do not let a first-render inherited gray masquerade as a kept swatch.
  head.format.font.color=COLORS.line;await context.sync();
  head.format.font.load('color');await context.sync();
  if(String(head.format.font.color).replace(/^#/,'').toUpperCase()!==COLORS.line.slice(1))throw new UserError('瀑布图[增长箭头颜色]：宿主未保持指定浅灰色。');
}
async function rebind(args) {
  const {context,chart,helper,binding,instance,previous}=args;
  // writeHelper checked the persisted component count before changing any formulas.
  bind(chart,helper,binding);await context.sync();labels(chart,instance.helperName,binding,false,previous.count+(binding.o.totals==='append'?1:0));await context.sync();
  if(hasSubtitle(binding)) {
    const point=chart.series.getItemAt(labelBase(binding)+4).points.getItemAt(1);
    point.load('hasDataLabel');await context.sync();
    // Upgrade the previously unlabelled R8 head only. Preserve an existing head's manual styling.
    if(!point.hasDataLabel) {
      point.hasDataLabel=true;
      const head=point.dataLabel;
      head.formula=`=${quote(instance.helperName)}!$P$3`;
      head.format.font.name='微软雅黑';head.format.font.size=10;
      head.format.font.bold=false;head.format.font.color=COLORS.line;
      head.format.fill.clear();head.format.border.lineStyle='None';
      head.geometricShapeType='Rectangle';head.textOrientation=0;
      
      head.setWidth(12);head.setHeight(12);head.position=Excel.ChartDataLabelPosition.center;
      await context.sync();
      head.format.font.color=COLORS.line;await context.sync();
    }
  }
  await connectors(args); // Existing connector styles are untouched; only extra rows get default connectors.
}

const inputSpec={category:{id:'change',title:'变化与分解'},description:'选择分类、1–8列数值，可选末列“类型”（变化/总计/小计）。有类型列时以该列为准；细分正负值围绕同一累计起点展开。',minVersion:'1.19',minRows:2,maxRows:MAX_ROWS,
  options:[{id:'title',label:'标题',type:'text',default:'图表标题',maxLength:120},{id:'variant',label:'版式',type:'select',default:'page058',choices:[{value:'page058',label:'七巨头原例（PDF058）'},{value:'standard',label:'常规瀑布'},{value:'stacked',label:'堆积瀑布'}]},{id:'palette',label:'原例配色',type:'select',default:'default',choices:[{value:'default',label:'默认蓝色'},{value:'pdf156',label:'外卖经济（PDF156）'},{value:'pdf165',label:'液化天然气（PDF165）'},{value:'pdf168',label:'工业机会（PDF168）'}]},{id:'coordinateMode',label:'数值坐标',type:'select',default:'scaled',choices:[{value:'scaled',label:'图册定位'},{value:'values',label:'原始数值（仅垂直）'}]},{id:'orientation',label:'方向',type:'select',default:'vertical',choices:[{value:'vertical',label:'垂直'},{value:'horizontal',label:'水平'}]},{id:'totals',label:'总计位置',type:'select',default:'append',choices:[{value:'append',label:'自动添加末尾总计'},{value:'last',label:'末行是绝对总计'},{value:'marked',label:'按“类型”列指定'},{value:'none',label:'不另加总计'}]},{id:'firstTotal',label:'首项作为绝对起点',type:'checkbox',default:true},{id:'totalLabel',label:'自动总计名称',type:'text',default:'2020-24 年',maxLength:40},{id:'connectors',label:'显示连接线',type:'checkbox',default:true},{id:'linkedHeadline',label:'使用原例动态标题（覆盖上方标题）',type:'checkbox',default:false},{id:'decimals',label:'标签小数位',type:'number',default:1,min:0,max:4,step:1}],
  storage:{marker:MARK,helperPrefix:'CC_瀑布_',chartPrefix:'CC_Waterfall_'}};
export default {id:'waterfall',title:'瀑布图',inputSpec,validate,writeHelper,createChart,style,rebind,finalizePresentation};
