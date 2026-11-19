import {UserError} from '../core/errors.js';

// One native Excel chart. All geometry and text depend on worksheet formulas.
// No error bars, shapes, event handlers, bitmap output, or refresh service.
const MARK = 'ChartCraft.Butterfly.v1';
const HELPER_PREFIX = 'CC_蝴蝶_';
const CHART_PREFIX = 'CC_Butterfly_';
const FONT = '微软雅黑';
const CAREER_FONT = FONT; // User preview standard: Microsoft YaHei on both platforms.
const fontFor = c => c.layout==='career049'?CAREER_FONT:FONT;
const MAX_ROWS = 200;
const COLORS = {left: ['#F692B6', '#BA3F70', '#EBC8D5', '#7F284C'], right: ['#40A1ED', '#0A1C2C', '#323FCA', '#A2E5FB'], neutral: '#BFBFBF', text: '#595959', line: '#595959'};
const PALETTES={career:COLORS,ai:{left:['#B2182B','#EF8A62','#FDDBC7','#7F284C'],right:['#D1E5F0','#67A9CF','#2166AC','#0A1C2C'],neutral:'#D9D9D9'},survey:{left:['#CA0020','#F4A582','#FDDBC7','#7F284C'],right:['#92C5DE','#0571B0','#2166AC','#0A1C2C'],neutral:'#D9D9D9'},mckinsey:{left:['#0A1C2C','#323FCA','#40A1ED','#A2E5FB'],right:['#40A1ED','#323FCA','#0A1C2C','#A2E5FB'],neutral:'#D9D9D9'}};
PALETTES.ai2021={left:['#246696','#4481B0','#79A6CD','#0A1C2C'],right:['#80EBFF','#45CDFD','#01B7F9','#40A1ED'],neutral:'#D9D9D9'};
PALETTES.transport={...PALETTES.mckinsey,left:['#595959','#0A1C2C','#323FCA','#40A1ED']};
PALETTES.green={left:['#A6611A','#DFC27D','#828C95','#000000'],right:['#80CDC1','#018571','#40A1ED','#0A1C2C'],neutral:'#D9D9D9'};
PALETTES.consumer={left:['#000000','#828C95','#2B5580','#051C2C'],right:['#99DCF9','#00A9F0','#40A1ED','#0A1C2C'],neutral:'#D9D9D9'};
PALETTES.jobs={left:['#051C2C','#103559','#2B5580','#0A1C2C'],right:['#40A1ED','#99DCF9','#00A9F0','#0A1C2C'],neutral:'#D9D9D9'};
const quote = name => `'${String(name).replace(/'/g, "''")}'`;
const literal = value => "'" + String(value);
const col = index => { let s = '', n = index + 1; while (n) { n--; s = String.fromCharCode(65 + n % 26) + s; n = Math.floor(n / 26); } return s; };
const RAW = 26, NORMAL = 36, LABEL = 46;
const blank = x => x === '' || x === null;
const categoryText = x => typeof x==='string' && Boolean(x.trim()) && !/^#(?:N\/A|REF!|VALUE!|DIV\/0!|NAME\?|NUM!|NULL!|SPILL!|CALC!|GETTING_DATA)$/i.test(x.trim());

function validate(values) {
  if (!Array.isArray(values) || values.length < 3 || !Array.isArray(values[0])) throw new UserError('请选择含表头和至少2行数据的分类、左值、右值。');
  const columns = values[0].length;
  if (columns < 3 || columns > 10 || values.some(row => !Array.isArray(row) || row.length !== columns)) throw new UserError('需要3–10列等宽数据：分类、左侧系列、右侧系列，以及可选的最后一列中立值。');
  if (values.length - 1 > MAX_ROWS) throw new UserError(`最多支持${MAX_ROWS}行。`);
  if (values[0].some(x => typeof x !== 'string' || !x.trim() || x.length > 255)) throw new UserError('每列需要非空文字表头（不超过255字）。');
  let populated = 0;
  const blankRows = [];
  // A number stored as text must not silently turn a 3-series table into city pairs.
  const dual = columns===4 && values.slice(1).some(row=>categoryText(row[2])&&!Number.isFinite(Number(row[2])));
  const numericColumns=dual?[1,3]:Array.from({length:columns-1},(_,i)=>i+1);
  values.slice(1).forEach((row, i) => {
    if (row.every(blank)) { blankRows.push(i); return; } // Explicit group spacer, never a missing numeric value.
    if (!categoryText(row[0])) throw new UserError(`第${i + 2}行分类为空、错误值或不是文字。`);
    if(dual&&(!categoryText(row[2])||Number.isFinite(Number(row[2]))))throw new UserError(`第${i+2}行右侧分类须为非数字文字。`);
    numericColumns.forEach(j => {
      const x=row[j];if (typeof x !== 'number' || !Number.isFinite(x) || x < 0) throw new UserError(`第${i + 2}行第${j + 1}列必须是非负数字；未知值不能填0。`);
    });
    populated++;
  });
  if (populated < 2) throw new UserError('至少需要2个有效分类，空行只能用作分组间隔。');
  return {count: values.length - 1, columnCount: columns, categoryHeader: values[0][0], seriesHeaders: numericColumns.map(k=>values[0][k]),numericColumns,inputSchema:dual?'dual-categories':'shared-category',blankRows};
}

function config(source, opts = {}) {
  const o = Object.fromEntries(inputSpec.options.map(x => [x.id, x.default])); Object.assign(o, opts);
  for (const item of inputSpec.options) {
    const v = o[item.id];
    if (item.type === 'select' && !item.choices.some(x => x.value === v)) throw new UserError(`无效选项：${item.label}`);
    if (item.type === 'checkbox' && typeof v !== 'boolean') throw new UserError(`${item.label}须为布尔值。`);
    if (item.type === 'number' && (typeof v !== 'number' || !Number.isFinite(v) || v < item.min || v > item.max)) throw new UserError(`${item.label}超出允许范围。`);
    if (item.type === 'text' && (typeof v !== 'string' || v.length > item.maxLength)) throw new UserError(`${item.label}文字超出允许长度。`);
  }
  if (!Number.isInteger(o.leftSeries) || !Number.isInteger(o.decimals)) throw new UserError('左侧系列数和小数位必须为整数。');
  const allNumericColumns=source.numericColumns??Array.from({length:source.columnCount-1},(_,i)=>i+1),dual=source.inputSchema==='dual-categories';
  const numericColumns=o.countRow?allNumericColumns.slice(0,-1):allNumericColumns,countColumn=o.countRow?allNumericColumns.at(-1):null;
  if(o.countRow&&(dual||o.direction!=='vertical'))throw new UserError('独立计数行仅用于纵向图，最后一列计数不参加堆叠。');
  if(o.countRow&&source.values?.slice(1).some(row=>!row.every(blank)&&!Number.isInteger(row[countColumn])))throw new UserError('底部计数须为非负整数。');
  if(['top','value-end'].includes(o.categoryPosition)&&o.direction!=='vertical')throw new UserError('顶部/柱端分类标签仅用于纵向图。');
  const numeric = numericColumns.length, middle = o.neutral !== 'none', right = numeric - o.leftSeries - Number(middle);
  if (right < 1 || right > 4 || o.leftSeries > 4) throw new UserError('请按左侧系列数分列，左右各1–4列；中立项须在最后一列。');
  if (o.scale === 'independent' && o.net !== 'none') throw new UserError('左右单位或尺度独立时不能相减；请关闭Net。');
  if (o.neutral !== 'none' && o.scale === 'independent') throw new UserError('中立项须使用共同尺度。');
  if (o.neutral === 'center' && o.categoryPosition === 'center') throw new UserError('居中中立项占用零轴，请把分类标签放到外侧或左侧。');
  if((o.leftStyle!=='bar'&&o.leftSeries!==1)||(o.rightStyle!=='bar'&&right!==1))throw new UserError('点线或大头针每侧限一个系列。');
  if((o.leftStyle!=='bar'||o.rightStyle!=='bar')&&o.neutral!=='none')throw new UserError('点线/大头针与中立项须分别创建。');
  if(dual&&(middle||o.leftSeries!==1||o.direction!=='horizontal'))throw new UserError('双分类格式仅用于横向左右各一列、无中立项。');
  if(!Number.isInteger(o.rightDecimals))throw new UserError('右侧小数位须为整数。');
  const fmt = (o.units==='dollars'?'"$"':'')+'0' + (o.decimals ? '.' + (o.trimZeros?'#':'0').repeat(o.decimals) : '') + (o.units === 'percent-points' ? '"%"' : o.units === 'fraction' ? '%' : '');
  const gap = o.neutral === 'center' ? 0 : o.centerGap;
  if(o.layout==='career049' && (o.direction!=='horizontal'||middle||o.net!=='none'||o.leftSeries!==1||right!==1||o.scale!=='shared'))throw new UserError('PDF049版式限横向、左右各1列、共同尺度且无中立或Net。');
  if(o.showAxisLabels && gap!==0)throw new UserError('数值刻度目前用于共用零轴，请把中心半间距设为0。');
  const leftExtent = o.layout==='career049'?507.5/429:1 + gap + (o.countRow?.45:o.direction==='horizontal'&&o.categoryPosition==='outside'?.85:.24);
  const rightExtent = o.layout==='career049'?(1146-507.5)/429:1 + gap + (o.neutral === 'separate' ? 1.45 : o.net === 'column' ? .52 : .24);
  const rightFmt=o.rightDecimals<0?fmt:(o.units==='dollars'?'"$"':'')+'0'+(o.rightDecimals?'.'+(o.trimZeros?'#':'0').repeat(o.rightDecimals):'')+(o.units==='percent-points'?'"%"':o.units==='fraction'?'%':'');
  return {...o,valuePosition:dual?'inside':o.valuePosition,left: o.leftSeries, right, numeric,numericColumns,countColumn,dual,schema:dual?'dual-categories':'shared-category',sideCount: numeric - Number(middle), middle, fmt,rightFmt,gap, min: -leftExtent, max: rightExtent};
}

// Optional decimal placeholders alone can leave a trailing decimal point in Excel.
function labelFormat(value,c,right=false) {
  const decimals=right&&c.rightDecimals>=0?c.rightDecimals:c.decimals,base=right?excelString(c.rightFmt):'$G$1';
  if(!c.trimZeros||decimals===0)return base;
  const integer=(c.units==='dollars'?'"$"':'')+'0'+(c.units==='percent-points'?'"%"':c.units==='fraction'?'%':'');
  return `IF(MOD(ROUND(${value}${c.units==='fraction'?'*100':''},${decimals}),1)=0,${excelString(integer)},${base})`;
}

// The inspectable pure plan is also used by the module-local formula tests.
function helperPlan(source, options = {}) {
  const c = config(source, options), n = source.count, last = n + 2;
  const q = quote(source.sheet), src = (r, k) => `${q}!$${col(source.columnIndex + k)}$${source.rowIndex + r + 1}`;
  const rows = [];
  const range = (k, r) => `${col(k)}${r}`;
  const sumRaw = (a, b, r) => `SUM(${col(RAW + a)}${r}:${col(RAW + b)}${r})`;
  for (let i = 0; i < n; i++) {
    const r = i + 3, sr = i + 1;
    const empty = `COUNTBLANK(${src(sr, 0)}:$${col(source.columnIndex+source.columnCount-1)}$${source.rowIndex+sr+1})=${source.columnCount}`;
    const validCategories=[src(sr,0),...(c.dual?[src(sr,2)]:[])].map(ref=>`ISTEXT(${ref}),LEN(TRIM(${ref}))>0`).join(',');
    const a = Array(56).fill('');
    a[0] = `=IFERROR(${src(sr, 0)}&"","数据无效")`;
    for (let k = 0; k < c.numeric; k++) {
      const ref = src(sr, c.numericColumns[k]), raw = range(RAW + k, r);
      a[RAW + k] = `=IF(${empty},0,IFERROR(IF(AND(ISNUMBER(${ref}),${ref}>=0,${validCategories}),${ref},NA()),NA()))`;
      a[LABEL + k] = `=IF(A${r}="","",IFERROR(TEXT(${raw},${labelFormat(raw,c,k>=c.left)}),"数据无效"))`;
      if (k < c.sideCount) a[NORMAL + k] = `=${k < c.left ? '-' : ''}${raw}/${k < c.left ? '$E$1' : '$F$1'}`;
    }
    a[1] = `=${sumRaw(0, c.left - 1, r)}`;
    a[2] = `=${sumRaw(c.left, c.sideCount - 1, r)}`;
    a[3] = `=C${r}-B${r}`;
    a[4] = c.middle ? `=${range(RAW + c.numeric - 1, r)}` : '=0';
    a[5] = c.direction === 'horizontal' ? `=1-(ROW()-2.5)/$C$1` : `=(ROW()-2.5)/$C$1`;
    a[6] = '=0';
    a[7] = '=-$D$1'; a[8] = '=$D$1';
    const half = c.neutral === 'center' ? `+E${r}/2/$E$1` : '';
    a[9] = `=-($D$1+B${r}/$E$1${half})`;
    a[10] = `=$D$1+C${r}/$F$1${half}`;
    for (const [k, value] of [[11, `B${r}`], [12, `C${r}`], [13, `D${r}`], [14, `E${r}`]]) {
      const fmt=labelFormat(value,c,k===12);
      a[k] = `=IF(A${r}="","",IFERROR(TEXT(${value},${k === 13 ? `"+"&${fmt}&";-"&${fmt}&";"&${fmt}` : fmt}),"数据无效"))`;
    }
    a[24]=c.dual?`=IFERROR(${src(sr,2)}&"","数据无效")`:`=A${r}`;
    if(c.dual)a[25]=`=F${r}+0.42/$C$1`;
    a[15] = c.categoryPosition === 'center' ? '=0' : c.categoryPosition === 'left' ? '=-$D$1+0.02' : c.categoryPosition==='value-end'?`=K${r}`:c.categoryPosition==='top'?`=${c.max-.02}`:`=${c.min+(c.countRow?.25:.03)}`;
    a[16] = c.net === 'marker' ? `=D${r}/$E$1` : `=${c.max - .14}`;
    a[17] = c.neutral === 'separate' ? `=MAX(${c.neutralStart}/$E$1,$D$1+IFERROR(AGGREGATE(14,6,$C$3:$C$${last},1),0)/$F$1+0.2)-K${r}` : '=0';
    a[18] = c.neutral === 'separate' ? `=E${r}/$E$1` : '=0';
    a[19] = `=IFERROR(IF(AND(COUNT(${col(RAW)}${r}:${col(RAW + c.numeric - 1)}${r})=${c.numeric},OR(A${r}<>"",${empty})${c.countRow?`,ISNUMBER(Z${r})`:''}),1,0),0)`;
    a[20] = c.neutral === 'center' ? `=-E${r}/2/$E$1` : '=0';
    a[21] = c.neutral === 'center' ? `=E${r}/2/$E$1` : '=0';
    a[22] = `=IF(T${r}=1,"","数据无效")`;
    a[23] = c.neutral === 'separate' ? `=R${r}+K${r}+S${r}` : '=0';
    if(c.countRow){const ref=src(sr,c.countColumn);a[25]=`=IF(${empty},0,IFERROR(IF(AND(ISNUMBER(${ref}),${ref}>=0,INT(${ref})=${ref},${validCategories}),${ref},NA()),NA()))`;a[54]=`=IF(A${r}="","",IFERROR(TEXT(Z${r},"0"),"数据无效"))`;a[55]=`=${c.min+.03}`;}
    rows.push(a);
  }
  // AGGREGATE ignores deliberate #N/A gaps from invalid edits; no text-to-zero coercion.
  const leftMax = `MAX(IFERROR(AGGREGATE(14,6,$B$3:$B$${last},1),0),${c.showReferences?c.leftReference:0})`, rightMax = `MAX(IFERROR(AGGREGATE(14,6,$C$3:$C$${last},1),0),${c.showReferences?c.rightReference:0})`;
  // Centered neutral can enlarge both sides. Bounding by full neutral guarantees space.
  const neutralMax = c.neutral === 'center' ? `+IFERROR(AGGREGATE(14,6,$E$3:$E$${last},1),0)/2` : '';
  const separateBound=c.neutral==='separate'?`,IFERROR(AGGREGATE(14,6,$E$3:$E$${last},1),0),(${c.neutralStart}+IFERROR(AGGREGATE(14,6,$E$3:$E$${last},1),0))/2.3`:'';
  const denominator = c.scale === 'shared' ? `MAX(IFERROR(${leftMax},0)${c.layout==='career049'?'/0.7':''},IFERROR(${rightMax},0)${separateBound})${neutralMax}` : null;
  const params = [[String(source.columnCount), String(n), String(c.gap), `=MAX(1E-12,${denominator ?? `IFERROR(${leftMax},0)`})`, `=MAX(1E-12,${denominator ?? `IFERROR(${rightMax},0)`})`, literal(c.fmt), literal(c.schema)]];
  const headers = Array(56).fill('');
  ['分类','左合计','右合计','Net','中立','行位置','零','左间隔','右间隔','左端','右端','左标签','右标签','Net标签','中立标签','分类位置','Net位置','中立间隔','中立宽','数据有效','中立左半','中立右半','输入错误','中立右端'].forEach((x,i)=>headers[i]=literal(x));
  c.numericColumns.forEach((_,i)=>{headers[RAW+i]=`=${src(0,c.numericColumns[i])}&""`;headers[NORMAL+i]=literal(`几何${i+1}`);headers[LABEL+i]=literal(`标签${i+1}`);});
  if(c.countRow){headers[25]=literal('独立计数');headers[54]=literal('计数标签');headers[55]=literal('计数标签位置');}
  return {config:c,last,rows,params,headers,errorsFormula:`=COUNTIF(T3:T${last},0)`, neutralDenominator:`=MAX(1E-12,IFERROR(AGGREGATE(14,6,$E$3:$E$${last},1),0))`};
}
function excelString(value){return '"'+String(value).replace(/"/g,'""')+'"';}

async function writeHelper({context,helper,source,options,previous}) {
  const plan = helperPlan(source,options);
  if (previous) {
    const old = helper.getRange('B1:H1'); old.load('values'); await context.sync();
    if (Number(old.values[0][0]) !== source.columnCount||old.values[0][6]!==plan.config.schema) throw new UserError('重绑须保持原系列列数和分类结构；要改左右或中立结构请新建图，避免覆盖手改格式。');
    helper.getRange(`A3:BD${Math.max(previous.count,source.count)+2}`).clear(Excel.ClearApplyTo.contents);
    helper.getRange(`BF3:BH${Math.max(previous.count,source.count)*3+2}`).clear(Excel.ClearApplyTo.contents);
    helper.getRange(`BL3:BM${Math.max(previous.count,source.count)*3+2}`).clear(Excel.ClearApplyTo.contents);
  }
  helper.getRange('B1:H1').formulas = plan.params;
  helper.getRange('N1:Q1').formulas = [[String(plan.config.left),String(plan.config.right),plan.errorsFormula,plan.neutralDenominator]];
  helper.getRange('A2:BD2').formulas = [plan.headers];
  helper.getRange(`A3:BD${plan.last}`).formulas = plan.rows;
  helper.getRange(`B3:E${plan.last}`).numberFormat = plan.rows.map(()=>Array(4).fill(plan.config.fmt));
  // Native series for inner edges, independent of count; reserved below the maximum data row.
  helper.getRange('A205:D206').formulas = [['=-$D$1','=$D$1','0','1'],['=-$D$1','=$D$1','1','0']];
  if(plan.config.leftStyle!=='bar'||plan.config.rightStyle!=='bar'){
    const lines=[];
    for(let i=0;i<source.count;i++){const r=i+3;lines.push(['=-$D$1','=$D$1',`=F${r}`],[`=J${r}`,`=K${r}`,`=F${r}`],['=NA()','=NA()','=NA()']);}
    helper.getRange(`BF3:BH${source.count*3+2}`).formulas=lines;
  }
  if(plan.config.showReferences)helper.getRange('A208:F209').formulas=[
    [`=-$D$1-${plan.config.leftReference}/$E$1`,`=$D$1+${plan.config.rightReference}/$F$1`,'0','1',`=TEXT(${plan.config.leftReference},${labelFormat(String(plan.config.leftReference),plan.config)})`,`=TEXT(${plan.config.rightReference},${labelFormat(String(plan.config.rightReference),plan.config,true)})`],
    [`=-$D$1-${plan.config.leftReference}/$E$1`,`=$D$1+${plan.config.rightReference}/$F$1`,'1','0','','']
  ];
  if(plan.config.showAxisLabels){
    const ticks=[];for(let i=-5;i<=5;i++){const x=i/5,value=`ABS(A${i+217})*${x<0?'$E$1':'$F$1'}`;ticks.push([String(x),'0',`=TEXT(${value},${labelFormat(value,plan.config,x>=0)})`]);}
    helper.getRange('A212:C222').formulas=ticks;
  }
  if(plan.config.showLeaders){const lines=[];for(let i=0;i<source.count;i++){const r=i+3;lines.push(['=-$D$1',`=F${r}`],['=$D$1',`=F${r}`],['=NA()','=NA()']);}helper.getRange(`BL3:BM${source.count*3+2}`).formulas=lines;}
  if(plan.config.layout==='career049')helper.getRange('A225:C227').formulas=[
    [literal(plan.config.subtitle),'0','1'],['=$AA$2','0','1'],['=$AB$2','0','1']
  ];
  return {last:plan.last,config:plan.config};
}

function descriptors(c) {
  const d = [{role:'spacer',name:'左间隔',value:7},{role:'spacer',name:'右间隔',value:8}];
  if(c.neutral==='center')d.push({role:'neutral',nameIndex:c.numeric-1,value:20},{role:'neutral',nameIndex:c.numeric-1,value:21});
  const leftOrder=Array.from({length:c.left},(_,i)=>i);if(c.leftOrder==='outside-in')leftOrder.reverse();
  for(const k of [...leftOrder,...Array.from({length:c.right},(_,i)=>i+c.left)])d.push({role:'bar',side:k<c.left?'left':'right',ordinal:k<c.left?k:k-c.left,nameIndex:k,value:NORMAL+k,label:LABEL+k});
  if(c.neutral==='separate')d.push({role:'spacer',name:'中立间隔',value:17},{role:'neutral',name:'中立',nameIndex:c.numeric-1,value:18});
  d.push({role:'category',name:'分类标签',x:15,label:0},{role:'value',name:'左合计标签',side:'left',x:9,label:11},{role:'value',name:'右合计标签',side:'right',x:10,label:12});
  if(c.neutral==='center')d.push({role:'neutralLabel',name:'中立标签',x:6,label:14});
  if(c.neutral==='separate')d.push({role:'neutralValue',name:'中立标签',side:'right',x:23,label:14});
  if(c.net!=='none')d.push({role:'net',name:'Net',x:16,label:13});
  if(c.showBaselines && c.neutral!=='center')d.push({role:'edge',name:'左内边界',edge:'A'},{role:'edge',name:'右内边界',edge:'B'});
  for(const side of ['left','right'])if(c[side+'Style']!=='bar')d.push({role:'pinLine',side,nameIndex:side==='left'?0:c.left});
  if(c.showReferences)d.push({role:'reference',side:'left',name:'左参考值',edge:'A',labelCell:'E208'},{role:'reference',side:'right',name:'右参考值',edge:'B',labelCell:'F208'});
  if(c.showAxisLabels)d.push({role:'axisLabels',name:'数值刻度'});
  if(c.showLeaders)d.push({role:'leader',name:'分类引导线'});
  if(c.dual)d.push({role:'dualCategory',side:'left',name:'左城市',x:7,y:25,label:0},{role:'dualCategory',side:'right',name:'右城市',x:8,y:25,label:24});
  if(c.countRow)d.push({role:'countRow',name:'独立计数行',x:55,label:54});
  // Text-only one-point series cannot expose labels on the 14 spacer bars.
  if(c.layout==='career049'){
    if(c.subtitle)d.push({role:'careerText',name:'副标题',text:'subtitle',row:225});
    if(c.showLegend)d.push({role:'careerText',name:'左侧标题',text:'left',row:226},{role:'careerText',name:'右侧标题',text:'right',row:227});
  }
  return d;
}
function bindSeries(chart,helper,binding) {
  const {last,config:c} = binding, range=k=>helper.getRange(`${col(k)}3:${col(k)}${last}`);
  descriptors(c).forEach((d,i)=>{
    const s=chart.series.getItemAt(i);
    if(d.value!==undefined){s.setXAxisValues(range(0));s.setValues(range(d.value));}
    else if(d.role==='careerText'){
      s.setXAxisValues(helper.getRange(`B${d.row}`));s.setValues(helper.getRange(`C${d.row}`));
    }else if(d.role==='edge'||d.role==='reference'){
      const row=d.role==='edge'?205:208;
      s.setXAxisValues(helper.getRange(c.direction==='horizontal'?`${d.edge}${row}:${d.edge}${row+1}`:`C${row}:C${row+1}`));
      s.setValues(helper.getRange(c.direction==='horizontal'?`C${row}:C${row+1}`:`${d.edge}${row}:${d.edge}${row+1}`));
    }else if(d.role==='pinLine'){
      const last=(binding.last-2)*3+2,which=d.side==='left'?'BF':'BG';
      s.setXAxisValues(helper.getRange(c.direction==='horizontal'?`${which}3:${which}${last}`:`BH3:BH${last}`));
      s.setValues(helper.getRange(c.direction==='horizontal'?`BH3:BH${last}`:`${which}3:${which}${last}`));
    }else if(d.role==='axisLabels'){
      s.setXAxisValues(helper.getRange(c.direction==='horizontal'?'A212:A222':'B212:B222'));
      s.setValues(helper.getRange(c.direction==='horizontal'?'B212:B222':'A212:A222'));
    }else if(d.role==='leader'){
      const last=(binding.last-2)*3+2;
      s.setXAxisValues(helper.getRange(c.direction==='horizontal'?`BL3:BL${last}`:`BM3:BM${last}`));
      s.setValues(helper.getRange(c.direction==='horizontal'?`BM3:BM${last}`:`BL3:BL${last}`));
    } else {
      s.setXAxisValues(range(c.direction==='horizontal'?d.x:5));
      s.setValues(range(c.direction==='horizontal'?(d.y??5):d.x));
    }
  });
}
function linkLabels(chart,helperName,n,c,{initial=false,oldCount=0}={}) {
  descriptors(c).forEach((d,i)=>{
    if(d.label===undefined)return;
    if(d.role==='category'&&c.dual)return;
    if(d.role==='bar' && c.valuePosition!=='inside'&&(!c.segmentLabels || c.sideCount===2))return;
    if(d.role==='value' && (!c.showValues||c.valuePosition==='inside'))return;
    if(d.role==='net'&&c.net==='marker'&&!c.showNetLabels)return;
    const series=chart.series.getItemAt(i);
    for(let j=0;j<n;j++){
      const point=series.points.getItemAt(j);
      if(initial || j>=oldCount)point.hasDataLabel=true;
      point.dataLabel.formula=`=${quote(helperName)}!$${col(d.label)}$${j+3}`;
      // Do not rewrite existing placement/font/fill during rebind.
      if(initial || j>=oldCount){
        const label=point.dataLabel;
        label.format.font.name=fontFor(c); label.format.font.size=c.fontSize; label.format.font.color=COLORS.text;
        if(d.role==='bar'){const hex=PALETTES[c.palette][d.side][d.ordinal];const rgb=[1,3,5].map(i=>parseInt(hex.slice(i,i+2),16));label.format.font.color=c.dual?(d.side==='left'?'#FFFFFF':COLORS.text):rgb[0]*.299+rgb[1]*.587+rgb[2]*.114<145?'#FFFFFF':COLORS.text;}
        label.position=d.role==='bar'?(c.dual?'InsideEnd':'Center'):d.role==='neutralLabel'?'Center':d.role==='countRow'?'Bottom':d.role==='category'?(c.categoryPosition==='center'?'Center':c.categoryPosition==='value-end'?'Top':(c.direction==='horizontal'?'Right':'Bottom')):d.role==='net'&&c.net==='marker'?'Right':c.direction==='horizontal'?(d.side==='left'?'Left':'Right'):(d.side==='left'?'Bottom':'Top');
        if(d.role==='category'){label.format.fill.setSolidColor(c.gapFill==='gray'?'#F2F2F2':'#FFFFFF');label.format.border.lineStyle='None';}
      }
    }
  });
}
async function createChart({context,sourceSheet,source,helper,instance,binding}) {
  const c=binding.config,ds=descriptors(c),type=c.direction==='horizontal'?'BarStacked':'ColumnStacked';
  const chart=sourceSheet.charts.add(type,helper.getRange(`H3:H${binding.last}`),'Columns');
  chart.name=instance.chartName;
  chart.plotVisibleOnly=false;
  chart.displayBlanksAs='NotPlotted';
  const anchor=sourceSheet.getRangeByIndexes(source.rowIndex,source.columnIndex+source.columnCount,1,1);anchor.load('left,top'); await context.sync();
  chart.left=anchor.left+12;chart.top=anchor.top;chart.width=c.width;chart.height=c.height;
  chart.series.getItemAt(0).name=ds[0].name;
  for(const d of ds.slice(1))chart.series.add(d.nameIndex===undefined?d.name:source.seriesHeaders[d.nameIndex]);
  await context.sync();
  ds.forEach((d,i)=>{
    const s=chart.series.getItemAt(i);
    s.chartType=d.value!==undefined?type:['edge','reference','pinLine','leader'].includes(d.role)?'XYScatterLinesNoMarkers':'XYScatter';
    s.axisGroup=d.value!==undefined?'Primary':'Secondary';
  });
  await context.sync();bindSeries(chart,helper,binding);await context.sync();
  return chart;
}
function hideAxis(axis){axis.tickLabelPosition='None';axis.majorTickMark='None';axis.minorTickMark='None';axis.format.line.lineStyle='None';axis.majorGridlines.visible=false;}
async function syncStage(context,stage){
  try{await context.sync();}
  catch(error){error.message=`蝴蝶图[${stage}]：${error.message}`;throw error;}
}
async function hideSecondaryAxis(context,axis,minimum,maximum,stage){
  // Mac can reject line/gridline formatting on an already hidden secondary axis.
  // Preserve its geometry, hide the axis itself, and verify the native readback.
  axis.minimum=minimum;axis.maximum=maximum;
  await syncStage(context,stage+'范围');
  axis.visible=false;
  await syncStage(context,stage+'隐藏');
  axis.load('minimum,maximum,visible');
  await syncStage(context,stage+'回读');
  if(axis.visible!==false||!Number.isFinite(axis.minimum)||!Number.isFinite(axis.maximum)||Math.abs(axis.minimum-minimum)>1e-9||Math.abs(axis.maximum-maximum)>1e-9)throw new UserError(`蝴蝶图[${stage}回读]：隐藏后未保留范围${minimum}–${maximum}或可见性，原生图未验收。`);
  return {axis:stage,minimum:axis.minimum,maximum:axis.maximum,visible:axis.visible};
}
async function style({context,chart,source,options,instance,binding}) {
  const c=binding.config,ds=descriptors(c),horizontal=c.direction==='horizontal',palette=PALETTES[c.palette],career=c.layout==='career049',font=fontFor(c),legendVisible=c.showLegend&&!career;
  chart.format.font.name=font;chart.format.font.size=c.fontSize;chart.format.font.color=COLORS.text;
  chart.format.fill.setSolidColor('#FFFFFF');chart.format.border.color='#D9D9D9';chart.format.border.weight=.75;
  chart.title.visible=Boolean(c.title);chart.title.text=c.title+(!career&&c.subtitle?'\n'+c.subtitle:'');chart.title.format.font.size=14;chart.title.format.font.bold=false;chart.title.horizontalAlignment='Left';
  if(career){
    chart.title.format.font.name=font;chart.title.format.font.color='#000000';chart.title.overlay=true;
  }
  chart.legend.visible=legendVisible;
  if(legendVisible){chart.legend.position='Top';chart.legend.format.font.size=c.fontSize;}
  ds.forEach((d,i)=>{
    const s=chart.series.getItemAt(i);s.format.line.lineStyle='None';
    if(d.value!==undefined){
      s.gapWidth=c.barGap;s.overlap=100;
      if(d.role==='spacer'&&[7,8].includes(d.value)&&c.gapFill==='gray')s.format.fill.setSolidColor('#F2F2F2');
      else if(d.role==='spacer'||d.role==='bar'&&c[d.side+'Style']!=='bar')s.format.fill.clear();else s.format.fill.setSolidColor(d.role==='neutral'?palette.neutral:palette[d.side][d.ordinal]);
    }else{
      s.markerStyle='None';
      if(d.role==='edge'){s.format.line.lineStyle='Continuous';s.format.line.color=COLORS.line;s.format.line.weight=.6;}
      if(d.role==='net'&&c.net==='marker'){s.markerStyle='Circle';s.markerSize=7;s.markerBackgroundColor='#FFFFFF';s.markerForegroundColor='#0A1C2C';}
      if(d.role==='pinLine'){s.format.line.lineStyle=c[d.side+'Style']==='dotted'?'RoundDot':'Continuous';s.format.line.color=palette[d.side][0];s.format.line.weight=1;}
      if(d.role==='reference'){s.format.line.lineStyle='Dash';s.format.line.color=COLORS.text;s.format.line.weight=.75;}
      if(d.role==='leader'){s.format.line.lineStyle='Continuous';s.format.line.color='#D9D9D9';s.format.line.weight=.4;}
    }
  });
  await syncStage(context,'图表与系列样式');
  const cat=chart.axes.getItem('Category','Primary'),val=chart.axes.getItem('Value','Primary');
  hideAxis(cat);cat.reversePlotOrder=horizontal;
  await syncStage(context,'主分类轴');
  hideAxis(val);val.minimum=c.min;val.maximum=c.max;val.majorUnit=.5;
  await syncStage(context,'主数值轴');
  const sx=chart.axes.getItem('Category','Secondary'),sy=chart.axes.getItem('Value','Secondary');
  const secondaryAxes=[];
  secondaryAxes.push(await hideSecondaryAxis(context,sx,horizontal?c.min:0,horizontal?c.max:1,'次散点X轴'));
  secondaryAxes.push(await hideSecondaryAxis(context,sy,horizontal?0:c.min,horizontal?1:c.max,'次散点Y轴'));
  linkLabels(chart,instance.helperName,source.count,c,{initial:true});
  linkSpecialPoints(chart,instance.helperName,source.count,c,{initial:true});
  await syncStage(context,'链接标签');
  if(legendVisible){
    const entries=chart.legend.legendEntries;entries.load('items');await syncStage(context,'读取图例');
    ds.forEach((d,i)=>{if(entries.items[i])entries.items[i].visible=d.role==='bar'&&c[d.side+'Style']==='bar'||d.role==='pinLine'||d.role==='neutral'&&!(c.neutral==='center'&&d.value===20)||d.role==='net';});
    await syncStage(context,'图例可见性');
  }
  // Native axis alignment stays count-independent; rebind never rewrites it.
  const plot=c.layout==='career049'?{left:0,top:c.height*133/702,width:c.width,height:c.height*546/702}:{left:horizontal?10:45,top:c.subtitle?80:65,width:c.width-(horizontal?20:65),height:c.height-(c.subtitle?95:80)};
  for(let k=0;k<2;k++){chart.plotArea.insideLeft=plot.left;chart.plotArea.insideTop=plot.top;chart.plotArea.insideWidth=plot.width;chart.plotArea.insideHeight=plot.height;await syncStage(context,`绘图区${k+1}`);}
  const textLayout=career?await placeCareerText(context,chart,c):null;
  return {notes:['原生Bar/ColumnStacked＋XYScatter；字体、标签碰撞和双轴对齐待双端实机验收。'],variant:{direction:c.direction,left:c.left,right:c.right,neutral:c.neutral,net:c.net},plot,secondaryAxes,textLayout};
}
async function measureCareerSubtitle(text){
  // Only measure glyph advances; the visible result remains a formula-linked
  // native Excel label. A full-chart width leaves Excel's centered text anchor
  // roughly 97pt to the right even when the paragraph reports Left alignment.
  const cssFont='normal 10pt "Microsoft YaHei", "微软雅黑"';
  if(globalThis.document?.fonts?.load)await document.fonts.load(cssFont,text);
  const canvas=globalThis.document?.createElement('canvas'),metrics=canvas?.getContext('2d');
  if(!metrics)throw new UserError('蝴蝶图[049副标题量宽]：当前宿主无法测量微软雅黑10pt文字。');
  metrics.font=cssFont;
  const advanceWidthPt=Math.max(...String(text).split(/\r?\n/).map(line=>metrics.measureText(line).width))*72/96;
  if(!Number.isFinite(advanceWidthPt)||advanceWidthPt<=0)throw new UserError('蝴蝶图[049副标题量宽]：宿主未返回有效文字宽度。');
  // Four points per side cover native label text margins. Quarter-point ceiling
  // avoids a borderline wrap without retaining the former 194pt empty space.
  return {font:CAREER_FONT,fontSize:10,advanceWidthPt,paddingPt:8,width:Math.ceil((advanceWidthPt+8)*4)/4,measurement:'host-canvas-font-metrics'};
}
async function placeCareerText(context,chart,c){
  const ds=descriptors(c),labelFor=kind=>{const i=ds.findIndex(d=>d.role==='careerText'&&d.text===kind);return i<0?null:chart.series.getItemAt(i).points.getItemAt(0).dataLabel;};
  const headings=c.showLegend?['left','right'].map(labelFor):[],subtitle=labelFor('subtitle');
  const rect=o=>o?Object.fromEntries(['left','top','width','height'].map(k=>[k,Number.isFinite(o[k])?o[k]:null])):null;
  const subtitleMetrics=subtitle?await measureCareerSubtitle(c.subtitle):null;
  const requested={title:{left:c.width*12/1146,top:c.height*8/702},subtitle:{left:c.width*12/1146,top:c.height*46/702,width:subtitleMetrics?Math.min(c.width-20,subtitleMetrics.width):0,height:24},headingCenters:[258,849].map(x=>c.width*x/1146),headingCenterY:c.height*115.5/702};
  if(c.title){chart.title.left=c.width*12/1146;chart.title.top=c.height*8/702;}
  // Resizing width alone preserves the old tall box on Mac. Set both native
  // dimensions explicitly; 24pt allows the final 10pt line plus margins. Visual clipping
  // still needs the native screenshot check; text readback alone cannot prove it.
  if(subtitle){subtitle.setWidth(requested.subtitle.width);subtitle.setHeight(requested.subtitle.height);}
  await syncStage(context,'049副标题宽高');
  // Excel's manual label positions depend on the rendered text rectangle.
  // Flush after the final typography, before using that rectangle as an anchor.
  chart.getImage();await syncStage(context,'049最终字体排版');
  for(const label of headings)label.load('width,height');
  if(subtitle)subtitle.load('left,top,width,height');
  if(c.title)chart.title.load('left,top,width,height');
  await syncStage(context,'049标题与系列标题尺寸');
  const resizedSubtitle=rect(subtitle);
  for(let i=0;i<headings.length;i++){
    const label=headings[i];
    if(!Number.isFinite(label.width)||!Number.isFinite(label.height))throw new UserError(`蝴蝶图[049系列标题]：无法取得原生标签尺寸；actual=${JSON.stringify(rect(label))}`);
    label.left=requested.headingCenters[i]-label.width/2;
    label.top=requested.headingCenterY-label.height/2;
  }
  if(subtitle){subtitle.left=requested.subtitle.left;subtitle.top=requested.subtitle.top;}
  await syncStage(context,'049副标题与系列标题定位');
  for(const label of [...headings,...(subtitle?[subtitle]:[])])label.load('left,top,width,height');
  if(c.title)chart.title.load('left,top,width,height');
  await syncStage(context,'049文字位置回读');
  const actual={title:c.title?rect(chart.title):null,subtitle:rect(subtitle),headings:headings.map(rect)},diagnostics={unit:'pt',requested,subtitleMetrics,resizedSubtitle,actual};
  const fail=(stage,message)=>{throw new UserError(`蝴蝶图[${stage}]：${message}；layout=${JSON.stringify(diagnostics)}`);};
  const validRect=r=>r&&Object.values(r).every(Number.isFinite)&&r.width>0&&r.height>0;
  const headingTop=headings.length?Math.min(...headings.map(l=>l.top)):c.height*100/702;
  if(subtitle&&(!validRect(actual.subtitle)||Object.keys(requested.subtitle).some(k=>Math.abs(subtitle[k]-requested.subtitle[k])>1)||subtitle.top+subtitle.height>headingTop-2))fail('049副标题回读','原生副标题宽高或位置不符，或侵入两侧标题区');
  if(actual.headings.some((r,i)=>!validRect(r)||Math.abs(r.left+r.width/2-requested.headingCenters[i])>1||Math.abs(r.top+r.height/2-requested.headingCenterY)>1))fail('049系列标题回读','原生两侧标题位置不符');
  if(c.title&&(!validRect(actual.title)||chart.title.top+chart.title.height>headingTop-2))fail('049主标题回读','主标题侵入系列标题区');
  return {font:CAREER_FONT,...diagnostics,visualVerified:false};
}
// Core calls this only after applyChartPresentation on create/restyle. Rebind
// must never call it: manual label positions and typography remain user-owned.
async function finalizePresentation({context,chart,binding}){
  if(binding.config.layout!=='career049')return null;
  return {textLayout:await placeCareerText(context,chart,binding.config),visualVerified:false};
}
async function rebind({context,chart,helper,source,instance,previous,binding}) {
  bindSeries(chart,helper,binding);await context.sync();
  linkLabels(chart,instance.helperName,source.count,binding.config,{oldCount:previous.count});
  linkSpecialPoints(chart,instance.helperName,source.count,binding.config,{oldCount:previous.count});await context.sync();
}
function linkSpecialPoints(chart,helperName,n,c,{initial=false,oldCount=0}={}){
  descriptors(c).forEach((d,i)=>{
    const s=chart.series.getItemAt(i);
    if(c.layout==='career049'&&d.role==='spacer'){
      // Never access spacer.dataLabels: collection formatting can turn on every
      // automatic numeric label. Rebinding must also suppress newly added points.
      s.hasDataLabels=false;
      for(let j=0;j<n;j++)s.points.getItemAt(j).hasDataLabel=false;
    }
    if(d.role==='careerText'){
      const p=s.points.getItemAt(0),subtitle=d.text==='subtitle';
      if(initial){
        s.dataLabels.showValue=false;s.dataLabels.showCategoryName=false;s.dataLabels.showSeriesName=false;
        s.dataLabels.showLeaderLines=false;s.dataLabels.textOrientation=0;
        s.dataLabels.horizontalAlignment=subtitle?'Left':'Center';
        p.hasDataLabel=true;p.dataLabel.position='Center';p.dataLabel.textOrientation=0;
        p.dataLabel.format.font.name=fontFor(c);p.dataLabel.format.font.size=10;p.dataLabel.format.font.bold=false;
        p.dataLabel.format.font.color=subtitle?'#000000':PALETTES[c.palette][d.text][0];
        // Mac Excel rejects point-label verticalAlignment here. The fixed
        // native label rectangle supplies placement without that setter.
      }
      p.dataLabel.formula=`=${quote(helperName)}!$A$${d.row}`;
    }
    if(d.role==='pinLine')for(let j=initial?0:oldCount;j<n;j++){const p=s.points.getItemAt(j*3+1);p.markerStyle='Circle';p.markerSize=7;p.markerBackgroundColor=PALETTES[c.palette][d.side][0];p.markerForegroundColor=PALETTES[c.palette][d.side][0];}
    if(d.role==='reference'){const p=s.points.getItemAt(1);if(initial){p.hasDataLabel=true;p.dataLabel.position=c.direction==='horizontal'?'Top':'Right';p.dataLabel.format.font.name=FONT;p.dataLabel.format.font.size=c.fontSize;}p.dataLabel.formula=`=${quote(helperName)}!$${d.labelCell[0]}$${d.labelCell.slice(1)}`;}
    if(d.role==='axisLabels')for(let j=0;j<11;j++){const p=s.points.getItemAt(j);if(initial){p.hasDataLabel=true;p.dataLabel.position=c.direction==='horizontal'?'Bottom':'Left';p.dataLabel.format.font.name=FONT;p.dataLabel.format.font.size=c.fontSize;}p.dataLabel.formula=`=${quote(helperName)}!$C$${212+j}`;}
  });
}

const choice=(value,label)=>({value,label});
const inputSpec={
  category:{id:'comparison',title:'比较'},
  description:'首行为表头：分类、左侧1–4列、右侧1–4列，可选中立项及末列独立计数；或4列：左分类、左值、右分类、右值。输入非负数字，整行留空可分组。公式自动联动，重绑保留格式。',
  minVersion:'1.19',columns:3,minRows:2,maxRows:MAX_ROWS,
  options:[
    {id:'title',label:'标题',type:'text',default:'蝴蝶图',maxLength:160},
    {id:'subtitle',label:'副标题',type:'text',default:'',maxLength:200},
    {id:'direction',label:'方向',type:'select',default:'horizontal',choices:[choice('horizontal','横向蝴蝶'),choice('vertical','纵向分歧柱')]},
    {id:'layout',label:'版式',type:'select',default:'balanced',choices:[choice('balanced','对称通用'),choice('career049','PDF049男女职业')]},
    {id:'leftSeries',label:'左侧系列数（纵向为下侧）',type:'number',default:1,min:1,max:4,step:1},
    {id:'leftOrder',label:'左侧输入列顺序',type:'select',default:'outside-in',choices:[choice('outside-in','从外侧到中心'),choice('inside-out','从中心到外侧')]},
    {id:'palette',label:'配色',type:'select',default:'career',choices:[choice('career','049粉蓝'),choice('ai','091红蓝分层'),choice('ai2021','093蓝青分层'),choice('survey','147/153红蓝'),choice('green','147/153棕绿'),choice('consumer','153黑灰蓝'),choice('jobs','151深浅蓝'),choice('mckinsey','154深蓝灰蓝'),choice('transport','054灰蓝点线')]},
    {id:'leftStyle',label:'左侧绘制',type:'select',default:'bar',choices:[choice('bar','实条'),choice('pin','细线圆点'),choice('dotted','虚线圆点')]},
    {id:'rightStyle',label:'右侧绘制',type:'select',default:'bar',choices:[choice('bar','实条'),choice('pin','细线圆点'),choice('dotted','虚线圆点')]},
    {id:'neutral',label:'中立项（独立计数前的末列）',type:'select',default:'none',choices:[choice('none','无中立项'),choice('center','以零轴居中分半'),choice('separate','右侧单独显示')]},
    {id:'neutralStart',label:'独立中立栏起点（原数值单位）',type:'number',default:100,min:0,max:1000000000000,step:1},
    {id:'net',label:'Net（右减左）',type:'select',default:'none',choices:[choice('none','不显示'),choice('column','单独标签列'),choice('marker','差额位置圆点')]},
    {id:'showNetLabels',label:'圆点显示推导Net数值',type:'checkbox',default:false},
    {id:'categoryPosition',label:'分类标签',type:'select',default:'center',choices:[choice('center','中间居中'),choice('left','中间靠左'),choice('outside','外侧'),choice('top','顶部一行（纵向）'),choice('value-end','正侧柱端（纵向）')]},
    {id:'countRow',label:'末列作为底部独立计数（纵向）',type:'checkbox',default:false},
    {id:'scale',label:'左右尺度',type:'select',default:'shared',choices:[choice('shared','共同尺度'),choice('independent','独立尺度（不同单位）')]},
    {id:'units',label:'数值格式',type:'select',default:'percent-points',choices:[choice('percent-points','17显示17%'),choice('fraction','0.17显示17%'),choice('number','普通数值'),choice('dollars','美元符号')]},
    {id:'decimals',label:'小数位',type:'number',default:0,min:0,max:4,step:1},
    {id:'rightDecimals',label:'右侧小数位（-1跟随通用）',type:'number',default:-1,min:-1,max:4,step:1},
    {id:'trimZeros',label:'省略末尾小数零',type:'checkbox',default:false},
    {id:'centerGap',label:'中心半间距（相对单侧最大值）',type:'number',default:.2972,min:0,max:.8,step:.01},
    {id:'gapFill',label:'中心间隔背景',type:'select',default:'none',choices:[choice('none','透明'),choice('gray','浅灰带')]},
    {id:'barGap',label:'行间距（%）',type:'number',default:50,min:0,max:500,step:1},
    {id:'fontSize',label:'字号',type:'number',default:10,min:6,max:24,step:1},
    {id:'width',label:'宽度（pt）',type:'number',default:700,min:300,max:1400,step:10},
    {id:'height',label:'高度（pt）',type:'number',default:429,min:220,max:1200,step:10},
    {id:'showValues',label:'显示两侧合计',type:'checkbox',default:true},
    {id:'valuePosition',label:'数值标签位置',type:'select',default:'outside',choices:[choice('outside','外侧合计'),choice('inside','段内居中')]},
    {id:'segmentLabels',label:'堆叠段数值',type:'checkbox',default:true},
    {id:'showBaselines',label:'显示内侧基线',type:'checkbox',default:true},
    {id:'showLeaders',label:'中心分类引导线',type:'checkbox',default:false},
    {id:'showAxisLabels',label:'显示数值刻度（中心间距为0）',type:'checkbox',default:false},
    {id:'showReferences',label:'显示左右参考线',type:'checkbox',default:false},
    {id:'leftReference',label:'左侧参考值',type:'number',default:37,min:0,max:1000000000000,step:.1},
    {id:'rightReference',label:'右侧参考值',type:'number',default:3.9,min:0,max:1000000000000,step:.1},
    {id:'showLegend',label:'显示图例',type:'checkbox',default:true}
  ],storage:{marker:MARK,helperPrefix:HELPER_PREFIX,chartPrefix:CHART_PREFIX}
};
export {validate,config,helperPlan,descriptors,MARK,HELPER_PREFIX,CHART_PREFIX,COLORS,FONT};
export default {id:'butterfly',title:'蝴蝶图',inputSpec,validate,writeHelper,createChart,style,finalizePresentation,rebind};
