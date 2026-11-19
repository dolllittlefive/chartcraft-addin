// F015/F076: native stacked-area rectangles on a date axis; all geometry is cell-formula driven.
// Dates encode continuous coordinates at 1e-6 resolution. No shapes, VBA, event handlers or image output.
import {UserError} from '../core/errors.js';
import {setChartTitleTypography} from '../core/presentation.js';

export const MARK = 'ChartCraft.VariableWidthBar.v1';
export const HELPER_PREFIX = 'CC_不等宽_';
export const CHART_PREFIX = 'CC_VariableWidth_';
export const FONT = '微软雅黑';
export const MAX_ROWS = 40;
export const DATE_ORIGIN = 40000;
export const DATE_SPAN = 1000000;
export const AREA_ROW = 50;
export const FIRST_BODY_SERIES = 5;
export const TICK_COUNT = 11;
export const COORDINATE_VERSION = 'date-x-v2';
const GRID_COORDINATE_VERSION = 'date-x-v2-grid-v2';
const usesCpiGrid = o => o.profile==='pdf171'&&o.direction==='vertical'&&o.variant==='single';
const CPI_SUBTITLE = '2023年7月消费者价格指数（CPI-U）同比变化。部分主要支出类别*';
const CPI_REFERENCE_INDEX = 2*TICK_COUNT;
const CPI_AXIS_NOTE_INDICES = {y:CPI_REFERENCE_INDEX+2,width:CPI_REFERENCE_INDEX+4};
const CPI_LEFT_BAND = 24;
const CPI_WIDTH_NOTE_WIDTH = 220;
const CPI_BOTTOM_BAND = 20;
export const COLORS = ['#002B49', '#007CB0', '#00A3AD', '#78BE20', '#F0AB00', '#C04C36', '#7C6992', '#A7A9AC'];
export const LAYOUT = {width: 720, height: 445, plot: {left: 68, top: 64, width: 612, height: 302}};
export function layoutFor(o) {
  if (o.profile !== 'pdf171') return {...LAYOUT,font:10,titleFont:14};
  const horizontal=o.direction==='horizontal', scale=horizontal ? 196.48/448 : 419.42/710;
  const bounds=horizontal ? [59,113,349,376] : [83,71,612,284];
  // Reserve a fixed band on both hosts, preserving the original plot dimensions.
  const header=usesCpiGrid(o)&&o.subtitle?80:0;
  const leftBand=usesCpiGrid(o)?CPI_LEFT_BAND:0;
  // Mac preserves the measured 37.31pt caption height after setWidth.
  // Add bottom whitespace for 10pt text; keep the original plot dimensions.
  const bottomBand=usesCpiGrid(o)?CPI_BOTTOM_BAND:0;
  return {width:(horizontal?196.48:419.42)+leftBand,height:231.57+header+bottomBand,font:10,titleFont:14,
    plot:{left:bounds[0]*scale+leftBand,top:bounds[1]*scale+header,width:bounds[2]*scale,height:bounds[3]*scale}};
}
const quote = value => `'${String(value).replace(/'/g, "''")}'`;
const literal = value => "'" + String(value);
export function columnLetter(index) { let out = '', n = index + 1; while (n) { out = String.fromCharCode(65 + (n - 1) % 26) + out; n = Math.floor((n - 1) / 26); } return out; }
const fmt = value => String(value).replace(/"/g, '""');
const rangeEnd = (row, count) => row + count - 1;

export function validate(values) {
  if (!Array.isArray(values) || values.length < 3 || !Array.isArray(values[0])) throw new UserError('请选择含表头、至少两行数据的区域。');
  const columns = values[0].length, count = values.length - 1;
  if (columns < 3 || columns > 8) throw new UserError('需要3至8列：分类、宽度权重、数值；堆积变体可追加最多6个数值列。');
  if (count > MAX_ROWS) throw new UserError(`本配方最多支持${MAX_ROWS}个分类。`);
  if (values[0].some(v => v === '' || v === null || v === undefined)) throw new UserError('每列都需要非空表头。');
  let totalWidth = 0;
  const labels = new Set();
  values.slice(1).forEach((row, index) => {
    if (!Array.isArray(row) || row.length !== columns) throw new UserError(`第${index + 2}行列数不一致。`);
    if (row[0] === '' || row[0] === null || row[0] === undefined || /^#(?:VALUE!|REF!|N\/A|DIV\/0!|NUM!|NAME\?|SPILL!)/.test(String(row[0]))) throw new UserError(`第${index + 2}行分类为空或为错误值。`);
    if (labels.has(String(row[0]))) throw new UserError(`分类“${row[0]}”重复；请使用唯一名称，以便保留分类格式。`);
    labels.add(String(row[0]));
    row.slice(1).forEach((v, offset) => { if (typeof v !== 'number' || !Number.isFinite(v)) throw new UserError(`第${index + 2}行第${offset + 2}列必须是有限数字。`); });
    if (row[1] < 0) throw new UserError(`第${index + 2}行宽度权重不能为负数。`);
    totalWidth += row[1];
  });
  if (!(totalWidth > 0) || !Number.isFinite(totalWidth)) throw new UserError('宽度权重合计必须为有限正数。');
  return {count, columns, measureCount: columns - 2, categoryHeader: String(values[0][0]), widthHeader: String(values[0][1]), valueHeader: String(values[0][2]), measureHeaders: values[0].slice(2).map(String), totalWidth};
}

export function normalizeOptions(options = {}) {
  const o = {title: '不等宽柱与条形', subtitle:'', direction: 'vertical', variant: 'single', gap: 0,
    showWidth: true, showValues: true, showLegend: false, axisSymmetric: false,
    valuePlacement: 'outside', colorMode: 'category', profile:'standard', showPlus:false, valueFormat: '0.0', widthFormat: '0.0',
    showReference:true, referenceValue:.032, referenceLabel:'All items CPI-U:', referenceCell:'', ...options};
  if (typeof o.title !== 'string' || o.title.length > 120) throw new UserError('标题必须是120字以内的文本。');
  if (typeof o.subtitle !== 'string' || o.subtitle.length > 160) throw new UserError('副标题必须是160字以内的文本。');
  if(usesCpiGrid(o)&&!o.subtitle)o.subtitle=CPI_SUBTITLE;
  if (typeof o.referenceValue!=='number'||!Number.isFinite(o.referenceValue)) throw new UserError('参考指标必须是有限数字，百分比按小数保存。');
  if (typeof o.referenceLabel!=='string'||o.referenceLabel.length>60||/[\r\n]/.test(o.referenceLabel)) throw new UserError('参考指标名称必须是60字以内的单行文本。');
  if (typeof o.referenceCell!=='string') throw new UserError('参考指标单元格必须是同一源表的单个地址，例如K2。');
  o.referenceCell=o.referenceCell.trim().toUpperCase();
  if(o.referenceCell){
    const match=/^\$?([A-Z]{1,3})\$?([1-9]\d{0,6})$/.exec(o.referenceCell);
    const column=match&&[...match[1]].reduce((n,c)=>n*26+c.charCodeAt(0)-64,0);
    if(!match||column>16384||Number(match[2])>1048576) throw new UserError('参考指标单元格必须是同一源表的单个有效地址，例如K2。');
    o.referenceCell=`$${match[1]}$${match[2]}`;
  }
  if (!['vertical', 'horizontal'].includes(o.direction)) throw new UserError('方向必须为纵向或横向。');
  if (!['single', 'stacked', 'percent', 'bubble'].includes(o.variant)) throw new UserError('未知不等宽变体。');
  if (typeof o.gap !== 'number' || !Number.isFinite(o.gap) || o.gap < 0 || o.gap > 0.3) throw new UserError('分类间隙必须在0%至30%之间。');
  for (const key of ['showWidth', 'showValues', 'showLegend', 'axisSymmetric', 'showPlus', 'showReference']) if (typeof o[key] !== 'boolean') throw new UserError(`${key}必须为布尔值。`);
  if (!['outside','inside'].includes(o.valuePlacement) || !['category','sign'].includes(o.colorMode)) throw new UserError('未知标签位置或配色选项。');
  if (!['standard','pdf171'].includes(o.profile)) throw new UserError('未知原例版式。');
  for (const key of ['valueFormat', 'widthFormat']) if (!['0', '0.0', '0.00', '0%', '0.0%', '0.00%'].includes(o[key])) throw new UserError('数值格式仅支持0、0.0、0.00及对应百分比格式。');
  return o;
}

function validateVariant(source, o, capacity) {
  const dataMeasures = source.measureCount ?? source.columnCount - 2;
  const single=['single','bubble'].includes(o.variant), measures=o.variant==='bubble'?1:dataMeasures;
  if (o.variant === 'single' && measures !== 1) throw new UserError('普通不等宽图需要正好3列；多数值列请选择堆积变体，不能静默忽略列。');
  if (o.variant==='bubble' && (dataMeasures!==2 || o.direction!=='vertical' || source.values.slice(1).some(r=>r[3]<0))) throw new UserError('气泡组合目前要求纵向4列：分类、宽度、高度、非负气泡大小；未使用静态标记替代。');
  if (!single && source.values.slice(1).some(row => row.slice(2).some(v => v < 0))) throw new UserError('堆积变体不接受负分段；带正负值的数据请使用普通变体。');
  if (o.variant === 'percent' && source.values.slice(1).some(row => row.slice(2).reduce((a,b) => a+b, 0) <= 0)) throw new UserError('百分比堆积的每个分类数值合计必须大于0。');
  if (o.direction === 'vertical' && source.values.slice(1).some(row => row[1] > 0 && row[1] / source.totalWidth * (1 - o.gap) < 1 / DATE_SPAN)) throw new UserError('存在小于日期轴分辨率的正宽度；请合并极小分类。未生成等宽替代图。');
  const segments = single && o.colorMode === 'sign' ? 2 : measures;
  const stride = segments + (o.direction === 'horizontal' ? 1 : 0);
  if (FIRST_BODY_SERIES + (o.variant==='bubble'?1:0) + capacity * stride > 240) throw new UserError('分类与堆积分段组合过多；请缩小选区，使原生系列数不超过240。');
  return measures;
}

/** Pure, auditable formula plan. Chart data never comes from JavaScript numeric snapshots. */
export function formulaPlan(source, options, capacity = source.count) {
  if (!Number.isInteger(capacity) || capacity < source.count || capacity > MAX_ROWS) throw new UserError('辅助表分类容量无效，未修改原图。');
  const o = normalizeOptions(options), measures = validateVariant(source, o, capacity);
  const n = source.count, last = n + 2, sheet = quote(source.sheet), single = ['single','bubble'].includes(o.variant);
  const dataMeasures=source.measureCount ?? source.columnCount-2, firstBody=FIRST_BODY_SERIES+(o.variant==='bubble'?1:0);
  const src = (i, c) => `${sheet}!$${columnLetter(source.columnIndex + c)}$${source.rowIndex + i + 2}`;
  const numeric = ref => `=IFERROR(IF(ISNUMBER(${ref}),${ref},0),0)`;
  const table = [];
  for (let i = 0; i < n; i++) {
    const r = i + 3, raw = Array.from({length: dataMeasures}, (_, j) => src(i, j + 2));
    const labelValue=single ? `C${r}` : o.variant==='percent' ? '1' : `I${r}`;
    const valid = [`${src(i,0)}<>""`, `ISNUMBER(${src(i,1)})`, `${src(i,1)}>=0`, ...raw.map(v => `ISNUMBER(${v})`)];
    if (!single) valid.push(...raw.map(v => `${v}>=0`));
    if (o.variant==='bubble') valid.push(`${raw[1]}>=0`);
    if (o.variant === 'percent') valid.push(`SUM(${raw.join(',')})>0`);
    if (o.direction === 'vertical') valid.push(`OR(B${r}=0,IFERROR(B${r}/$E$1*(1-${o.gap}),0)>=1/${DATE_SPAN})`);
    const row = [`=IFERROR(${src(i,0)}&"","")`, numeric(src(i,1)), ...raw.map(numeric)];
    while (row.length < 8) row.push('=0');
    row.push(
      `=SUM(C${r}:H${r})`, // I raw total
      single ? `=MIN(0,C${r})` : '=0', // J raw rectangle lower value
      single ? `=ABS(C${r})` : o.variant === 'percent' ? `=IF(I${r}>0,1,0)` : `=I${r}`, // K extent
      i === 0 ? '=0' : `=IFERROR(SUM($B$3:B${r - 1})/$E$1,0)`, // L cumulative start
      `=IFERROR(SUM($B$3:B${r})/$E$1,0)`, // M cumulative end
      `=(L${r}+M${r})/2`, // N category centre
      `=IFERROR(IF(AND(${valid.join(',')}),1,0),0)`, // O valid source, including later user edits
      `=IF($F$1,IF(B${r}>0,A${r}${o.showWidth ? `&CHAR(10)&TEXT(B${r},"${fmt(o.widthFormat)}")` : ''},""),${i === 0 ? '"数据无效：请检查源数据"' : '""'})`,
      `=IF(AND($F$1,B${r}>0),${o.showPlus?`IF(${labelValue}>0,"+","")&`:''}TEXT(${labelValue},"${fmt(o.valueFormat)}"),"")`,
      `=IF(AND($F$1,B${r}>0),TEXT(B${r},"${fmt(o.widthFormat)}"),"")`,
      `=IFERROR((J${r}-$C$1)/$G$1,0)`, // S lower normalized value
      `=IFERROR((${single ? `C${r}` : `K${r}`}-$C$1)/$G$1,0)`, // T end normalized value
      '=IFERROR(-$C$1/$G$1,0)', // U zero normalized value
      `=IFERROR(B${r}/$E$1,0)` // V category weight fraction
    );
    table.push(row);
  }
  let rawMin = single ? `MIN(0,$C$3:$C$${last})` : '0';
  let rawMax = single ? `MAX(0,$C$3:$C$${last})` : `MAX(0,$K$3:$K$${last})`;
  if (o.axisSymmetric) { const extent=`MAX(ABS(${rawMin}),ABS(${rawMax}))`; rawMin=`-${extent}`; rawMax=extent; }
  const rawSpan = `(${rawMax}-${rawMin})`, power = `10^INT(LOG10(${rawSpan}/5))`, unit = `(${rawSpan}/5/${power})`;
  const stat = {
    B1: capacity, C1: `=IFERROR(INT(${rawMin}/$H$1)*$H$1,0)`,
    D1: `=IFERROR(IF(${rawMax}=${rawMin},$C$1+1,-INT(-${rawMax}/$H$1)*$H$1),1)`,
    E1: `=SUM($B$3:$B$${last})`, F1: `=IFERROR(AND(COUNTIF($O$3:$O$${last},1)=${n},$E$1>0),FALSE)`,
    G1: '=MAX(0.000000000001,$D$1-$C$1)',
    H1: `=IFERROR(IF(${rawSpan}>0,${power}*IF(${unit}<=1,1,IF(${unit}<=2,2,IF(${unit}<=5,5,10))),1),1)`,
    N1: literal(o.direction), O1: dataMeasures, P1: n, Q1: literal(o.variant), AZ2:literal(usesCpiGrid(o)?GRID_COORDINATE_VERSION:COORDINATE_VERSION)
  };
  if(usesCpiGrid(o)) {
    // The original 3.2% is an external all-items indicator, not an average of
    // measured category widths. A user-selected source cell stays formula-live.
    stat.AU1=o.referenceCell?`=IFERROR(IF(ISNUMBER(${sheet}!${o.referenceCell}),${sheet}!${o.referenceCell},""),"")`:o.referenceValue;
    stat.AV1=literal(o.referenceLabel);stat.AW1=literal(o.referenceCell);stat.AX1=o.showReference;
  }
  const endpoints = [['=0']];
  for (let i = 0; i < n; i++) for (let j = 0; j < measures; j++) {
    const r = i + 3, componentEnd = `SUM($C${r}:${columnLetter(2+j)}${r})`;
    endpoints.push([single ? `=C${r}` : o.variant === 'percent' ? `=IFERROR(${componentEnd}/I${r},0)` : `=${componentEnd}`]);
  }
  const segments = single && o.colorMode === 'sign' ? 2 : measures;
  const stride = segments + (o.direction === 'horizontal' ? 1 : 0), area = [];
  const bodyColumns = capacity * stride;
  const height = (i,j) => {
    if (i >= n) return '0';
    const r = i+3, c = columnLetter(2+j);
    return single ? (segments === 2 ? `IF(C${r}${j===0?'>=':'<'}0,ABS(C${r})/$G$1,0)` : `ABS(C${r})/$G$1`) : o.variant === 'percent' ? `IFERROR(${c}${r}/I${r},0)/$G$1` : `${c}${r}/$G$1`;
  };
  if (o.direction === 'vertical') {
    for (let i = 0; i < n; i++) for (let k = 0; k < 4; k++) {
      const r = i+3, x = k < 2 ? `L${r}+V${r}*${o.gap}/2` : `M${r}-V${r}*${o.gap}/2`;
      const row = [`=${DATE_ORIGIN}+ROUND((${x})*${DATE_SPAN},0)`, `=IF($F$1,S${r},0)`];
      for (let c = 0; c < bodyColumns; c++) row.push(`=IF($F$1,${Math.floor(c / segments) === i && (k === 1 || k === 2) ? height(i,c % segments) : '0'},0)`);
      area.push(row);
    }
  } else {
    // The sorted endpoints and duplicate left/right limits create exact steps, never sloped ramps.
    const points = [{raw: '$C$1', side: 1}];
    for (let k = 1; k <= endpoints.length; k++) for (let side = 0; side <= 1; side++) points.push({raw:`SMALL($ZZ$3:$ZZ$${endpoints.length+2},${k})`, side});
    points.push({raw: '$D$1', side: 0});
    for (const point of points) {
      const row = [`=${DATE_ORIGIN}+ROUND((${point.raw}-$C$1)/$G$1*${DATE_SPAN},0)`, '=0'];
      for (let i = 0; i < capacity; i++) {
        const r = i+3, active = [];
        for (let j = 0; j < segments; j++) {
          let lo = single ? `MIN(0,C${r})` : j === 0 ? '0' : `SUM($C${r}:${columnLetter(1+j)}${r})`;
          let hi = single ? `MAX(0,C${r})` : `SUM($C${r}:${columnLetter(2+j)}${r})`;
          if (o.variant === 'percent') { lo = `IFERROR(${lo}/I${r},0)`; hi = `IFERROR(${hi}/I${r},0)`; }
          const within = `AND(${point.raw}${point.side ? '>=' : '>'}${lo},${point.raw}${point.side ? '<' : '<='}${hi})`;
          const sign = single && segments===2 ? `,C${r}${j===0?'>=':'<'}0` : '';
          const draw = i < n ? `IF(AND($F$1,${within}${sign}),V${r}*(1-${o.gap}),0)` : '0';
          active.push(draw); row.push(`=${draw}`);
        }
        // Transparent remainder keeps each category in its own cumulative-weight slot at every X.
        row.push(i < n ? `=IF($F$1,MAX(0,V${r}-(${active.join('+')})),0)` : '=0');
      }
      area.push(row);
    }
  }
  // Fixed label-series indices survive row growth; body slots are appended, never reordered or restyled.
  // Keep every scatter X in the same date-number domain as the area chart. Mac's hidden
  // secondary-axis combination omitted the old 0..1-X labels while its data-label API still returned text.
  const dateX = expression => `=${DATE_ORIGIN}+(${expression})*${DATE_SPAN}`;
  const labels = table.map((_,i) => {
    const r = i+3;
    const value = o.valuePlacement==='inside' ? `S${r}+K${r}/$G$1/2` : `T${r}+IF(${single?`C${r}`:'1'}<0,-0.025,0.025)`;
    // PDF171's two narrow adjacent positive bars need a small native-label gap.
    // Three original-image pixels each side; positions still follow live weighted midpoints.
    const cpiNudge=o.profile==='pdf171' && o.variant==='single' && o.valuePlacement==='outside';
    const valueX=`N${r}${cpiNudge && i===3?'-3/612':cpiNudge && i===4?'+3/612':''}`;
    return o.direction === 'vertical'
      ? [dateX(`N${r}`), `=U${r}`, dateX(valueX), `=IF($F$1,${value},0)`]
      : [dateX(`U${r}`), `=1-L${r}-V${r}*(1-${o.gap})/2`, dateX(`IF($F$1,${value},0)`), `=1-L${r}-V${r}*(1-${o.gap})/2`];
  });
  let guide = o.direction === 'vertical' ? [[dateX('0'),'=IFERROR(-$C$1/$G$1,0)'],[dateX('1'),'=IFERROR(-$C$1/$G$1,0)']]
    : [[dateX('IFERROR(-$C$1/$G$1,0)'),'=0'],[dateX('IFERROR(-$C$1/$G$1,0)'),'=1']];
  if(usesCpiGrid(o)) {
    // The original zero line has a short tick at every live weighted midpoint.
    // Retrace each tick in one native line series; no detached drawing objects.
    const zero='IFERROR(-$C$1/$G$1,0)', depth=2.5/layoutFor(o).plot.height;
    guide=[guide[0],...table.flatMap((_,i)=>[
      [dateX(`N${i+3}`),`=${zero}`],
      [dateX(`N${i+3}`),`=${zero}-${depth}`],
      [dateX(`N${i+3}`),`=${zero}`]
    ]),guide[1]];
  }
  // Nice raw-value steps remain formula driven when source values change. Unused points have empty labels.
  const ticks = Array.from({length:TICK_COUNT},(_,i) => {
    const position=`MIN(1,${i}*$H$1/$G$1)`;
    const label=`=IF(AND($F$1,${i}*$H$1<=$G$1+0.000000000001),TEXT($C$1+${i}*$H$1,"${fmt(o.valueFormat)}"),"")`;
    return o.direction==='vertical' ? [dateX('0'),`=${position}`,label] : [dateX(position),'=0',label];
  });
  const tickStride=usesCpiGrid(o)?2:1;
  // Masked paths can exceed AREA_ROW; keep their columns beyond all 40 single
  // category body slots, so later row growth cannot overwrite either recipe.
  const tickColumns=tickStride===2?{x:'EA',y:'EB',label:'EC'}:{x:'AL',y:'AM',label:'AN'};
  // Keep the eleven label anchors first, each followed by an NA separator.
  // Stable indices survive added rows; the masked grid geometry follows them.
  let tickPath=ticks;
  if(tickStride===2) {
    const blank=()=>['=NA()','=NA()','=""'];
    tickPath=ticks.flatMap(([x,y,label],i)=>{
      const active=`AND($F$1,${i}*$H$1<=$G$1+0.000000000001)`;
      return [[`=IF(${active},${x.slice(1)},NA())`,`=IF(${active},${y.slice(1)},NA())`,label],blank()];
    });
    // One fixed anchor plus a separator: annotation point22 never moves when
    // row count or gap changes. Reuse the native grid series, without adding a
    // detached shape or changing any body-series index/capacity.
    tickPath.push([
      `=IF(AND($F$1,$AX$1),${DATE_ORIGIN}+0.2*${DATE_SPAN},NA())`,
      '=IF(AND($F$1,$AX$1),0.2,NA())',
      '=IF(AND($F$1,$AX$1),$AV$1&CHAR(10)&IF(ISNUMBER($AU$1),IF($AU$1>0,"▲+",IF($AU$1<0,"▼",""))&TEXT($AU$1,"0.0%"),"参考指标无效"),"")'
    ],blank());
    // Native formula labels, separated from the grid: retain point22 and all
    // body-series indices. The width caption keeps its original image centre.
    tickPath.push([dateX('0'),'=0.5','="YoY"'],blank(),
      [dateX('(318-83)/612'),'=0','="在 CPI-U 中的相对权重"'],blank());
    ticks.forEach(([,y],i)=>table.forEach((_,j)=>{
      const r=j+3,level=`($C$1+${i}*$H$1)`;
      const active=`AND($F$1,${i}*$H$1<=$G$1+0.000000000001,B${r}>0)`;
      const blocked=`AND(C${r}<>0,${level}>=MIN(0,C${r}),${level}<=MAX(0,C${r}))`;
      const date=expr=>`${DATE_ORIGIN}+ROUND((${expr})*${DATE_SPAN},0)`;
      const left=date(`L${r}`),right=date(`M${r}`);
      const bodyLeft=date(`L${r}+V${r}*${o.gap}/2`),bodyRight=date(`M${r}-V${r}*${o.gap}/2`);
      const segment=(a,b,condition)=>{
        const show=`AND(${active},${condition},(${b})>(${a}))`;
        tickPath.push([`=IF(${show},${a},NA())`,`=IF(${show},${y.slice(1)},NA())`,'=""'],
          [`=IF(${show},${b},NA())`,`=IF(${show},${y.slice(1)},NA())`,'=""'],blank());
      };
      // Clear rows retain the whole category slot; occupied rows retain only
      // left/right gap slivers. Endpoints use the exact area date quantization.
      segment(left,`IF(${blocked},${bodyLeft},${right})`,'TRUE');
      if(o.gap>0)segment(bodyRight,right,blocked);
    }));
  }
  const bubbles=table.map((_,i)=>{const r=i+3;return [dateX(`N${r}`),'=-0.23',`=IF($F$1,D${r},0)`,`=IF($F$1,TEXT(D${r},"0"),"")`];});
  return {options:o, count:n, capacity, measures, dataMeasures, firstBody, segments, stride, bodyColumns, areaLast:rangeEnd(AREA_ROW,area.length),
    stat, table, area, endpoints, labels, guide, ticks, tickPath, tickStride, tickColumns, bubbles,
    referenceIndex:usesCpiGrid(o)?CPI_REFERENCE_INDEX:null,
    axisNoteIndices:usesCpiGrid(o)?{...CPI_AXIS_NOTE_INDICES}:null};
}

export async function writeHelper({context,helper,source,options,previous}) {
  let capacity = source.count;
  if (previous) {
    const old = helper.getRange('B1'), oldMeasures = helper.getRange('O1'), coordinateVersion=helper.getRange('AZ2'); old.load('values'); oldMeasures.load('values'); coordinateVersion.load('values'); await context.sync();
    const expectedVersion=usesCpiGrid(normalizeOptions(options))?GRID_COORDINATE_VERSION:COORDINATE_VERSION;
    if (coordinateVersion.values[0][0]!==expectedVersion) throw new UserError('该图使用旧版标签坐标，请用原数据新建修正版图表；原图和手工格式未修改。');
    if (!Number.isInteger(old.values[0][0]) || old.values[0][0]<previous.count || oldMeasures.values[0][0]!==source.measureCount) throw new UserError('原辅助表容量或数值列数不匹配；增加/删除数值列请新建图表，原图未修改。');
    capacity = Math.max(capacity,old.values[0][0]);
  }
  const binding = formulaPlan(source, options, capacity);
  if (previous) {
    const used = helper.getUsedRange(); used.load('rowCount,columnCount'); await context.sync();
    // Only recipe rows; core's A1/I1:M1/R1 metadata and all Chart formatting remain intact.
    if (used.rowCount > 1) helper.getRangeByIndexes(1,0,used.rowCount-1,Math.max(702,used.columnCount)).clear(Excel.ClearApplyTo.contents);
  }
  for (const [address,value] of Object.entries(binding.stat)) helper.getRange(address).formulas = [[value]];
  helper.getRange('A2:V2').values = [['分类','宽度','值1','值2','值3','值4','值5','值6','合计','下界','幅度','累计起点','累计终点','中点','有效','分类标签','值标签','宽标签','下界位置','端点位置','零点位置','宽度占比']];
  helper.getRange('C2:H2').formulas = [Array.from({length:6},(_,j)=>j<binding.dataMeasures
    ? `=IFERROR(${quote(source.sheet)}!$${columnLetter(source.columnIndex+j+2)}$${source.rowIndex+1}&"","")` : '=""')];
  helper.getRange(`A3:V${source.count+2}`).formulas = binding.table;
  helper.getRange(`AD3:AG${source.count+2}`).formulas = binding.labels;
  helper.getRange(`AI3:AJ${binding.guide.length+2}`).formulas = binding.guide;
  helper.getRange(`${binding.tickColumns.x}3:${binding.tickColumns.label}${binding.tickPath.length+2}`).formulas = binding.tickPath;
  if (binding.options.variant==='bubble') helper.getRange(`AP3:AS${source.count+2}`).formulas=binding.bubbles;
  helper.getRange(`ZZ3:ZZ${binding.endpoints.length+2}`).formulas = binding.endpoints;
  helper.getRangeByIndexes(AREA_ROW-1,0,binding.area.length,binding.bodyColumns+2).formulas = binding.area;
  return binding;
}

function bindOne(series, helper, x, y) { series.setXAxisValues(helper.getRange(x)); series.setValues(helper.getRange(y)); }
function hideLine(series) { series.format.line.lineStyle = 'None'; series.markerStyle = 'None'; }
function bodyName(slot, binding, source) {
  const category = Math.floor(slot / binding.stride), segment = slot % binding.stride;
  return segment === binding.segments ? `间隔${category+1}` : `${source.values[category+1]?.[0] ?? `保留${category+1}`} · ${source.measureHeaders?.[segment] ?? `值${segment+1}`}`;
}
function setBodyDefaults(series, slot, binding) {
  series.chartType = 'AreaStacked'; series.axisGroup = 'Primary'; series.format.line.lineStyle = 'None';
  const segment = slot % binding.stride, category = Math.floor(slot / binding.stride);
  if (segment === binding.segments) series.format.fill.clear();
  else series.format.fill.setSolidColor(['single','bubble'].includes(binding.options.variant) && binding.options.colorMode==='sign'
    ? ['#0A1C2C','#40A1ED'][segment] : COLORS[(['single','bubble'].includes(binding.options.variant) ? category : segment) % COLORS.length]);
}
async function bindAxisNotes({context,chart,instance,binding},initialize) {
  if(!binding.axisNoteIndices)return;
  for(const [role,index] of Object.entries(binding.axisNoteIndices)) {
    const point=chart.series.getItemAt(4).points.getItemAt(index);
    if(!initialize){point.load('hasDataLabel');await context.sync();}
    // A legacy data-only rebind must preserve its existing dimensions. Its
    // missing notes are added by explicit style refresh, which adds left space.
    if(!initialize&&point.hasDataLabel!==true)continue;
    const label=point.dataLabel;
    label.formula=`=${quote(instance.helperName)}!$${binding.tickColumns.label}$${index+3}`;
    if(initialize) {
      point.hasDataLabel=true;label.position='Center';label.textOrientation=role==='y'?-90:0;
      label.format.font.name=FONT;label.format.font.size=10;label.format.font.color='#000000';
      label.format.fill.clear();label.format.border.lineStyle='None';
    }
  }
}
async function restrictCpiGridLabels({context,chart,binding}) {
  if(!usesCpiGrid(binding.options))return;
  const grid=chart.series.getItemAt(4);
  const roles=new Set(Array.from({length:TICK_COUNT},(_,i)=>i*binding.tickStride));
  if(binding.options.showReference)roles.add(binding.referenceIndex);
  for(const index of Object.values(binding.axisNoteIndices||{}))roles.add(index);
  const suppressed=[];
  for(let index=0;index<binding.tickPath.length;index++)if(!roles.has(index)) {
    const point=grid.points.getItemAt(index);point.hasDataLabel=false;suppressed.push({index,point});
  }
  // E's saved XML had series-level showVal=1 with only 14 explicit labels:
  // every other grid point inherited its numeric Y value as a visible label.
  // Explicit formula labels retain their point settings; disable the fallback.
  grid.dataLabels.showValue=false;
  await context.sync();
  grid.dataLabels.load('showValue');
  for(const {point} of suppressed)point.load('hasDataLabel');
  await context.sync();
  const visible=suppressed.filter(({point})=>point.hasDataLabel!==false).map(({index})=>index);
  if(grid.dataLabels.showValue!==false||visible.length)throw new UserError('171网格辅助点标签未关闭：'+JSON.stringify({showValue:grid.dataLabels.showValue,unexpectedPoints:visible}));
}

async function bindChart({context, chart, helper, binding, source, instance}, isCreate) {
  chart.series.load('items'); await context.sync();
  const previousSeries = chart.series.items.length;
  const needed = binding.firstBody + binding.bodyColumns;
  for (let i = previousSeries; i < needed; i++) {
    const slot = i-binding.firstBody, series = chart.series.add(bodyName(slot,binding,source));
    setBodyDefaults(series,slot,binding);
    // A native series name formula keeps category text live, including while the add-in is closed.
    if (slot%binding.stride<binding.segments) {
      const single=['single','bubble'].includes(binding.options.variant), segment=slot%binding.stride;
      series.name=single && binding.options.colorMode==='sign' ? ['正值','负值'][segment]
        : single ? `=${quote(instance.helperName)}!$A$${Math.floor(slot/binding.stride)+3}`
        : `=${quote(instance.helperName)}!$${columnLetter(segment+2)}$2`;
    }
  }
  await context.sync();
  const s = i => chart.series.getItemAt(i), last = source.count+2;
  bindOne(s(0),helper,`A${AREA_ROW}:A${binding.areaLast}`,`B${AREA_ROW}:B${binding.areaLast}`);
  bindOne(s(1),helper,`AD3:AD${last}`,`AE3:AE${last}`);
  bindOne(s(2),helper,`AF3:AF${last}`,`AG3:AG${last}`);
  bindOne(s(3),helper,`AI3:AI${binding.guide.length+2}`,`AJ3:AJ${binding.guide.length+2}`);
  bindOne(s(4),helper,`${binding.tickColumns.x}3:${binding.tickColumns.x}${binding.tickPath.length+2}`,`${binding.tickColumns.y}3:${binding.tickColumns.y}${binding.tickPath.length+2}`);
  if (binding.options.variant==='bubble') {
    bindOne(s(5),helper,`AP3:AP${last}`,`AQ3:AQ${last}`); s(5).setBubbleSizes(helper.getRange(`AR3:AR${last}`));
  }
  for (let i = 0; i < binding.bodyColumns; i++) {
    const col = columnLetter(i+2); bindOne(s(binding.firstBody+i),helper,`A${AREA_ROW}:A${binding.areaLast}`,`${col}${AREA_ROW}:${col}${binding.areaLast}`);
  }
  await context.sync();
  const previousCount = isCreate ? 0 : Math.max(0, Number(instance.previousCount) || source.count);
  for (let i = 0; i < source.count; i++) for (const [series,col] of [[1,'P'],[2,'Q']]) {
    const point = s(series).points.getItemAt(i);
    if (isCreate || i >= previousCount) {
      point.hasDataLabel = series === 1 || binding.options.showValues;
      const label = point.dataLabel;
      const negative = ['single','bubble'].includes(binding.options.variant) && source.values[i+1][2]<0;
      label.position = series === 1 ? (binding.options.direction === 'vertical' ? (negative?'Top':'Bottom') : (negative?'Right':'Left')) : 'Center';
      // Native upright vertical Chinese text, as in the source image; energy stays horizontal.
      if (series===1 && binding.options.profile==='pdf171' && binding.options.direction==='vertical' && binding.options.variant==='single' && (i===5 || i===6)) label.textOrientation=180;
      label.format.font.name = FONT; label.format.font.size = layoutFor(binding.options).font;
      label.format.font.color = series === 2 && binding.options.valuePlacement==='inside' ? '#FFFFFF' : '#333333';
    }
    if (series === 1 || binding.options.showValues) point.dataLabel.formula = `=${quote(instance.helperName)}!$${col}$${i+3}`;
  }
  if (binding.options.variant==='bubble') for (let i=0;i<source.count;i++) {
    const point=s(5).points.getItemAt(i);
    if (isCreate || i>=previousCount) { point.hasDataLabel=true; point.dataLabel.position='Center'; point.dataLabel.format.font.name=FONT; point.dataLabel.format.font.size=8; }
    point.dataLabel.formula=`=${quote(instance.helperName)}!$AS$${i+3}`;
  }
  if (isCreate) for (let i=0;i<TICK_COUNT;i++) {
    const point = s(4).points.getItemAt(i*binding.tickStride); point.hasDataLabel = true;
    point.dataLabel.formula = `=${quote(instance.helperName)}!$${binding.tickColumns.label}$${i*binding.tickStride+3}`;
    point.dataLabel.position = binding.options.direction === 'vertical' ? 'Left' : 'Bottom';
    point.dataLabel.format.font.name = FONT; point.dataLabel.format.font.size = layoutFor(binding.options).font; point.dataLabel.format.font.color = '#666666';
  }
  if(binding.referenceIndex!==null) {
    const point=s(4).points.getItemAt(binding.referenceIndex);
    if(!isCreate){point.load('hasDataLabel');await context.sync();}
    const existing=isCreate?false:point.hasDataLabel===true;
    if(isCreate||existing!==binding.options.showReference)point.hasDataLabel=binding.options.showReference;
    if(binding.options.showReference) {
      const label=point.dataLabel;
      label.formula=`=${quote(instance.helperName)}!$${binding.tickColumns.label}$${binding.referenceIndex+3}`;
      if(isCreate||!existing) {
        label.position='Right';label.textOrientation=0;
        label.format.font.name=FONT;label.format.font.size=10;label.format.font.color='#000000';
        label.format.fill.clear();label.format.border.lineStyle='None';
        s(4).dataLabels.horizontalAlignment='Left';
      }
    }
  }
  await bindAxisNotes({context,chart,instance,binding},isCreate);
  if (!isCreate && needed>previousSeries) {
    chart.legend.load('visible'); await context.sync();
    if (chart.legend.visible===true) {
      const entries=chart.legend.legendEntries; entries.load('items'); await context.sync();
      for (let i=previousSeries;i<entries.items.length;i++) entries.items[i].visible=bodyLegendVisible(i,binding);
    }
  }
  await context.sync();
  await restrictCpiGridLabels({context,chart,binding});
}

function bodyLegendVisible(i,binding) {
  if (i<binding.firstBody) return false;
  const slot=i-binding.firstBody;
  return slot%binding.stride<binding.segments &&
    (['single','bubble'].includes(binding.options.variant) && binding.options.colorMode!=='sign' || Math.floor(slot/binding.stride)===0);
}

export async function createChart(args) {
  const {context,sourceSheet,helper,source,instance,binding} = args;
  const chart = sourceSheet.charts.add('AreaStacked',helper.getRange(`B${AREA_ROW}:B${binding.areaLast}`),'Columns');
  chart.name = instance.chartName; chart.series.getItemAt(0).name = '透明底座';
  for (const name of ['分类与宽度','数值','零线','刻度']) chart.series.add(name);
  if (binding.options.variant==='bubble') chart.series.add('规模气泡');
  await context.sync();
  for (let i=1;i<FIRST_BODY_SERIES;i++) {
    const series=chart.series.getItemAt(i); series.chartType='XYScatter'; series.axisGroup='Secondary'; hideLine(series);
  }
  chart.series.getItemAt(3).chartType='XYScatterLinesNoMarkers';
  if(binding.tickStride===2)chart.series.getItemAt(4).chartType='XYScatterLinesNoMarkers';
  if (binding.options.variant==='bubble') {
    const bubble=chart.series.getItemAt(5); bubble.chartType='Bubble'; bubble.axisGroup='Secondary';
    bubble.bubbleScale=65; bubble.format.fill.setSolidColor('#40A1ED'); bubble.format.line.color='#0A1C2C'; bubble.format.line.weight=.75;
  }
  const anchor=sourceSheet.getRangeByIndexes(source.rowIndex,source.columnIndex+source.columnCount+1,1,1); anchor.load('left,top'); await context.sync();
  const layout=layoutFor(binding.options);
  chart.left=anchor.left; chart.top=anchor.top; chart.width=layout.width; chart.height=layout.height;
  chart.plotVisibleOnly=false;
  if(binding.tickStride===2)chart.displayBlanksAs='NotPlotted';
  await bindChart({...args,chart},true);
  return chart;
}

export async function style({context,chart,helper,options,binding,instance}) {
  const o=normalizeOptions(options);
  const layout=layoutFor(o);
  if(binding.axisNoteIndices) {
    // Explicit format refresh also upgrades the legacy 171 left margin.
    chart.width=layout.width;chart.height=layout.height;
    await bindAxisNotes({context,chart,instance,binding},true);
  }
  const titleText=usesCpiGrid(o)&&o.subtitle?`${o.title}\n${o.subtitle}`:o.title;
  chart.title.visible=!!titleText; chart.title.text=titleText; chart.title.horizontalAlignment='Left';
  chart.title.format.font.name=FONT; chart.title.format.font.size=layout.titleFont;
  setChartTitleTypography(chart,{text:titleText});
  chart.format.font.name=FONT; chart.format.font.size=layout.font; chart.format.fill.setSolidColor('#FFFFFF');
  chart.legend.visible=o.showLegend; // Chart border is owned by the common presentation helper.
  chart.series.getItemAt(0).format.fill.clear(); chart.series.getItemAt(0).format.line.lineStyle='None';
  const primaryX=chart.axes.categoryAxis, primaryY=chart.axes.valueAxis;
  primaryX.categoryType='DateAxis'; primaryX.baseTimeUnit='Days'; primaryX.minimum=DATE_ORIGIN; primaryX.maximum=DATE_ORIGIN+DATE_SPAN;
  primaryX.isBetweenCategories=false; primaryX.visible=false;
  primaryY.minimum=o.variant==='bubble'?-.4:0; primaryY.maximum=1; primaryY.majorUnit=.2; primaryY.visible=false;
  // Horizontal stacked areas are mirrored by the primary numeric Y axis; scatter label Y already uses 1-centre.
  primaryY.reversePlotOrder=o.direction==='horizontal';
  const secondaryX=chart.axes.getItem('Category','Secondary'), secondaryY=chart.axes.getItem('Value','Secondary');
  secondaryX.minimum=DATE_ORIGIN; secondaryX.maximum=DATE_ORIGIN+DATE_SPAN; secondaryX.visible=false;
  secondaryY.minimum=o.variant==='bubble'?-.4:0; secondaryY.maximum=1; secondaryY.visible=false;
  for (const axis of [primaryX,primaryY,secondaryX,secondaryY]) { axis.majorGridlines.visible=false; axis.minorGridlines.visible=false; axis.majorTickMark='None'; axis.minorTickMark='None'; }
  if(binding.tickStride===2) {
    const grid=chart.series.getItemAt(4);
    grid.format.line.lineStyle='Continuous';grid.format.line.color='#F2F2F2';grid.format.line.weight=.5;
  }
  const guide=chart.series.getItemAt(3); guide.format.line.lineStyle='Continuous';guide.format.line.color='#808080'; guide.format.line.weight=.75;
  await context.sync();
  // Mac throws InvalidOperation on legendEntries while the legend is hidden.
  if (o.showLegend) {
    chart.legend.position='Top';
    const entries=chart.legend.legendEntries; entries.load('items'); await context.sync();
    for (let i=0;i<entries.items.length;i++) entries.items[i].visible=bodyLegendVisible(i,binding);
  }
  const p=layout.plot;
  for (let i=0;i<2;i++) { chart.plotArea.insideLeft=p.left; chart.plotArea.insideTop=p.top; chart.plotArea.insideWidth=p.width; chart.plotArea.insideHeight=p.height; await context.sync(); }
  if(usesCpiGrid(o)&&o.subtitle){chart.title.left=8;chart.title.top=6;await context.sync();}
  if(binding.axisNoteIndices) {
    const notes=Object.entries(binding.axisNoteIndices).map(([role,index])=>({role,label:chart.series.getItemAt(4).points.getItemAt(index).dataLabel}));
    // Mac auto-sized this 10pt caption to 79.46 x 37.31pt (two lines).
    // width is read-only; ExcelApi 1.19 setWidth() reserves sufficient text width.
    // The host can retain its prior height: read it instead of assuming reflow;
    // never clip the label or reduce its font to pass the bottom-margin check.
    notes.find(note=>note.role==='width').label.setWidth(CPI_WIDTH_NOTE_WIDTH);
    await context.sync();
    chart.getImage();await context.sync();
    for(const {label} of notes)label.load('width,height');
    await context.sync();
    for(const {role,label} of notes) {
      if(![label.width,label.height].every(v=>Number.isFinite(v)&&v>0))throw new UserError('171轴说明未返回可核对的原生标签边界');
      if(role==='width'&&Math.abs(label.width-CPI_WIDTH_NOTE_WIDTH)>.25)throw new UserError('171宿主未保留底注单行宽度：'+JSON.stringify({actual:label.width,expected:CPI_WIDTH_NOTE_WIDTH,height:label.height}));
      // Original PDF171 centres: YoY x=27px, width caption x=318px;
      // both follow the translated plot, retaining their source-image offsets.
      const left=p.left+((role==='y'?27:318)-83)/612*p.width-label.width/2;
      const top=role==='y'?p.top+p.height/2-label.height/2:p.top+p.height+2;
      if(left<1||top<1||left+label.width>layout.width-1||top+label.height>layout.height-1)throw new UserError('171轴说明留白不足，未认定视觉通过：'+JSON.stringify({role,label:{width:label.width,height:label.height},requested:{left,top},chart:{width:layout.width,height:layout.height},plot:p}));
      Object.assign(notes.find(note=>note.label===label),{left,top});
      label.left=left;label.top=top;
    }
    chart.series.getItemAt(4).dataLabels.showLeaderLines=false;
    await context.sync();
    chart.getImage();await context.sync();
    for(const {label} of notes)label.load('left,top,width,height');
    await context.sync();
    if(notes.some(({label,left,top})=>!Number.isFinite(label.left)||!Number.isFinite(label.top)||Math.abs(label.left-left)>.25||Math.abs(label.top-top)>.25))throw new UserError('171宿主未保留轴说明位置，未认定视觉通过');
    if(notes.some(({label})=>!Number.isFinite(label.width)||!Number.isFinite(label.height)||label.width<=0||label.height<=0||label.left<1||label.top<1||label.left+label.width>layout.width-1||label.top+label.height>layout.height-1))throw new UserError('171轴说明定位后留白不足，未认定视觉通过：'+JSON.stringify(notes.map(({role,label})=>({role,left:label.left,top:label.top,width:label.width,height:label.height}))));
  }
  if (o.profile==='pdf171' && o.direction==='vertical' && o.variant==='single') {
    // The user's 10pt text is larger than the old scaled 5.9pt candidate.
    // Measure actual native labels; stagger crowded values without moving
    // their data points, and keep adjacent short category names separated.
    // Force Excel to finish font/plot layout before reading native label boxes.
    chart.getImage();await context.sync();
    const groups = [1, ...(o.showValues ? [2] : [])].map(index => ({index, labels:Array.from({length:binding.count},(_,i)=>chart.series.getItemAt(index).points.getItemAt(i).dataLabel)}));
    for (const group of groups) {
      chart.series.getItemAt(group.index).dataLabels.showLeaderLines=false;
      for (const label of group.labels) label.load('left,top,width,height,text');
    }
    await context.sync();
    if(o.subtitle) {
      chart.title.load('top,height');await context.sync();
      const firstLabelTop=Math.min(...groups.flatMap(g=>g.labels.map(l=>l.top)));
      if(!Number.isFinite(chart.title.height)||chart.title.top+chart.title.height+6>firstLabelTop) throw new UserError('不等宽柱图标题与正文留白不足，未认定通过：'+JSON.stringify({titleTop:chart.title.top,titleHeight:chart.title.height,firstLabelTop}));
    }
    for (const {labels} of groups.filter(g=>g.index===1)) for (let i=1;i<labels.length;i++) {
      const a=labels[i-1], b=labels[i];
      if (![a.left,a.top,a.width,a.height,b.left,b.top,b.width,b.height].every(Number.isFinite)) throw new UserError('不等宽柱图未返回可核对的标签边界');
      const overlapX=a.left+a.width+3-b.left;
      if (overlapX>0 && a.top < b.top+b.height+2 && b.top < a.top+a.height+2) {
        if (a.height<25 && b.height<25) {a.left-=overlapX/2;b.left+=overlapX/2;}
      }
    }
    await context.sync();
    // Category spacing can move a name into the neighbouring negative value.
    // Re-read after those changes and reserve space across both label groups.
    chart.getImage();await context.sync();
    for (const group of groups) for (const label of group.labels) label.load('left,top,width,height,text');
    await context.sync();
    // Native 10pt label frames include about 3pt padding on each side. Use a
    // conservative 13.2pt centered ink height (width for vertical categories),
    // then translate the final ink coordinates back to the unchanged frames.
    const inkBox=(label,vertical=false)=>{
      if(![label.left,label.top,label.width,label.height].every(Number.isFinite)||label.width<=6||label.height<=6) throw new UserError('不等宽柱图未返回可核对的标签边界');
      const width=vertical?Math.min(13.2,label.width):label.width-6;
      const height=vertical?label.height-6:Math.min(13.2,label.height);
      const insetX=(label.width-width)/2,insetY=(label.height-height)/2;
      return {left:label.left+insetX,top:label.top+insetY,width,height,insetX,insetY,vertical};
    };
    const occupied=groups[0].labels.map((l,i)=>({...inkBox(l,i===5||i===6),category:true}));
    const barCoordinates=helper.getRange(`L3:U${binding.count+2}`);barCoordinates.load('values');await context.sync();
    for(const row of barCoordinates.values){
      const [start,end]=row,topValue=row[8],zero=row[9];
      if([start,end,topValue,zero].every(Number.isFinite)&&topValue>zero)occupied.push({left:p.left+start*p.width,top:p.top+(1-topValue)*p.height,width:(end-start)*p.width,height:(topValue-zero)*p.height,bar:true});
    }
    const requested=[];
    for (const label of groups.find(g=>g.index===2)?.labels || []) {
      const box=inkBox(label);
      const negative=/^[-−]/.test(label.text);
      let downward=negative;
      const overlaps=b=>box.left<b.left+b.width+2 && box.left+box.width+2>b.left && box.top<b.top+b.height+2 && box.top+box.height+2>b.top;
      for (let pass=0;pass<occupied.length+1;pass++) {
        const hit=occupied.find(overlaps); if(!hit)break;
        // A tall vertical category must not push +3.2% above the chart. Clear
        // its left edge horizontally first, retaining a 2pt text-to-text gap.
        if(hit.category&&hit.vertical) box.left=hit.left-box.width-2;
        else {
          if(hit.category) downward=true;
          box.top=downward?hit.top+hit.height+2:hit.top-box.height-2;
        }
      }
      if(box.left<4 || box.left+box.width>layout.width-4 || box.top<24 || box.top+box.height>layout.height-4 || occupied.some(overlaps)) throw new UserError('不等宽柱图标签留白不足，未认定视觉通过：'+JSON.stringify({text:label.text,box,limit:layout.height,conflicts:occupied.filter(overlaps)}));
      const left=box.left-box.insetX,top=box.top-box.insetY;
      label.left=left;label.top=top;
      occupied.push(box);requested.push({label,left,top});
    }
    await context.sync();
    chart.getImage();await context.sync();
    for(const {label} of requested)label.load('left,top');
    await context.sync();
    const failed=requested.filter(r=>!Number.isFinite(r.label.left)||!Number.isFinite(r.label.top)||Math.abs(r.label.left-r.left)>.25||Math.abs(r.label.top-r.top)>.25);
    if(failed.length) throw new UserError('不等宽柱图宿主未保留标签避让位置，未认定通过：'+JSON.stringify(failed.map(r=>({text:r.label.text,requested:[r.left,r.top],actual:[r.label.left,r.label.top]}))));
  }
  await restrictCpiGridLabels({context,chart,binding});
  return {notes:['原生面积系列绘制不等宽矩形；日期轴边界按百万分之一绘图区取整，双端实机及原图视觉待验。','微软雅黑为替代字体，未取得原版字体参数。'], nativeValidation:'not-run', recipe:'date-axis-stacked-area', dateResolution:DATE_SPAN};
}

// Core invokes this after shared typography on create and explicit refresh.
export async function finalizePresentation(args) { await restrictCpiGridLabels(args); }

export async function rebind(args) {
  // Rebinding updates formulas/references and styles only newly appended series/labels.
  // Existing chart dimensions, axes, colors, labels, line weights and title formatting are never reset.
  return bindChart({...args,instance:{...args.instance,previousCount:args.previous.count}},false);
}

export const inputSpec={
  category:{id:'comparison',title:'比较'},
  description:'选择分类、宽度权重和数值。宽度与高度分别随源数据重算；堆积可追加数值列，普通变体支持正负值。',
  minVersion:'1.9',columns:3,minRows:2,maxRows:MAX_ROWS,
  options:[
    {id:'title',label:'标题',type:'text',default:'不等宽柱与条形',maxLength:120},
    {id:'subtitle',label:'副标题（171留空使用原文）',type:'text',default:'',maxLength:160},
    {id:'direction',label:'方向',type:'select',default:'vertical',choices:[{value:'vertical',label:'纵向柱形'},{value:'horizontal',label:'横向条形'}]},
    {id:'variant',label:'变体',type:'select',default:'single',choices:[{value:'single',label:'普通（可含正负值）'},{value:'stacked',label:'堆积数值'},{value:'percent',label:'百分比堆积'},{value:'bubble',label:'纵向柱形＋规模气泡（4列）'}]},
    {id:'gap',label:'分类间隙',type:'number',default:0,min:0,max:30,step:1,scale:.01,suffix:'%'},
    {id:'showWidth',label:'显示宽度标签行',type:'checkbox',default:true},
    {id:'showValues',label:'显示数值',type:'checkbox',default:true},
    {id:'showLegend',label:'显示图例',type:'checkbox',default:false},
    {id:'profile',label:'版式',type:'select',default:'standard',choices:[{value:'standard',label:'通用'},{value:'pdf171',label:'图册171：CPI'}]},
    {id:'showReference',label:'171：显示整体参考指标（不是分类加权平均）',type:'checkbox',default:true},
    {id:'referenceLabel',label:'171：参考指标名称',type:'text',default:'All items CPI-U:',maxLength:60},
    {id:'referenceValue',label:'171：整体参考指标（原例3.2%）',type:'number',default:3.2,min:-10000,max:10000,step:.1,scale:.01,suffix:'%'},
    {id:'referenceCell',label:'171：跟随同表单元格（可选，如K2；优先于上方固定值）',type:'text',default:'',maxLength:12},
    {id:'showPlus',label:'正数显示加号',type:'checkbox',default:false},
    {id:'axisSymmetric',label:'正负轴对称',type:'checkbox',default:false},
    {id:'valuePlacement',label:'数值位置',type:'select',default:'outside',choices:[{value:'outside',label:'条柱外侧'},{value:'inside',label:'条柱内部'}]},
    {id:'colorMode',label:'配色',type:'select',default:'category',choices:[{value:'category',label:'按分类或堆积分段'},{value:'sign',label:'正深蓝、负浅蓝（普通变体）'}]},
    {id:'valueFormat',label:'数值格式',type:'select',default:'0.0',choices:['0','0.0','0.00','0%','0.0%','0.00%'].map(value=>({value,label:value}))},
    {id:'widthFormat',label:'宽度格式',type:'select',default:'0.0',choices:['0','0.0','0.00','0%','0.0%','0.00%'].map(value=>({value,label:value}))}
  ],
  storage:{marker:MARK,helperPrefix:HELPER_PREFIX,chartPrefix:CHART_PREFIX}
};
export default {id:'variablewidthbar',title:'不等宽柱与条形',inputSpec,validate,writeHelper,createChart,style,finalizePresentation,rebind};
