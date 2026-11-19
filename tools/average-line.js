import {UserError} from '../core/errors.js';
// F107 is a modifier of an existing chart, not a core-v1 chart module.
// Bounded candidate: one ordinary categorical series backed by a local contiguous A1 range.
export const REQUIRED_API='1.15',MARKER='ChartCraft.Native.AverageLine.v1';
const supportedTypes=['ColumnClustered','ColumnStacked','Area'];
const pointCoordinates=type=>type==='Area';
const quote=s=>`'${s.replace(/'/g,"''")}'`;
export class AverageLineError extends UserError {}
const fail=message=>{throw new AverageLineError(message);};
export function parseLocalRange(text){
 if(typeof text!=='string')fail('系列未返回可解析的本地地址。');
 const match=/^=?('(?:[^']|'')+'|[^'!=]+)!\$?([A-Z]{1,3})\$?([1-9]\d*)(?::\$?([A-Z]{1,3})\$?([1-9]\d*))?$/i.exec(text.trim());
 if(!match)fail('仅支持本工作簿的连续单行或单列单元格来源；命名公式、数组及多区域暂未实现。');
 const sheet=match[1][0]==="'"?match[1].slice(1,-1).replace(/''/g,"'"):match[1];
 if(/[\[\]\\/:?*]/.test(sheet))fail('外部工作簿来源暂不支持。');
 const column=x=>[...x.toUpperCase()].reduce((n,c)=>n*26+c.charCodeAt(0)-64,0),c1=column(match[2]),c2=column(match[4]||match[2]),r1=+match[3],r2=+(match[5]||match[3]);
 if(c2<c1||r2<r1||c2>16384||r2>1048576||(c2!==c1&&r2!==r1))fail('数值来源须为有效的单行或单列区域。');
 const count=(r2-r1+1)*(c2-c1+1);if(count>5000)fail('目前支持最多 5000 个源数据点。');
 const start=`$${match[2].toUpperCase()}$${r1}`,end=`$${(match[4]||match[2]).toUpperCase()}$${r2}`,address=start===end?start:`${start}:${end}`;
 return {sheet,address,formulaRef:`${quote(sheet)}!${address}`,count,orientation:c1===c2?'column':'row'};
}
// REA: Enhance 201 0600022f calls Share 060004e4 GetFmt0(first source cell).
// The label format is chosen when the tool runs, just as the original does.
export function averageNumberFormat({text,numberFormat,decimalSeparator='.'}){
 const displayed=String(text),format=String(numberFormat);
 const decimals=percent=>{
  const at=displayed.indexOf(decimalSeparator);
  const count=at<0?0:displayed.length-at-(percent?2:1);
  return count===0?0:count===1?1:2;
 };
 if(['General','G','通用格式'].includes(format)){
  const n=decimals(false);return n===0?'#,##0':n===1?'0.0':'0.00';
 }
 if(displayed.endsWith('%')){const n=decimals(true);return '0'+(n?'.'+'0'.repeat(n):'')+'%';}
 return format.replaceAll('[Red]','').replaceAll('[红色]','');
}
export function helperPlan(source,numberFormat='#,##0',points=false){
 const fmt=numberFormat.replaceAll('"','""');
 return {address:'A3:C4',formulas:[[points?'=1':'=0.5',`=IFERROR(AVERAGE(${source.formulaRef}),0)`,`="◀ Avg= "&TEXT(B3,"${fmt}")`],[`=${source.orientation==='column'?'ROWS':'COLUMNS'}(${source.formulaRef})${points?'':'+0.5'}`,'=$B$3','=$C$3']]};
}
export async function inspectAverageLineTarget({context,chart}){
 if(!globalThis.Office?.context?.requirements?.isSetSupported('ExcelApi',REQUIRED_API))fail('当前 Excel 版本不能读取图表的数据来源，暂无法添加动态平均线。');
 chart??=context.workbook.getActiveChartOrNullObject();chart.load('isNullObject');await context.sync();
 if(chart.isNullObject)fail('请先选择一个支持的现有原生图表。');
 // Resolve the selection before helper creation; subsequent operations must not depend on the active chart.
 chart.load('id,name');chart.worksheet.load('id');await context.sync();
 const selectedId=chart.id;
 chart=context.workbook.worksheets.getItem(chart.worksheet.id).charts.getItem(chart.name);
 chart.load('id,name,plotVisibleOnly');chart.legend.load('visible');chart.series.load('items/chartType,items/axisGroup,items/name');await context.sync();
 if(chart.id!==selectedId)fail('图表名称重复或目标已变化，未添加平均线。请为所选图表设置唯一名称后重试。');
 // No documented isPivotChart getter: do not treat a hidden field-button state as proof.
 const pivots=context.workbook.pivotTables;pivots.load('items');await context.sync();if(pivots.items.length)fail('当前基础版本不处理含数据透视表的工作簿。请使用普通工作簿中的受支持单系列图表。');
 if(chart.series.items.length!==1)fail('基础平均线仅支持单系列图表；已有辅助线或其他系列时不重复添加。');
 const original=chart.series.items[0];if(!supportedTypes.includes(original.chartType)||original.axisGroup!=='Primary')fail('目前支持主坐标轴上的单系列簇柱、堆积柱、面积图。');
 if(original.chartType!=='ColumnClustered'&&globalThis.Office?.context?.platform!=='PC')fail('该图型平均线目前仅供Windows验收，Mac尚未验证。');
 // These retrieve source metadata, never static point values. Do not substitute getDimensionValues.
 const sourceType=original.getDimensionDataSourceType('Values'),sourceString=original.getDimensionDataSourceString('Values');await context.sync();
 if(sourceType.value!=='LocalRange')fail(`系列来源类型${sourceType.value}暂不支持；只接受本地单元格区域。`);
 const source=parseLocalRange(sourceString.value),range=context.workbook.worksheets.getItem(source.sheet).getRange(source.address);
 const firstCell=range.getCell(0,0),culture=context.application.cultureInfo.numberFormat;
 firstCell.load('text,numberFormat');culture.load('numberDecimalSeparator');
 range.load('address,rowCount,columnCount,rowHidden,columnHidden');chart.axes.categoryAxis.load('categoryType,isBetweenCategories');
 chart.plotArea.load('insideLeft,insideTop,insideWidth,insideHeight');await context.sync();
 if(range.rowCount*range.columnCount!==source.count)fail('来源地址与实际区域尺寸不一致。');
 if(range.rowHidden!==false||range.columnHidden!==false)fail('源区包含隐藏行列；平均线辅助表需允许隐藏数据，暂不改变原图隐藏数据语义。');
 if(chart.axes.categoryAxis.categoryType==='DateAxis'||(!pointCoordinates(original.chartType)&&chart.axes.categoryAxis.isBetweenCategories!==true))fail('日期轴或非间隔分类轴暂不支持。');
 const plot=Object.fromEntries(['insideLeft','insideTop','insideWidth','insideHeight'].map(k=>[k,chart.plotArea[k]]));
 if(Object.values(plot).some(v=>!Number.isFinite(v))||plot.insideWidth<=0||plot.insideHeight<=0)fail('原图未返回有效绘图区。');
 const numberFormat=averageNumberFormat({text:firstCell.text[0][0],numberFormat:firstCell.numberFormat[0][0],decimalSeparator:culture.numberDecimalSeparator});
 return {chart,source,numberFormat,sourceType:sourceType.value,sourceString:sourceString.value,originalName:original.name,originalType:original.chartType,plot,previousPlotVisibleOnly:chart.plotVisibleOnly};
}
export async function addAverageLine({context,chart,helperName=`_CC_AVG_${Date.now().toString(36)}`}){
 if(!/^_CC_AVG_[A-Za-z0-9_]{1,23}$/.test(helperName)||helperName.length>31)fail('辅助表名称无效。');
 const target=await inspectAverageLineTarget({context,chart}),old=context.workbook.worksheets.getItemOrNullObject(helperName);old.load('isNullObject');await context.sync();if(!old.isNullObject)fail('辅助表已存在，未覆盖或重复添加。');
 // REA BackupChart: keep the original and operate on a native duplicate.
 // Windows probe confirmed that Unsupported chart shapes still support copyTo.
 const originalChart=target.chart,copyName=`CC_Avg_${helperName.slice(8)}`;
 const existing=originalChart.worksheet.charts.getItemOrNullObject(copyName);existing.load('isNullObject');await context.sync();if(!existing.isNullObject)fail('平均线副本名称已存在，未覆盖原图。');
 const charts=originalChart.worksheet.charts;charts.load('items/id');await context.sync();const beforeIds=new Set(charts.items.map(c=>c.id));
 originalChart.worksheet.shapes.getItem(originalChart.name).copyTo();await context.sync();
 // copyTo's returned shape can retain the source name. Resolve the new native chart by ID.
 charts.load('items/id');await context.sync();const added=charts.items.filter(c=>!beforeIds.has(c.id));
 if(added.length!==1)fail('原生图表复制结果不唯一，未修改原图。');
 target.chart=charts.getItem(added[0].id);target.chart.name=copyName;await context.sync();
 target.chart=originalChart.worksheet.charts.getItem(copyName);target.chart.load('id,width,name');target.chart.legend.load('visible');target.chart.series.load('items/chartType');await context.sync();
 if(target.chart.id===originalChart.id||target.chart.series.items.length!==1||target.chart.series.items[0].chartType!==target.originalType)fail('宿主未保留可编辑的原生柱形图副本；未修改原图。');
 const helper=context.workbook.worksheets.add(helperName),plan=helperPlan(target.source,target.numberFormat,pointCoordinates(target.originalType));
 helper.getRange('A1:B1').values=[[MARKER,JSON.stringify({version:1,chart:target.chart.name,source:target.source.formulaRef,originalType:target.originalType,coordinates:pointCoordinates(target.originalType)?'points':'interval',originalAxis:'Primary',sourceKind:'LocalRange',previousPlotVisibleOnly:target.previousPlotVisibleOnly})]];
 helper.getRange('A2:C2').values=[['分类坐标','平均值','标签']];helper.getRange(plan.address).formulas=plan.formulas;helper.visibility='Hidden';
 let series=target.chart.series.add('◀ Avg=');series.chartType='XYScatterLinesNoMarkers';series.axisGroup='Primary';series.markerStyle='None';series.hasDataLabels=false;
 series.setXAxisValues(helper.getRange('A3:A4'));series.setValues(helper.getRange('B3:B4'));
 series.format.line.lineStyle='Dash';series.format.line.color='#000000';series.format.line.weight=.75;target.chart.plotVisibleOnly=false;
 if(pointCoordinates(target.originalType))target.chart.axes.categoryAxis.isBetweenCategories=false;
 await context.sync();
 // add() returns a creation-path proxy; reacquire the committed second series before using its children.
 series=target.chart.series.getItemAt(1);
 const endpoint=series.points.getItemAt(1);endpoint.hasDataLabel=true;endpoint.dataLabel.formula=`=${quote(helperName)}!$C$4`;endpoint.dataLabel.position='Right';endpoint.dataLabel.format.font.name='微软雅黑';endpoint.dataLabel.format.font.size=10;endpoint.dataLabel.format.font.color='#000000';
 if(target.chart.legend.visible){const entries=target.chart.legend.legendEntries;entries.load('items');await context.sync();if(entries.items.length!==2)fail('新增平均线图例结构异常，保留对象供检查。');entries.items[1].visible=false;}
 // Match the original conditional label fit, rather than pinning the old plot box.
 const label=endpoint.dataLabel;label.load('left,width');target.chart.plotArea.load('insideLeft,insideWidth');await context.sync();
 if(label.left<target.chart.plotArea.insideLeft+target.chart.plotArea.insideWidth)
  target.chart.plotArea.insideWidth=target.chart.width-target.chart.plotArea.insideLeft-label.width-5;
 target.chart.activate();
 series.load('chartType,axisGroup');target.chart.series.load('items/chartType,items/axisGroup');await context.sync();
 if(series.chartType!=='XYScatterLinesNoMarkers'||series.axisGroup!=='Primary'||target.chart.series.items[0].chartType!==target.originalType)fail('宿主未保持原图与同轴平均线拓扑；已创建对象保留供检查，不认定成功。');
 if(['PC','Mac'].includes(globalThis.Office?.context?.platform)){
  if(Office.context.platform==='Mac'){target.chart.width=640;target.chart.height=360;await context.sync();}
  context.workbook.load('name');target.chart.worksheet.load('name');await context.sync();
  const response=await fetch('/native-average-dash',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({workbook:context.workbook.name,worksheet:target.chart.worksheet.name,chart:target.chart.name,helper:helperName}),credentials:'omit'});
  const result=await response.json();if(!response.ok||result.ok!==true)fail('原生短虚线未完成：'+(result.error||response.status));
 }
 return {helperName,chartName:target.chart.name,source:target.source.formulaRef,seriesIndex:1,sourcePoints:target.source.count,nativeValidation:'requires-host-acceptance',limitations:['仅不含数据透视表工作簿中的普通单系列簇柱/堆积柱/面积图；未覆盖条、折线、雷达、散点十字线或选中子系列自动识别。','仅本工作簿连续A1范围；命名公式、结构化引用、外链、数组、多区域、隐藏源行列拒绝。','新增或减少数据行后可按选区更新；原区域改数与插行重定位由Excel公式负责。','原厂命名公式转为隐藏辅助单元格；字体为用户覆盖10号微软雅黑。']};
}

// Existing average-line instances have a separate marker from chart recipes.
export async function listAverageLines(context){
 const sheets=context.workbook.worksheets;sheets.load('items/name');await context.sync();
 const helpers=sheets.items.filter(s=>/^_CC_AVG_[A-Za-z0-9_]+$/.test(s.name));
 const cells=helpers.map(s=>s.getRange('A1:B1'));cells.forEach(c=>c.load('values'));await context.sync();
 const result=[];
 for(let i=0;i<helpers.length;i++){
  if(cells[i].values[0][0]!==MARKER)continue;
  let meta;try{meta=JSON.parse(cells[i].values[0][1]);}catch{continue;}
  if(meta.version!==1||typeof meta.chart!=='string')continue;
  const source=parseLocalRange(meta.source);
  const charts=[];for(const sheet of sheets.items){const chart=sheet.charts.getItemOrNullObject(meta.chart);chart.load('isNullObject');charts.push({sheet,chart});}
  await context.sync();const matches=charts.filter(c=>!c.chart.isNullObject);if(matches.length!==1)continue;
  result.push({moduleId:'average-line',helperName:helpers[i].name,chartName:meta.chart,chartSheet:matches[0].sheet.name,title:meta.chart,source:source.formulaRef,count:source.count});
 }
 return result;
}
export async function selectAverageSource(context,helperName){
 const items=await listAverageLines(context),item=items.find(x=>x.helperName===helperName);if(!item)fail('未找到唯一平均线图表，未修改。');
 const source=parseLocalRange(item.source),sheet=context.workbook.worksheets.getItem(source.sheet);sheet.activate();sheet.getRange(source.address).select();await context.sync();return item;
}
export async function rebindAverageLine(context,helperName){
 const items=await listAverageLines(context),item=items.find(x=>x.helperName===helperName);if(!item)fail('未找到唯一平均线图表，未修改。');
 const active=context.workbook.getActiveChartOrNullObject();active.load('isNullObject');await context.sync();if(!active.isNullObject)fail('请先选择含表头的两列数据区域。');
 const selected=context.workbook.getSelectedRange();selected.load('address,rowCount,columnCount,values,rowHidden,columnHidden');await context.sync();
 if(selected.columnCount!==2||selected.rowCount<3||selected.rowCount>5001||selected.rowHidden!==false||selected.columnHidden!==false)fail('请选择含表头的两列连续区域，至少两行数据，最多5000行，不含隐藏行列。');
 if(selected.values.slice(1).some(r=>typeof r[1]!=='number'||!Number.isFinite(r[1])||r[0]===''))fail('分类不可为空，数值须为有效数字。未更改原图。');
 const n=selected.rowCount-1,categories=selected.getCell(1,0).getResizedRange(n-1,0),values=selected.getCell(1,1).getResizedRange(n-1,0);values.load('address');await context.sync();
 const source=parseLocalRange(values.address),chart=context.workbook.worksheets.getItem(item.chartSheet).charts.getItem(item.chartName),helper=context.workbook.worksheets.getItem(helperName);
 chart.series.load('items/chartType,items/axisGroup');const metadata=helper.getRange('B1');metadata.load('values');await context.sync();
 if(chart.series.items.length!==2||!supportedTypes.includes(chart.series.items[0].chartType)||chart.series.items[1].chartType!=='XYScatterLinesNoMarkers')fail('图表结构已改变，未更新来源。');
 const meta=JSON.parse(metadata.values[0][0]);
 if(chart.series.items[0].chartType!==meta.originalType||chart.series.items.some(s=>s.axisGroup!=='Primary'))fail('图表类型或坐标轴已改变，未更新来源。');
 // Reuse the original creation-time label format. Only source formulas change.
 chart.series.items[0].setXAxisValues(categories);chart.series.items[0].setValues(values);
 helper.getRange('B3').formulas=[[`=IFERROR(AVERAGE(${source.formulaRef}),0)`]];
 helper.getRange('A4').formulas=[[`=ROWS(${source.formulaRef})${pointCoordinates(meta.originalType)?'':'+0.5'}`]];
 meta.source=source.formulaRef;metadata.values=[[JSON.stringify(meta)]];await context.sync();
 return {...item,source:source.formulaRef,oldCount:item.count,count:n};
}
