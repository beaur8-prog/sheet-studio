import {cloneData} from './clone.mjs';
import {conditionalFormula,shiftWorksheetFeatures} from './conditional.mjs';
import {nativeFilterPlan} from './filter-native.mjs';
import {shiftAutoFilter} from './filters.mjs';
import {chartSeries} from './chart-model.mjs';
import {address, parseAddress, transformFormulaReferences} from './engine.mjs';

const clone=value=>cloneData(value);
const equal=(a,b)=>JSON.stringify(a??null)===JSON.stringify(b??null);
const nativeTypes={bar:[51,52,53],'bar-horizontal':[57,58,59],line:[4,63,64],area:[1,76,77],pie:[5],doughnut:[-4120],scatter:[-4169],bubble:[15],radar:[-4151],combo:[51]};
function chartPlan(sheet,before,workbook){
  const oldCharts=before?.charts||[],charts=[];
  for(const chart of sheet.charts||[]){
    const old=oldCharts.find(c=>c.id===chart.id),patch={name:old?.nativeName||null};
    const rebuild=!old||chart.rebuilt&&!old.rebuilt||chart.type!==old.type||chart.grouping!==old.grouping||!chart.source&&chart.range!==old.range||!chart.source&&['firstRowAsHeader','firstColumnAsLabels','seriesBy'].some(key=>!equal(chart[key],old[key]));
    if(rebuild){
      if(chart.source&&!old)throw Error('가져온 차트의 시트 복제는 XLSX 저장을 사용하거나 차트 설정에서 새 차트로 변환하세요.');
      const types=nativeTypes[chart.type];if(!types)throw Error('이 차트 유형은 Excel 원본을 유지한 상태에서만 XLSM에 저장할 수 있습니다.');
      patch.chartType=types[['stacked','percentStacked'].includes(chart.grouping)?['','stacked','percentStacked'].indexOf(chart.grouping):0]||types[0];patch.range=chart.range;
      if(chart.type==='combo')patch.combo=true;
      patch.secondaryAxis=chart.secondaryAxis!==false;
      // Bind the same series ranges as the renderer/XLSX exporter instead of
      // letting Excel guess headers and silently omit a numeric first column.
      patch.series=chartSeries(chart,sheet,workbook).slice(0,chart.type==='pie'?1:32).map(s=>({valuesRange:s.valRef.range,...(s.nameRef?.range?{nameRange:s.nameRef.range}:{}),...((s.xRef||s.catRef)?.range?{categoriesRange:(s.xRef||s.catRef).range}:{}),...(s.sizeRef?.range?{sizesRange:s.sizeRef.range}:{})}));
    }
    for(const key of ['x','y','width','height','title','legend','labels','plotVisibleOnly','xTitle','yTitle'])if(rebuild||!equal(chart[key],old?.[key]))patch[key]=chart[key]??({x:40,y:25,width:500,height:300,title:'',legend:true,labels:false,plotVisibleOnly:true,xTitle:'',yTitle:''}[key]);
    if(patch.labels!==undefined&&['pie','doughnut'].includes(chart.type))patch.percentLabels=true;
    if(patch.legend!==undefined&&!chart.source&&chart.type==='bar')patch.legend=chart.legend!==false&&(patch.series||chartSeries(chart,sheet,workbook)).length>1;
    if(Object.keys(patch).length>1)charts.push(patch);
  }
  const deletedCharts=oldCharts.filter(c=>c.nativeName&&!sheet.charts?.some(now=>now.id===c.id)).map(c=>c.nativeName);
  return {charts,deletedCharts};
}
export function attachNativeSource(workbook, data, filename) {
  const bytes=data instanceof Uint8Array?data:new Uint8Array(data);
  if(bytes.length>15*1024*1024)throw new Error('매크로 통합 문서는 현재 15 MB까지 지원합니다.');
  let binary='';for(let at=0;at<bytes.length;at+=32768)binary+=String.fromCharCode(...bytes.subarray(at,at+32768));
  workbook.nativeSource={format:'xlsm',filename,base64:btoa(binary),baseline:clone(workbook.sheets),operations:[]};
  return workbook;
}

// Mirror structural edits in the comparison baseline. Excel performs the real
// row/column operations first, preserving named ranges, controls and charts.
export function recordNativeStructure(workbook,sheetId,axis,index,delta){
  const source=workbook.nativeSource;if(!source)return;
  source.operations.push({sheetId,axis,index,delta});
  const target=source.baseline.find(s=>s.id===sheetId);if(!target)return;
  const cells={};for(const [key,cell] of Object.entries(target.cells)){
    const p=parseAddress(key);if(delta<0&&p[axis]===index)continue;
    if(p[axis]>=index)p[axis]+=delta;
    const next=address(p.r,p.c);if(next!=='#REF!')cells[next]=cell;
  }target.cells=cells;
  for(const owner of source.baseline)for(const cell of Object.values(owner.cells))if(cell.raw.startsWith('='))cell.raw=transformFormulaReferences(cell.raw,{axis,index,delta,targetSheet:target.name,ownerSheet:owner.name});
  target.merges=(target.merges||[]).map(value=>transformFormulaReferences(value,{axis,index,delta,targetSheet:target.name,ownerSheet:target.name})).filter(value=>!value.includes('#REF!'));
  const prop=axis==='r'?'rowHeights':'columnWidths',sizes={};for(const [key,value] of Object.entries(target[prop]||{})){let n=+key;if(delta<0&&n===index)continue;if(n>=index)n+=delta;sizes[n]=value;}target[prop]=sizes;shiftAutoFilter(target,axis,index,delta);shiftWorksheetFeatures(target,axis,index,delta);
}

export function buildNativePlan(workbook){
  const source=workbook.nativeSource;if(source?.format!=='xlsm')throw new Error('먼저 XLSM 통합 문서를 여세요.');
  const baseline=new Map(source.baseline.map(s=>[s.id,s]));
  const sheets=workbook.sheets.map(sheet=>{
    const before=baseline.get(sheet.id),cells=[];
    for(const key of new Set([...Object.keys(before?.cells||{}),...Object.keys(sheet.cells)])){
      const old=before?.cells[key]||{raw:''},now=sheet.cells[key]||{raw:''},patch={address:key};
      if(old.raw!==now.raw)patch.raw=now.raw;
      const style={};for(const name of new Set([...Object.keys(old.style||{}),...Object.keys(now.style||{})]))if(!equal(old.style?.[name],now.style?.[name]))style[name]=now.style?.[name]??null;
      if(Object.keys(style).length)patch.style=style;
      if(!equal(old.note,now.note))patch.note=now.note||'';
      if(Object.keys(patch).length>1)cells.push(patch);
    }
    const dimensions=prop=>Object.fromEntries([...new Set([...Object.keys(before?.[prop]||{}),...Object.keys(sheet[prop]||{})])].filter(key=>!equal(before?.[prop]?.[key],sheet[prop]?.[key])).map(key=>[key,sheet[prop]?.[key]??null]));
    const conditionalFormats=!equal(sheet.conditionalRules,before?.conditionalRules)||sheet.unsupportedConditionalCount!==before?.unsupportedConditionalCount?(sheet.conditionalRules||[]).map(r=>({range:r.range,type:r.type==='cellIs'?1:2,operator:r.type==='cellIs'?({between:1,notBetween:2,equal:3,notEqual:4,greaterThan:5,lessThan:6,greaterThanOrEqual:7,lessThanOrEqual:8})[r.operator]:null,formula1:r.type==='cellIs'?r.formulae[0]:conditionalFormula(r),formula2:r.formulae[1]||null,style:clone(r.style),stopIfTrue:!!r.stopIfTrue})):undefined;
    if(conditionalFormats&&sheet.unsupportedConditionalCount)throw Error('미지원 조건부 서식은 Excel에서 편집하거나 모두 지운 뒤 저장하세요.');
    return {...(conditionalFormats!==undefined?{conditionalFormats}:{}),...(!equal(sheet.freeze,before?.freeze)?{freeze:sheet.freeze||{rows:0,cols:0}}:{}),...(!equal(sheet.autoFilter,before?.autoFilter)?{autoFilter:nativeFilterPlan(sheet.autoFilter,workbook),manualHiddenRows:Object.entries(sheet.rowHeights||{}).filter(([,h])=>h===0).map(([r])=>+r)}:{}),id:sheet.id,sourceName:before?.name??null,name:sheet.name,cells,columnWidths:dimensions('columnWidths'),rowHeights:dimensions('rowHeights'),unmerge:(before?.merges||[]).filter(v=>!(sheet.merges||[]).includes(v)),merge:(sheet.merges||[]).filter(v=>!(before?.merges||[]).includes(v)),...chartPlan(sheet,before,workbook)};
  });
  return {pivots:workbook.sheets.filter(s=>s.pivot).map(s=>({sheet:s.name,source:workbook.sheets.find(source=>source.id===s.pivot.sourceSheetId)?.name||null,spec:clone(s.pivot)})),sheets,deleted:source.baseline.filter(s=>!workbook.sheets.some(v=>v.id===s.id)).map(s=>s.name),operations:clone(source.operations||[]),activeSheetId:workbook.activeSheetId};
}

export function nativeSourceChanged(workbook){const p=buildNativePlan(workbook);return !!(p.deleted.length||p.operations.length||p.sheets.some(s=>!s.sourceName||s.name!==s.sourceName||s.cells.length||s.autoFilter!==undefined||s.freeze!==undefined||s.conditionalFormats!==undefined||s.merge.length||s.unmerge.length||s.charts.length||s.deletedCharts.length||Object.keys(s.columnWidths).length||Object.keys(s.rowHeights).length));}

// This is a labelled display fallback, never a replacement for VBA evaluation.
export function nativeCachedValue(workbook,sheetId,key,computed){
  if(computed!=='#NAME?'||!workbook.nativeSource)return undefined;
  const cell=workbook.sheets.find(s=>s.id===sheetId)?.cells[key];
  const original=workbook.nativeSource.baseline.find(s=>s.id===sheetId)?.cells[key];
  return cell?.raw===original?.raw&&cell?.raw?.startsWith('=')?cell.cachedValue:undefined;
}
