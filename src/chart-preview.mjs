import {chartData,chartRangeError,refFormula,referenceSize} from './chart-model.mjs';
import {renderChartSVG} from './chart-render.mjs';
import {validateChart} from './chart-schema.mjs';
import {formatValue} from './engine.mjs';

/** Read-only dialog preview. Run in the workbook worker so ranges and formula
 * results use the same revision as the editor, without copying the document. */
export function buildChartPreview(book,sheetId,{chartId,retain=false,patch={}}={}){
  const sheet=book.sheets.find(s=>s.id===sheetId);if(!sheet)throw Error('차트 원본 시트를 찾을 수 없습니다.');
  const original=chartId?sheet.charts?.find(c=>c.id===chartId):null;if(retain&&!original)throw Error('원본 차트가 변경되었습니다. 설정 창을 다시 여세요.');
  const draft={width:600,height:360,...original,...patch,...(!retain?{series:undefined,source:undefined}:{})};validateChart(draft);
  if(!retain){const error=chartRangeError(draft,sheet,book);if(error)throw Error(error);}
  const data=chartData(draft,sheet,book),plot=renderChartSVG(draft,sheet,book),prefix='chart-preview-'+crypto.randomUUID();
  const rows=data.slice(0,8).flatMap((s,index)=>s.values.slice(0,8).map((value,i)=>({series:s.name,item:s.labels[i]??String(i+1),value:value===null?(s.blankValues[i]?'빈 값':'숫자 아님 / 오류'):String(formatValue(value,s.valueFormat,book))+(s.blankValues[i]?' (빈 셀)':''),x:['scatter','bubble'].includes(draft.type)?s.xs[i]??'빈 값':undefined,size:draft.type==='bubble'?s.sizes[i]??'빈 값':undefined,used:draft.type!=='pie'||index===0})));
  const reference=(ref,fallback)=>ref&&(ref.formula||ref.sheetId&&ref.range)?refFormula(ref,book):fallback;
  const references=data.slice(0,8).map(s=>({series:s.name,values:reference(s.valRef,'파일의 저장값'),items:reference(s.catRef,'1, 2, 3…'),...(s.xRef?{x:reference(s.xRef,'1, 2, 3…')}:{}),...(s.sizeRef?{size:reference(s.sizeRef,'파일의 저장값')}:{})}));
  return {...plot,svg:plot.svg.replace(/chart-plot-\d+/g,prefix),title:String(draft.title||'차트'),type:draft.type,note:draft.previewNote||draft.unsupportedReason||'',seriesCount:data.length,points:data.map(s=>({series:s.name,shown:s.values.length,total:referenceSize(s.valRef)})),rows,references,truncated:data.some(s=>s.truncated),usesCachedValues:data.some(s=>s.usesCachedValues)};
}
