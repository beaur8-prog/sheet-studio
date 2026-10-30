import {address,parseAddress,createEvaluator,formatValue} from './engine.mjs';
import {tableBounds} from './table-schema.mjs';
import {storedArrayIndex} from './stored-arrays.mjs';
import {worksheetCellStyle} from './cell-style.mjs';
import {numberFormatStyle} from './number-formats.mjs';
import {createConditionalEvaluator} from './conditional.mjs';
import {bordersCss,cellBorders,edgeCss} from './borders.mjs';
import {renderChartSVG} from './chart-render.mjs';
import {chartRect} from './chart-geometry.mjs';
import {drawingPreviews} from './xlsx-drawings.mjs';
import {nativeCachedValue} from './native-plan.mjs';
import {readPrintSettings,planPrintPages,checkPrintRange,printRowHeight,printColWidth} from './print-layout.mjs';

const esc=value=>String(value??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
function printRange(sheet,selection,spills=[]){
  if(selection){
    checkPrintRange(selection);
    return {...selection};
  }
  let range=null;
  const include=b=>{if(!b)return;if(!range)range={...b};else {range.r1=Math.min(range.r1,b.r1);range.c1=Math.min(range.c1,b.c1);range.r2=Math.max(range.r2,b.r2);range.c2=Math.max(range.c2,b.c2);}};
  for(const ref in sheet.cells){const cell=sheet.cells[ref];if(cell.raw===''&&!Object.keys(cell.style||{}).length)continue;const p=parseAddress(ref);if(p)include({r1:p.r,r2:p.r,c1:p.c,c2:p.c});}
  for(const table of sheet.tables||[])include(tableBounds(table.range));
  for(const array of [...storedArrayIndex(sheet).regions,...spills])include(array);
  if(!range)return null;
  // Merge bounds are meaningful when their content lies in the used region.
  for(const ref of sheet.merges||[]){const b=tableBounds(ref);if(b&&b.r1<=range.r2&&b.r2>=range.r1&&b.c1<=range.c2&&b.c2>=range.c1)include(b);}
  return range;
}
function gridHTML(book,sheet,{rows,cols,allRows,allCols,scale},{gridLines=false,evaluate}){
  if(!rows.length||!cols.length)return {html:'<p>인쇄할 표시 셀이 없습니다.</p>',count:0};
  const merges=new Map();
  for(const ref of sheet.merges||[]){
    const b=tableBounds(ref);if(!b)continue;const rr=rows.filter(r=>r>=b.r1&&r<=b.r2),cc=cols.filter(c=>c>=b.c1&&c<=b.c2);if(!rr.length||!cc.length)continue;
    const merge={...b,firstR:rr[0],firstC:cc[0],rowspan:rr.length,colspan:cc.length};
    for(const r of rr)for(const c of cc){const key=address(r,c);if(merges.has(key))throw Error('겹친 병합 셀은 인쇄할 수 없습니다.');merges.set(key,merge);}
  }
  const conditional=createConditionalEvaluator(book,sheet,evaluate),width=cols.reduce((n,c)=>n+printColWidth(sheet,c),0);let cachedCount=0;
  const tr=rows.map(r=>{
    const td=cols.map(c=>{
      const merge=merges.get(address(r,c));if(merge&&(r!==merge.firstR||c!==merge.firstC))return '';
      const row=merge?.r1??r,col=merge?.c1??c,ref=address(row,col),cell=sheet.cells[ref]||{};let result=evaluate.result(ref);const cached=result.cached&&!result.stale?result.cachedValue:result.error?nativeCachedValue(book,sheet.id,ref,result.value):undefined;if(cached!==undefined){result={value:cached,error:result.cachedError|| (cell.cachedError?cached:null)};cachedCount++;}const value=result.value;
      const style={...numberFormatStyle(value,worksheetCellStyle(book,sheet,row,col)),...conditional(row,col,value)};
      const css=['background:'+(style.background||'white'),'color:'+(style.color||'#222'),'font-family:'+(style.fontFamily||'Arial'),'font-size:'+(style.fontSize||11)+'pt','font-weight:'+(style.bold?'bold':'normal'),'font-style:'+(style.italic?'italic':'normal'),'text-decoration:'+(style.underline?'underline':'none'),'text-align:'+(style.align|| (typeof value==='number'?'right':'left')),'vertical-align:'+(style.verticalAlign||'middle'),'white-space:'+(style.wrap?'pre-wrap':'pre'),bordersCss(style)].join(';');
      const height=merge?rows.filter(r=>r>=merge.r1&&r<=merge.r2).reduce((n,r)=>n+printRowHeight(sheet,r),0):printRowHeight(sheet,r),edges=cellBorders(style),border=Math.max(gridLines?1:0,...['top','bottom'].map(side=>parseFloat(edgeCss(edges[side]))||0));
      // A fixed inner box clips text to the saved row height. Leave room for
      // collapsed borders so they cannot grow every row beyond the page plan.
      const continuation=merge&&(!rows.includes(allRows.find(r=>r>=merge.r1&&r<=merge.r2))||!cols.includes(allCols.find(c=>c>=merge.c1&&c<=merge.c2)));
      return '<td data-ref="'+ref+'"'+(merge?' rowspan="'+merge.rowspan+'" colspan="'+merge.colspan+'"':'')+' style="'+esc(css)+'"><div class="print-cell" style="height:'+Math.max(0,height-border*2)+'px">'+(continuation?'':esc(formatValue(value,style,{...book,error:result.error})))+'</div></td>';
    }).join('');
    return '<tr data-row="'+r+'" style="height:'+printRowHeight(sheet,r)+'px">'+td+'</tr>';
  }).join('');
  const columns=cols.map(c=>'<col style="width:'+printColWidth(sheet,c)+'px">').join('');
  const html='<section class="print-grid" style="zoom:'+scale+'"><table class="'+(gridLines?'with-grid':'')+'" style="width:'+width+'px"><colgroup>'+columns+'</colgroup><tbody>'+tr+'</tbody></table></section>';
  return {html,count:rows.length*cols.length,cachedCount};
}
/** Print cells and supported objects in a paginated flow. Charts follow the
 * data, instead of being silently omitted or clipped at their screen position.
 * This is an immutable document, built with the worker's calculation session.
 */
export function buildSheetPrintDocument(book,sheetId,{scope='sheet',selection,orientation,gridLines,ignorePrintArea=false,includeCharts=true,includeImages=true,evaluate}={}){
  const sheet=book.sheets.find(s=>s.id===sheetId);if(!sheet)throw Error('인쇄할 시트를 찾을 수 없습니다.');
  const settings=readPrintSettings(book,sheet,{orientation,ignorePrintArea:ignorePrintArea||scope!=='sheet',ignorePrintTitles:scope!=='sheet'});orientation=settings.orientation;gridLines??=settings.gridLines;
  if(!['sheet','selection','charts'].includes(scope)||!['portrait','landscape'].includes(orientation))throw Error('인쇄 설정이 올바르지 않습니다.');
  if(scope==='selection'&&!selection)throw Error('인쇄할 선택 영역을 지정하세요.');
  const {contentWidth,contentHeight}=settings,calculate=evaluate||createEvaluator(book,sheetId),chartOnly=scope==='charts'||sheet.kind==='chart';
  // Selection printing deliberately ignores saved areas and repeating titles.
  if(scope==='selection'){settings.repeatRows=null;settings.repeatCols=null;}
  const ranges=chartOnly?[]:settings.areas.length?settings.areas:[printRange(sheet,scope==='selection'?selection:null,calculate.spills?.()||[])],plan=planPrintPages(sheet,ranges,settings);
  let cachedCount=0;
  const pages=plan.pages.map((page,i)=>{const grid=gridHTML(book,sheet,page,{gridLines,evaluate:calculate});cachedCount+=grid.cachedCount||0;return '<section class="print-page" data-page="'+(i+1)+'" style="height:'+(contentHeight-1)+'px">'+grid.html+'</section>';});
  const charts=(includeCharts||chartOnly?sheet.charts||[]:[]).map(chart=>{
    const rect=chartRect(chart),plot=renderChartSVG(chart,sheet,book),scale=Math.min(1,contentWidth/rect.width,Math.max(20,contentHeight-40)/rect.height),width=rect.width*scale,height=rect.height*scale;
    return '<figure class="print-chart" data-chart="'+esc(chart.id)+'" style="width:'+width+'px;height:'+height+'px"><h2>'+esc(chart.title||'차트')+'</h2><svg xmlns="http://www.w3.org/2000/svg" viewBox="'+esc(plot.viewBox)+'" role="img" aria-label="'+esc(chart.title||'차트')+'">'+plot.svg+'</svg></figure>';
  });
  const previews=drawingPreviews(sheet),images=includeImages&&scope!=='charts'?previews.images.map(image=>{const scale=Math.min(1,contentWidth/image.width,Math.max(20,contentHeight-40)/image.height);return '<figure class="print-image"><img src="'+esc(image.src)+'" alt="'+esc(image.name)+'" style="width:'+(image.width*scale)+'px;height:'+(image.height*scale)+'px"></figure>';}):[];
  const notes=[...settings.notes,...(previews.other&&scope!=='charts'?['미리 보기를 지원하지 않는 이미지·도형 '+previews.other+'개는 인쇄에서 제외했습니다. 원본은 XLSX에 유지됩니다.']:[])];
  if(cachedCount)notes.push('직접 계산할 수 없는 수식 '+cachedCount+'개는 Excel 파일에 저장된 계산값을 사용합니다.');
  if(chartOnly&&!charts.length)throw Error('인쇄할 차트가 없습니다.');
  const style=[
    '@page{size:'+(settings.paper.name==='B5'?settings.paper.width+'mm '+settings.paper.height+'mm':settings.paper.name+' '+orientation)+';margin:'+['top','right','bottom','left'].map(k=>settings.margin[k]+'mm').join(' ')+'}*{box-sizing:border-box}body{margin:0;color:#222;font:11pt Arial}table{table-layout:fixed;border-collapse:collapse}td{padding:0;overflow:hidden;border:0}table.with-grid td{border:1px solid #ccc}.print-cell{padding:2px 4px;overflow:hidden}tr{break-inside:avoid}.print-page{break-after:page;overflow:hidden}.print-page:last-child{break-after:auto}.print-chart,.print-image{margin:18px 0;break-inside:avoid;max-width:100%;background:white}.print-chart{display:flex;flex-direction:column;border:1px solid #ddd;padding:12px 18px}.print-chart h2{font-size:13pt;text-align:center;margin:0 0 6px}.print-chart svg{display:block;flex:1;min-height:0;width:100%;overflow:visible}.print-note{font-size:9pt;color:#555}img{display:block}.print-objects{clear:both}'
  ].join('');
  const content=(pages.join('')||(!chartOnly?'<p>인쇄할 표시 셀이 없습니다.</p>':''))+(charts.length||images.length?'<section class="print-objects">'+charts.join('')+images.join('')+'</section>':'')+notes.map(note=>'<p class="print-note">'+esc(note)+'</p>').join('');
  return {html:'<!doctype html><html lang="ko"><head><meta charset="utf-8"><meta http-equiv="Content-Security-Policy" content="default-src \'none\'; style-src \'unsafe-inline\'; img-src data:"><title>'+esc(book.name)+' · '+esc(sheet.name)+'</title><style>'+style+'</style></head><body>'+content+'</body></html>',dataCells:plan.dataCells,pageCount:pages.length,chartCount:charts.length,imageCount:images.length,notes};
}
