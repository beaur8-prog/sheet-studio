import {renderPageHeaders,HEADER_FIELDS} from './print-headers.mjs';
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
import {readPrintSettings,planPrintPages,checkPrintRange,printRowHeight,printColWidth,MAX_PRINT_PAGES} from './print-layout.mjs';
import {printObjectAxes,objectCellBounds,printObjectFragments} from './print-objects.mjs';
import {drawingPrintsWithSheet} from './anchor-position.mjs';

const esc=value=>String(value??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
function printRange(sheet,selection,spills=[],objects=[]){
  if(selection){
    checkPrintRange(selection);
    return {...selection};
  }
  let range=null;
  const include=b=>{if(!b)return;if(!range)range={...b};else {range.r1=Math.min(range.r1,b.r1);range.c1=Math.min(range.c1,b.c1);range.r2=Math.max(range.r2,b.r2);range.c2=Math.max(range.c2,b.c2);}};
  for(const ref in sheet.cells){const cell=sheet.cells[ref];if(cell.raw===''&&!Object.keys(cell.style||{}).length)continue;const p=parseAddress(ref);if(p)include({r1:p.r,r2:p.r,c1:p.c,c2:p.c});}
  for(const table of sheet.tables||[])include(tableBounds(table.range));
  for(const array of [...storedArrayIndex(sheet).regions,...spills])include(array);
  for(const bounds of objects)include(bounds);
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
/** Immutable output using one calculation session. Drawings can retain their
 * worksheet positions or print on individual object pages. */
export function buildSheetPrintDocument(book,sheetId,{scope='sheet',selection,orientation,gridLines,ignorePrintArea=false,includeCharts=true,includeImages=true,objectLayout='sheet',evaluate,now=new Date()}={}){
  const sheet=book.sheets.find(s=>s.id===sheetId);if(!sheet)throw Error('인쇄할 시트를 찾을 수 없습니다.');
  const settings=readPrintSettings(book,sheet,{orientation,ignorePrintArea:ignorePrintArea||scope!=='sheet',ignorePrintTitles:scope!=='sheet'});orientation=settings.orientation;gridLines??=settings.gridLines;
  if(!['sheet','selection','charts'].includes(scope)||!['portrait','landscape'].includes(orientation)||!['sheet','pages'].includes(objectLayout))throw Error('인쇄 설정이 올바르지 않습니다.');
  if(scope==='selection'&&!selection)throw Error('인쇄할 선택 영역을 지정하세요.');
  const {contentWidth,contentHeight}=settings,calculate=evaluate||createEvaluator(book,sheetId),chartOnly=scope==='charts'||sheet.kind==='chart';
  const positioned=objectLayout==='sheet'&&!chartOnly,axes=positioned?printObjectAxes(sheet):null,previews=drawingPreviews(sheet),printedObjects=new Set();let objectFragments=0,objectCharacters=0;
  const acceptObject=html=>{if(++objectFragments>2000||(objectCharacters+=html.length)>32000000)throw Error('개체의 인쇄 분량이 너무 큽니다. 인쇄 영역을 줄이거나 차트·이미지 포함을 해제하세요.');};
  const charts=(includeCharts||chartOnly?sheet.charts||[]:[]).filter(chart=>chartOnly||drawingPrintsWithSheet(chart.source?.anchor));
  const images=includeImages&&scope!=='charts'?previews.images.filter(image=>image.printWithSheet!==false):[];
  const chartHTML=(chart,width,height)=>{const plot=renderChartSVG(chart,sheet,book);return '<figure class="print-chart" data-chart="'+esc(chart.id)+'" style="width:'+width+'px;height:'+height+'px"><h2>'+esc(chart.title||'차트')+'</h2><svg xmlns="http://www.w3.org/2000/svg" viewBox="'+esc(plot.viewBox)+'" role="img" aria-label="'+esc(chart.title||'차트')+'">'+plot.svg+'</svg></figure>';};
  const imageHTML=(image,width,height)=>'<figure class="print-image" style="width:'+width+'px;height:'+height+'px"><img src="'+esc(image.src)+'" alt="'+esc(image.name)+'" style="width:100%;height:100%"></figure>';
  const objects=[...charts.map(chart=>({kind:'chart',rect:chartRect(chart),item:chart})),...images.map(image=>({kind:'image',rect:image,item:image}))];
  const fullObjectHTML=object=>object.html??=(object.kind==='chart'?chartHTML(object.item,object.rect.width,object.rect.height):imageHTML(object.item,object.rect.width,object.rect.height));
  // Selection printing deliberately ignores saved areas and repeating titles.
  if(scope==='selection'){settings.repeatRows=null;settings.repeatCols=null;}
  const ranges=chartOnly?[]:settings.areas.length?settings.areas:[printRange(sheet,scope==='selection'?selection:null,calculate.spills?.()||[],positioned?objects.map(item=>objectCellBounds(item.rect,axes)):[])],plan=planPrintPages(sheet,ranges,settings);
  if(!positioned&&plan.pages.length+objects.length>MAX_PRINT_PAGES)throw Error('차트·이미지를 포함한 인쇄 페이지가 500개를 넘습니다. 인쇄 대상을 줄이세요.');
  let cachedCount=0;
  const pages=plan.pages.map(page=>{
    const grid=gridHTML(book,sheet,positioned?{...page,scale:1}:page,{gridLines,evaluate:calculate});cachedCount+=grid.cachedCount||0;if(!positioned)return {html:grid.html,scale:page.scale};
    const fragments=objects.map((item,index)=>({item,index})).sort((a,b)=>(a.item.kind==='chart')-(b.item.kind==='chart')).map(({item,index})=>printObjectFragments(item.rect,page,axes).map(f=>{const html=fullObjectHTML(item);acceptObject(html);printedObjects.add(index);return '<div class="print-object-clip" data-object="'+index+'" style="left:'+f.x+'px;top:'+f.y+'px;width:'+f.width+'px;height:'+f.height+'px"><div style="position:absolute;left:'+f.contentX+'px;top:'+f.contentY+'px">'+html+'</div></div>';}).join('')).join('');
    return {html:'<div class="print-sheet" style="zoom:'+page.scale+';width:'+page.cols.reduce((n,c)=>n+printColWidth(sheet,c),0)+'px;height:'+page.rows.reduce((n,r)=>n+printRowHeight(sheet,r),0)+'px">'+grid.html+fragments+'</div>',scale:page.scale};
  });
  const notes=[...settings.notes,...(previews.other&&scope!=='charts'?['미리 보기를 지원하지 않는 이미지·도형 '+previews.other+'개는 인쇄에서 제외했습니다. 원본은 XLSX에 유지됩니다.']:[])];
  if(cachedCount)notes.push('직접 계산할 수 없는 수식 '+cachedCount+'개는 Excel 파일에 저장된 계산값을 사용합니다.');
  if(chartOnly&&!charts.length)throw Error('인쇄할 차트가 없습니다.');
  const objectPages=positioned?[]:objects.map((object,index)=>{const {rect,item}=object,scale=Math.min(1,contentWidth/rect.width,Math.max(20,contentHeight-40)/rect.height),html=object.kind==='chart'?chartHTML(item,rect.width*scale,rect.height*scale):imageHTML(item,rect.width*scale,rect.height*scale);acceptObject(html);printedObjects.add(index);return {html,scale:1,object:true};});
  const bodies=[...pages,...objectPages];
  if(!bodies.length)bodies.push({html:'<p>인쇄할 표시 셀이 없습니다.</p>',scale:1});
  if(bodies.length>MAX_PRINT_PAGES)throw Error('차트·이미지를 포함한 인쇄 페이지가 500개를 넘습니다. 인쇄 대상을 줄이세요.');
  for(const [kind,margin] of [['Header','top'],['Footer','bottom']])if(HEADER_FIELDS.filter(f=>f.endsWith(kind)).some(f=>settings.headerFooter[f])&&settings.margin[kind.toLowerCase()]>settings.margin[margin])throw Error('머리글·바닥글 위치가 인쇄 여백 밖에 있습니다. 페이지 설정에서 위치와 여백을 조정하세요.');
  const printed=bodies.map((body,index)=>{const headers=renderPageHeaders(settings.headerFooter,{book,sheet,index,pageNumber:settings.firstPageNumber+index,totalPages:bodies.length,now,scale:settings.headerFooter.scaleWithDoc?body.scale:1});for(const note of headers.notes)if(!notes.includes(note))notes.push(note);return '<section class="print-page'+(body.object?' print-object-page':'')+'" data-page="'+(index+1)+'" style="height:'+(settings.paper.height*96/25.4-1)+'px;width:'+(settings.paper.width*96/25.4)+'px">'+headers.header+'<div class="print-page-body">'+body.html+'</div>'+headers.footer+'</section>';});
  const mmToPx=96/25.4,m= settings.margin,headerSide='left:'+(settings.headerFooter.alignWithMargins?m.left:7.62)+'mm;right:'+(settings.headerFooter.alignWithMargins?m.right:7.62)+'mm;';

  const style=[
    '@page{size:'+(settings.paper.name==='B5'?settings.paper.width+'mm '+settings.paper.height+'mm':settings.paper.name+' '+orientation)+';margin:0}*{box-sizing:border-box}body{margin:0;color:#222;font:11pt Arial}table{table-layout:fixed;border-collapse:collapse}td{padding:0;overflow:hidden;border:0}table.with-grid td{border:1px solid #ccc}.print-cell{padding:2px 4px;overflow:hidden}tr{break-inside:avoid}.print-page{position:relative;break-after:page;padding:'+['top','right','bottom','left'].map(k=>settings.margin[k]+'mm').join(' ')+';}.print-page:last-of-type{break-after:auto}.print-page-body{height:100%;overflow:hidden}.print-object-page .print-page-body{display:flex;justify-content:center;align-items:flex-start}.print-object-page figure{margin:0}.print-chart,.print-image{margin:18px 0;break-inside:avoid;max-width:100%;background:white}.print-chart{display:flex;flex-direction:column;border:1px solid #ddd;padding:12px 18px}.print-chart h2{font-size:13pt;text-align:center;margin:0 0 6px}.print-chart svg{display:block;flex:1;min-height:0;width:100%;overflow:visible}.print-note{font-size:9pt;color:#555}img{display:block}.print-objects{clear:both}'
    ,'.print-sheet{position:relative}.print-object-clip{position:absolute;overflow:hidden}.print-object-clip figure{margin:0;max-width:none}'
    ,'.print-header,.print-footer{position:absolute;display:grid;grid-template-columns:repeat(3,minmax(0,1fr));gap:8px;'+headerSide+'}.print-header{top:'+m.header+'mm;height:'+Math.max(0,(m.top-m.header)*mmToPx)+'px;align-items:start}.print-footer{bottom:'+m.footer+'mm;height:'+Math.max(0,(m.bottom-m.footer)*mmToPx)+'px;align-items:end}.print-header-cell{font-size:0;white-space:pre;max-height:100%;overflow:hidden;line-height:1.15}.print-header-cell.center{text-align:center}.print-header-cell.right{text-align:right}@media print{.print-notes{display:none}}'
  ].join('');
  const content=printed.join('')+(notes.length?'<aside class="print-notes">'+notes.map(note=>'<p class="print-note">'+esc(note)+'</p>').join('')+'</aside>':'');
  return {html:'<!doctype html><html lang="ko"><head><meta charset="utf-8"><meta http-equiv="Content-Security-Policy" content="default-src \'none\'; style-src \'unsafe-inline\'; img-src data:"><title>'+esc(book.name)+' · '+esc(sheet.name)+'</title><style>'+style+'</style></head><body>'+content+'</body></html>',dataCells:plan.dataCells,pageCount:bodies.length,chartCount:[...printedObjects].filter(i=>objects[i].kind==='chart').length,imageCount:[...printedObjects].filter(i=>objects[i].kind==='image').length,notes};
}
