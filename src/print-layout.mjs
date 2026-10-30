import {colIndex} from './engine.mjs';
import {xml,child,children,attr} from './chart-xml.mjs';
import {filteredRowsForSheet,tableBounds} from './table-schema.mjs';

const mmToPx=96/25.4;
const truth=value=>value==='1'||value==='true';
const papers={1:{name:'Letter',width:215.9,height:279.4},5:{name:'Legal',width:215.9,height:355.6},8:{name:'A3',width:297,height:420},9:{name:'A4',width:210,height:297},11:{name:'A5',width:148,height:210},13:{name:'B5',width:182,height:257}};
export const MAX_PRINT_CELLS=20000;
export const MAX_PRINT_PAGES=500;
const limitMessage='셀 인쇄는 한 번에 20,000셀까지 지원합니다. 인쇄창에서 선택 영역을 사용하거나 차트만 인쇄하세요.';
export function checkPrintRange(range){
  if(!range)return;
  if(!['r1','r2','c1','c2'].every(k=>Number.isInteger(range[k]))||range.r1<0||range.c1<0||range.r2<range.r1||range.c2<range.c1||range.r2>=1048576||range.c2>=16384)throw Error('인쇄 범위가 올바르지 않습니다.');
  if((range.r2-range.r1+1)*(range.c2-range.c1+1)>MAX_PRINT_CELLS)throw Error(limitMessage);
}
function feature(sheet,name){
  if(sheet.xlsxFeatures?.[name])return xml(sheet.xlsxFeatures[name]).documentElement;
  if(sheet.kind==='chart'&&sheet.xlsxChartSheet?.xml)return child(xml(sheet.xlsxChartSheet.xml).documentElement,name);
  return null;
}
function number(node,name,fallback,min,max){
  const text=node&&attr(node,name),value=text!==''&&text!=null?Number(text):fallback;
  if(!Number.isFinite(value)||value<min||value>max)throw Error('저장된 인쇄 설정이 올바르지 않습니다: '+name);
  return value;
}
// Commas and exclamation marks inside quoted sheet names are ordinary text.
function splitReferences(formula){
  const refs=[];let start=0,quoted=false;
  for(let i=0;i<formula.length;i++){
    if(formula[i]==="'"){if(quoted&&formula[i+1]==="'")i++;else quoted=!quoted;}
    else if(formula[i]===','&&!quoted){refs.push(formula.slice(start,i));start=i+1;}
  }
  if(quoted)throw Error('저장된 인쇄 참조가 올바르지 않습니다.');
  refs.push(formula.slice(start));return refs;
}
function localReference(text,sheet){
  text=text.trim();let bang=-1,quoted=false;
  for(let i=0;i<text.length;i++)if(text[i]==="'"){if(quoted&&text[i+1]==="'")i++;else quoted=!quoted;}else if(text[i]==='!'&&!quoted)bang=i;
  if(bang>=0){const owner=text.slice(0,bang),name=owner.startsWith("'")?owner.slice(1,-1).replace(/''/g,"'"):owner;if(name.toLocaleLowerCase()!==sheet.name.toLocaleLowerCase())throw Error('다른 시트를 참조하는 인쇄 영역은 지원하지 않습니다.');text=text.slice(bang+1);}
  return text.replace(/\$/g,'').toUpperCase();
}
function namedReferences(book,sheet,name){
  const local=(book.chartNames||[]).find(n=>n.name.toLowerCase()===name.toLowerCase()&&n.sheetId===sheet.id);
  if(!local)return [];
  return splitReferences(local.formula.replace(/^=/,'')).map(ref=>localReference(ref,sheet));
}
function area(text){
  let range=tableBounds(text);if(range)return range;
  const row=/^(\d+):(\d+)$/.exec(text),col=/^([A-Z]+):([A-Z]+)$/.exec(text);
  if(row)range={r1:Number(row[1])-1,r2:Number(row[2])-1,c1:0,c2:16383};
  if(col)range={r1:0,r2:1048575,c1:colIndex(col[1]),c2:colIndex(col[2])};
  if(!range)throw Error('저장된 인쇄 영역을 해석할 수 없습니다. 선택 영역을 인쇄하거나 인쇄 영역 무시를 선택하세요.');
  return range;
}
export function readPrintSettings(book,sheet,{orientation,ignorePrintArea=false,ignorePrintTitles=false}={}){
  const setup=feature(sheet,'pageSetup'),margins=feature(sheet,'pageMargins'),options=feature(sheet,'printOptions'),props=child(feature(sheet,'sheetPr'),'pageSetUpPr'),notes=[];
  orientation??=setup&&attr(setup,'orientation')==='landscape'?'landscape':'portrait';
  if(!['portrait','landscape'].includes(orientation))throw Error('인쇄 방향이 올바르지 않습니다.');
  const paperId=number(setup,'paperSize',9,0,10000),paper={...(papers[paperId]||papers[9])};
  if(!papers[paperId])notes.push('지원하지 않는 용지 크기('+paperId+')는 이번 출력에서 A4로 사용합니다. 원래 설정은 파일에 유지됩니다.');
  if(orientation==='landscape')[paper.width,paper.height]=[paper.height,paper.width];
  const margin=Object.fromEntries(['top','right','bottom','left'].map(k=>[k,Number((number(margins,k,12/25.4,0,100)*25.4).toFixed(6))]));
  const contentWidth=(paper.width-margin.left-margin.right)*mmToPx,contentHeight=(paper.height-margin.top-margin.bottom)*mmToPx;
  if(contentWidth<10*mmToPx||contentHeight<10*mmToPx)throw Error('용지보다 여백이 큽니다. Excel의 페이지 설정을 확인하세요.');
  const fit=truth(props&&attr(props,'fitToPage'));
  // New documents retain the app's existing shrink-to-width default. Imported
  // explicit scale and fit settings take precedence over that default.
  const scale=setup&&attr(setup,'scale')!==''?number(setup,'scale',100,10,400)/100:null;
  const fitWidth=fit?number(setup,'fitToWidth',1,0,32767):!setup?1:0,fitHeight=fit?number(setup,'fitToHeight',1,0,32767):0;
  const areas=ignorePrintArea?[]:namedReferences(book,sheet,'_xlnm.Print_Area').map(area);
  if(areas.length>32)throw Error('인쇄 영역은 한 번에 32개까지 지원합니다.');
  let repeatRows=null,repeatCols=null;
  for(const ref of ignorePrintTitles?[]:namedReferences(book,sheet,'_xlnm.Print_Titles')){
    const row=/^(\d+):(\d+)$/.exec(ref),col=/^([A-Z]+):([A-Z]+)$/.exec(ref);
    if(row){const start=Number(row[1])-1,end=Number(row[2])-1;if(start<0||end<start||end>=1048576)throw Error('인쇄 반복 행이 올바르지 않습니다.');repeatRows={start,end};}
    else if(col){const start=colIndex(col[1]),end=colIndex(col[2]);if(start<0||end<start||end>=16384)throw Error('인쇄 반복 열이 올바르지 않습니다.');repeatCols={start,end};}
    else throw Error('저장된 인쇄 반복 제목을 해석할 수 없습니다.');
  }
  const breaks=name=>children(feature(sheet,name),'brk').filter(n=>truth(attr(n,'man'))).map(n=>({id:number(n,'id',0,0,1048576),min:number(n,'min',0,0,1048576),max:number(n,'max',1048576,0,1048576)}));
  return {orientation,paper,margin,contentWidth,contentHeight,scale:fit?null:scale,fitToPage:fit,fitWidth,fitHeight,areas,repeatRows,repeatCols,rowBreaks:breaks('rowBreaks'),colBreaks:breaks('colBreaks'),pageOrder:setup&&attr(setup,'pageOrder')==='overThenDown'?'overThenDown':'downThenOver',gridLines:truth(options&&attr(options,'gridLines')),notes};
}
export const printRowHeight=(sheet,r)=>sheet.rowHeights?.[r]??sheet.defaultRowHeight??25;
export const printColWidth=(sheet,c)=>sheet.columnWidths?.[c]??sheet.defaultColumnWidth??90;
function visible(start,end,hidden){const values=[];for(let i=start;i<=end;i++)if(!hidden(i))values.push(i);return values;}
function axisPages(body,titles,size,capacity,breaks){
  const titleSet=new Set(titles),data=body.filter(i=>!titleSet.has(i)),titleSize=titles.reduce((n,i)=>n+size(i),0),pages=[];
  if(titleSize>=capacity&&data.length)throw Error('반복 제목이 한 페이지보다 큽니다. 반복 행·열 또는 배율을 줄이세요.');
  breaks=[...new Set(breaks)].sort((a,b)=>a-b);let current=[],used=titleSize,nextBreak=0,expanded=0;
  const emit=()=>{expanded+=titles.length+current.length;if(expanded>100000||pages.length>=MAX_PRINT_PAGES)throw Error('반복 제목과 페이지 나눔으로 인쇄 분량이 너무 커졌습니다. 인쇄 영역을 줄이세요.');pages.push([...titles,...current]);};
  for(const i of data){const next=size(i);if(next>capacity-titleSize+.01)throw Error('셀 크기가 인쇄 가능한 한 페이지보다 큽니다. 용지·배율 또는 셀 크기를 조정하세요.');
    let manual=false;while(nextBreak<breaks.length&&breaks[nextBreak]<=i){if(current.length&&breaks[nextBreak]>current.at(-1))manual=true;nextBreak++;}
    if(current.length&&(used+next>capacity+.01||manual)){emit();current=[];used=titleSize;}
    current.push(i);used+=next;
  }
  if(current.length||titles.length)emit();return pages;
}
function fitsAxis(body,titles,size,capacity,target){
  if(!target)return true;
  const titleSet=new Set(titles),titleSize=titles.reduce((n,i)=>n+size(i),0);let pages=1,used=titleSize,items=0;
  if(titleSize>capacity)return false;
  for(const i of body){if(titleSet.has(i))continue;const next=size(i);if(titleSize+next>capacity)return false;
    if(items&&used+next>capacity){if(++pages>target)return false;used=titleSize;items=0;}used+=next;items++;
  }
  return true;
}
export function planPrintPages(sheet,ranges,settings){
  const hidden=filteredRowsForSheet(sheet),rowHidden=r=>hidden.has(r)||printRowHeight(sheet,r)===0,colHidden=c=>printColWidth(sheet,c)===0;
  let cells=0,renderedCells=0,dataCells=0;const pages=[];
  for(const range of ranges){
    checkPrintRange(range);if(!range)continue;
    cells+=(range.r2-range.r1+1)*(range.c2-range.c1+1);if(cells>MAX_PRINT_CELLS)throw Error(limitMessage);
    const rows=visible(range.r1,range.r2,rowHidden),cols=visible(range.c1,range.c2,colHidden);if(!rows.length||!cols.length)continue;dataCells+=rows.length*cols.length;
    if(settings.repeatRows&&settings.repeatRows.end-settings.repeatRows.start>=MAX_PRINT_CELLS||settings.repeatCols&&settings.repeatCols.end-settings.repeatCols.start>=MAX_PRINT_CELLS)throw Error('인쇄 반복 제목이 너무 큽니다.');
    const titleRows=settings.repeatRows?visible(settings.repeatRows.start,settings.repeatRows.end,rowHidden):[],titleCols=settings.repeatCols?visible(settings.repeatCols.start,settings.repeatCols.end,colHidden):[];
    const allRows=[...new Set([...titleRows,...rows])],allCols=[...new Set([...titleCols,...cols])];if(allRows.length*allCols.length>MAX_PRINT_CELLS)throw Error(limitMessage);
    const totalWidth=allCols.reduce((n,c)=>n+printColWidth(sheet,c),0),totalHeight=allRows.reduce((n,r)=>n+printRowHeight(sheet,r),0),availableWidth=settings.contentWidth-4,availableHeight=settings.contentHeight-4;
    let scale=settings.scale??1;
    if(settings.fitWidth)scale=Math.min(scale,settings.fitWidth*availableWidth/totalWidth);
    if(settings.fitHeight)scale=Math.min(scale,settings.fitHeight*availableHeight/totalHeight);
    // Repeated titles consume space on every page. A simple total-size ratio
    // can therefore exceed a requested two-page width/height. Find a scale
    // that satisfies the actual discrete row and column page counts.
    if(settings.fitWidth||settings.fitHeight){
      const fits=value=>fitsAxis(rows,titleRows,r=>printRowHeight(sheet,r),availableHeight/value,settings.fitHeight)&&fitsAxis(cols,titleCols,c=>printColWidth(sheet,c),availableWidth/value,settings.fitWidth);
      if(!fits(scale)){let low=0,high=scale;for(let i=0;i<28;i++){const mid=(low+high)/2;if(fits(mid))low=mid;else high=mid;}scale=low;}
    }
    if(scale<.01)throw Error('인쇄 배율이 너무 작습니다. 인쇄 영역 또는 페이지 맞춤을 조정하세요.');
    // Fit-to-page ignores manual breaks, as Excel does. Otherwise a partial
    // break is relevant only when it intersects this area's other axis.
    const activeBreaks=(breaks,min,max)=>settings.fitToPage?[]:breaks.filter(b=>b.min<=max&&b.max>=min).map(b=>b.id);
    const rp=axisPages(rows,titleRows,r=>printRowHeight(sheet,r),availableHeight/scale,activeBreaks(settings.rowBreaks,range.c1,range.c2));
    const cp=axisPages(cols,titleCols,c=>printColWidth(sheet,c),availableWidth/scale,activeBreaks(settings.colBreaks,range.r1,range.r2));
    const add=(rr,cc)=>{renderedCells+=rr.length*cc.length;if(renderedCells>100000||pages.length>=MAX_PRINT_PAGES)throw Error('반복 제목과 페이지 나눔으로 인쇄 분량이 너무 커졌습니다. 인쇄 영역을 줄이세요.');pages.push({rows:rr,cols:cc,allRows,allCols,scale,range});};
    if(settings.pageOrder==='overThenDown')for(const rr of rp)for(const cc of cp)add(rr,cc);else for(const cc of cp)for(const rr of rp)add(rr,cc);
  }
  return {pages,dataCells};
}
