import {address,colName,colIndex} from './engine.mjs';
import {xml,stringify,child,children,attr} from './chart-xml.mjs';
import {tableBounds} from './table-schema.mjs';
import {readPrintSettings} from './print-layout.mjs';
import {HEADER_FIELDS,readHeaderFooter} from './print-headers.mjs';

const NS='http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const number=(value,name,min,max,{integer=false}={})=>{const n=Number(value);if(value===''||!Number.isFinite(n)||n<min||n>max||integer&&!Number.isInteger(n))throw Error(name+' 값이 올바르지 않습니다.');return n;};
const document=(features,name)=>xml(features[name]||'<'+name+' xmlns="'+NS+'"/>');
const set=(root,key,value)=>root.setAttribute(key,String(value));
const absolute=text=>{
  const b=tableBounds(text);if(b)return '$'+colName(b.c1)+'$'+(b.r1+1)+':$'+colName(b.c2)+'$'+(b.r2+1);
  const rows=/^(\d+):(\d+)$/.exec(text),cols=/^([A-Z]+):([A-Z]+)$/.exec(text);
  if(rows){const start=Number(rows[1]),end=Number(rows[2]);if(start>=1&&end>=start&&end<=1048576)return '$'+start+':$'+end;}
  if(cols){const start=colIndex(cols[1]),end=colIndex(cols[2]);if(start>=0&&end>=start&&end<16384)return '$'+cols[1]+':$'+cols[2];}
  throw Error('인쇄 영역은 A1:C30처럼 셀 범위로 입력하세요.');
};
function printNames(book,sheet,areas,repeatRows,repeatCols){
  const owner="'"+sheet.name.replace(/'/g,"''")+"'!",definitions=new Map();
  if(areas.trim()){
    const refs=areas.replace(/\$/g,'').toUpperCase().split(',').map(s=>s.trim());if(refs.length>32)throw Error('인쇄 영역은 최대 32개까지 지정할 수 있습니다.');
    definitions.set('_xlnm.Print_Area',refs.map(ref=>owner+absolute(ref)).join(','));
  }
  const titles=[];
  if(repeatRows.trim()){const ref=repeatRows.trim().replace(/\$/g,'');if(!/^\d+:\d+$/.test(ref))throw Error('반복 행은 1:2처럼 입력하세요.');titles.push(owner+absolute(ref));}
  if(repeatCols.trim()){const ref=repeatCols.trim().replace(/\$/g,'').toUpperCase();if(!/^[A-Z]+:[A-Z]+$/.test(ref))throw Error('반복 열은 A:B처럼 입력하세요.');titles.push(owner+absolute(ref));}
  if(titles.length)definitions.set('_xlnm.Print_Titles',titles.join(','));
  const result=[];
  for(const old of book.chartNames||[]){const key=[...definitions.keys(),'_xlnm.Print_Area','_xlnm.Print_Titles'].find(n=>n.toLowerCase()===old.name.toLowerCase());if(old.sheetId!==sheet.id||!key){result.push(old);continue;}if(definitions.has(key)){result.push({...old,formula:definitions.get(key)});definitions.delete(key);}}
  for(const [name,formula] of definitions)result.push({name,sheetId:sheet.id,formula});return result;
}
export function editablePrintSettings(book,sheet){
  let s,referenceWarning='';try{s=readPrintSettings(book,sheet);}catch(error){s=readPrintSettings(book,sheet,{ignorePrintArea:true,ignorePrintTitles:true});referenceWarning=error.message;}
  const localName=name=>{const item=(book.chartNames||[]).find(n=>n.sheetId===sheet.id&&n.name.toLowerCase()===name.toLowerCase());return (item?.formula||'').replace(/^=/,'').replace(/(^|,)\s*(?:'(?:[^']|'')*'|[\p{L}_][\p{L}\p{N}_.]*)!/gu,'$1').replace(/\$/g,'');};
  const rawTitles=referenceWarning?localName('_xlnm.Print_Titles').split(','):[],badTitles=rawTitles.filter(r=>r&&!/^(?:\d+:\d+|[A-Z]+:[A-Z]+)$/i.test(r));
  const chartSetup=sheet.kind==='chart'&&sheet.xlsxChartSheet?.xml?child(xml(sheet.xlsxChartSheet.xml).documentElement,'pageSetup'):null,raw=sheet.xlsxFeatures?.pageSetup||(chartSetup?stringify(chartSetup):'');
  return {orientation:s.orientation,paperId:s.paperId,mode:s.scale!==null?'scale':'fit',scale:s.scale===null?100:s.scale*100,fitWidth:s.fitWidth,fitHeight:s.fitHeight,margin:{...s.margin},areas:referenceWarning?localName('_xlnm.Print_Area'):s.areas.map(b=>address(b.r1,b.c1)+':'+address(b.r2,b.c2)).join(','),repeatRows:referenceWarning?rawTitles.find(r=>/^\d+:\d+$/.test(r))||badTitles.join(','):s.repeatRows?(s.repeatRows.start+1)+':'+(s.repeatRows.end+1):'',repeatCols:referenceWarning?rawTitles.find(r=>/^[A-Z]+:[A-Z]+$/i.test(r))||'':s.repeatCols?colName(s.repeatCols.start)+':'+colName(s.repeatCols.end):'',referenceWarning,gridLines:s.gridLines,pageOrder:s.pageOrder,firstPageNumber:s.firstPageNumber,automaticPageNumber:!raw||!['1','true'].includes(attr(xml(raw).documentElement,'useFirstPageNumber')),headerFooter:readHeaderFooter(sheet)};
}
/** Build an atomic metadata-only plan; validation never mutates the document. */
export function printSettingsPlan(book,sheet,values){
  if(book.nativeSource)throw Error('XLSM의 페이지 설정 저장은 Excel에서 변경하거나 XLSX 사본을 사용하세요.');
  const features={...sheet.xlsxFeatures},setup=document(features,'pageSetup'),props=document(features,'sheetPr'),margins=document(features,'pageMargins'),options=document(features,'printOptions');
  if(!['portrait','landscape'].includes(values.orientation))throw Error('용지 방향을 확인하세요.');
  const paperId=number(values.paperId,'용지 크기',0,10000,{integer:true});set(setup.documentElement,'paperSize',paperId);set(setup.documentElement,'orientation',values.orientation);
  let pageProps=child(props.documentElement,'pageSetUpPr');if(!pageProps){pageProps=props.createElementNS(NS,'pageSetUpPr');props.documentElement.appendChild(pageProps);}
  if(sheet.kind==='chart'){/* Chart sheets fit their chart to the output page. */}
  else if(values.mode==='fit'){
    set(pageProps,'fitToPage',1);set(setup.documentElement,'fitToWidth',number(values.fitWidth,'가로 페이지 수',0,32767,{integer:true}));set(setup.documentElement,'fitToHeight',number(values.fitHeight,'세로 페이지 수',0,32767,{integer:true}));setup.documentElement.removeAttribute('scale');
  }else if(values.mode==='scale'){
    set(pageProps,'fitToPage',0);set(setup.documentElement,'scale',number(values.scale,'인쇄 배율',10,400,{integer:true}));
  }else throw Error('배율 방식을 확인하세요.');
  for(const name of ['top','bottom','left','right','header','footer'])set(margins.documentElement,name,number(values.margin?.[name],name==='header'?'머리글 위치':name==='footer'?'바닥글 위치':'여백',0,500)/25.4);
  if(!['downThenOver','overThenDown'].includes(values.pageOrder))throw Error('페이지 순서를 확인하세요.');if(sheet.kind!=='chart')set(setup.documentElement,'pageOrder',values.pageOrder);set(setup.documentElement,'useFirstPageNumber',values.automaticPageNumber?0:1);set(setup.documentElement,'firstPageNumber',values.automaticPageNumber?1:number(values.firstPageNumber,'시작 페이지 번호',1,32767,{integer:true}));set(options.documentElement,'gridLines',values.gridLines?1:0);
  Object.assign(features,{pageSetup:stringify(setup),pageMargins:stringify(margins)});if(sheet.kind!=='chart')Object.assign(features,{sheetPr:stringify(props),printOptions:stringify(options)});
  const current=readHeaderFooter(sheet),header=document(features,'headerFooter');
  for(const field of HEADER_FIELDS){const value=String(values.headerFooter?.[field]??current[field]);if(value.length>255&&value!==current[field])throw Error('각 머리글·바닥글은 서식과 필드를 포함해 255자까지 입력할 수 있습니다.');let item=child(header.documentElement,field);if(!item){item=header.createElementNS(NS,field);header.documentElement.insertBefore(item,children(header.documentElement).find(n=>HEADER_FIELDS.indexOf(n.localName)>HEADER_FIELDS.indexOf(field))||null);}item.textContent=value;}
  for(const key of ['differentFirst','differentOddEven','scaleWithDoc','alignWithMargins'])set(header.documentElement,key,(values.headerFooter?.[key]??current[key])?1:0);features.headerFooter=stringify(header);
  const chartNames=sheet.kind==='chart'?[...(book.chartNames||[])]:printNames(book,sheet,String(values.areas||''),String(values.repeatRows||''),String(values.repeatCols||'')),target={...sheet,xlsxFeatures:features},snapshot={...book,chartNames,sheets:book.sheets.map(s=>s.id===sheet.id?target:s)};
  const checked=readPrintSettings(snapshot,target);if(HEADER_FIELDS.filter(f=>f.endsWith('Header')).some(f=>checked.headerFooter[f])&&checked.margin.header>checked.margin.top||HEADER_FIELDS.filter(f=>f.endsWith('Footer')).some(f=>checked.headerFooter[f])&&checked.margin.footer>checked.margin.bottom)throw Error('머리글·바닥글 위치는 위·아래 여백보다 작거나 같아야 합니다.');
  return {features,chartNames};
}
