import {numberFormatKind} from './engine.mjs';
import {tableAt,tableBounds} from './table-schema.mjs';
import {xml,stringify,child,children,all,attr} from './chart-xml.mjs';
import {createColorResolver,xmlColor} from './spreadsheet-colors.mjs';
const NS='http://schemas.openxmlformats.org/spreadsheetml/2006/main',cache=new Map(),dxfCache=new Map(),colorCache=new Map();
export const tablePresets={green:{label:'녹색',header:'#107c41',stripe:'#e2efda'},blue:{label:'파랑',header:'#2f5597',stripe:'#d9e2f3'},gray:{label:'회색',header:'#50565d',stripe:'#e7e6e6'}};
function parse(source,read){let doc=cache.get(source);if(!doc){doc=xml(source);if(cache.size>500)cache.delete(cache.keys().next().value);cache.set(source,doc);}return read(doc);}
function dxfStyle(book,id){
  const source=book.xlsxDxfs?.[id];if(!source)return {};const key=source+'\0'+(book.chartTheme||'');if(dxfCache.has(key))return dxfCache.get(key);let resolve=colorCache.get(book.chartTheme||'');if(!resolve){resolve=createColorResolver(book.chartTheme||'',null);if(colorCache.size>10)colorCache.clear();colorCache.set(book.chartTheme||'',resolve);}
  const node=parse(source,doc=>doc.documentElement),font=child(node,'font'),fill=child(child(node,'fill'),'patternFill'),style={};
  const rgb=n=>xmlColor(n,resolve);if(attr(fill,'patternType')==='solid'&&rgb(child(fill,'fgColor')))style.background=rgb(child(fill,'fgColor'));
  if(rgb(child(font,'color')))style.color=rgb(child(font,'color'));
  for(const [tag,key] of [['b','bold'],['i','italic'],['u','underline']]){const part=child(font,tag);if(part)style[key]=!['0','false','none'].includes(attr(part,'val','1'));}
  const format=child(node,'numFmt');if(format){style.numberFormat=attr(format,'formatCode');style.format=numberFormatKind(style.numberFormat);}
  if(child(font,'sz'))style.fontSize=Number(attr(child(font,'sz'),'val'));if(child(font,'name'))style.fontFamily=attr(child(font,'name'),'val');if(dxfCache.size>500)dxfCache.delete(dxfCache.keys().next().value);dxfCache.set(key,style);return style;
}
function styleElements(book,table){
  const source=book.xlsxTableStyles?.find(source=>parse(source,doc=>attr(doc.documentElement,'name'))===table.style?.name);
  if(source)return parse(source,doc=>children(doc.documentElement,'tableStyleElement').map(node=>({type:attr(node,'type'),dxf:Number(attr(node,'dxfId')),size:Math.max(1,Number(attr(node,'size','1')))})));
  // Built-in styles are Excel resources, not embedded definitions. Keep their
  // exact name in XLSX and use a restrained preview in the local grid.
  return null;
}
export function tableCellStyle(book,sheet,r,c){
  const table=tableAt(sheet,r,c);if(!table||!table.style?.name)return {};const b=tableBounds(table.range),header=table.headerRow&&r===b.r1,total=table.totalsRow&&r===b.r2,style={},parts=styleElements(book,table);
  if(!parts){if(header)return {background:'#4472c4',color:'#ffffff',bold:true};if(total)return {background:'#d9e2f3',bold:true};if(table.style?.showRowStripes&&(r-b.r1-(table.headerRow?1:0))%2===0)style.background='#eef3fa';if(table.style?.showFirstColumn&&c===b.c1||table.style?.showLastColumn&&c===b.c2)style.bold=true;return style;}
  const get=type=>parts.find(part=>part.type===type),apply=type=>{const part=get(type);if(part)Object.assign(style,dxfStyle(book,part.dxf));};apply('wholeTable');
  if(!header&&!total){
    if(table.style?.showRowStripes){const a=get('firstRowStripe')?.size||1,z=get('secondRowStripe')?.size||1;apply((r-b.r1-(table.headerRow?1:0))%(a+z)<a?'firstRowStripe':'secondRowStripe');}
    if(table.style?.showColumnStripes){const a=get('firstColumnStripe')?.size||1,z=get('secondColumnStripe')?.size||1;apply((c-b.c1)%(a+z)<a?'firstColumnStripe':'secondColumnStripe');}
    if(table.style?.showFirstColumn&&c===b.c1)apply('firstColumn');if(table.style?.showLastColumn&&c===b.c2)apply('lastColumn');
  }
  if(header){apply('headerRow');if(c===b.c1)apply('firstHeaderCell');if(c===b.c2)apply('lastHeaderCell');}
  if(total){apply('totalRow');if(c===b.c1)apply('firstTotalCell');if(c===b.c2)apply('lastTotalCell');}return style;
}
export function ensureTablePreset(book,key){
  const preset=tablePresets[key];if(!preset)throw Error('표 스타일을 선택하세요.');book.xlsxDxfs||=[];book.xlsxTableStyles||=[];
  const base='SheetStudio'+key[0].toUpperCase()+key.slice(1)+'V1';for(const source of book.xlsxTableStyles){const doc=parse(source,doc=>doc),name=attr(doc.documentElement,'name');if(!name.startsWith(base))continue;const parts=children(doc.documentElement,'tableStyleElement'),header=parts.find(p=>attr(p,'type')==='headerRow'),stripe=parts.find(p=>attr(p,'type')==='firstRowStripe');if(dxfStyle(book,Number(attr(header,'dxfId'))).background===preset.header&&dxfStyle(book,Number(attr(stripe,'dxfId'))).background===preset.stripe)return name;}
  const style=(fill,font)=>`<dxf xmlns="${NS}">${font?`<font>${font.bold?'<b/>':''}${font.color?`<color rgb="FF${font.color.slice(1).toUpperCase()}"/>`:''}</font>`:''}${fill?`<fill><patternFill patternType="solid"><fgColor rgb="FF${fill.slice(1).toUpperCase()}"/><bgColor indexed="64"/></patternFill></fill>`:''}</dxf>`;
  const add=source=>{let id=book.xlsxDxfs.indexOf(source);if(id<0){id=book.xlsxDxfs.length;book.xlsxDxfs.push(source);}return id;};
  const header=add(style(preset.header,{color:'#ffffff',bold:true})),stripe=add(style(preset.stripe)),total=add(style(preset.stripe,{bold:true})),name='SheetStudio'+key[0].toUpperCase()+key.slice(1)+'V1';
  let candidate=name,index=1;const make=name=>`<tableStyle xmlns="${NS}" name="${name}" pivot="0" table="1" count="3"><tableStyleElement type="headerRow" dxfId="${header}"/><tableStyleElement type="firstRowStripe" dxfId="${stripe}" size="1"/><tableStyleElement type="totalRow" dxfId="${total}"/></tableStyle>`;
  while(book.xlsxTableStyles.some(s=>parse(s,doc=>attr(doc.documentElement,'name'))===candidate&&s!==make(candidate)))candidate=name+'_'+ ++index;
  const source=make(candidate);if(!book.xlsxTableStyles.includes(source))book.xlsxTableStyles.push(source);return candidate;
}
