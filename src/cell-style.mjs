import {tableCellStyle} from './table-style.mjs';
const fontFields=['fontFamily','fontSize','bold','italic','underline','color'],alignFields=['align','verticalAlign','wrap'];
export const normalCellStyle=book=>book.xlsxStyles?.[book.xlsxDefaultStyle]?.baseline||{};
export function inheritedStyles(book,sheet,r,c){if(sheet.cells?.[cellRef(r,c)]?.xlsxNormal)return [];return [sheet.xlsxStyleDefaults?.columns?.[c],sheet.xlsxStyleDefaults?.rows?.[r]].map(i=>book.xlsxStyles?.[i]).filter(Boolean);}
function cellRef(r,c){let name='';for(let n=c+1;n;n=Math.floor((n-1)/26))name=String.fromCharCode(65+(n-1)%26)+name;return name+(r+1);}
function templateStyle(base,entry,normal){
  const next={...base},style=entry.style||{},visible=entry.baseline||{};
  if(style.font)for(const key of fontFields)next[key]=visible[key]??(['bold','italic','underline'].includes(key)?false:key==='color'?'#000000':normal[key]);
  if(style.alignment)for(const key of alignFields)next[key]=visible[key];
  if(style.fill)next.background=visible.background;
  if(style.border){next.border=false;next.borders=visible.borders;}
  return {...next,...visible};
}
export function worksheetCellStyle(book,sheet,r,c,{table=true,cell=true}={}){
  const normal=normalCellStyle(book),record=sheet.cells?.[cellRef(r,c)],original=cell?book.xlsxStyles?.[record?.xlsxStyle]:null;
  const base=original?templateStyle(normal,original,normal):normal;
  if(original&&!original.style?.numFmt){base.format='general';base.numberFormat='General';}
  let style={...base,...(table?tableCellStyle(book,sheet,r,c):{})};
  // An explicit cell XF already contains its resolved font/fill/border model.
  if(!cell||record?.xlsxStyle===undefined&&!record?.xlsxNormal)for(const entry of inheritedStyles(book,sheet,r,c))style=templateStyle(style,entry,normal);
  return {...style,...(cell?record?.style:{})};
}
// Excel's "Automatic" font filter is distinct from explicitly painted black.
export function worksheetFilterStyle(book,sheet,r,c){
  const style=worksheetCellStyle(book,sheet,r,c),record=sheet.cells?.[cellRef(r,c)],original=book.xlsxStyles?.[record?.xlsxStyle],explicit=record?.style?.color;
  if(explicit!==undefined&&(!original||explicit!==original.baseline?.color))return style;
  const templates=original||record?.xlsxNormal?[]:inheritedStyles(book,sheet,r,c),font=original?.style?.font||templates.toReversed().find(t=>t.style?.font)?.style.font||book.xlsxStyles?.[book.xlsxDefaultStyle]?.style?.font,color=font?.color;
  const automatic=!color||color.auto||color.theme===1&&!color.tint;
  return automatic&&/^#000000$/i.test(style.color||'#000000')?{...style,color:undefined}:style;
}
