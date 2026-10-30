import {parseAddress} from './engine.mjs';
import {createColorResolver} from './spreadsheet-colors.mjs';
const message='XLSM 저장 경로가 이 셀의 상세 서식·링크·문자별 서식 이동을 지원하지 않습니다. 값/수식만 붙여넣거나 XLSX 사본을 사용하세요.';
const ordered=value=>Array.isArray(value)?value.map(ordered):value&&typeof value==='object'?Object.fromEntries(Object.keys(value).sort().filter(k=>value[k]!==undefined).map(k=>[k,ordered(value[k])])):value;
const signature=value=>JSON.stringify(ordered(value));
function fullStyle(book,sheet,ref,cell,styles){
 const point=parseAddress(ref),original=styles?.[cell.xlsxStyle],normal=book.xlsxStyles?.[book.xlsxDefaultStyle]?.style||{};
 const templates=original||cell.xlsxNormal?[]:[sheet.xlsxStyleDefaults?.columns?.[point.c],sheet.xlsxStyleDefaults?.rows?.[point.r]].map(id=>book.xlsxStyles?.[id]?.style);
 return {...Object.assign({},normal,...templates,original?.style),...(cell.xlsxProtection?{protection:cell.xlsxProtection}:{})};
}
function unsupportedFeatures(book,sheet,ref,cell,styles,resolve){
 const full=fullStyle(book,sheet,ref,cell,styles),font=Object.fromEntries(['strike','vertAlign','outline','shadow','condense','extend'].filter(key=>full.font?.[key]&&full.font[key]!=='baseline').map(key=>[key,full.font[key]])),alignment=Object.fromEntries(Object.entries(full.alignment||{}).filter(([key,value])=>!['horizontal','vertical','wrapText'].includes(key)&&value!==undefined&&value!==false&&value!==0));
 const colors=value=>{if(Array.isArray(value))return value.map(colors);if(!value||typeof value!=='object')return value;if(['argb','rgb','theme','indexed','auto'].some(key=>Object.hasOwn(value,key)))return resolve(value)||value;return Object.fromEntries(Object.entries(value).map(([key,item])=>[key,Array.isArray(item)?item.map(colors):colors(item)]));};
 const fill=full.fill?.type==='gradient'||full.fill?.type==='pattern'&&!['none','solid'].includes(full.fill.pattern)?colors(full.fill):undefined,diagonal=full.border?.diagonal;
 const payload=cell.xlsxValue?.value;
 return {font,alignment,fill,...(diagonal?.style?{diagonal:{...diagonal,color:diagonal.color?resolve(diagonal.color)||diagonal.color:undefined}}:{}),protection:{locked:full.protection?.locked!==false,hidden:full.protection?.hidden===true},...(payload?.richText?{richText:colors(payload.richText)}:{}),...(payload?.hyperlink?{hyperlink:payload.hyperlink,tooltip:payload.tooltip}: {})};
}
export function nativeCellTransferError(book,sheet,ref,old,now,{styles=book.xlsxStyles,resolve}={}){
 if(!book.nativeSource)return null;
 if(old.xlsxStyle===now.xlsxStyle&&old.xlsxNormal===now.xlsxNormal&&signature(old.xlsxProtection)===signature(now.xlsxProtection)&&signature(old.xlsxValue?.value)===signature(now.xlsxValue?.value))return null;
 resolve||=createColorResolver(book.chartTheme,book.xlsxColors?'<styleSheet>'+book.xlsxColors+'</styleSheet>':undefined);
 return signature(unsupportedFeatures(book,sheet,ref,old,book.xlsxStyles,resolve))===signature(unsupportedFeatures(book,sheet,ref,now,styles,resolve))?null:message;
}
export function assertNativeCellTransfers(book,sheet,updates,options={}){
 if(!book.nativeSource)return;
 const resolve=createColorResolver(book.chartTheme,book.xlsxColors?'<styleSheet>'+book.xlsxColors+'</styleSheet>':undefined);
 for(const {key,cell} of updates){const error=nativeCellTransferError(book,sheet,key,sheet.cells[key]||{raw:''},cell||{raw:''},{...options,resolve});if(error)throw Error(error);}
}
