import {cloneData} from './clone.mjs';
export function resetCellFormat(cell={raw:''},{protection}={}){
 const next={...cloneData(cell),style:{},xlsxNormal:true};delete next.xlsxStyle;delete next.xlsxProtection;
 if(protection!==undefined)next.xlsxProtection=cloneData(protection);
 if(next.xlsxValue?.value?.richText)next.xlsxValue.value.richText=next.xlsxValue.value.richText.map(({text})=>({text}));
 return next;
}
export function stripRichTextFormat(cell){
 if(cell.xlsxValue?.value?.richText)cell.xlsxValue.value.richText=cell.xlsxValue.value.richText.map(({text})=>({text}));return cell;
}
