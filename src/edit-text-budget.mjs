import {MAX_DOCUMENT_TEXT_CHARS} from './limits.mjs';
// Release all overwritten destinations first: a valid replacement must not
// depend on the order of long and short cells. Call before deep-cloning a cell.
export function createEditTextBudget(textChars,overwrittenChars=0){
 let total=Number.isSafeInteger(textChars)?Math.max(0,textChars-overwrittenChars):0;
 return raw=>{total+=String(raw??'').length;if(total>MAX_DOCUMENT_TEXT_CHARS)throw Error('문서의 셀 텍스트 총량이 100,000,000자를 초과합니다. 범위를 줄이세요.');};
}
