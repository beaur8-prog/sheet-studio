import {workbookMetadata} from './recovery-store.mjs';
import {estimateDataBytes} from './data-size.mjs';
import {MAX_HISTORY_BYTES} from './limits.mjs';
export function cellDescription(cells){
  let textLength=0,count=0,bytes=4;
  for(const ref in cells){count++;textLength+=cells[ref].raw.length;if(bytes<=MAX_HISTORY_BYTES)bytes+=estimateDataBytes(cells[ref],MAX_HISTORY_BYTES-bytes)+(ref.length+4)*2;}
  return {count,textLength,bytes:Math.min(bytes,MAX_HISTORY_BYTES+1)};
}
export function describeWorkbook(book){
  const sheets=book.sheets.map(sheet=>({id:sheet.id,...cellDescription(sheet.cells)}));
  return {bytes:Math.min(MAX_HISTORY_BYTES+1,estimateDataBytes(workbookMetadata(book),MAX_HISTORY_BYTES)+sheets.reduce((n,s)=>n+s.bytes+32,0)),sheets};
}
