import {cellBorders,edgeCss} from './borders.mjs';
import {tableToClipboard,tableFromHTML} from './studio-table.mjs';
import {address,formatValue} from './engine.mjs';
import {scalarRaw} from './cell-paste.mjs';
export function copyTable(clip,{columnWidths=[],rowHeights=[],displayStyles}={}){
  const cells=clip.data.map((row,r)=>row.map((cell,c)=>{
    const style={...cell.style,...displayStyles?.(r,c)},borders={};
    for(const [side,edge] of Object.entries(cellBorders(style))){const [width,line,color]=edgeCss(edge).split(' ');borders[side]=edge?{width:parseFloat(width),line,color}:null;}
    return {text:formatValue(clip.values[r][c],style),style:{...style,borders,align:style.align||(typeof clip.values[r][c]==='number'?'right':'left')}};
  }));
  const merges=clip.merges.map(ref=>{const [a,b]=ref.split(':').map(key=>{const m=key.match(/^([A-Z]+)(\d+)$/);let c=0;for(const letter of m[1])c=c*26+letter.charCodeAt(0)-64;return {r:Number(m[2])-1,c:c-1};});return {row:a.r-clip.r,column:a.c-clip.c,rowSpan:b.r-a.r+1,columnSpan:b.c-a.c+1};});
  return tableToClipboard({cells,merges,columnWidths,rowHeights});
}
export function spreadsheetTable(html){
  const table=tableFromHTML(html);if(!table)return null;
  return {r:0,c:0,data:table.cells.map(row=>row.map(cell=>{
    const style={...cell.style,border:false,borders:{}};for(const [side,edge] of Object.entries(cell.style.borders||{}))style.borders[side]=edge?{style:edge.line==='double'?'double':edge.line==='dotted'?'dotted':edge.line==='dashed'?'dashed':edge.width>=3?'thick':edge.width>=2?'medium':'thin',color:edge.color==='transparent'?'#ffffff':edge.color}:null;
    return {raw:scalarRaw(cell.text),style};
  })),rowHeights:table.rowHeights,merges:table.merges.map(m=>`${address(m.row,m.column)}:${address(m.row+m.rowSpan-1,m.column+m.columnSpan-1)}`)};
}
