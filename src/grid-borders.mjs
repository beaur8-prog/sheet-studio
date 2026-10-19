import {cellBorders,edgeCss} from './borders.mjs';

const gridEdge={style:'thin',color:'#d6d6d6'};
const lineRank={double:4,solid:3,dashed:2,dotted:1};
// Cells retain all four stored sides for XLSX/clipboard exchange. Only the view
// collapses coincident sides; splitting intervals also handles merged neighbours.
export function gridBorderSegments(cells,{showGrid=true}={}){
  const lines=new Map();
  for(const cell of cells){
    const {x,y,w,h,pane='',z=1,order=0,minX=0,minY=0,originX=0,originY=0}=cell,borders=cellBorders(cell.style);
    for(const side of ['top','right','bottom','left']){
      const explicit=!!borders[side],edge=borders[side]||(showGrid&&['right','bottom'].includes(side)?gridEdge:null);
      if(!edge)continue;
      const horizontal=side==='top'||side==='bottom',pos=Math.round(horizontal?y+(side==='bottom'?h:0):x+(side==='right'?w:0));
      const start=Math.round(horizontal?x:y),end=Math.round(horizontal?x+w:y+h);
      if(end<=start)continue;
      const css=edgeCss(edge),[width,line]=css.split(' '),weight=(explicit?100:0)+parseInt(width)*10+lineRank[line];
      const key=JSON.stringify([pane,horizontal,pos]);
      if(!lines.has(key))lines.set(key,[]);
      lines.get(key).push({horizontal,pos,start,end,css,width:parseInt(width),weight,order,z,originX:Math.round(originX),originY:Math.round(originY),minX:Math.ceil(minX),minY:Math.ceil(minY)});
    }
  }
  const result=[];
  for(const edges of lines.values()){
    const stops=[...new Set(edges.flatMap(e=>[e.start,e.end]))].sort((a,b)=>a-b);
    let previous;
    for(let i=0;i<stops.length-1;i++){
      const start=stops[i],end=stops[i+1];let winner;
      for(const e of edges)if(e.start<=start&&e.end>=end&&(!winner||e.weight>winner.weight||e.weight===winner.weight&&e.order>winner.order))winner=e;
      if(!winner){previous=undefined;continue;}
      if(previous&&previous.end===start&&previous.css===winner.css&&previous.z===winner.z&&previous.minX===winner.minX&&previous.minY===winner.minY&&previous.originX===winner.originX&&previous.originY===winner.originY)previous.end=end;
      else {previous={...winner,start,end};result.push(previous);}
    }
  }
  return result;
}

export function gridBordersHtml(cells,options){
  return gridBorderSegments(cells,options).map(e=>{
    // Keep strokes inside the upper/left cell, as the original gridlines were.
    // Sheet origins are the exception: their outer sides must stay below/right
    // of the headers. Extend the leading cap to join corners of equal weight.
    const x=e.horizontal?e.start-e.width:e.pos===e.originX?e.pos:e.pos-e.width;
    const y=e.horizontal?(e.pos===e.originY?e.pos:e.pos-e.width):e.start-e.width;
    const w=e.horizontal?e.end-e.start+e.width:e.width,h=e.horizontal?e.width:e.end-e.start+e.width;
    const clipX=Math.max(0,e.minX-x),clipY=Math.max(0,e.minY-y);
    if(clipX>=w||clipY>=h)return '';
    return `<i class="cell-border" aria-hidden="true" style="left:${x}px;top:${y}px;width:${w}px;height:${h}px;border-${e.horizontal?'top':'left'}:${e.css};clip-path:inset(${clipY}px 0 0 ${clipX}px);z-index:${e.z}"></i>`;
  }).join('');
}
