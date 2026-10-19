export const borderSides=['top','right','bottom','left'];
export const borderStyles=['thin','medium','thick','dotted','dashed','double','hair','mediumDashed','dashDot','mediumDashDot','dashDotDot','mediumDashDotDot','slantDashDot'];
export function validateBorders(borders){
  if(!borders||typeof borders!=='object'||Array.isArray(borders)||Object.keys(borders).some(k=>!borderSides.includes(k)))throw Error('테두리 설정이 올바르지 않습니다.');
  for(const edge of Object.values(borders))if(edge!==null&&(!edge||typeof edge!=='object'||Array.isArray(edge)||Object.keys(edge).some(k=>!['style','color'].includes(k))||!borderStyles.includes(edge.style)||!/^#[\da-f]{6}$/i.test(edge.color)))throw Error('테두리 선 또는 색상이 올바르지 않습니다.');
}
export function cellBorders(style={}){
  const fallback=style.border?{style:'thin',color:'#8a8a8a'}:null;
  return Object.fromEntries(borderSides.map(side=>[side,style.borders&&Object.hasOwn(style.borders,side)?style.borders[side]:fallback]));
}
export function edgeCss(edge){
  if(!edge)return '';
  const width=edge.style==='thick'?3:edge.style==='medium'||edge.style.startsWith('medium')?2:edge.style==='double'?3:1;
  const line=edge.style==='double'?'double':edge.style==='dotted'?'dotted':/dash/i.test(edge.style)?'dashed':'solid';
  return `${width}px ${line} ${edge.color}`;
}
export function bordersCss(style){return Object.entries(cellBorders(style)).filter(([,edge])=>edge).map(([side,edge])=>`border-${side}:${edgeCss(edge)};`).join('');}
// Undefined decisions keep an existing edge. Null explicitly removes it.
export function borderPatch(style,cell,range,decisions,line,color){
  const borders=cellBorders(style),{r1,r2,c1,c2}=cell;
  const edge={style:line,color};
  for(const side of borderSides){
    const atEdge=side==='top'?r1===range.r1:side==='bottom'?r2===range.r2:side==='left'?c1===range.c1:c2===range.c2;
    const key=atEdge?side:(side==='top'||side==='bottom'?'insideH':'insideV');
    if(decisions[key]!==undefined)borders[side]=decisions[key]?{...edge}:null;
  }
  return {border:false,borders};
}

export function borderSummary(targets,range){
  const groups=Object.fromEntries([...borderSides,'insideH','insideV'].map(side=>[side,[]]));
  for(const {cell,rect:m} of targets){for(const [side,edge] of Object.entries(cellBorders(cell.style))){
    const outer=side==='top'?m.r1===range.r1:side==='bottom'?m.r2===range.r2:side==='left'?m.c1===range.c1:m.c2===range.c2;
    groups[outer?side:side==='top'||side==='bottom'?'insideH':'insideV'].push(edge);
  }}
  return Object.fromEntries(Object.entries(groups).map(([side,edges])=>[side,!edges.length?false:edges.every(Boolean)?true:edges.every(e=>!e)?false:null]));
}
