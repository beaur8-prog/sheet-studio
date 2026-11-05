import {xml,first,val} from './chart-xml.mjs';
const policies=new Set(['gap','zero','span']),cache=new WeakMap();
export function chartBlankPolicy(chart){
  if(policies.has(chart.blankCells))return chart.blankCells;
  const source=chart.source;if(!source)return 'gap';if(policies.has(source.original?.blankCells))return source.original.blankCells;if(cache.has(source))return cache.get(source);
  let policy='gap';try{const raw=source.parts?.[source.root]?.data;if(raw){const doc=xml(new TextDecoder().decode(Uint8Array.from(atob(raw),c=>c.charCodeAt(0)))),value=val(first(doc,'chart'),'dispBlanksAs','gap');if(policies.has(value))policy=value;}}catch{}cache.set(source,policy);return policy;
}
/** Keep indexes so area baselines and scatter coordinates follow the same gaps.
 * Only physical blanks can be spanned; invalid values and log exclusions break. */
export function chartLineSegments(points,{skip=[],closed=false}={}){
  let order=points.map((_,i)=>i).filter(i=>points[i]||!skip[i]);if(!order.length)return [];
  if(closed){const gap=order.findIndex(i=>!points[i]);if(gap<0)order.push(order[0]);else order=[...order.slice(gap+1),...order.slice(0,gap+1)];}
  const segments=[];let current=[];for(const i of [...order,-1]){if(i>=0&&points[i])current.push(i);else if(current.length){segments.push(current);current=[];}}return segments;
}
