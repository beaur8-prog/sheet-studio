export const CHART_LIMITS={minWidth:280,minHeight:180,maxWidth:2400,maxHeight:1800};
export const chartRect=chart=>({x:chart.x??40,y:chart.y??25,width:Math.max(100,Math.min(CHART_LIMITS.maxWidth,chart.width??500)),height:Math.max(100,Math.min(CHART_LIMITS.maxHeight,chart.height??300))});
const clamp=(value,min,max)=>Math.max(min,Math.min(max,value));
export function resizeChart(rect,edge,dx,dy){
  const next={...rect},limits=CHART_LIMITS;
  if(edge.includes('e'))next.width=clamp(rect.width+dx,limits.minWidth,limits.maxWidth);
  if(edge.includes('s'))next.height=clamp(rect.height+dy,limits.minHeight,limits.maxHeight);
  if(edge.includes('w')){next.width=clamp(rect.width-dx,limits.minWidth,Math.min(limits.maxWidth,rect.x+rect.width));next.x=rect.x+rect.width-next.width;}
  if(edge.includes('n')){next.height=clamp(rect.height-dy,limits.minHeight,Math.min(limits.maxHeight,rect.y+rect.height));next.y=rect.y+rect.height-next.height;}
  return next;
}
