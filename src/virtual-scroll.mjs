// Blink caps an element's physical height. Keep logical worksheet coordinates
// independent from a bounded scrollbar, while cells keep their true pixel size.
export const MAX_SCROLL_EXTENT=30_000_000;
export function scrollGeometry(total,viewport,cap=MAX_SCROLL_EXTENT){
  total=Math.max(0,Number(total)||0);viewport=Math.max(0,Number(viewport)||0);
  const extent=Math.min(total,cap),physical=Math.max(0,extent-viewport),logical=Math.max(0,total-viewport),ratio=physical?logical/physical:1;
  const clamp=(n,max)=>Math.max(0,Math.min(max,Number(n)||0));
  return {extent,physical,logical,ratio,toLogical:n=>physical?clamp(n,physical)*ratio:0,toPhysical:n=>logical?clamp(n,logical)/ratio:0};
}
export function createVerticalScroll(grid,surface){
  let total=0;
  const geometry=()=>scrollGeometry(total,grid.clientHeight),top=()=>geometry().toLogical(grid.scrollTop),setTop=value=>{grid.scrollTop=geometry().toPhysical(value);};
  grid.addEventListener('wheel',event=>{
    const g=geometry();if(g.ratio<=1||event.defaultPrevented||event.ctrlKey||grid.ownerDocument.body.dataset.dimensionResize)return;
    const unit=event.deltaMode===1?25:event.deltaMode===2?grid.clientHeight:1;
    event.preventDefault();
    if(event.shiftKey&&!event.deltaX)grid.scrollLeft+=event.deltaY*unit;
    else {grid.scrollLeft+=event.deltaX*unit;setTop(top()+event.deltaY*unit);}
  },{passive:false});
  return {
    top,setTop,offset:()=>top()-grid.scrollTop,
    extent(height){const previous=top();total=Math.max(0,height);surface.style.height=geometry().extent+'px';setTop(previous);},
  };
}
