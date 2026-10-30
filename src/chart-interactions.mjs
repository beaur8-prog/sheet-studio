import {chartRect,resizeChart} from './chart-geometry.mjs';

export function installChartInteractions({grid,getChart,getZoom,canEdit=()=>true,commitEdit,checkpoint,changed,render}){
  let gesture=null;
  function restore(chart,rect){for(const key of ['x','y','width','height']){if(rect[key]===undefined)delete chart[key];else chart[key]=rect[key];}}
  function finish(cancel=false){
    if(!gesture)return;
    const g=gesture;gesture=null;const chart=getChart(g.id);
    if(chart){const next=chartRect(chart);restore(chart,g.original);
      if(!cancel&&JSON.stringify(next)!==JSON.stringify(g.rect)){checkpoint();Object.assign(chart,next);changed();}else if(g.captured)render();
    }
    if(grid.hasPointerCapture(g.pointerId))grid.releasePointerCapture(g.pointerId);
  }
  grid.addEventListener('pointerdown',event=>{
    const card=event.target.closest('.chart-card'),handle=event.target.closest('[data-chart-resize]');
    if(event.button!==0||!card||(!handle&&event.target.closest('button')))return;
    const chart=getChart(card.dataset.chart);if(!chart||!canEdit())return;
    event.preventDefault();event.stopPropagation();if(commitEdit()===false)return;
    gesture={id:chart.id,edge:handle?.dataset.chartResize,pointerId:event.pointerId,startX:event.clientX,startY:event.clientY,scrollX:grid.scrollLeft,scrollY:grid.scrollTop,rect:chartRect(chart),original:{x:chart.x,y:chart.y,width:chart.width,height:chart.height}};
    grid.focus();
  });
  grid.addEventListener('pointermove',event=>{
    if(!gesture||gesture.pointerId!==event.pointerId)return;
    const g=gesture,chart=getChart(g.id);if(!chart)return finish(true);
    const scale=getZoom()/100,dx=(event.clientX-g.startX+grid.scrollLeft-g.scrollX)/scale,dy=(event.clientY-g.startY+grid.scrollTop-g.scrollY)/scale;
    if(!g.captured){if(Math.abs(dx)+Math.abs(dy)<3)return;grid.setPointerCapture(event.pointerId);g.captured=true;}
    Object.assign(chart,g.edge?resizeChart(g.rect,g.edge,dx,dy):{...g.rect,x:Math.max(0,g.rect.x+dx),y:Math.max(0,g.rect.y+dy)});render();
  });
  grid.addEventListener('pointerup',()=>finish());
  grid.addEventListener('pointercancel',()=>finish(true));
  grid.addEventListener('lostpointercapture',()=>finish());
  window.addEventListener('blur',()=>finish());
  document.addEventListener('keydown',event=>{
    if(gesture){if(event.key==='Escape')finish(true);event.preventDefault();event.stopImmediatePropagation();return;}
    const handle=event.target.closest('[data-chart-resize]');
    if(!handle||!['ArrowLeft','ArrowRight','ArrowUp','ArrowDown'].includes(event.key))return;
    event.preventDefault();event.stopPropagation();const id=handle.closest('[data-chart]').dataset.chart,chart=getChart(id),edge=handle.dataset.chartResize;if(!chart||!canEdit())return;
    const step=event.shiftKey?1:10,dx=event.key==='ArrowLeft'?-step:event.key==='ArrowRight'?step:0,dy=event.key==='ArrowUp'?-step:event.key==='ArrowDown'?step:0;
    const before=chartRect(chart),next=resizeChart(before,edge,dx,dy);if(JSON.stringify(before)===JSON.stringify(next))return;
    checkpoint();Object.assign(chart,next);changed();
    requestAnimationFrame(()=>[...grid.querySelectorAll('[data-chart-resize]')].find(el=>el.closest('[data-chart]').dataset.chart===id&&el.dataset.chartResize===edge)?.focus());
  },true);
}
