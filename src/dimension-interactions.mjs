// The headers are virtualized. Capture on their stable parent, never a handle
// that is replaced while the sheet is redrawn. On Android, only the guide moves
// until release: rebuilding all million row positions on every move blocks input.
export function installDimensionInteractions({host,grid,getSheet,getZoom,getBoundary,deferResize=false,canResize=()=>true,onDimensionsChanged=()=>{},commitEdit,checkpoint,preview,changed}){
  let gesture=null,guideFrame=0;
  const guide=document.createElement('div'),line=document.createElement('div'),label=document.createElement('div');
  guide.className='dimension-guide';guide.hidden=true;
  line.className='dimension-guide-line';label.className='dimension-guide-label';
  guide.append(line,label);host.append(guide);

  function showGuide(){
    const g=gesture,column=g.axis==='c',position=deferResize?g.boundary+(g.value-g.size)*g.scale:getBoundary(g.axis,g.index);
    guide.dataset.axis=g.axis;guide.hidden=false;
    line.style.cssText=column?`left:${position}px;top:0;bottom:0`:`top:${position}px;left:0;right:0`;
    label.textContent=`${column?'열 너비':'행 높이'}: ${g.value}px${getZoom()===100?'':` (화면 ${Math.round(g.value*getZoom()/100)}px · ${getZoom()}%)`}`;
    label.style.left=`${column?Math.max(42,Math.min(position+8,(deferResize?g.hostWidth:host.clientWidth)-150)):45}px`;
    label.style.top=`${column?28:Math.max(28,Math.min(position+8,(deferResize?g.hostHeight:host.clientHeight)-28))}px`;
  }
  function restore(g){
    if(g.original===undefined)delete g.sheet[g.property][g.index];
    else g.sheet[g.property][g.index]=g.original;
  }
  function queueGuide(){if(!guideFrame)guideFrame=requestAnimationFrame(()=>{guideFrame=0;if(gesture)showGuide();});}
  function finish(cancel=false){
    if(!gesture)return;
    const g=gesture;gesture=null;cancelAnimationFrame(guideFrame);guideFrame=0;guide.hidden=true;delete document.body.dataset.dimensionResize;
    if(g.previewed){
      if(!deferResize)restore(g);
      if(!cancel&&g.value!==g.size){const before={...g.sheet,[g.property]:{...g.sheet[g.property]}};checkpoint();g.sheet[g.property]||={};g.sheet[g.property][g.index]=g.value;onDimensionsChanged(g.sheet,before);changed({rebuild:true});}
      else if(!deferResize)preview();
    }
    if(host.hasPointerCapture(g.pointerId))host.releasePointerCapture(g.pointerId);
  }
  host.addEventListener('pointerdown',event=>{
    const handle=event.target.closest('[data-resize-col],[data-resize-row]');
    if(event.button!==0||!handle||gesture)return;
    event.preventDefault();event.stopPropagation();if(commitEdit()===false)return;grid.focus();
    const axis=handle.hasAttribute('data-resize-col')?'c':'r',index=Number(axis==='c'?handle.dataset.resizeCol:handle.dataset.resizeRow);
    if(!canResize(axis))return;
    const sheet=getSheet(),property=axis==='c'?'columnWidths':'rowHeights';if(!deferResize)sheet[property]||={};
    const original=sheet[property]?.[index],size=Math.max(axis==='c'?26:18,original||(axis==='c'?sheet.defaultColumnWidth||90:sheet.defaultRowHeight||25));
    gesture={sheet,axis,index,property,original,size,value:size,start:axis==='c'?event.clientX:event.clientY,scale:getZoom()/100,pointerId:event.pointerId,previewed:false,boundary:deferResize?getBoundary(axis,index):0,hostWidth:deferResize?host.clientWidth:0,hostHeight:deferResize?host.clientHeight:0};
    document.body.dataset.dimensionResize=axis;showGuide();
  });
  function move(event){
    const g=gesture;if(!g||g.pointerId!==event.pointerId)return;
    const delta=((g.axis==='c'?event.clientX:event.clientY)-g.start)/g.scale;
    if(!g.previewed&&Math.abs(delta)<2)return;
    const value=Math.max(g.axis==='c'?26:18,Math.min(1000,Math.round(g.size+delta)));
    if(value===g.value)return;
    if(!g.previewed)host.setPointerCapture(event.pointerId);
    g.value=value;g.previewed=true;if(deferResize)queueGuide();else{g.sheet[g.property][g.index]=value;preview();showGuide();}
  }
  document.addEventListener('pointermove',move);
  document.addEventListener('pointerup',event=>{if(event.pointerId===gesture?.pointerId){if(deferResize)move(event);finish();}});
  host.addEventListener('pointercancel',()=>finish(true));
  host.addEventListener('lostpointercapture',()=>finish(true));
  host.addEventListener('wheel',event=>{if(gesture)event.preventDefault();},{passive:false});
  window.addEventListener('blur',()=>finish(true));
  document.addEventListener('keydown',event=>{
    if(!gesture)return;
    if(event.key==='Escape')finish(true);
    event.preventDefault();event.stopImmediatePropagation();
  },true);
}
