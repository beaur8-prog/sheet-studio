export const MIN_ZOOM=50,MAX_ZOOM=180;
export const clampZoom=value=>Math.max(MIN_ZOOM,Math.min(MAX_ZOOM,value));

// Pixel, line and page wheel units all feed the same bounded sheet scale.
export function wheelZoom(value,event,height){
  const unit=event.deltaMode===1?16:event.deltaMode===2?height:1;
  // Fine pinch deltas retain their native logarithmic sensitivity; a coarse
  // mouse-wheel tick is capped at a 10% change instead of a large jump.
  const limit=100*Math.log(1.1),delta=Math.max(-limit,Math.min(limit,event.deltaY*unit));
  return Number.isFinite(delta)?clampZoom(value*Math.exp(-delta/100)):value;
}
export function zoomedScroll(anchor,point,frozen,scale){
  return Math.max(0,anchor.fixed?anchor.scroll*scale:anchor.content*scale-Math.max(point,frozen));
}
// Scaling cached axes avoids rescanning a million rows and their filters for
// every pinch frame. Hidden rows/columns and variable sizes remain unchanged.
export function rescaleLayout(layout,ratio){
  for(const axis of [layout.rows,layout.cols])for(let i=1;i<axis.length;i++)axis[i]*=ratio;
}

export function installZoomInteractions({root,host,getZoom,setZoom,getContext,canZoom,captureAnchor=p=>p,getViewport=()=>'',nativeAndroid=false}){
  const win=root.defaultView;
  let frame=0,pending=null,wheel=null,gesture=null;
  const point=event=>{
    const r=host.getBoundingClientRect();
    return Number.isFinite(event.clientX)&&Number.isFinite(event.clientY)&&event.clientX>=r.left&&event.clientX<r.right&&event.clientY>=r.top&&event.clientY<r.bottom
      ?{clientX:event.clientX,clientY:event.clientY}:{clientX:r.left+r.width/2,clientY:r.top+r.height/2};
  };
  function cancel(){win.cancelAnimationFrame(frame);frame=0;pending=null;wheel=null;gesture=null;}
  function flush(){
    frame=0;const next=pending;pending=null;
    if(!next||!canZoom()||next.context!==getContext()||next.base!==getZoom()||next.viewport!==getViewport()){wheel=null;return;}
    setZoom(Math.round(next.value),next.point);
    if(wheel){wheel.actual=getZoom();wheel.viewport=getViewport();}
  }
  function queue(value,anchor){
    pending={value:clampZoom(value),point:anchor,context:getContext(),base:getZoom(),viewport:getViewport()};
    if(!frame)frame=win.requestAnimationFrame(flush);
  }
  function onWheel(event){
    if(!event.ctrlKey||event.defaultPrevented||!Number.isFinite(event.deltaY)||!event.deltaY)return;
    // Chromium also sends trackpad pinch as Ctrl+wheel. Cancel the browser's
    // page zoom, including over chrome/dialogs, so the app UI stays its size.
    event.preventDefault();
    if(gesture)return;
    if(!event.cancelable||event.buttons||!canZoom()){cancel();return;}
    const now=win.performance.now(),context=getContext(),actual=getZoom(),viewport=getViewport();
    if(!wheel||now-wheel.time>250||wheel.context!==context||wheel.actual!==actual||wheel.viewport!==viewport)wheel={value:actual,actual,context,viewport,time:now};
    wheel.time=now;wheel.value=wheelZoom(wheel.value,event,host.clientHeight);
    const at=point(event);
    if(!wheel.anchor||wheel.anchor.clientX!==at.clientX||wheel.anchor.clientY!==at.clientY)wheel.anchor=captureAnchor(at);
    queue(wheel.value,wheel.anchor);
  }
  function onGestureStart(event){
    event.preventDefault();cancel();
    if(canZoom())gesture={zoom:getZoom(),context:getContext(),point:captureAnchor(point(event))};
  }
  function onGestureChange(event){
    event.preventDefault();
    if(!gesture||!canZoom()||gesture.context!==getContext()||!Number.isFinite(event.scale)||event.scale<=0)return;
    queue(gesture.zoom*event.scale,gesture.point);
  }
  function onGestureEnd(event){if(gesture)event.preventDefault();gesture=null;}
  root.addEventListener('wheel',onWheel,{capture:true,passive:false});
  // Safari's trackpad reports cumulative scale through GestureEvent instead.
  root.addEventListener('gesturestart',onGestureStart,{passive:false});
  root.addEventListener('gesturechange',onGestureChange,{passive:false});
  root.addEventListener('gestureend',onGestureEnd,{passive:false});
  root.addEventListener('pointerdown',cancel,{capture:true});
  win.addEventListener('blur',cancel);
  win.addEventListener('pagehide',cancel);
  win.addEventListener('resize',cancel);
  if(nativeAndroid)win.addEventListener('sheet-studio-android-wheel',event=>{
    const d=event.detail;if(!d||d.ctrlKey!==true||![d.deltaY,d.x,d.y].every(Number.isFinite)||d.x<0||d.x>1||d.y<0||d.y>1)return;
    const v=win.visualViewport;
    onWheel({ctrlKey:true,deltaY:d.deltaY,deltaMode:0,buttons:d.buttons||0,cancelable:true,defaultPrevented:false,preventDefault(){},clientX:(v?.offsetLeft||0)+d.x*(v?.width||win.innerWidth),clientY:(v?.offsetTop||0)+d.y*(v?.height||win.innerHeight)});
  });
  root.addEventListener('visibilitychange',()=>{if(root.hidden)cancel();});
  return {cancel};
}
