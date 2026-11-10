export const MIN_ZOOM=50,MAX_ZOOM=180;
export const clampZoom=value=>Math.max(MIN_ZOOM,Math.min(MAX_ZOOM,value));

// Pixel, line and page wheel units all feed the same bounded sheet scale.
export function wheelZoom(value,event,height){
  const unit=event.deltaMode===1?16:event.deltaMode===2?height:1;
  const delta=Math.max(-100,Math.min(100,event.deltaY*unit));
  return Number.isFinite(delta)?clampZoom(value*Math.exp(-delta*Math.log(1.1)/100)):value;
}
export function zoomedScroll(scroll,point,nextPoint,frozen,nextFrozen,ratio){
  if(point<Math.max(0,frozen))return Math.max(0,scroll*ratio);
  return Math.max(0,(scroll+point)*ratio-Math.max(nextPoint,nextFrozen));
}
// Scaling cached axes avoids rescanning a million rows and their filters for
// every pinch frame. Hidden rows/columns and variable sizes remain unchanged.
export function rescaleLayout(layout,ratio){
  for(const axis of [layout.rows,layout.cols])for(let i=1;i<axis.length;i++)axis[i]*=ratio;
}

export function installZoomInteractions({root,host,getZoom,setZoom,getContext,canZoom}){
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
    if(!next||!canZoom()||next.context!==getContext()||next.base!==getZoom()){wheel=null;return;}
    setZoom(Math.round(next.value),next.point);
    if(wheel)wheel.actual=getZoom();
  }
  function queue(value,anchor){
    pending={value:clampZoom(value),point:anchor,context:getContext(),base:getZoom()};
    if(!frame)frame=win.requestAnimationFrame(flush);
  }
  function onWheel(event){
    if(!event.ctrlKey||event.defaultPrevented||!Number.isFinite(event.deltaY)||!event.deltaY)return;
    // Chromium also sends trackpad pinch as Ctrl+wheel. Cancel the browser's
    // page zoom, including over chrome/dialogs, so the app UI stays its size.
    event.preventDefault();
    if(gesture)return;
    if(!event.cancelable||event.buttons||!canZoom()){cancel();return;}
    const now=win.performance.now(),context=getContext(),actual=getZoom();
    if(!wheel||now-wheel.time>250||wheel.context!==context||wheel.actual!==actual)wheel={value:actual,actual,context,time:now};
    wheel.time=now;wheel.value=wheelZoom(wheel.value,event,host.clientHeight);
    queue(wheel.value,point(event));
  }
  function onGestureStart(event){
    event.preventDefault();cancel();
    if(canZoom())gesture={zoom:getZoom(),context:getContext(),point:point(event)};
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
  root.addEventListener('visibilitychange',()=>{if(root.hidden)cancel();});
  return {cancel};
}
