import test from 'node:test';
import assert from 'node:assert/strict';
import {wheelZoom,zoomedScroll,rescaleLayout,installZoomInteractions} from '../src/zoom-interactions.mjs';

test('wheel zoom normalizes units, direction and bounds without accepting invalid deltas',()=>{
  assert.ok(Math.abs(wheelZoom(100,{deltaY:-100,deltaMode:0},600)-110)<1e-9);
  assert.equal(wheelZoom(100,{deltaY:1,deltaMode:1},600),wheelZoom(100,{deltaY:16,deltaMode:0},600));
  assert.equal(wheelZoom(100,{deltaY:1,deltaMode:2},600),wheelZoom(100,{deltaY:100,deltaMode:0},600));
  assert.equal(wheelZoom(180,{deltaY:-100,deltaMode:0},600),180);
  assert.equal(wheelZoom(50,{deltaY:100,deltaMode:0},600),50);
  assert.equal(wheelZoom(100,{deltaY:NaN,deltaMode:0},600),100);
});
test('zoom retains the sheet point under the pointer and respects frozen panes/heading changes',()=>{
  const before=(1234+210)/1,after=zoomedScroll({content:1444,scroll:1234,fixed:false},203,108,1.2);
  assert.ok(Math.abs((after+203)/1.2-before)<1e-9);
  assert.equal(zoomedScroll({content:1264,scroll:1234,fixed:true},30,108,1.2),1234*1.2);
  assert.equal(zoomedScroll({content:200,scroll:0,fixed:false},200,0,.5),0);
  assert.equal(zoomedScroll({content:200,scroll:100,fixed:false},100,180,2),220);
});
test('cached variable and hidden axes scale reversibly without editing their source dimensions',()=>{
  const dimensions={rowHeights:{1:0,2:50}},layout={rows:new Float64Array([0,25,25,75]),cols:new Float64Array([0,90,230])};
  rescaleLayout(layout,1.8);assert.deepEqual([...layout.rows],[0,45,45,135]);rescaleLayout(layout,1/1.8);
  assert.deepEqual([...layout.rows],[0,25,25,75]);assert.deepEqual(dimensions,{rowHeights:{1:0,2:50}});
});
function harness(nativeAndroid=false){
  const root=new EventTarget(),win=new EventTarget(),frames=new Map();let serial=0,zoom=100,context='first',allowed=true,time=0,calls=0;
  Object.assign(win,{innerWidth:900,innerHeight:600,performance:{now:()=>time},requestAnimationFrame:fn=>{frames.set(++serial,fn);return serial;},cancelAnimationFrame:id=>frames.delete(id)});root.defaultView=win;
  const control=installZoomInteractions({root,host:{clientHeight:600,getBoundingClientRect:()=>({left:0,top:0,right:900,bottom:600,width:900,height:600})},getZoom:()=>zoom,setZoom:value=>{zoom=value;calls++;},getContext:()=>context,canZoom:()=>allowed,nativeAndroid});
  const event=(name,props={})=>{const e=new Event(name,{cancelable:true});Object.assign(e,{clientX:200,clientY:200,ctrlKey:true,deltaY:-1,deltaMode:0,buttons:0,...props});root.dispatchEvent(e);return e;};
  const flush=()=>{const f=[...frames.values()];frames.clear();f.forEach(fn=>fn());time+=16;};
  return {event,flush,control,native:detail=>{const e=new Event('sheet-studio-android-wheel');e.detail=detail;win.dispatchEvent(e);},blur:()=>win.dispatchEvent(new Event('blur')),pagehide:()=>win.dispatchEvent(new Event('pagehide')),get zoom(){return zoom;},get calls(){return calls;},set context(v){context=v;},set allowed(v){allowed=v;}};
}
test('high-rate wheel is coalesced and sub-percent trackpad deltas accumulate across frames',()=>{
  const h=harness();for(let i=0;i<100;i++)h.event('wheel',{deltaY:-.1});h.flush();assert.equal(h.zoom,111);assert.equal(h.calls,1);
  const small=harness();for(let i=0;i<20;i++){small.event('wheel',{deltaY:-.1});small.flush();}assert.equal(small.zoom,102);
  assert.equal(h.event('wheel',{ctrlKey:false}).defaultPrevented,false);
});
test('pending zoom is cancelled for sheet changes, modals, pointer gestures and window blur',()=>{
  const h=harness();h.event('wheel',{deltaY:-100});h.context='second';h.flush();assert.equal(h.zoom,100);
  h.event('wheel',{deltaY:-100});h.allowed=false;h.flush();assert.equal(h.zoom,100);
  h.allowed=true;h.event('wheel',{deltaY:-100});h.event('pointerdown');h.flush();assert.equal(h.zoom,100);
  h.event('wheel',{deltaY:-100});h.blur();h.flush();assert.equal(h.zoom,100);
  h.event('wheel',{deltaY:-100});h.pagehide();h.flush();assert.equal(h.zoom,100);
});
test('WebKit cumulative gesture scale does not multiply updates or duplicate Ctrl+wheel',()=>{
  const h=harness();h.event('gesturestart',{scale:1});h.event('gesturechange',{scale:1.3});h.event('wheel',{deltaY:-100});h.flush();assert.equal(h.zoom,130);
  h.event('gesturechange',{scale:1.1});h.event('gestureend');h.flush();assert.equal(h.zoom,110);
});

test('Android native Ctrl-wheel uses the common zoom and rejects other sources or invalid coordinates',()=>{
 const a=harness(true),detail={ctrlKey:true,deltaY:-100,x:.5,y:.5,buttons:0};a.native(detail);a.flush();assert.equal(a.zoom,110);
 a.native({...detail,ctrlKey:false});a.flush();assert.equal(a.zoom,110);a.native({...detail,x:NaN});a.flush();assert.equal(a.zoom,110);
 const desktop=harness();desktop.native(detail);desktop.flush();assert.equal(desktop.zoom,100);
});
