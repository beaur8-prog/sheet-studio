const edges=['top','right','bottom','left'];
const zero=()=>Object.fromEntries(edges.map(edge=>[edge,0]));
const read=value=>Object.fromEntries(edges.map(edge=>[edge,Number.isFinite(value?.[edge])&&value[edge]>=0&&value[edge]<500?value[edge]:0]));
export function androidHeaderCutout(detail,width){
  if(detail?.displayCutout?.top===0)return null;
  const top=detail?.displayCutout?.top??48;
  const rects=Array.isArray(detail?.displayCutoutRects)?detail.displayCutoutRects.filter(r=>r&&['top','right','bottom','left'].every(k=>Number.isFinite(r[k]))&&r.top>=0&&r.top<top&&r.bottom>r.top&&r.left>=0&&r.right>r.left&&r.right<=width):[];
  if(!rects.length||detail.consumedByWebView)return null;
  const left=Math.max(0,Math.min(...rects.map(r=>r.left))-8),right=Math.min(width,Math.max(...rects.map(r=>r.right))+8),height=Math.max(48,...rects.map(r=>r.bottom));
  const end=width>=600?192:144;
  if(left>=52&&width-right>=end)return {mode:'split',left:left-4,gap:right-left,height};
  if(width-right>=280)return {mode:'left',inset:right,height};
  if(left>=280)return {mode:'right',inset:width-left,height};
  return null;
}

// Insets are snapshots in CSS pixels, not additions to the viewport size.
// Old flat messages describe system bars; typed messages retain real cutouts
// after those bars disappear. Insets already used by the WebView add nothing.
export function androidInsets(detail={}){
  detail=detail&&typeof detail==='object'?detail:{};
  if(detail.consumedByWebView===true)return {bars:zero(),cutout:zero()};
  if(detail.systemBars||detail.displayCutout)return {bars:read(detail.systemBars),cutout:read(detail.displayCutout)};
  return {bars:read(detail),cutout:null};
}

export function androidViewportFrame(win){
  const v=win.visualViewport,unzoomed=v&&Math.abs((v.scale??1)-1)<.01;
  return unzoomed?{width:v.width||win.innerWidth,height:v.height||win.innerHeight,top:Math.max(0,v.offsetTop||0),left:Math.max(0,v.offsetLeft||0)}:
    {width:win.innerWidth,height:win.innerHeight,top:0,left:0};
}
