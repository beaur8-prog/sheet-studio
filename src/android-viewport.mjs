const edges=['top','right','bottom','left'];
const zero=()=>Object.fromEntries(edges.map(edge=>[edge,0]));
const read=value=>Object.fromEntries(edges.map(edge=>[edge,Number.isFinite(value?.[edge])&&value[edge]>=0&&value[edge]<500?value[edge]:0]));

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
