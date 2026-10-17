// Pixel geometry shared by cells, selection, reference outlines and pointer picking.
export function offsetIndex(arr, p) {
  let l=0,h=arr.length-1;
  while(l<h){const m=Math.ceil((l+h)/2);if(arr[m]<=p)l=m;else h=m-1;}
  return Math.min(l,arr.length-2);
}
export function visibleAxis(arr, scroll, size, frozen=0) {
  const result=[], end=arr.length-2, fixed=Math.min(size,arr[frozen]||0);
  for(let i=0;i<frozen&&i<=end&&arr[i]<size;i++)if(arr[i+1]>arr[i])result.push(i);
  if(fixed>=size)return result;
  const first=Math.max(frozen,offsetIndex(arr,scroll+fixed)),last=Math.min(end,offsetIndex(arr,scroll+size)+1);
  for(let i=first;i<=last;i++)if(arr[i+1]>arr[i])result.push(i);
  return result;
}
export function axisPoint(arr,index,scroll,frozen){return arr[index]-(index<frozen?0:scroll);}
export function axisPick(arr,pixel,scroll,frozen){return offsetIndex(arr,Math.max(0,pixel+(pixel<(arr[frozen]||0)?0:scroll)));}
export function paneRects(range, layout, freeze, scroll, width, height) {
  const out=[],fr=freeze?.rows||0,fc=freeze?.cols||0,fw=layout.cols[fc],fh=layout.rows[fr];
  for(const [rlo,rhi,sy,ylo,yhi] of [[0,fr-1,0,0,fh],[fr,layout.rows.length-2,scroll.top,fh,height]])
    for(const [clo,chi,sx,xlo,xhi] of [[0,fc-1,0,0,fw],[fc,layout.cols.length-2,scroll.left,fw,width]]) {
      const r1=Math.max(rlo,range.r1),r2=Math.min(rhi,range.r2),c1=Math.max(clo,range.c1),c2=Math.min(chi,range.c2);
      if(r1>r2||c1>c2)continue;
      const x=Math.max(xlo,layout.cols[c1]-sx),y=Math.max(ylo,layout.rows[r1]-sy),right=Math.min(xhi,width,layout.cols[c2+1]-sx),bottom=Math.min(yhi,height,layout.rows[r2+1]-sy);
      if(right>x&&bottom>y)out.push({x,y,w:right-x,h:bottom-y,r1,r2,c1,c2});
    }
  return out;
}
