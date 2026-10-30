import test from 'node:test';
import assert from 'node:assert/strict';
import {scrollGeometry,MAX_SCROLL_EXTENT} from '../src/virtual-scroll.mjs';
import {axisPick,axisPoint,visibleAxis,rowHeadingWidth} from '../src/worksheet-view.mjs';
test('large logical extents reach both ends without changing visible cell dimensions',()=>{
  const rows=1048576,height=45,viewport=700,geometry=scrollGeometry(rows*height+25,viewport);assert.equal(geometry.extent,MAX_SCROLL_EXTENT);assert.ok(geometry.ratio>1);
  const positions=new Float64Array(rows+1);for(let r=0;r<=rows;r++)positions[r]=r*height;
  const top=geometry.toLogical(geometry.physical);assert.equal(top,geometry.logical);assert.ok(visibleAxis(positions,top,viewport-25).includes(rows-1));
  const y=axisPoint(positions,rows-1,top,0),offset=top-geometry.toPhysical(top);assert.ok(y>=0&&y+height<=viewport);assert.ok(Math.abs((positions[rows-1]-offset)-geometry.toPhysical(top)-y)<1e-6);
  assert.equal(axisPick(positions,y+20,top,0),rows-1);
});
test('ordinary extents retain native scrolling and compressed positions roundtrip with variable viewports',()=>{
  assert.equal(rowHeadingWidth(0),40);assert.equal(rowHeadingWidth(9999),47);assert.equal(rowHeadingWidth(99999),54);assert.equal(rowHeadingWidth(1048575),61);
  const ordinary=scrollGeometry(26214425,700);assert.equal(ordinary.ratio,1);assert.equal(ordinary.toPhysical(200),200);assert.equal(ordinary.toLogical(200),200);
  for(const viewport of [400,700,1400])for(const logical of [0,25,500000,26214425,47185200]){const g=scrollGeometry(47185945,viewport),target=Math.min(logical,g.logical);assert.ok(Math.abs(g.toLogical(g.toPhysical(logical))-target)<1e-6);}
  assert.equal(scrollGeometry(100,700).toLogical(100),0);assert.equal(scrollGeometry(100,700).toPhysical(100),0);
});
