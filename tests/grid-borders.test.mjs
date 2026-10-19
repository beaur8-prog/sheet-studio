import test from 'node:test';
import assert from 'node:assert/strict';
import {gridBorderSegments,gridBordersHtml} from '../src/grid-borders.mjs';
const edge=(style='thin',color='#ff0000')=>({style,color});
const cell=(x,y,w=90,h=25,borders={})=>({x,y,w,h,style:{borders}});
const all=(style='thin')=>Object.fromEntries(['top','bottom','left','right'].map(side=>[side,edge(style)]));

test('shared borders collapse once, retaining the requested thickness and all stored sides',()=>{
  for(const style of ['thin','medium','thick','double','dotted','dashed']){
    const cells=[cell(0,0,90,25,all(style)),cell(90,0,90,25,all(style)),cell(0,25,90,25,all(style)),cell(90,25,90,25,all(style))],before=JSON.stringify(cells);
    const lines=gridBorderSegments(cells,{showGrid:false});
    assert.equal(lines.length,6);assert.ok(lines.every(l=>l.css.includes(edge(style).color)));
    assert.equal(lines.filter(l=>l.horizontal&&l.pos===25).length,1);
    assert.equal(lines.filter(l=>!l.horizontal&&l.pos===90).length,1);assert.equal(JSON.stringify(cells),before);
  }
});
test('merged neighbours split shared sides and preserve stronger, differently coloured borders',()=>{
  const merged=cell(0,0,90,50,{right:edge()}),upper=cell(90,0,90,25,{left:edge('medium','#0000ff')}),lower=cell(90,25,90,25,{left:edge('thick','#00ff00')});
  const lines=gridBorderSegments([merged,upper,lower],{showGrid:false});
  assert.deepEqual(lines.map(l=>[l.pos,l.start,l.end,l.css]),[[90,0,25,'2px solid #0000ff'],[90,25,50,'3px solid #00ff00']]);
  // A one-sided explicit border replaces the adjacent default gridline.
  const shared=gridBorderSegments([cell(0,0),upper]).find(l=>!l.horizontal&&l.pos===90);
  assert.equal(shared.css,'2px solid #0000ff');
});
test('grid visibility, frozen panes and clipping do not remove explicit borders or make strokes interactive',()=>{
  assert.equal(gridBorderSegments([cell(0,0)],{showGrid:false}).length,0);
  assert.equal(gridBorderSegments([cell(0,0)]).length,2);
  const fixed={...cell(40,25,90,25,all()),pane:'fixed',z:3,minX:40,minY:25,originX:40,originY:25};
  const scrolling={...cell(0,0,90,25,all()),pane:'scroll',z:1,minX:130,minY:50};
  const html=gridBordersHtml([fixed,scrolling],{showGrid:false});
  assert.equal((html.match(/cell-border/g)||[]).length,4);assert.match(html,/aria-hidden="true"/);assert.match(html,/z-index:3/);assert.doesNotMatch(html,/z-index:1/);
});
