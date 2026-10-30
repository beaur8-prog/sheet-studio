import test from 'node:test';
import assert from 'node:assert/strict';
import JSZip from 'jszip';
import {drawingFixture,png} from './fixtures/drawings.mjs';
import {readSpreadsheet,writeSpreadsheet} from '../src/io.mjs';
import {validateWorkbook} from '../src/engine.mjs';
import {drawingPreviews,drawingAnchorRect} from '../src/xlsx-drawings.mjs';
import {xml,all,attr,readRelationships,resolvePart,R} from '../src/chart-xml.mjs';
test('ordinary image and shape anchors, media and external hyperlinks survive editing and repeated saves alongside charts',async()=>{
  let bytes=await drawingFixture(),book,counts;
  for(let round=0;round<4;round++){
    book=validateWorkbook(await readSpreadsheet(bytes,'pictures.xlsx'));const s=book.sheets[0],preview=drawingPreviews(s);assert.equal(preview.images.length,1);assert.equal(preview.other,1);assert.equal(preview.images[0].src,'data:image/png;base64,'+png);assert.equal(preview.images[0].width,140);assert.equal(preview.images[0].height,80);assert.equal(s.xlsxDrawings[0].rels.find(r=>r.mode==='External').target,'https://example.com');assert.ok(s.xlsxDrawings[0].anchors.some(a=>a.includes('원본 도형')));
    s.cells.A2.raw=String(round+10);if(!s.charts.length)s.charts.push({id:'new',type:'bar',range:'A1:A2',title:'Chart',width:300,height:200,x:0,y:300});bytes=(await writeSpreadsheet(book)).data;
    const zip=await JSZip.loadAsync(bytes),drawing=zip.file('xl/drawings/sheetStudioDrawing1.xml'),doc=xml(await drawing.async('string')),ids=all(doc,'cNvPr').map(n=>attr(n,'id'));assert.equal(ids.length,new Set(ids).size);assert.equal(all(doc,'pic').length,1);assert.equal(all(doc,'sp').length,1);assert.equal(all(doc,'chart').length,1);
    const rels=await readRelationships(zip,drawing.name);for(const rel of rels)if(rel.mode!=='External')assert.ok(zip.file(resolvePart(drawing.name,rel.target)),rel.target);
    const current=Object.keys(zip.files).filter(p=>!zip.files[p].dir).length;if(round===1)counts=current;else if(round>1)assert.equal(current,counts);
  }
  assert.equal(book.sheets[0].charts.length,1);
  const cropped=structuredClone(book),group=cropped.sheets[0].xlsxDrawings[0],index=group.anchors.findIndex(a=>a.includes('pic'));group.anchors[index]=group.anchors[index].replace(/(<(?:\w+:)?blipFill[^>]*>)/,'$1<a:srcRect xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" l="10000"/>');assert.equal(drawingPreviews(cropped.sheets[0]).images.length,0);assert.equal(drawingPreviews(cropped.sheets[0]).other,2);
});
test('drawing positions respect hidden dimensions and do not loop through the entire row extent',()=>{
  const anchor=xml('<oneCellAnchor><from><col>3</col><row>1000000</row><colOff>0</colOff><rowOff>0</rowOff></from><ext cx="952500" cy="952500"/></oneCellAnchor>').documentElement,rect=drawingAnchorRect(anchor,{defaultColumnWidth:90,defaultRowHeight:25,columnWidths:{0:0,2:100},rowHeights:{1:0}});assert.equal(rect.x,190);assert.equal(rect.y,24999975);assert.equal(rect.width,100);
});
