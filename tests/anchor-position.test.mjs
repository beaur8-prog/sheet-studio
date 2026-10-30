import test from 'node:test';
import assert from 'node:assert/strict';
import JSZip from 'jszip';
import {xml,child,all,attr,D} from '../src/chart-xml.mjs';
import {drawingAnchorRect,anchorPlacement,positionAnchor,transformAnchor} from '../src/anchor-position.mjs';
import {transformAnchoredObjects,dimensionSnapshot} from '../src/anchored-objects.mjs';
import {createWorkbook} from '../src/engine.mjs';
import {readSpreadsheet,writeSpreadsheet} from '../src/io.mjs';
import {structuralWorkbookPlan} from '../src/workbook-structure.mjs';
import {drawingFixture} from './fixtures/drawings.mjs';
const blank=()=>({defaultColumnWidth:90,defaultRowHeight:25,columnWidths:{},rowHeights:{}}),source=`<xdr:absoluteAnchor xmlns:xdr="${D}"><xdr:clientData/></xdr:absoluteAnchor>`,rect={x:185,y:55,width:360,height:200},read=(a,s)=>drawingAnchorRect(xml(a).documentElement,s),near=(a,b)=>{for(const k of ['x','y','width','height'])assert.ok(Math.abs(a[k]-b[k])<.001,k+': '+a[k]+' != '+b[k]);};
test('anchor conversions preserve exact offsets, placement and millionth-row geometry using sparse axis dimensions',()=>{
 const s=blank();s.columnWidths={0:0,2:110};s.rowHeights={0:0,10:45};const large={...rect,y:25000005};for(const mode of ['fixed','move','move-size']){const a=positionAnchor(source,large,s,mode);assert.equal(anchorPlacement(xml(a).documentElement),mode);near(read(a,s),large);assert.ok(child(xml(a).documentElement,'clientData'));}
});
test('row/column insertion and deletion obey all three object placement modes and preserve fixed pixel sizes',()=>{
 const s=blank(),after=blank();for(const mode of ['fixed','move','move-size']){
  let a=positionAnchor(source,rect,s,mode);a=transformAnchor(a,s,after,{axis:'r',index:0,delta:3});near(read(a,after),{...rect,y:mode==='fixed'?55:130});
  a=positionAnchor(source,rect,s,mode);a=transformAnchor(a,s,after,{axis:'r',index:4,delta:2});near(read(a,after),{...rect,height:mode==='move-size'?250:200});
  a=positionAnchor(source,rect,s,mode);a=transformAnchor(a,s,after,{axis:'c',index:0,delta:-1});near(read(a,after),{...rect,x:mode==='fixed'?185:95});
 }
 const deleted=transformAnchor(positionAnchor(source,rect,s,'move-size'),s,after,{axis:'r',index:2,delta:-3});assert.equal(read(deleted,after).y,50);assert.equal(child(child(xml(deleted).documentElement,'from'),'rowOff').textContent,'0');
});
test('dimension resizing and hidden columns move/resize anchored objects while fixed placement stays put',()=>{
 const before=blank(),after=blank();after.columnWidths={0:0,3:150};after.rowHeights={0:50,4:75};for(const mode of ['fixed','move','move-size']){const a=transformAnchor(positionAnchor(source,rect,before,mode),before,after);near(read(a,after),mode==='fixed'?rect:{...rect,x:95,y:80,width:mode==='move-size'?420:360,height:mode==='move-size'?250:200});}
});
test('two-cell editAs oneCell/absolute semantics survive normalization',()=>{
 const s=blank(),after=blank();after.columnWidths={0:120};for(const [edit,mode] of [['oneCell','move'],['absolute','fixed']]){const a=positionAnchor(source,rect,s,'move-size').replace('<xdr:twoCellAnchor','<xdr:twoCellAnchor editAs="'+edit+'"'),next=transformAnchor(a,s,after,{axis:'r',index:0,delta:2});assert.equal(anchorPlacement(xml(next).documentElement),mode);near(read(next,after),mode==='fixed'?rect:{...rect,x:215,y:105});}
});
test('imported image/shape and chart anchors follow bulk structure, survive manual chart moves and three XLSX reopen cycles',async()=>{
 let book=await readSpreadsheet(await drawingFixture(),'drawings.xlsx'),s=book.sheets[0];const originalImage=read(s.xlsxDrawings[0].anchors[0],s),originalShape=read(s.xlsxDrawings[0].anchors[1],s);s.charts.push({id:'chart',type:'bar',range:'A1:A2',title:'Test',placement:'move-size',...rect});
 book=structuralWorkbookPlan(book,s.id,{axis:'r',index:0,count:3,delta:1});s=book.sheets[0];near(read(s.xlsxDrawings[0].anchors[0],s),{...originalImage,y:originalImage.y+3*(s.defaultRowHeight||25)});near(read(s.xlsxDrawings[0].anchors[1],s),originalShape);assert.equal(s.charts[0].y,rect.y+3*(s.defaultRowHeight||25));
 for(let i=0;i<3;i++){book=await readSpreadsheet((await writeSpreadsheet(book)).data,'round.xlsx');s=book.sheets[0];assert.equal(s.charts[0].placement,'move-size');near(read(s.charts[0].source.anchor,s),s.charts[0]);s.charts[0].x+=17;s.charts[0].height+=10;const before=dimensionSnapshot(s);s.rowHeights[0]=(s.defaultRowHeight||25)+15;transformAnchoredObjects(s,before);assert.ok(s.charts[0].source.anchor.includes('twoCellAnchor'));}
 const bytes=(await writeSpreadsheet(book)).data,zip=await JSZip.loadAsync(bytes),drawing=await zip.file('xl/drawings/sheetStudioDrawing1.xml').async('string');assert.equal(all(xml(drawing),'chart').length,1);assert.equal(all(xml(drawing),'pic').length,1);assert.equal(all(xml(drawing),'sp').length,1);const round=await readSpreadsheet(bytes,'end.xlsx');near(round.sheets[0].charts[0],s.charts[0]);
});
