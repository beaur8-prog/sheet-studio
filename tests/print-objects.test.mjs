import test from 'node:test';
import assert from 'node:assert/strict';
import {createWorkbook,createEvaluator,setExternalEvaluator} from '../src/engine.mjs';
import {printObjectAxes,objectCellBounds,printObjectFragments} from '../src/print-objects.mjs';
import {drawingPrintsWithSheet} from '../src/anchor-position.mjs';
import {buildSheetPrintDocument} from '../src/sheet-print.mjs';
import {readPrintSettings,planPrintPages} from '../src/print-layout.mjs';
import {readSpreadsheet,writeSpreadsheet} from '../src/io.mjs';
import {drawingFixture} from './fixtures/drawings.mjs';
const chart=(id,x=270,y=50)=>({id,type:'bar',range:'A1:B2',title:id,firstRowAsHeader:false,firstColumnAsLabels:true,x,y,width:400,height:200});
function area(book,sheet,text){book.chartNames=[{name:'_xlnm.Print_Area',sheetId:sheet.id,formula:"'"+sheet.name+"'!"+text}];}

test('sparse drawing print axes collapse hidden dimensions and locate distant objects without expanding the sheet',()=>{
 const s={defaultRowHeight:25,defaultColumnWidth:90,rowHeights:{0:50,2:0,999998:100},columnWidths:{1:0,3:120},autoFilter:{range:{r1:0,r2:3,c1:0,c2:1},columns:{},hiddenRows:[1]}};
 const a=printObjectAxes(s);assert.equal(a.rows.offset(3),50);assert.equal(a.rows.offset(4),75);assert.equal(a.cols.offset(4),300);assert.equal(a.rows.at(50),3);assert.equal(a.cols.at(90),2);
 const y=a.rows.offset(999998);assert.deepEqual(objectCellBounds({x:0,y,width:90,height:100},a),{r1:999998,r2:999998,c1:0,c2:0});
});
test('worksheet-position objects share a page with data and extend the automatic used range',()=>{
 const b=createWorkbook(),s=b.sheets[0];s.cells={A1:{raw:'Jan'},B1:{raw:'2'},A2:{raw:'Feb'},B2:{raw:'5'}};s.charts=[chart('sales')];const before=JSON.stringify(b),out=buildSheetPrintDocument(b,s.id);
 assert.equal(out.pageCount,1);assert.equal(out.chartCount,1);assert.equal(out.dataCells,80);assert.match(out.html,/class="print-object-clip"[^>]*left:270px;top:50px;width:400px;height:200px/);assert.doesNotMatch(out.html,/class="print-page print-object-page"/);assert.equal(JSON.stringify(b),before);
});
test('saved print areas and selections clip crossing objects and exclude unrelated outside drawings',()=>{
 const b=createWorkbook(),s=b.sheets[0];s.cells.A1={raw:'in area'};s.charts=[chart('cross',100,20),chart('outside',700,500)];area(b,s,'$A$1:$B$4');
 for(const options of [{},{scope:'selection',selection:{r1:0,r2:3,c1:0,c2:1}}]){const out=buildSheetPrintDocument(b,s.id,options);assert.equal(out.pageCount,1);assert.equal(out.chartCount,1);assert.match(out.html,/left:100px;top:20px;width:80px;height:80px/);assert.match(out.html,/data-chart="cross"/);assert.doesNotMatch(out.html,/data-chart="outside"/);}
 const separate=buildSheetPrintDocument(b,s.id,{objectLayout:'pages'});assert.equal(separate.chartCount,2);assert.equal(separate.pageCount,3);
});
test('page boundaries tile a chart in place while repeating titles do not clone drawings',()=>{
 const b=createWorkbook(),s=b.sheets[0];s.cells.A1={raw:'Title'};s.xlsxFeatures={pageSetup:'<pageSetup scale="100"/>',rowBreaks:'<rowBreaks><brk id="4" man="1"/></rowBreaks>',colBreaks:'<colBreaks><brk id="3" man="1"/></colBreaks>'};area(b,s,'A1:F8');b.chartNames.push({name:'_xlnm.Print_Titles',sheetId:s.id,formula:"'"+s.name+"'!$1:$1,'"+s.name+"'!$A:$A"});
 const axes=printObjectAxes(s),plan=planPrintPages(s,[{r1:0,r2:7,c1:0,c2:5}],readPrintSettings(b,s)),rect={x:200,y:70,width:200,height:100};assert.equal(plan.pages.length,4);
 const parts=plan.pages.map(p=>printObjectFragments(rect,p,axes));assert.deepEqual(parts.map(p=>p.length),[1,1,1,1]);assert.equal(parts.flat().reduce((n,p)=>n+p.width*p.height,0),rect.width*rect.height);assert.equal(parts[1][0].y,25);assert.equal(parts[2][0].x,90);
 s.charts=[{...chart('title',0,0),width:100,height:100}];const out=buildSheetPrintDocument(b,s.id);assert.equal((out.html.match(/data-chart="title"/g)||[]).length,1);
});
test('nonadjacent areas remain separate and only their intersecting object portions are printed',()=>{
 const b=createWorkbook(),s=b.sheets[0];s.charts=[chart('wide',0,0)];area(b,s,"$A$1:$B$4,'"+s.name+"'!$E$1:$F$4");const out=buildSheetPrintDocument(b,s.id);assert.equal(out.pageCount,2);assert.equal(out.chartCount,1);assert.equal((out.html.match(/data-chart="wide"/g)||[]).length,2);assert.match(out.html,/left:-360px/);
});
test('source drawing print flags survive repeated XLSX saves and apply to charts and images',async()=>{
 assert.equal(drawingPrintsWithSheet('<absoluteAnchor><clientData fPrintsWithSheet="false"/></absoluteAnchor>'),false);assert.equal(drawingPrintsWithSheet('<absoluteAnchor><clientData/></absoluteAnchor>'),true);
 let b=await readSpreadsheet(await drawingFixture(),'pictures.xlsx'),s=b.sheets[0];s.xlsxDrawings[0].anchors=s.xlsxDrawings[0].anchors.map(text=>text.replace(/<((?:\w+:)?clientData)(?=[\s/>])/,'<$1 fPrintsWithSheet="0"'));s.charts=[{...chart('excluded'),source:undefined}];
 const image=buildSheetPrintDocument(b,s.id);assert.equal(image.imageCount,0);assert.equal(image.chartCount,1);
 for(let round=0;round<2;round++){b=await readSpreadsheet((await writeSpreadsheet(b)).data,'again.xlsx');s=b.sheets[0];assert.equal(buildSheetPrintDocument(b,s.id).imageCount,0);}
 s.charts[0].source.anchor=s.charts[0].source.anchor.replace(/<((?:\w+:)?clientData)(?=[\s/>])/,'<$1 fPrintsWithSheet="false"');assert.equal(buildSheetPrintDocument(b,s.id).chartCount,0);assert.equal(buildSheetPrintDocument(b,s.id,{scope:'charts'}).chartCount,1);b=await readSpreadsheet((await writeSpreadsheet(b)).data,'print-flag.xlsx');assert.equal(buildSheetPrintDocument(b,b.sheets[0].id).chartCount,0);
});

test('outside-area charts do not calculate unrelated volatile data and object expansion has an explicit budget',()=>{
 const b=createWorkbook(),s=b.sheets[0];s.cells.A1={raw:'=RAND()'};s.cells.B1={raw:'selected'};s.charts=[{...chart('outside',1000,1000),range:'A1:A1',firstColumnAsLabels:false}];
 const requests=[],evaluate=createEvaluator(b,s.id),snapshot={...b},shared=(id,ref)=>{requests.push(ref);return evaluate(ref);};shared.result=(id,ref)=>{requests.push(ref);return evaluate.result(ref);};setExternalEvaluator(snapshot,shared);
 const out=buildSheetPrintDocument(snapshot,s.id,{scope:'selection',selection:{r1:0,r2:0,c1:1,c2:1}});assert.equal(out.chartCount,0);assert.ok(!requests.includes('A1'));
 s.charts=Array.from({length:2001},(_,i)=>({...chart('chart '+i,0,0),width:100,height:100}));assert.throws(()=>buildSheetPrintDocument(b,s.id),/개체의 인쇄 분량/);
});
