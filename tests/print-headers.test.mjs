import test from 'node:test';
import assert from 'node:assert/strict';
import {createWorkbook} from '../src/engine.mjs';
import {splitHeaderSections,headerEditorText,headerTextFromEditor,renderPageHeaders,readHeaderFooter} from '../src/print-headers.mjs';
import {editablePrintSettings,printSettingsPlan} from '../src/print-settings.mjs';
import {readPrintSettings} from '../src/print-layout.mjs';
import {writeSpreadsheet,readSpreadsheet} from '../src/io.mjs';
import {xml,child,children,attr} from '../src/chart-xml.mjs';
const context={book:{name:'Book <name>'},sheet:{name:'Report & data'},index:0,pageNumber:3,totalPages:7,now:new Date('2026-10-31T03:04:00Z'),scale:1};

test('headers split escaped ampersands and quoted fonts, retaining raw formatting when edited text is unchanged',()=>{
 assert.deepEqual(splitHeaderSections('&LLeft &&C&C&"Arial,Bold"Center&RRight'),{left:'Left &&C',center:'&"Arial,Bold"Center',right:'Right'});
 assert.deepEqual(splitHeaderSections('Default center&L1&L2'),{left:'12',center:'Default center',right:''});
 assert.equal(headerEditorText('&BTitle&B && &P / &N &A'),'Title & {페이지} / {전체 페이지} {시트 이름}');
 assert.equal(headerTextFromEditor('Title & {페이지} / {전체 페이지} {시트 이름}'),'Title && &P / &N &A');
});
test('first/even headers, page offsets, dates, escaped HTML and rich text codes render safely',()=>{
 const settings={oddHeader:'&L&BTitle <script>&B&C&P+2 / &N&R&A',evenHeader:'EVEN',firstHeader:'FIRST',oddFooter:'&C&"Georgia,Italic"&14&KFF0000Footer && &F',differentFirst:true,differentOddEven:true};
 let rendered=renderPageHeaders(settings,context);assert.match(rendered.header,/FIRST/);assert.doesNotMatch(rendered.header,/Title/);rendered=renderPageHeaders(settings,{...context,index:2,pageNumber:5});assert.match(rendered.footer,/font-family:&#39;Georgia&#39;/);assert.match(rendered.footer,/font-size:14pt/);assert.match(rendered.footer,/color:#FF0000/);assert.match(rendered.footer,/Footer &amp; Book &lt;name&gt;/);
 rendered=renderPageHeaders(settings,{...context,index:1,pageNumber:4});assert.match(rendered.header,/EVEN/);rendered=renderPageHeaders(settings,{...context,index:2,pageNumber:5});assert.match(rendered.header,/7 \/ 7/);assert.match(rendered.header,/Report &amp; data/);assert.match(rendered.header,/Title &lt;script&gt;/);assert.doesNotMatch(rendered.header,/<script>/);
 const notes=renderPageHeaders({oddHeader:'&G &Z'},context).notes;assert.equal(notes.size,2);
});
test('page settings plans are atomic, retain unrelated XML and names, and survive XLSX round trips',async()=>{
 let book=createWorkbook(),s=book.sheets[0];s.name="Q1, Bob's!";s.cells.A1={raw:'1'};s.xlsxFeatures={sheetPr:'<sheetPr codeName="Report"><outlinePr summaryBelow="0"/></sheetPr>',pageSetup:'<pageSetup paperSize="9" orientation="portrait" horizontalDpi="600"/>',headerFooter:'<headerFooter differentFirst="1"><firstHeader>ORIGINAL</firstHeader></headerFooter>'};book.chartNames=[{name:'UserName',formula:'1'},{name:'_xlnm.Print_Area',sheetId:'other',formula:'Other!A1:B2'}];
 const v=editablePrintSettings(book,s);Object.assign(v,{orientation:'landscape',paperId:8,mode:'scale',scale:75,areas:'A1:D40,F1:G5',repeatRows:'1:2',repeatCols:'A:B',firstPageNumber:3,automaticPageNumber:false});v.margin={top:20,bottom:20,left:12,right:12,header:8,footer:8};v.headerFooter.oddHeader='&C&A';v.headerFooter.oddFooter='&C&P / &N';
 const before=JSON.stringify(book),bad={...v,scale:500};assert.throws(()=>printSettingsPlan(book,s,bad),/배율/);assert.equal(JSON.stringify(book),before);
 const plan=printSettingsPlan(book,s,v);s.xlsxFeatures=plan.features;book.chartNames=plan.chartNames;let settings=readPrintSettings(book,s);assert.equal(settings.orientation,'landscape');assert.equal(settings.paperId,8);assert.equal(settings.scale,.75);assert.equal(settings.firstPageNumber,3);assert.equal(settings.areas.length,2);assert.deepEqual(settings.repeatRows,{start:0,end:1});assert.deepEqual(settings.repeatCols,{start:0,end:1});assert.match(s.xlsxFeatures.sheetPr,/codeName="Report"/);assert.match(s.xlsxFeatures.sheetPr,/summaryBelow="0"/);assert.match(s.xlsxFeatures.pageSetup,/horizontalDpi="600"/);assert.equal(book.chartNames.filter(n=>n.name==='UserName').length,1);assert.equal(book.chartNames.filter(n=>n.sheetId==='other').length,1);
 assert.deepEqual(children(xml(s.xlsxFeatures.headerFooter).documentElement).map(n=>n.localName),['oddHeader','oddFooter','evenHeader','evenFooter','firstHeader','firstFooter']);
 for(let round=0;round<3;round++){book=await readSpreadsheet((await writeSpreadsheet(book)).data,'settings.xlsx');s=book.sheets[0];settings=readPrintSettings(book,s);assert.equal(settings.paperId,8);assert.equal(settings.firstPageNumber,3);assert.equal(settings.headerFooter.firstHeader,'ORIGINAL');assert.equal(settings.headerFooter.oddFooter,'&C&P / &N');assert.equal(settings.areas.length,2);assert.equal(settings.repeatCols.end,1);}
});
test('print settings reject invalid references and margins without changing the source',()=>{
 const b=createWorkbook(),s=b.sheets[0],v=editablePrintSettings(b,s),before=JSON.stringify(b);
 for(const changes of [{areas:'A0:B1'},{areas:'XFE1'},{repeatRows:'0:2'},{repeatCols:'B:A'},{fitWidth:-1},{firstPageNumber:0,automaticPageNumber:false},{margin:{...v.margin,left:200,right:200}},{headerFooter:{...v.headerFooter,oddHeader:'x'.repeat(256)}}])assert.throws(()=>printSettingsPlan(b,s,{...v,...changes}));assert.equal(JSON.stringify(b),before);
 b.nativeSource={};assert.throws(()=>printSettingsPlan(b,s,v),/XLSM/);
});

test('fit scaling, explicit empty first/even headers and unsupported effects retain their meaning',()=>{
 const base={oddHeader:'&C&10Title',firstHeader:'',differentFirst:true,differentOddEven:true};
 assert.equal(renderPageHeaders(base,context).header,'');assert.match(renderPageHeaders(base,{...context,index:1,pageNumber:4,scale:.5}).header,/font-size:5pt/);
 assert.equal(renderPageHeaders({...base,evenHeader:''},{...context,index:1,pageNumber:4}).header,'');
 const unknown=renderPageHeaders({oddHeader:'&C&OOutline &K02+050Theme'},context);assert.equal(unknown.notes.size,1);assert.match(unknown.header,/Theme/);assert.doesNotMatch(unknown.header,/&amp;O/);
});
test('chart sheet page margins and first/even headers round trip without worksheet-only setup attributes',async()=>{
 let b=createWorkbook(),s={id:'chart',name:'Chart',kind:'chart',cells:{},merges:[],rowHeights:{},columnWidths:{},charts:[{id:'c1',type:'bar',series:[{name:'Value',values:[1,2],labels:['A','B']}],width:500,height:300}]};b.sheets.push(s);const v=editablePrintSettings(b,s);v.orientation='landscape';v.margin.top=25;v.headerFooter.oddHeader='&CChart header';v.headerFooter.oddFooter='&R&P';const plan=printSettingsPlan(b,s,v);s.xlsxFeatures=plan.features;b.chartNames=plan.chartNames;assert.doesNotMatch(s.xlsxFeatures.pageSetup,/fitToWidth|pageOrder|scale=/);
 b=await readSpreadsheet((await writeSpreadsheet(b)).data,'chart-settings.xlsx');s=b.sheets.find(s=>s.kind==='chart');const settings=readPrintSettings(b,s);assert.equal(settings.orientation,'landscape');assert.equal(settings.margin.top,25);assert.equal(settings.headerFooter.oddHeader,'&CChart header');assert.equal(settings.headerFooter.oddFooter,'&R&P');
});

test('OOXML escaped literal sequences are not decoded twice',()=>{assert.equal(headerEditorText('_x005F_x000A_'),'_x000A_');assert.equal(headerEditorText('Line1_x000A_Line2'),'Line1\nLine2');});


test('invalid saved print references remain visible and repairable through page settings',()=>{
 const b=createWorkbook(),s=b.sheets[0];b.chartNames=[{name:'_xlnm.Print_Area',sheetId:s.id,formula:"'Sheet1'!#REF!"},{name:'_xlnm.Print_Titles',sheetId:s.id,formula:"'Sheet1'!#REF!"}];const before=JSON.stringify(b),v=editablePrintSettings(b,s);assert.equal(v.areas,'#REF!');assert.equal(v.repeatRows,'#REF!');assert.ok(v.referenceWarning);assert.equal(JSON.stringify(b),before);v.areas='A1:C10';v.repeatRows='1:1';const plan=printSettingsPlan(b,s,v);assert.equal(plan.chartNames.length,2);assert.match(plan.chartNames[0].formula,/\$A\$1:\$C\$10/);
});
