import {readSpreadsheet} from '../../src/io.mjs';
import {drawingFixture} from './drawings.mjs';
export async function androidPrintFixture(){
 const book=await readSpreadsheet(await drawingFixture(),'Android print.xlsx');book.name='인쇄 검증 문서';const s=book.sheets[0];s.name='인쇄 원본';
 s.cells={A1:{raw:'항목',style:{bold:true,background:'#d9ead3'}},B1:{raw:'금액',style:{bold:true,background:'#d9ead3'}},C1:{raw:'검증',style:{bold:true}},A2:{raw:'사과'},B2:{raw:'10'},C2:{raw:'선택 영역'},A3:{raw:'배'},B3:{raw:'20'},C3:{raw:'수식 결과'},A4:{raw:'합계'},B4:{raw:'=SUM(B2:B3)'},A5:{raw:'둘째 페이지'},B5:{raw:'-12',style:{numberFormat:'0;[Red](0)'}},A6:{raw:'가로 출력'},B6:{raw:'50'},A7:{raw:'병합 제목',style:{bold:true}},A8:{raw:'마지막 행'},B8:{raw:'=AVERAGE(B2:B3)'},E8:{raw:'123456789012345',style:{numberFormat:'0'}},H1:{raw:'PRIVATE_OUTSIDE_PRINT_AREA'}};
 s.columnWidths={0:130,1:100,2:120,3:80,4:90,5:80};s.rowHeights={0:30,6:32};s.merges=['A7:C7'];
 s.xlsxFeatures={...s.xlsxFeatures,pageSetup:'<pageSetup paperSize="9" scale="100" orientation="portrait"/>',sheetPr:'<sheetPr><pageSetUpPr fitToPage="0"/></sheetPr>',rowBreaks:'<rowBreaks><brk id="4" min="0" max="16383" man="1"/></rowBreaks>',headerFooter:'<headerFooter><oddHeader>&amp;C인쇄 검증 &amp;P / &amp;N</oddHeader><oddFooter>&amp;L원본 보존</oddFooter></headerFooter>'};
 book.chartNames=[{name:'_xlnm.Print_Area',sheetId:s.id,formula:"'인쇄 원본'!$A$1:$F$8"},{name:'_xlnm.Print_Titles',sheetId:s.id,formula:"'인쇄 원본'!$1:$1"}];
 s.charts=[{id:'print-chart',type:'bar',title:'매출 차트',range:'A1:B3',x:40,y:60,width:480,height:280}];
 return book;
}
