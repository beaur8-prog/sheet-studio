import {createWorkbook,address} from '../../src/engine.mjs';
export function chartFixture(){
  const book=createWorkbook(),sheet=book.sheets[0];sheet.name="매출 '원본'";book.name='차트 검증';
  const rows=[['항목','매출','비용','크기'],['가',40,22,8],['나',25,18,12],['다',-10,5,5],['라',35,10,15]];
  rows.forEach((row,r)=>row.forEach((v,c)=>sheet.cells[address(r,c)]={raw:String(v)}));
  return book;
}
export const chartCases=[...['bar','bar-horizontal','line','area'].flatMap(type=>['clustered','stacked','percentStacked'].map(grouping=>({type,grouping}))),...['pie','doughnut','scatter','bubble','radar','combo'].map(type=>({type,grouping:'clustered'}))];
export function addChart(book,patch={}){const chart={id:crypto.randomUUID(),type:'bar',range:'A1:C5',title:'분기 매출',x:250,y:30,width:620,height:360,legend:true,labels:false,grouping:'clustered',...patch};book.sheets[0].charts.push(chart);return chart;}
