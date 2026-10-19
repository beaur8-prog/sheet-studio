import {createWorkbook,address} from '../../src/engine.mjs';
export function colorFixture(){
 const b=createWorkbook();b.name='색 기준 필터 확인';const s=b.sheets[0];
 const rows=[['상품','금액','구분'],['사과',10,'국내'],['사과',20,'수입'],['배',30,'국내'],['포도',40,'국내'],['귤',50,'수입'],['',60,'국내'],['딸기',70,'국내']];
 rows.forEach((row,r)=>row.forEach((raw,c)=>s.cells[address(r,c)]={raw:String(raw)}));
 const styles={A1:{background:'#ff0000'},A2:{background:'#FF0000',color:'#008000'},A3:{background:'#0000ff',color:'#ff0000'},A4:{background:'#f00'},A5:{color:'#ff0000'},A6:{background:'#ffffff',color:'#000000'},A8:{background:'#0000ff'}};
 for(const [key,style] of Object.entries(styles))s.cells[key].style=style;
 s.conditionalRules=[{id:'color-rule',range:'A2:A8',type:'expression',formulae:['$B2>=60'],style:{background:'#ff0000',color:'#008000'}}];
 s.autoFilter={range:{r1:0,r2:7,c1:0,c2:2},columns:{},hiddenRows:[]};s.columnWidths={0:130,1:100,2:100};return b;
}
