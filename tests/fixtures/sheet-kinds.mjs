import {createWorkbook} from '../../src/engine.mjs';
export function chartsheetBook(){
  const book=createWorkbook(),source=book.sheets[0];source.name='Source';source.state='hidden';source.cells={A1:{raw:'Item'},B1:{raw:'Revenue'},A2:{raw:'a'},B2:{raw:'10'},A3:{raw:'b'},B3:{raw:'20'}};
  const chart={id:'chart',name:'Dashboard',kind:'chart',cells:{},rowHeights:{},columnWidths:{},merges:[],charts:[{id:'bars',type:'bar',title:'Dashboard',range:'A1:B3',x:0,y:0,width:700,height:400,series:[{name:'Revenue',catRef:{sheetId:source.id,range:'A2:A3'},valRef:{sheetId:source.id,range:'B2:B3'}}]}]},tail={id:'tail',name:'Tail',cells:{A1:{raw:'=Local'}},rowHeights:{},columnWidths:{},merges:[],charts:[]};book.sheets=[chart,source,tail];book.activeSheetId=chart.id;book.chartNames=[{name:'Local',sheetId:tail.id,formula:"'Source'!$B$3"}];return book;
}
