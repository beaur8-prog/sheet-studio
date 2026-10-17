import {createWorkbook} from '../../src/engine.mjs';
export function worksheetFixture(){
  const book=createWorkbook();book.name='Worksheet feature verification';const s=book.sheets[0];
  [['상품','수량','단가','금액'],['Alpha',2,100,'=B2*C2'],['Beta',3,200,'=B3*C3'],['Alpha',-1,50,'=B4*C4'],['Gamma',10,30,'=B5*C5'],['Delta',0,25,'=B6*C6']].forEach((row,r)=>row.forEach((v,c)=>s.cells[String.fromCharCode(65+c)+(r+1)]={raw:String(v),style:r?{}:{bold:true,background:'#107c41',color:'#ffffff'}}));
  s.freeze={rows:1,cols:1};s.columnWidths={0:130,1:90,2:90,3:100,5:0};s.rowHeights={8:0};
  s.conditionalRules=[
    {id:'negative',range:'B2:B6',type:'cellIs',operator:'lessThan',formulae:['0'],style:{background:'#ffc7ce',color:'#9c0006'},stopIfTrue:true},
    {id:'large',range:'D2:D6',type:'cellIs',operator:'greaterThan',formulae:['250'],style:{background:'#c6efce',color:'#006100'}},
    {id:'duplicate',range:'A2:A6',type:'duplicateValues',formulae:[],style:{background:'#ffeb9c'}},
    {id:'relative',range:'C2:C6',type:'expression',formulae:['$B2<0'],style:{color:'#9c0006'}},
  ];return book;
}
