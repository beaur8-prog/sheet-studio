import ExcelJS from 'exceljs';
import JSZip from 'jszip';
import {xml,stringify,child,children,all,attr} from '../src/chart-xml.mjs';
export async function tableFixture(){
  const source=new ExcelJS.Workbook(),s=source.addWorksheet('Sales Data');
  s.addTable({name:'Sales',ref:'A1',headerRow:true,totalsRow:true,style:{theme:'TableStyleMedium2',showRowStripes:true},columns:[{name:'Qty',totalsRowLabel:'합계',filterButton:true},{name:'Price',filterButton:true},{name:'Amount',totalsRowFunction:'sum',filterButton:true}],rows:[[2,3,{formula:'[@Qty]*[@Price]',result:6}],[4,5,{formula:'[@Qty]*[@Price]',result:20}]]});
  s.getCell('E1').value={formula:'SUM(Sales[Amount])',result:26};
  const zip=await JSZip.loadAsync(await source.xlsx.writeBuffer()),doc=xml(await zip.file('xl/tables/table1.xml').async('string')),column=all(doc,'tableColumn').find(c=>attr(c,'name')==='Amount'),formula=doc.createElementNS(doc.documentElement.namespaceURI,'calculatedColumnFormula');formula.textContent='[@Qty]*[@Price]';column.appendChild(formula);zip.file('xl/tables/table1.xml',stringify(doc));return zip.generateAsync({type:'uint8array'});
}
