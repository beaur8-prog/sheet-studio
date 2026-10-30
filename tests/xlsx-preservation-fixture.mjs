import ExcelJS from 'exceljs';
export async function preservationFixture(){
  const source=new ExcelJS.Workbook(),s=source.addWorksheet('Details');
  s.getCell('A1').value={richText:[{font:{bold:true,color:{argb:'FFFF0000'}},text:'중요 '},{font:{italic:true},text:'메모'}]};
  s.getCell('A1').alignment={textRotation:45,indent:2,shrinkToFit:true,horizontal:'left'};
  s.getCell('A1').font={name:'Arial',size:13,strike:true};
  s.getCell('A2').value={text:'문서 링크',hyperlink:'https://example.org/manual?q=a&x=b',tooltip:'원본 도움말'};
  s.getCell('A3').value=10;s.getCell('A3').fill={type:'gradient',gradient:'angle',degree:30,stops:[{position:0,color:{argb:'FF000000'}},{position:1,color:{argb:'FFFFFFFF'}}]};
  s.getCell('A3').border={diagonal:{style:'double',color:{argb:'FF008800'},up:true}};
  s.getCell('A4').value='입력 가능';s.getCell('A4').protection={locked:false,hidden:true};
  s.getCell('B1').value={formula:'SUM(A3:A3)',result:10,shareType:'array',ref:'B1:B2'};s.getCell('B2').value=10;
  s.getCell('C1').value={formula:'_xlfn._xlws.FILTER(A3:A3,A3:A3>0)',result:10,shareType:'array',ref:'C1:C1'};
  s.getCell('D1').dataValidation={type:'list',allowBlank:true,formulae:['"예,아니오"'],showErrorMessage:true,errorStyle:'stop',errorTitle:'목록',error:'목록에서 선택하세요.'};
  await s.protect('secret',{spinCount:1000,selectUnlockedCells:true});
  source.definedNames.add("'Details'!$A$3",'Revenue');
  return new Uint8Array(await source.xlsx.writeBuffer());
}
