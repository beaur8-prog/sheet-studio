import {cloneData} from './clone.mjs';
import {address} from './engine.mjs';
import {worksheetCellStyle,inheritedStyles} from './cell-style.mjs';
// A clipboard outlives its workbook. Its style indices must be private, portable
// definitions rather than indices into a workbook that may since be closed.
export function captureClipboardStyles(book,sheet,data,rows,columns){
  const styles=[],known=new Map(),normal=book.xlsxStyles?.[book.xlsxDefaultStyle],defaults={fontFamily:'맑은 고딕',fontSize:11,bold:false,italic:false,underline:false,color:'#222222',background:'#ffffff',verticalAlign:'middle',wrap:false,format:'general',numberFormat:'General'},baseStyle={font:{name:'맑은 고딕',size:11,color:{argb:'FF222222'}},numFmt:'General',alignment:{vertical:'middle'},fill:{type:'pattern',pattern:'none'},border:{}};
  for(let r=0;r<rows.length;r++)for(let c=0;c<columns.length;c++){
    const row=rows[r],col=columns[c],ref=address(row,col),record=sheet.cells[ref]||{raw:''},copied=data[r][c];
    const original=book.xlsxStyles?.[record.xlsxStyle],templates=original||record.xlsxNormal?[]:inheritedStyles(book,sheet,row,col);
    const full=Object.assign({},baseStyle,normal?.style,...(original?[original.style]:templates.map(t=>t.style)));
    const bare={...sheet,cells:{[ref]:{...record,style:{}}}},baseline={...defaults,...worksheetCellStyle(book,bare,row,col,{table:false})};
    const entry=JSON.parse(JSON.stringify({style:full,baseline})),key=JSON.stringify(entry);
    if(!known.has(key)){known.set(key,styles.length);styles.push(entry);}
    copied.style=cloneData({...defaults,...worksheetCellStyle(book,sheet,row,col)});copied.xlsxStyle=known.get(key);delete copied.xlsxNormal;
  }
  return styles;
}
export function adoptClipboardStyles(book,clip,updates,{mode='all'}={}){
  if(!clip?.xlsxStyles||!['all','formats'].includes(mode))return [];
  const existing=book.xlsxStyles||[],known=new Map(existing.map((entry,i)=>[JSON.stringify(entry),i])),mapped=new Map(),added=[];
  for(const {cell} of updates){const id=cell.xlsxStyle;if(id===undefined)continue;
    if(!mapped.has(id)){const entry=clip.xlsxStyles[id];if(!entry)throw Error('복사한 셀의 서식 정보가 없습니다. 다시 복사하세요.');const key=JSON.stringify(entry);if(!known.has(key)){known.set(key,existing.length+added.length);added.push(cloneData(entry));}mapped.set(id,known.get(key));}
    cell.xlsxStyle=mapped.get(id);
  }
  if(existing.length+added.length>65536)throw Error('통합 문서의 셀 서식 수가 한도를 초과합니다. 일부 셀을 값으로 붙여넣으세요.');
  return added;
}
