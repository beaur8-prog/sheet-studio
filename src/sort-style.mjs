import {address} from './engine.mjs';
import {worksheetCellStyle} from './cell-style.mjs';
// A row format stays on the worksheet when records are sorted. Resolve its
// contribution only for cells that actually cross different row templates.
// Keep table banding dynamic and reuse XF definitions across all moved cells.
export function sortStyleResolver(book,sheet,bounds){
  const rows=sheet.xlsxStyleDefaults?.rows||{};
  if(!Object.keys(rows).some(r=>+r>=bounds.r1&&+r<=bounds.r2))return cell=>cell;
  if(book.nativeSource)throw Error('행 기본 서식이 있는 XLSM 범위는 Excel에서 정렬하거나 XLSX 사본을 사용하세요.');
  const known=new Map((book.xlsxStyles||[]).map((entry,i)=>[JSON.stringify(entry),i])),templates=new Map();
  return (cell,sourceRow,targetRow,column)=>{
    if(rows[sourceRow]===rows[targetRow]||cell?.xlsxStyle!==undefined||cell?.xlsxNormal)return cell;
    const ref=address(sourceRow,column),record=cell||{raw:''},inherited=[sheet.xlsxStyleDefaults?.columns?.[column],rows[sourceRow]].map(id=>book.xlsxStyles?.[id]).filter(Boolean);
    if(!inherited.length)return {...record,xlsxNormal:true};
    const key=String(sheet.xlsxStyleDefaults?.columns?.[column]??'')+'/'+String(rows[sourceRow]??'');
    let id=templates.get(key);
    if(id===undefined){
      const bare={...sheet,cells:{[ref]:{raw:''}}},style=Object.assign({},book.xlsxStyles?.[book.xlsxDefaultStyle]?.style,...inherited.map(t=>t.style));
      const entry=JSON.parse(JSON.stringify({style,baseline:worksheetCellStyle(book,bare,sourceRow,column,{table:false})})),signature=JSON.stringify(entry);
      id=known.get(signature);
      if(id===undefined){book.xlsxStyles||=[];if(book.xlsxStyles.length>=65536)throw Error('통합 문서의 셀 서식 수가 한도를 초과하여 정렬하지 않았습니다.');id=book.xlsxStyles.length;book.xlsxStyles.push(entry);known.set(signature,id);}
      templates.set(key,id);
    }
    const baseline=book.xlsxStyles[id].baseline;
    return {...record,xlsxStyle:id,style:Object.keys(record.style||{}).length?{...baseline,...record.style}:baseline};
  };
}
