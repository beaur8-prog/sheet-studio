import {address,parseAddress,shiftFormula,transformFormulaReferences} from './engine.mjs';
import {cloneData,unwrap} from './clone.mjs';
import {tableBounds} from './table-schema.mjs';
import {editableTableDefinition,renameTableReferences} from './tables.mjs';
import {shiftAutoFilter,applyFilters} from './filters.mjs';
import {xml,stringify,child,children,all,attr} from './chart-xml.mjs';
const label=b=>address(b.r1,b.c1)+':'+address(b.r2,b.c2);
export function assertTableStructureAllowed(book,sheet,axis,index,count,delta){
  const removing=[];
  for(const table of sheet.tables||[]){
    const b=tableBounds(table.range),lo=axis==='r'?b.r1:b.c1,hi=axis==='r'?b.r2:b.c2,end=index+count-1;
    if(delta>0&&index<=hi&&hi+count>=(axis==='r'?1048576:16384))throw Error('표가 시트의 마지막 행·열을 넘어갑니다.');
    const intersects=delta>0?index>lo&&index<=hi:index<=hi&&end>=lo;
    if(!intersects)continue;editableTableDefinition(book,sheet,table);
    if(delta<0&&index<=lo&&end>=hi){removing.push(table.id);continue;}
    if(axis==='r'&&delta<0&&table.headerRow&&index<=b.r1&&end>=b.r1)throw Error('표의 머리글 행 일부를 삭제할 수 없습니다. 표 전체를 삭제하거나 일반 범위로 변환하세요.');
  }
  return removing;
}
export function removeTablesForStructure(book,sheet,ids){
  for(const id of ids){const table=sheet.tables.find(t=>t.id===id);if(table)renameTableReferences(book,sheet,table,{name:null});}
  if(ids.length)sheet.tables=sheet.tables.filter(t=>!ids.includes(t.id));
}
function nextColumn(table,position){const used=new Set(table.columns.map(c=>c.name.toLocaleLowerCase())),base='열'+(position+1);let name=base,n=2;while(used.has(name.toLocaleLowerCase()))name=base+n++;return {id:Math.max(0,...table.columns.map(c=>c.id||0))+1,name,filterButton:true};}
function styleCell(sheet,r,c,neighbor){
  const source=sheet.cells[address(neighbor.r,neighbor.c)]||{},cell={raw:'',style:cloneData(source.style||{})};for(const key of ['xlsxStyle','xlsxProtection'])if(source[key]!==undefined)cell[key]=cloneData(source[key]);
  return cell;
}
function updateSource(table,b,old,axis,index,delta){
  if(!table.sourceXML)return;
  const doc=xml(table.sourceXML),root=doc.documentElement,filter=child(root,'autoFilter');root.setAttribute('ref',label(b));
  if(filter&&table.filter){filter.setAttribute('ref',label(table.filter.range));
    if(axis==='c'&&index>=old.c1&&index<=old.c2){const at=index-old.c1;for(const node of children(filter,'filterColumn')){const column=Number(attr(node,'colId'));if(delta<0&&column===at)filter.removeChild(node);else if(column>=at)node.setAttribute('colId',String(column+delta));}}
    for(const state of all(filter,'sortState'))if(table.filter.range.r2>table.filter.range.r1)state.setAttribute('ref',label({...table.filter.range,r1:table.filter.range.r1+1}));else state.parentNode.removeChild(state);
    for(const node of all(filter,'sortCondition')){const first=parseAddress(attr(node,'ref').split(':')[0]);if(first){let c=first.c;if(axis==='c'&&c>=index)c+=delta;node.setAttribute('ref',address(table.filter.range.r1+1,c)+':'+address(table.filter.range.r2,c));}}
  }
  table.sourceXML=stringify(doc);
}
/** Called after the sheet's cells and ordinary A1 references have moved. Table
 * templates are rebased to their new first data row, then inserted rows inherit
 * the calculated columns and formatting without cloning existing notes/values.
 */
export function shiftTables(book,sheetId,axis,index,delta){
  const target=book.sheets.find(s=>s.id===sheetId);if(!target)return;
  for(const owner of book.sheets){
    const tables=[...(owner.tables||[])];
    for(const table of tables){
      const old=tableBounds(table.range),local=owner.id===sheetId,b={...old},lo=axis==='r'?'r1':'c1',hi=axis==='r'?'r2':'c2';
      if(local){
        if(delta>0){if(index<=b[lo]){b[lo]++;b[hi]++;}else if(index<=b[hi])b[hi]++;}
        else if(index<b[lo]){b[lo]--;b[hi]--;}else if(index<=b[hi])b[hi]--;
      }
      if(local&&axis==='c'&&delta<0&&index>=old.c1&&index<=old.c2){const removed=table.columns[index-old.c1];renameTableReferences(book,owner,table,{columns:{[removed.name]:null}});}
      const next=cloneData(table),origins=next.columns.map((column,i)=>({column,index:i}));
      if(local&&axis==='c'&&index>old.c1&&index<=old.c2&&delta>0){const at=index-old.c1;next.columns.splice(at,0,nextColumn(next,at));}
      if(local&&axis==='c'&&index>=old.c1&&index<=old.c2&&delta<0)next.columns.splice(index-old.c1,1);
      if(local&&axis==='r'&&delta<0&&table.totalsRow&&index===old.r2)next.totalsRow=false;
      next.range=label(b);
      for(let columnIndex=0;columnIndex<next.columns.length;columnIndex++){
        const column=next.columns[columnIndex],original=origins.find(entry=>entry.column===column);if(!original)continue;
        const transform=f=>transformFormulaReferences(f,{axis,index,delta,targetSheet:target.name,ownerSheet:owner.name});
        if(column.formula!==undefined){
          let formula='='+column.formula.replace(/^=/,''),origin={r:old.r1+(table.headerRow?1:0),c:old.c1+original.index};
          if(local&&axis==='r'&&delta<0&&index===origin.r){formula=shiftFormula(formula,1,0);origin.r++;}
          const moved={...origin};if(local&&moved[axis]>=index)moved[axis]+=delta;
          formula=transform(formula);if(local)formula=shiftFormula(formula,b.r1+(next.headerRow?1:0)-moved.r,b.c1+columnIndex-moved.c);column.formula=formula.slice(1);
        }
        if(column.totalsFormula!==undefined)column.totalsFormula=transform(column.totalsFormula);
      }
      if(next.filter){const view={autoFilter:next.filter};if(local)shiftAutoFilter(view,axis,index,delta);next.filter=view.autoFilter;if(next.filter)next.filter.range={...b,r2:b.r2-(next.totalsRow?1:0)};}
      if(local)updateSource(next,b,old,axis,index,delta);
      const at=owner.tables.findIndex(t=>t.id===table.id);owner.tables[at]=next;
      if(local&&axis==='c'&&delta>0&&index>old.c1&&index<=old.c2&&next.headerRow){const key=address(b.r1,index),cell=styleCell(owner,b.r1,index,{r:b.r1,c:index+1});cell.raw=next.columns[index-b.c1].name;owner.cells[key]=cell;}
      const first=b.r1+(next.headerRow?1:0),last=b.r2-(next.totalsRow?1:0);
      if(local&&axis==='r'&&delta>0&&index>=first&&index<=last){
        for(let i=0;i<next.columns.length;i++){const c=b.c1+i,cell=styleCell(owner,index,c,{r:index<last?index+1:index-1,c}),column=next.columns[i];if(column.formula!==undefined)cell.raw=shiftFormula('='+column.formula.replace(/^=/,''),index-first,0);
          if(cell.raw||Object.keys(cell.style).length||cell.xlsxStyle!==undefined||cell.xlsxProtection!==undefined)owner.cells[address(index,c)]=cell;
        }
      }
    }
  }
  const snapshot={...unwrap(book),sheets:unwrap(book).sheets.map(s=>({...s}))};
  for(const table of target.tables||[])if(table.filter&&!table.filter.unsupported?.length)applyFilters(snapshot,{...target,autoFilter:table.filter});book.referenceCachesValid=false;
}
