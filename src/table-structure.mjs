import {address,parseAddress,shiftFormula,transformFormulaReferences} from './engine.mjs';
import {cloneData,unwrap} from './clone.mjs';
import {tableBounds} from './table-schema.mjs';
import {editableTableDefinition,renameTableReferences} from './tables.mjs';
import {structureIndex} from './structure-coordinates.mjs';
import {MAX_DOCUMENT_CELLS} from './limits.mjs';
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
function newColumns(table,position,count){const used=new Set(table.columns.map(c=>c.name.toLocaleLowerCase()));let id=table.columns.reduce((n,c)=>Math.max(n,c.id||0),0);return Array.from({length:count},(_,i)=>{const base='열'+(position+i+1);let name=base,n=2;while(used.has(name.toLocaleLowerCase()))name=base+n++;used.add(name.toLocaleLowerCase());return {id:++id,name,filterButton:true};});}
function styleCell(sheet,r,c,neighbor){
  const source=sheet.cells[address(neighbor.r,neighbor.c)]||{},cell={raw:'',style:cloneData(source.style||{})};for(const key of ['xlsxStyle','xlsxProtection','xlsxNormal'])if(source[key]!==undefined)cell[key]=cloneData(source[key]);
  return cell;
}
function updateSource(table,b,old,axis,index,delta){
  if(!table.sourceXML)return;
  const doc=xml(table.sourceXML),root=doc.documentElement,filter=child(root,'autoFilter');root.setAttribute('ref',label(b));
  if(filter&&table.filter){filter.setAttribute('ref',label(table.filter.range));
    if(axis==='c')for(const node of children(filter,'filterColumn')){const column=structureIndex(old.c1+Number(attr(node,'colId')),index,delta);if(column===null)filter.removeChild(node);else node.setAttribute('colId',String(column-b.c1));}
    for(const state of all(filter,'sortState'))if(table.filter.range.r2>table.filter.range.r1)state.setAttribute('ref',label({...table.filter.range,r1:table.filter.range.r1+1}));else state.parentNode.removeChild(state);
    for(const node of all(filter,'sortCondition')){const first=parseAddress(attr(node,'ref').split(':')[0]);if(first){const c=axis==='c'?structureIndex(first.c,index,delta):first.c;if(c===null)node.parentNode.removeChild(node);else node.setAttribute('ref',address(table.filter.range.r1+1,c)+':'+address(table.filter.range.r2,c));}}
  }
  table.sourceXML=stringify(doc);
}
/** Called after the sheet's cells and ordinary A1 references have moved. Table
 * templates are rebased to their new first data row, then inserted rows inherit
 * the calculated columns and formatting without cloning existing notes/values.
 */
export function shiftTables(book,sheetId,axis,index,delta){
  const target=book.sheets.find(s=>s.id===sheetId);if(!target)return;
  const count=Math.abs(delta),end=index+count-1;let cells=book.sheets.reduce((n,s)=>n+Object.keys(s.cells).length,0);
  const put=(owner,key,cell)=>{if(!Object.hasOwn(owner.cells,key)&&++cells>MAX_DOCUMENT_CELLS)throw Error('새 표 행·열의 수식과 서식을 포함하면 1,000,000셀 한도를 초과합니다.');owner.cells[key]=cell;};
  for(const owner of book.sheets){
    const tables=[...(owner.tables||[])];
    for(const table of tables){
      const old=tableBounds(table.range),local=owner.id===sheetId,b=local?tableBounds(transformFormulaReferences(table.range,{axis,index,delta,targetSheet:target.name,ownerSheet:owner.name})):{...old};
      if(!b)throw Error('표 전체를 삭제할 때는 표 정의를 먼저 제거해야 합니다.');
      const removedColumns=local&&axis==='c'&&delta<0?table.columns.filter((_,i)=>old.c1+i>=index&&old.c1+i<=end):[];
      if(removedColumns.length)renameTableReferences(book,owner,table,{columns:Object.fromEntries(removedColumns.map(c=>[c.name,null]))});
      const next=cloneData(table),origins=next.columns.map((column,i)=>({column,index:i}));
      if(local&&axis==='c'&&index>old.c1&&index<=old.c2&&delta>0){const at=index-old.c1;next.columns.splice(at,0,...newColumns(next,at,count));}
      if(removedColumns.length)next.columns.splice(Math.max(0,index-old.c1),removedColumns.length);
      if(local&&axis==='r'&&delta<0&&table.totalsRow&&old.r2>=index&&old.r2<=end)next.totalsRow=false;
      next.range=label(b);
      for(let columnIndex=0;columnIndex<next.columns.length;columnIndex++){
        const column=next.columns[columnIndex],original=origins.find(entry=>entry.column===column);if(!original)continue;
        const transform=f=>transformFormulaReferences(f,{axis,index,delta,targetSheet:target.name,ownerSheet:owner.name});
        if(column.formula!==undefined){
          let formula='='+column.formula.replace(/^=/,''),origin={r:old.r1+(table.headerRow?1:0),c:old.c1+original.index};
          if(local&&axis==='r'&&delta<0&&origin.r>=index&&origin.r<=end){const amount=end+1-origin.r;formula=shiftFormula(formula,amount,0);origin.r+=amount;}
          const moved={...origin};if(local)moved[axis]=structureIndex(moved[axis],index,delta)??index;
          formula=transform(formula);if(local)formula=shiftFormula(formula,b.r1+(next.headerRow?1:0)-moved.r,b.c1+columnIndex-moved.c);column.formula=formula.slice(1);
        }
        if(column.totalsFormula!==undefined)column.totalsFormula=transform(column.totalsFormula);
      }
      if(next.filter){const view={autoFilter:next.filter};if(local)shiftAutoFilter(view,axis,index,delta);next.filter=view.autoFilter;if(next.filter)next.filter.range={...b,r2:b.r2-(next.totalsRow?1:0)};}
      if(local)updateSource(next,b,old,axis,index,delta);
      const at=owner.tables.findIndex(t=>t.id===table.id);owner.tables[at]=next;
      if(local&&axis==='c'&&delta>0&&index>old.c1&&index<=old.c2&&next.headerRow)for(let c=index;c<index+count;c++){const key=address(b.r1,c),cell=styleCell(owner,b.r1,c,{r:b.r1,c:index+count});cell.raw=next.columns[c-b.c1].name;put(owner,key,cell);}
      const first=b.r1+(next.headerRow?1:0),last=b.r2-(next.totalsRow?1:0);
      if(local&&axis==='r'&&delta>0&&index>=first&&index<=last){
        for(let i=0;i<next.columns.length;i++){const c=b.c1+i,cell=styleCell(owner,index,c,{r:index+count<=last?index+count:index-1,c}),column=next.columns[i];
          if(column.formula!==undefined||Object.keys(cell.style).length||cell.xlsxStyle!==undefined||cell.xlsxProtection!==undefined)for(let r=Math.max(index,first);r<=Math.min(index+count-1,last);r++){const added=cloneData(cell);if(column.formula!==undefined)added.raw=shiftFormula('='+column.formula.replace(/^=/,''),r-first,0);put(owner,address(r,c),added);}
        }
      }
    }
  }
  const snapshot={...unwrap(book),sheets:unwrap(book).sheets.map(s=>({...s}))};
  for(const table of target.tables||[])if(table.filter&&!table.filter.unsupported?.length){if(Object.keys(table.filter.columns).length)applyFilters(snapshot,{...target,autoFilter:table.filter});else table.filter.hiddenRows=[];}book.referenceCachesValid=false;
}
