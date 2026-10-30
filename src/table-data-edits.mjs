import {address,parseAddress,shiftFormula} from './engine.mjs';
import {cloneData} from './clone.mjs';
import {editableTableDefinition} from './tables.mjs';
import {tableBounds,overlaps} from './table-schema.mjs';
import {arrayRegions} from './cell-constraints.mjs';
const label=b=>address(b.r1,b.c1)+':'+address(b.r2,b.c2);
/** Excel-style vertical growth and first formula autofill. Explicit values and
 * formula exceptions are never overwritten by these automatic operations.
 */
export function prepareTableDataEdits(book,sheet,plan){
  const updates=new Map(plan.updates.map(u=>[u.key,u])),changes=new Map(plan.tables.map(t=>[t.id,t])),input=plan.updates.map(u=>({...u,position:parseAddress(u.key)}));
  for(const original of sheet.tables||[]){
    if(original.filter?.unsupported?.length)continue;
    try{editableTableDefinition(book,sheet,original);}catch{continue;}
    const prior=changes.get(original.id),table=cloneData(prior?.table||original),old=tableBounds(table.range),b={...old},first=b.r1+(table.headerRow?1:0);let changed=false;
    if(!table.totalsRow){
      const rows=new Set(input.filter(u=>u.cell.raw&&u.position.c>=b.c1&&u.position.c<=b.c2&&u.position.r>b.r2).map(u=>u.position.r));while(rows.has(b.r2+1))b.r2++;
      const added=(b.r2-old.r2)*(b.c2-b.c1+1),conflict=(sheet.tables||[]).some(t=>t.id!==table.id&&overlaps(tableBounds(t.range),b))||(sheet.merges||[]).some(m=>overlaps(tableBounds(m),b))||arrayRegions(sheet).some(a=>overlaps(a,b));
      if(added&&added+updates.size<=100000&&!conflict){
        table.range=label(b);if(table.filter)table.filter.range={...b};changed=true;
        for(let r=old.r2+1;r<=b.r2;r++)for(let c=b.c1;c<=b.c2;c++){
          const key=address(r,c),existing=updates.get(key)?.cell||sheet.cells[key],neighbor=sheet.cells[address(Math.max(first,old.r2),c)]||{},cell={...cloneData(existing||{raw:''}),style:{...cloneData(neighbor.style||{}),...existing?.style}};
          for(const field of ['xlsxStyle','xlsxProtection'])if(cell[field]===undefined&&neighbor[field]!==undefined)cell[field]=cloneData(neighbor[field]);
          const column=table.columns[c-b.c1];if(!cell.raw&&column.formula!==undefined)cell.raw=shiftFormula('='+column.formula.replace(/^=/,''),r-first,0);
          if(cell.raw||Object.keys(cell.style).length||cell.xlsxStyle!==undefined||cell.xlsxProtection!==undefined||updates.has(key))updates.set(key,{key,cell});
        }
      }else Object.assign(b,old);
    }
    const last=b.r2-(table.totalsRow?1:0);
    if(input.length===1){
      const update=input[0],p=update.position,column=table.columns[p.c-b.c1];
      if(column&&column.formula===undefined&&update.cell.raw?.startsWith('=')&&p.r>=first&&p.r<=last&&last-first+1+updates.size<=100000){
        let empty=true;for(let r=first;r<=last;r++)if(r!==p.r&&(updates.get(address(r,p.c))?.cell||sheet.cells[address(r,p.c)])?.raw){empty=false;break;}
        if(empty){column.formula=shiftFormula(update.cell.raw,first-p.r,0).slice(1);changed=true;
          for(let r=first;r<=last;r++){const key=address(r,p.c),cell=updates.get(key)?.cell||sheet.cells[key]||{};updates.set(key,{key,cell:{...cloneData(cell),raw:shiftFormula(update.cell.raw,r-p.r,0)}});}
        }
      }
    }
    if(changed)changes.set(table.id,{id:table.id,table,renames:prior?.renames||{}});
  }
  return {updates:[...updates.values()],tables:[...changes.values()]};
}
