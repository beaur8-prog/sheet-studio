import {address,parseAddress,shiftFormula} from './engine.mjs';
import {cloneData} from './clone.mjs';
import {editableTableDefinition,appendTableColumnsPlan} from './tables.mjs';
import {tableBounds,overlaps} from './table-schema.mjs';
import {arrayRegions} from './cell-constraints.mjs';
const label=b=>address(b.r1,b.c1)+':'+address(b.r2,b.c2);
/** Adjacent vertical/horizontal growth and first formula autofill. Explicit values and
 * formula exceptions are never overwritten by these automatic operations.
 */
export function prepareTableDataEdits(book,sheet,plan){
  const updates=new Map(plan.updates.map(u=>[u.key,u])),changes=new Map(plan.tables.map(t=>[t.id,t])),input=plan.updates.map(u=>({...u,position:parseAddress(u.key)}));
  for(const original of sheet.tables||[]){
    if(original.filter?.unsupported?.length)continue;
    try{editableTableDefinition(book,sheet,original);}catch{continue;}
    const prior=changes.get(original.id);let table=cloneData(prior?.table||original);const old=tableBounds(table.range),b={...old},first=b.r1+(table.headerRow?1:0);let changed=false;
    // A wider paste below the table adds rows at the old width. Right-hand
    // columns grow only when input touches a row already inside the table.
    const cols=new Set(input.filter(u=>u.cell.raw&&u.position.r>=old.r1&&u.position.r<=old.r2-(table.totalsRow?1:0)&&u.position.c>old.c2).map(u=>u.position.c));
    while(cols.has(b.c2+1))b.c2++;
    if(b.c2>old.c2){
      const conflict=(sheet.tables||[]).some(t=>t.id!==table.id&&overlaps(tableBounds(t.range),b))||(sheet.merges||[]).some(m=>overlaps(tableBounds(m),b))||arrayRegions(sheet).some(a=>overlaps(a,b));
      const occupiedTotal=table.totalsRow&&input.some(u=>u.position.r===old.r2&&u.position.c>old.c2&&u.position.c<=b.c2&&u.cell.raw)||table.totalsRow&&Array.from({length:b.c2-old.c2},(_,i)=>sheet.cells[address(old.r2,old.c2+i+1)]?.raw).some(Boolean);
      if(!conflict&&!occupiedTotal&&b.c2-old.c2+updates.size<=100000){const expanded=appendTableColumnsPlan(book,sheet,table,b.c2,new Map([...updates].map(([key,u])=>[key,u.cell])));table=expanded.table;for(const u of expanded.updates)updates.set(u.key,u);changed=true;}else b.c2=old.c2;
    }
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
      }else b.r2=old.r2;
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
