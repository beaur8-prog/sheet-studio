import {ensureTablePreset} from './table-style.mjs';
import {applyFilters} from './filters.mjs';
import {address,parseAddress,formatValue,createEvaluator,shiftFormula,rewriteTableFormula} from './engine.mjs';
import {cloneData,unwrap} from './clone.mjs';
import {scalarRaw} from './cell-paste.mjs';
import {nameError} from './named-ranges.mjs';
import {tableBounds,tableAt,overlaps,validateTables} from './table-schema.mjs';
import {escapeColumn,resolveStructuredReference,structuredParts} from './table-references.mjs';
import {arrayRegions,sheetIsProtected} from './cell-constraints.mjs';
import {transformPreservedFeatures} from './xlsx-feature-transforms.mjs';
const label=b=>address(b.r1,b.c1)+':'+address(b.r2,b.c2),inside=(b,p)=>p.r>=b.r1&&p.r<=b.r2&&p.c>=b.c1&&p.c<=b.c2;
export function editableTableDefinition(book,sheet,table){
  if(book.nativeSource)throw Error('XLSM의 표 구성 편집은 현재 지원하지 않습니다. Excel에서 편집하거나 XLSX 사본을 사용하세요.');
  if(sheetIsProtected(sheet))throw Error('보호된 시트에서는 표 구성을 변경할 수 없습니다.');
  if(table?.sourceRels?.some(r=>r.type.endsWith('/queryTable'))||/tableType=["']queryTable["']/.test(table?.sourceXML||''))throw Error('외부 데이터에 연결된 표의 구성은 Excel에서 편집하세요. 표의 셀 값은 편집할 수 있습니다.');
}
function checkRange(sheet,b,original){
  if(!b)throw Error('표 범위를 확인하세요. 예: A1:C10');
  if((sheet.tables||[]).some(t=>t.id!==original?.id&&overlaps(tableBounds(t.range),b)))throw Error('다른 표와 범위가 겹칩니다.');
  if((sheet.merges||[]).some(range=>overlaps(tableBounds(range),b)))throw Error('병합 셀이 있는 범위에는 표를 만들 수 없습니다.');
  if(arrayRegions(sheet).some(a=>overlaps(a,b)))throw Error('배열 수식이 있는 범위에는 표를 만들 수 없습니다.');
}
function checkName(book,name,original){
  const probe={...book,sheets:book.sheets.map(s=>({...s,tables:s.tables?.filter(t=>t.id!==original?.id)}))},error=nameError(probe,name);
  if(error)throw Error(error);
}
const unique=(name,used,fallback)=>{const base=String(name||fallback).slice(0,32760)||fallback;let value=base,n=2;while(used.has(value.toLocaleLowerCase()))value=base+n++;used.add(value.toLocaleLowerCase());return value;};
export function nextTableName(book){let n=1;while(true){const name='Table'+n++;try{checkName(book,name);return name;}catch{}}}
export function createTablePlan(book,sheet,{range,name,headerRow=true}){
  editableTableDefinition(book,sheet);const b=tableBounds(range);checkRange(sheet,b);name=String(name||nextTableName(book)).trim();checkName(book,name);
  const evaluate=createEvaluator(book,sheet.id),used=new Set(),columns=[],updates=[];
  for(let c=b.c1;c<=b.c2;c++){
    const key=address(b.r1,c),old=sheet.cells[key],value=headerRow?evaluate(key):'',text=value===''?'':formatValue(value,old?.style,book),name=unique(text,used,'열'+(c-b.c1+1));columns.push({id:c-b.c1+1,name,filterButton:true});
    if(headerRow)updates.push({key,cell:{...cloneData(old||{}),raw:scalarRaw(name)}});
  }
  const table={id:crypto.randomUUID(),name,range:label(b),headerRow:!!headerRow,totalsRow:false,columns,style:{name:'TableStyleMedium2',showRowStripes:true,showColumnStripes:false,showFirstColumn:false,showLastColumn:false},...(headerRow?{filter:{range:{...b},columns:{},hiddenRows:[]}}:{})};validateTables([table]);return {table,updates,create:true};
}
const totalCodes={average:101,countNums:102,count:103,max:104,min:105,stdDev:107,var:110,sum:109};
export function totalFormula(table,column){if(column.totalsFunction==='custom')return column.totalsFormula?'='+column.totalsFormula.replace(/^=/,''):null;const code=totalCodes[column.totalsFunction];return code?`=SUBTOTAL(${code},${table.name}[${escapeColumn(column.name)}])`:null;}
function assertEditCount(count){if(count>100000)throw Error('한 번에 최대 100,000개 셀의 표 수식을 적용할 수 있습니다. 범위를 나누거나 수식 적용 없이 범위를 변경하세요.');}
export function updateTablePlan(book,sheet,table,values){
  editableTableDefinition(book,sheet,table);const old=tableBounds(table.range),b=tableBounds(values.range||table.range);checkRange(sheet,b,table);
  if(b.r1!==old.r1||b.c1!==old.c1||b.c2!==old.c2)throw Error('표 크기 변경은 시작 셀과 열 수를 유지하고 마지막 행을 조절하세요.');
  const name=String(values.name||table.name).trim();checkName(book,name,table);const next={...cloneData(table),name,range:label(b),totalsRow:values.totalsRow??table.totalsRow},updates=[];
  if(b.r2-b.r1+1<(next.headerRow?1:0)+(next.totalsRow?1:0))throw Error('머리글과 요약 행을 포함할 범위가 필요합니다.');
  const relocating=next.totalsRow&&(!table.totalsRow||b.r2!==old.r2);
  if(relocating)for(let c=b.c1;c<=b.c2;c++){const key=address(b.r2,c);if(sheet.cells[key]?.raw&&!(table.totalsRow&&b.r2===old.r2))throw Error('새 요약 행 위치에 데이터가 있습니다. 아래에 빈 행을 확보하세요.');}
  if(table.totalsRow&&(!next.totalsRow||old.r2!==b.r2))for(let c=old.c1;c<=old.c2;c++){const key=address(old.r2,c);updates.push({key,cell:{...cloneData(sheet.cells[key]||{}),raw:''}});}
  if(relocating){
    const hadTotals=next.columns.some(c=>c.totalsFunction!==undefined||c.totalsLabel!==undefined||c.totalsFormula!==undefined);
    const evaluate=createEvaluator(book,sheet.id);next.columns.forEach((column,i)=>{
      if(!hadTotals&&column.totalsFunction===undefined&&!column.totalsLabel){if(i===0)column.totalsLabel='합계';else if(Array.from({length:Math.min(8,Math.max(0,b.r2-b.r1-(next.headerRow?1:0)))},(_,r)=>evaluate(address(b.r1+(next.headerRow?1:0)+r,b.c1+i))).some(v=>typeof v==='number'))column.totalsFunction='sum';}
      const key=address(b.r2,b.c1+i),raw=totalFormula(next,column)||scalarRaw(column.totalsLabel||'');updates.push({key,cell:{...cloneData(sheet.cells[key]||{}),raw}});
    });
  }
  const first=b.r1+(next.headerRow?1:0),last=b.r2-(next.totalsRow?1:0),oldLast=old.r2-(table.totalsRow?1:0);
  if(last<oldLast)for(const key in sheet.cells){const p=parseAddress(key),cell=sheet.cells[key];if(p.r>last&&p.r<=oldLast&&p.c>=old.c1&&p.c<=old.c2&&cell.raw?.startsWith('=')){const raw=materializeLeavingFormula(book,sheet,table,cell.raw,p);if(raw!==cell.raw)updates.push({key,cell:{...cloneData(cell),raw}});}}
  const added=Math.max(0,last-oldLast),formulaColumns=next.columns.filter(c=>c.formula!==undefined);assertEditCount(added*formulaColumns.length+updates.length);
  for(let r=Math.max(first,oldLast+1);r<=last;r++)for(let i=0;i<next.columns.length;i++){const column=next.columns[i],key=address(r,b.c1+i);if(column.formula===undefined||sheet.cells[key]?.raw&&!(table.totalsRow&&r===old.r2))continue;const raw=shiftFormula('='+column.formula.replace(/^=/,''),r-first,0);updates.push({key,cell:{...cloneData(sheet.cells[key]||{}),raw}});}
  if(next.filter){next.filter.range={...b,r2:b.r2-(next.totalsRow?1:0)};next.filter.hiddenRows=(next.filter.hiddenRows||[]).filter(r=>r>next.filter.range.r1&&r<=next.filter.range.r2);}
  validateTables([next]);return {table:next,updates,...(name!==table.name?{rename:{from:table.name,to:name}}:{})};
}
export function calculatedColumnPlan(book,sheet,table,{column,formula}){
  editableTableDefinition(book,sheet,table);if(!Number.isInteger(column)||column<0||column>=table.columns.length)throw Error('계산할 표 열을 선택하세요.');
  formula=String(formula||'').trim();if(!formula.startsWith('=')||formula.length>8192)throw Error('계산 열의 수식은 =로 시작하고 8,192자 이하로 입력하세요.');
  const next=cloneData(table),b=tableBounds(table.range),first=b.r1+(table.headerRow?1:0),last=b.r2-(table.totalsRow?1:0),updates=[];assertEditCount(Math.max(0,last-first+1));next.columns[column].formula=formula.slice(1);
  for(let r=first;r<=last;r++){const key=address(r,b.c1+column);updates.push({key,cell:{...cloneData(sheet.cells[key]||{}),raw:shiftFormula(formula,r-first,0)}});}return {table:next,updates};
}
export function materializeLeavingFormula(book,sheet,table,formula,position){
  const format=ref=>{const same=ref.sheet.id===sheet.id,one=ref.rows===1,first=address(ref.top,ref.left),last=address(ref.top+ref.rows-1,ref.left+ref.cols-1),absolute=value=>value.replace(/^([A-Z]+)(\d+)$/,'$$$1$$$2');
    return same&&one?(first===last?first:first+':'+last):"'"+ref.sheet.name.replace(/'/g,"''")+"'!"+absolute(first)+(first===last?'':':'+absolute(last));};
  return rewriteTableFormula(formula,{tableName:table.name,implicit:true,allowBare:false,materialize:({tableName,selector})=>{
    if(tableName&&!structuredParts(selector)?.flags.has('#thisrow')&&!/\[\s*@/.test(selector))return tableName+selector;
    const ref=resolveStructuredReference(book,sheet,position,table.name,selector);if(ref.error)return ref.error;if(ref.empty)return '0';return ref.areas?'('+ref.areas.map(format).join(',')+')':format(ref);
  }});
}
/** Rename table references as metadata changes. Scan raw cells, then journal only
 * formulas that actually change; a large sheet is not cloned on the UI thread.
 */
export function renameTableReferences(book,owner,table,{name=table.name,columns}={}){
  const raw=unwrap(book),rewrite=(formula,implicit=false,allowBare=true)=>rewriteTableFormula(formula,{tableName:table.name,newTableName:name,columns,implicit,allowBare});
  for(const sheet of book.sheets){const source=raw.sheets.find(s=>s.id===sheet.id),hasLocal=(book.chartNames||[]).some(n=>n.sheetId===sheet.id&&n.name.toLocaleLowerCase()===table.name.toLocaleLowerCase());
    for(const key in source.cells){const cell=source.cells[key];if(cell.raw?.startsWith('=')){const p=parseAddress(key),implicit=sheet.id===owner.id&&inside(tableBounds(table.range),p),formula=rewrite(cell.raw,implicit,!hasLocal);if(formula!==cell.raw)sheet.cells[key]={...cell,raw:formula};}}
    for(const rule of sheet.conditionalRules||[])rule.formulae=rule.formulae.map(f=>rewrite(f));transformPreservedFeatures(sheet,{formula:f=>rewrite(f)});
    for(const current of sheet.tables||[])for(const column of current.columns)for(const property of ['formula','totalsFormula'])if(column[property]!==undefined)column[property]=rewrite(column[property],current.id===table.id);
    for(const chart of sheet.charts||[]){for(const series of chart.series||[])for(const key of ['nameRef','catRef','valRef','xRef','sizeRef'])if(series[key]?.formula)series[key].formula=rewrite(series[key].formula);for(const ref of chart.source?.formulas||[])ref.formula=rewrite(ref.formula);}
  }
  for(const named of book.chartNames||[])named.formula=rewrite(named.formula);book.referenceCachesValid=false;
}
export function prepareTableHeaderEdits(book,sheet,updates){
  const changed=new Map(updates.map(u=>[u.key,u])),tables=[];
  for(const table of sheet.tables||[]){if(!table.headerRow)continue;const b=tableBounds(table.range),edits=updates.filter(u=>{const p=parseAddress(u.key);return p.r===b.r1&&p.c>=b.c1&&p.c<=b.c2;});if(!edits.length)continue;
    const next=cloneData(table),edited=new Set(edits.map(u=>parseAddress(u.key).c-b.c1)),used=new Set(table.columns.filter((_,i)=>!edited.has(i)).map(c=>c.name.toLocaleLowerCase())),renames={};
    for(const update of edits){const i=parseAddress(update.key).c-b.c1;if(update.cell.raw.startsWith('='))throw Error('표 머리글에는 수식 대신 열 이름을 입력하세요.');const text=update.cell.raw.replace(/^'/,''),name=unique(text,used,'열'+(i+1));if(table.columns[i].name!==name)renames[table.columns[i].name]=name;next.columns[i].name=name;changed.set(update.key,{...update,cell:{...update.cell,raw:scalarRaw(name)}});}
    if(Object.keys(renames).length){editableTableDefinition(book,sheet,table);tables.push({id:table.id,table:next,renames});}
  }
  for(const change of tables){const original=sheet.tables.find(t=>t.id===change.id),b=tableBounds(original.range);for(const [key,update] of changed)if(update.cell.raw.startsWith('=')){const raw=rewriteTableFormula(update.cell.raw,{tableName:original.name,columns:change.renames,implicit:inside(b,parseAddress(key))});if(raw!==update.cell.raw)changed.set(key,{...update,cell:{...update.cell,raw}});}}
  return {updates:[...changed.values()],tables};
}
export function applyTableHeaderEdits(book,sheet,plan){for(const change of plan.tables){const current=sheet.tables.find(t=>t.id===change.id);renameTableReferences(book,sheet,current,{columns:change.renames});const index=sheet.tables.findIndex(t=>t.id===change.id),next=change.table;for(const column of next.columns)for(const key of ['formula','totalsFormula'])if(column[key]!==undefined)column[key]=rewriteTableFormula(column[key],{tableName:current.name,columns:change.renames,implicit:true});sheet.tables[index]=next;}}

export function finalizeTablePlan(book,sheet,plan,{preset}={}){
  const snapshot={...book,sheets:[],...(preset?{xlsxDxfs:[...(book.xlsxDxfs||[])],xlsxTableStyles:[...(book.xlsxTableStyles||[])]}:{})};
  if(preset){plan.table.style.name=ensureTablePreset(snapshot,preset);plan.styleMetadata={xlsxDxfs:snapshot.xlsxDxfs,xlsxTableStyles:snapshot.xlsxTableStyles};}
  const overrides=new Map(plan.updates.map(update=>[update.key,update.cell])),renamed=plan.rename,rewrite=formula=>renamed?rewriteTableFormula(formula,{tableName:renamed.from,newTableName:renamed.to}):formula;
  snapshot.chartNames=book.chartNames?.map(named=>({...named,formula:rewrite(named.formula)}));
  snapshot.sheets=book.sheets.map(source=>{const target=source.id===sheet.id,updates=target?overrides:new Map(),cells=new Proxy(source.cells,{get:(base,key)=>{const cell=updates.has(key)?updates.get(key):Reflect.get(base,key);if(renamed&&cell?.raw?.startsWith('=')){const raw=rewrite(cell.raw);if(raw!==cell.raw)return {...cell,raw};}return cell;},ownKeys:base=>[...new Set([...Reflect.ownKeys(base),...updates.keys()])],getOwnPropertyDescriptor:(base,key)=>updates.has(key)?{value:updates.get(key),enumerable:true,configurable:true,writable:true}:Reflect.getOwnPropertyDescriptor(base,key)});
    return {...source,cells,...(target?{tables:plan.create?[...(source.tables||[]),plan.table]:(source.tables||[]).map(t=>t.id===plan.table.id?plan.table:t)}:{})};});
  if(plan.table.filter&&!plan.table.filter.unsupported?.length)applyFilters(snapshot,{...snapshot.sheets.find(s=>s.id===sheet.id),autoFilter:plan.table.filter});return plan;
}
