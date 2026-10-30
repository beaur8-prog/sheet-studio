import {address,parseAddress,rewriteTableFormula} from './engine.mjs';
import {cloneData,unwrap} from './clone.mjs';
import {editableTableDefinition} from './tables.mjs';
import {tableBounds} from './table-schema.mjs';
import {worksheetCellStyle} from './cell-style.mjs';
import {resolveStructuredReference} from './table-references.mjs';
import {transformPreservedFeatures} from './xlsx-feature-transforms.mjs';
const inside=(b,p)=>p.r>=b.r1&&p.r<=b.r2&&p.c>=b.c1&&p.c<=b.c2;
/** A read-only planning snapshot shares cells until they actually change. */
export function tableConversionPlan(book,owner,table){
  editableTableDefinition(book,owner,table);const b=tableBounds(table.range);
  if(b.r2-b.r1+1-(table.headerRow?1:0)-(table.totalsRow?1:0)<1)throw Error('데이터 행이 없는 표는 먼저 데이터 행을 추가한 뒤 변환하세요.');
  if((b.r2-b.r1+1)*(b.c2-b.c1+1)>100000)throw Error('서식을 유지하며 일반 범위로 변환하는 작업은 한 번에 최대 100,000셀까지 지원합니다.');
  const updates=new Map(),snapshot={...unwrap(book),sheets:[]};
  for(const source of unwrap(book).sheets){
    const changed=new Map();updates.set(source.id,changed);
    const cells=new Proxy(source.cells,{get:(base,key)=>changed.has(key)?changed.get(key):Reflect.get(base,key),set:(_base,key,value)=>{changed.set(key,value);return true;},ownKeys:base=>[...new Set([...Reflect.ownKeys(base),...changed.keys()])],getOwnPropertyDescriptor:(base,key)=>changed.has(key)?{value:changed.get(key),enumerable:true,configurable:true,writable:true}:Reflect.getOwnPropertyDescriptor(base,key)});
    const next={...source,cells};for(const key of ['tables','charts','conditionalRules','xlsxFeatures'])if(source[key]!==undefined)next[key]=cloneData(source[key]);snapshot.sheets.push(next);
  }
  const target=snapshot.sheets.find(s=>s.id===owner.id),targetChanges=updates.get(owner.id);
  for(let r=b.r1;r<=b.r2;r++)for(let c=b.c1;c<=b.c2;c++){
    const key=address(r,c),cell=target.cells[key]||{raw:'',style:{}},style=worksheetCellStyle(book,owner,r,c);
    if(Object.keys(style).length)targetChanges.set(key,{...cell,style:{...style,...cell.style}});
  }
  const rewrite=(formula,sheet,position={r:0,c:0},implicit=false)=>{
    const local=(book.chartNames||[]).some(n=>n.sheetId===sheet.id&&n.name.toLocaleLowerCase()===table.name.toLocaleLowerCase());
    return rewriteTableFormula(formula,{tableName:table.name,implicit,allowBare:!local,materialize:({selector,sheetName})=>{
      const ref=resolveStructuredReference(book,sheet,position,table.name,selector,sheetName);if(ref.error)return ref.error;if(ref.empty)return '#REF!';
      const format=area=>{const first=address(area.top,area.left),last=address(area.top+area.rows-1,area.left+area.cols-1),absolute=value=>value.replace(/^([A-Z]+)(\d+)$/,'$$$1$$$2'),prefix="'"+area.sheet.name.replace(/'/g,"''")+"'!",isRow=/@|#thisrow/i.test(selector)&&sheet.id===area.sheet.id;
        return isRow?first+(first===last?'':':'+last):prefix+absolute(first)+(first===last?'':':'+absolute(last));};
      return ref.areas?'('+ref.areas.map(format).join(',')+')':format(ref);
    }});
  };
  for(const sheet of snapshot.sheets){
    for(const key in sheet.cells){const cell=sheet.cells[key],p=parseAddress(key),implicit=sheet.id===owner.id&&inside(b,p),raw=cell.raw?.startsWith('=')?rewrite(cell.raw,sheet,p,implicit):cell.raw;
      if(raw!==cell.raw||cell.xlsxFormula)sheet.cells[key]={...cell,raw,...(cell.xlsxFormula?{xlsxFormula:cloneData(cell.xlsxFormula)}:{})};
    }
    for(const rule of sheet.conditionalRules||[])rule.formulae=rule.formulae.map(f=>rewrite(f,sheet));
    transformPreservedFeatures(sheet,{formula:f=>rewrite(f,sheet)});
    for(const current of sheet.tables||[]){const cb=tableBounds(current.range);for(let i=0;i<current.columns.length;i++)for(const property of ['formula','totalsFormula'])if(current.columns[i][property]!==undefined)current.columns[i][property]=rewrite(current.columns[i][property],sheet,{r:property==='formula'?cb.r1+(current.headerRow?1:0):cb.r2,c:cb.c1+i},current.id===table.id);}
    for(const chart of sheet.charts||[]){for(const series of chart.series||[])for(const key of ['nameRef','catRef','valRef','xRef','sizeRef'])if(series[key]?.formula)series[key].formula=rewrite(series[key].formula,sheet);for(const ref of chart.source?.formulas||[])ref.formula=rewrite(ref.formula,sheet);}
  }
  target.tables=target.tables.filter(t=>t.id!==table.id);const hidden=new Set(table.filter?.hiddenRows||[]),otherHidden=new Set([...(target.autoFilter?.hiddenRows||[]),...(target.tables||[]).flatMap(t=>t.filter?.hiddenRows||[])]);
  const rowHeights={...(target.rowHeights||{})};for(const row of hidden)if(!otherHidden.has(row)&&rowHeights[row]!==0)rowHeights[row]=0;
  const chartNames=book.chartNames?.map(n=>({...cloneData(n),formula:rewrite(n.formula,snapshot.sheets.find(s=>s.id===n.sheetId)||target)}));
  return {updates:[...updates].flatMap(([sheetId,changed])=>{const source=book.sheets.find(s=>s.id===sheetId);return [...changed].filter(([key,cell])=>JSON.stringify(cell)!==JSON.stringify(source.cells[key])).map(([key,cell])=>({sheetId,key,cell}));}),metadata:snapshot.sheets.map(s=>({sheetId:s.id,...Object.fromEntries(['tables','charts','conditionalRules','xlsxFeatures'].filter(key=>s[key]!==undefined).map(key=>[key,s[key]])),...(s.id===owner.id?{rowHeights}:{})})),...(chartNames?{chartNames}:{})};
}
