import {cloneData} from './clone.mjs';
import {address,parseAddress,transformFormulaReferences,validateWorkbook} from './engine.mjs';
import {assertStructureAllowed} from './cell-constraints.mjs';
import {assertTableStructureAllowed,removeTablesForStructure,shiftTables} from './table-structure.mjs';
import {structureIndex,shiftIndexMap} from './structure-coordinates.mjs';
import {shiftAutoFilter} from './filters.mjs';
import {shiftWorksheetFeatures} from './conditional.mjs';
import {updateChartReferences} from './chart-model.mjs';
import {shiftPivotSources} from './pivot.mjs';
import {recordNativeStructure} from './native-plan.mjs';

/** Build a complete, validated replacement without mutating the current book.
 * Cells and A1 formulas are each visited once, independent of the row count.
 */
export function structuralWorkbookPlan(book,sheetId,{axis,index,count=1,delta}){
  const original=book.sheets.find(s=>s.id===sheetId);if(!original)throw Error('시트를 찾을 수 없습니다.');
  if(original.pivot)throw Error('피벗 결과에서는 행·열을 변경할 수 없습니다.');
  assertStructureAllowed(original,axis,index,count,delta);
  const removed=assertTableStructureAllowed(book,original,axis,index,count,delta),amount=count*delta,limit=axis==='r'?1048576:16384;
  if(book.nativeSource){const baseline=book.nativeSource.baseline.find(s=>s.id===sheetId);if(baseline)assertStructureAllowed(baseline,axis,index,count,delta);}
  const next=cloneData(book),sheet=next.sheets.find(s=>s.id===sheetId);removeTablesForStructure(next,sheet,removed);
  recordNativeStructure(next,sheetId,axis,index,amount);shiftPivotSources(next,sheetId,axis,index,amount);
  const cells={};for(const ref in sheet.cells){const p=parseAddress(ref),n=structureIndex(p[axis],index,amount);if(n===null)continue;p[axis]=n;if(n<0||n>=limit)throw Error('시트 범위를 벗어나는 셀이 있습니다.');cells[address(p.r,p.c)]=sheet.cells[ref];}sheet.cells=cells;
  const transform=(formula,ownerSheet)=>transformFormulaReferences(formula,{axis,index,delta:amount,targetSheet:sheet.name,ownerSheet});
  for(const owner of next.sheets)for(const ref in owner.cells){const cell=owner.cells[ref];if(cell.raw?.startsWith('='))cell.raw=transform(cell.raw,owner.name);}
  sheet.merges=(sheet.merges||[]).map(ref=>transform(ref,sheet.name)).filter(ref=>!ref.includes('#REF!')&&ref.split(':')[0]!==ref.split(':')[1]);
  updateChartReferences(next,sheetId,axis,index,amount);sheet.charts=(sheet.charts||[]).filter(chart=>chart.source||!chart.range?.includes('#REF!'));
  const prop=axis==='r'?'rowHeights':'columnWidths';sheet[prop]=shiftIndexMap(sheet[prop],index,amount,limit);
  shiftAutoFilter(sheet,axis,index,amount);shiftTables(next,sheetId,axis,index,amount);
  for(const owner of next.sheets)shiftWorksheetFeatures(owner,axis,index,amount,sheet.name);
  next.referenceCachesValid=false;
  return validateWorkbook(next);
}
