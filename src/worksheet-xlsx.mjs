import JSZip from 'jszip';
import {xml,stringify,all,attr,child,readRelationships,resolvePart} from './chart-xml.mjs';
async function parts(zip){
  const wb=xml(await zip.file('xl/workbook.xml').async('string')),rels=await readRelationships(zip,'xl/workbook.xml');
  return all(wb,'sheet').map(n=>{const rel=rels.find(r=>r.id===attr(n,'r:id'));return {name:attr(n,'name'),path:rel&&rel.mode!=='External'?resolvePart('xl/workbook.xml',rel.target):null};}).filter(p=>p.path);
}
// ExcelJS omits empty hidden rows and cfRule.stopIfTrue. Restore these from OOXML.
export async function restoreWorksheetParts(bytes,book){
  const zip=await JSZip.loadAsync(bytes);
  const workbook=xml(await zip.file('xl/workbook.xml').async('string')),dateOffset=['1','true'].includes(attr(all(workbook,'workbookPr')[0],'date1904'))?1462:0;
  for(const p of await parts(zip)){
    const s=book.sheets.find(s=>s.name===p.name);if(!s)continue;
    const doc=xml(await zip.file(p.path).async('string'));
    for(const cell of all(doc,'c'))if((!attr(cell,'t')||attr(cell,'t')==='n')&&s.cells[attr(cell,'r')]?.style?.format==='date'){const v=Number(child(cell,'v')?.textContent);if(Number.isFinite(v)&&!child(cell,'f'))s.cells[attr(cell,'r')].raw=String(v+dateOffset);}
    for(const cell of all(doc,'c'))if(attr(cell,'t')==='d'){const date=new Date(child(cell,'v')?.textContent),time=date.getTime();if(Number.isFinite(time)){const key=attr(cell,'r'),serial=time/86400000+25569-(time<Date.UTC(1900,2,1)?1:0);s.cells[key]={...s.cells[key],raw:String(Math.round(serial*1e8)/1e8),style:{...s.cells[key]?.style,format:'date'}};}}
    for(const row of all(doc,'row'))if(attr(row,'hidden')==='1'){
      const r=Number(attr(row,'r'))-1;if(!Number.isInteger(r)||r<0||r>=1048576)continue;
      const h=Number(attr(row,'ht'));s.hiddenSizes||={rows:{},cols:{}};s.hiddenSizes.rows[r]=h?h*4/3:s.defaultRowHeight||25;s.rowHeights[r]=0;
    }
    const conditions=new Map(all(doc,'cfRule').map(n=>[Number(attr(n,'priority')),n]));
    for(const rule of s.conditionalRules||[]){rule.stopIfTrue=attr(conditions.get(rule.sourcePriority),'stopIfTrue')==='1';delete rule.sourcePriority;}
  }
}
export async function addWorksheetParts(bytes,book,names){
  if(!book.sheets.some(s=>s.conditionalRules?.length))return bytes;
  const zip=await JSZip.loadAsync(bytes);
  for(const p of await parts(zip)){
    const s=book.sheets.find(s=>(names?.get(s.id)||s.name)===p.name);if(!s?.conditionalRules?.length)continue;
    const doc=xml(await zip.file(p.path).async('string'));
    for(const node of all(doc,'cfRule')){const rule=s.conditionalRules[Number(attr(node,'priority'))-1];if(rule)node.setAttribute('stopIfTrue',rule.stopIfTrue?'1':'0');}
    zip.file(p.path,stringify(doc));
  }
  return zip.generateAsync({type:'uint8array',compression:'DEFLATE'});
}
