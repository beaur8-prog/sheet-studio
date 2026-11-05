import {resolvePivotSpec,restorePivotTableBinding} from './pivot-source.mjs';
import JSZip from 'jszip';
import {buildPivot,applyPivot,aggregations,scalarKey} from './pivot.mjs';
import {validatePivot} from './pivot-schema.mjs';
import {address,parseAddress,transformFormulaReferences} from './engine.mjs';
import {shiftWorksheetFeatures} from './conditional.mjs';
import {updateChartReferences} from './chart-model.mjs';
import {xml} from './chart-xml.mjs';
import {nativePivotIndex,nativePivotLayout,legacyPivotMatches,preserveNativePivots,nativePivotNames,pivotHasExtraCells} from './native-pivots.mjs';

const NS='http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const REL='http://schemas.openxmlformats.org/officeDocument/2006/relationships/';
const DECL='<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
const esc=value=>String(value).replace(/[&<>"'\u0000-\u0008\u000b\u000c\u000e-\u001f]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&apos;'}[c]||''));
const unesc=value=>value.replace(/&(lt|gt|quot|apos|amp);/g,(_,v)=>({lt:'<',gt:'>',quot:'"',apos:"'",amp:'&'}[v]));
const rel=(id,type,target)=>`<Relationship Id="${id}" Type="${REL+type}" Target="${target}"/>`;
const rels=body=>`${DECL}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${body}</Relationships>`;
const META='customXml/sheetStudioPivots.xml';
const META_NS='urn:sheet-studio:pivots:1';
const typed=value=>value===''||value===null?'<m/>':typeof value==='number'?`<n v="${value}"/>`:typeof value==='boolean'?`<b v="${value?1:0}"/>`:`<s v="${esc(value)}"/>`;

// Zero-match reports have no native PivotTable to compare. Restore their app
// settings only while the static zero-result cells and placement still match.
function unchangedEmptyReport(target,entry){
  const s=entry.spec,labels=s.values.map(v=>`${aggregations[v.aggregate]}: ${v.field}`),header=s.rows.length?[...s.rows]:['행 레이블'];
  if(!s.columns.length)header.push(...labels);else if(s.totals)header.push(...labels.map(v=>'총합계 · '+v));
  const matrix=[header];if(!s.rows.length||s.totals){const row=s.rows.length?['총합계',...Array(s.rows.length-1).fill('')]:['전체'];if(!s.columns.length||s.totals)row.push(...s.values.map(v=>['sum','count','countNumbers'].includes(v.aggregate)?0:''));matrix.push(row);}
  const offset=entry.rowOffset||0,range=address(offset,0)+':'+address(offset+matrix.length-1,header.length-1);if(pivotHasExtraCells(target,entry,null,range))return false;
  return matrix.every((row,r)=>row.every((value,c)=>{const raw=target.cells[address(r+offset,c)]?.raw||'';return typeof value==='number'?raw===String(value):!raw.startsWith('=')&&raw.replace(/^'/,'')===value;}));
}

async function addRel(zip,path,body){const before=zip.file(path)?await zip.file(path).async('string'):rels('');zip.file(path,before.replace('</Relationships>',body+'</Relationships>'));}

/** App definitions are separate custom XML; native Excel PivotTable/cache parts
 * are emitted for XLSX as well. No Excel executable or macro runtime is needed. */
export async function addPivotParts(bytes,workbook,names){
  const pivots=workbook.sheets.filter(s=>s.pivot);
  if(!pivots.length)return bytes;
  const zip=await JSZip.loadAsync(bytes),definitions=[],usedNames=nativePivotNames(workbook);
  let types=await zip.file('[Content_Types].xml').async('string');
  const content=(path,type)=>{const part=`/${path}`;if(!types.includes(`PartName="${part}"`))types=types.replace('</Types>',`<Override PartName="${part}" ContentType="${type}"/></Types>`);};
  let caches='';
  for(let i=0;i<pivots.length;i++){
    const target=pivots[i],spec=resolvePivotSpec(workbook,target.pivot),sourceName=names.get(spec.sourceSheetId);
    const pages=spec.filters.filter(f=>!spec.rows.includes(f.field)&&!spec.columns.includes(f.field)).length;
    const entry={sheet:names.get(target.id),source:sourceName||null,spec,rowOffset:pages?pages+1:0};definitions.push(entry);
    const output=buildPivot(workbook,spec),{headers,records}=output.data,n=i+1;
    if(!output.matched){entry.nativeAbsent=true;continue;} // Excel cannot represent an all-items-excluded filter.
    let pivotName=`SheetStudioPivot${n}`;for(let suffix=2;usedNames.has(pivotName);suffix++)pivotName=`SheetStudioPivot${n}_${suffix}`;usedNames.add(pivotName);
    const fieldIndex=name=>headers.indexOf(name);
    const shared=headers.map((_,c)=>[...new Map(records.map(row=>[scalarKey(row[c]),row[c]])).values()]);
    const sharedFields=headers.map(name=>spec.rows.includes(name)||spec.columns.includes(name)||spec.filters.some(f=>f.field===name));
    const indexes=shared.map(values=>new Map(values.map((v,index)=>[scalarKey(v),index])));
    const cacheFields=headers.map((name,c)=>{
      const values=shared[c],number=values.some(v=>typeof v==='number'),string=values.some(v=>typeof v==='string'&&v!==''),blank=values.some(v=>v===''||v===null),other=values.some(v=>typeof v==='boolean');
      return `<cacheField name="${esc(name)}" numFmtId="0"><sharedItems count="${sharedFields[c]?values.length:0}" containsString="${+string}" containsNumber="${+number}" containsBlank="${+blank}" containsMixedTypes="${+(number&&(string||other))}" containsSemiMixedTypes="${+(string||blank||other)}">${sharedFields[c]?values.map(typed).join(''):''}</sharedItems></cacheField>`;
    }).join('');
    const definition=`${DECL}<pivotCacheDefinition xmlns="${NS}" xmlns:r="${REL.slice(0,-1)}" r:id="rIdRecords" refreshOnLoad="0" createdVersion="3" refreshedVersion="3" minRefreshableVersion="3" recordCount="${records.length}"><cacheSource type="worksheet"><worksheetSource ${spec.sourceTableName?`name="${esc(spec.sourceTableName)}"`:`ref="${esc(spec.sourceRange)}" sheet="${esc(sourceName)}"`}/></cacheSource><cacheFields count="${headers.length}">${cacheFields}</cacheFields></pivotCacheDefinition>`;
    const cachePath=`xl/pivotCache/pivotCacheDefinition${n}.xml`,recordPath=`xl/pivotCache/pivotCacheRecords${n}.xml`,tablePath=`xl/pivotTables/pivotTable${n}.xml`;
    zip.file(cachePath,definition);
    zip.file(recordPath,`${DECL}<pivotCacheRecords xmlns="${NS}" count="${records.length}">${records.map(row=>'<r>'+row.map((v,c)=>sharedFields[c]?`<x v="${indexes[c].get(scalarKey(v))}"/>`:typed(v)).join('')+'</r>').join('')}</pivotCacheRecords>`);
    zip.file(`xl/pivotCache/_rels/pivotCacheDefinition${n}.xml.rels`,rels(rel('rIdRecords','pivotCacheRecords',`pivotCacheRecords${n}.xml`)));
    zip.file(`xl/pivotTables/_rels/pivotTable${n}.xml.rels`,rels(rel('rIdCache','pivotCacheDefinition',`../pivotCache/pivotCacheDefinition${n}.xml`)));
    const pageFilters=spec.filters.filter(f=>!spec.rows.includes(f.field)&&!spec.columns.includes(f.field));
    const fields=headers.map((name,c)=>{
      const axis=spec.rows.includes(name)?'axisRow':spec.columns.includes(name)?'axisCol':pageFilters.some(f=>f.field===name)?'axisPage':'';
      const filters=spec.filters.filter(f=>f.field===name),isData=spec.values.some(v=>v.field===name);
      const items=axis?`<items count="${shared[c].length}">${shared[c].map((v,j)=>`<item x="${j}"${filters.some(f=>!f.values.some(w=>scalarKey(w)===scalarKey(v)))?' h="1"':''}/>`).join('')}</items>`:'';
      return `<pivotField${axis?` axis="${axis}"`:''}${isData?' dataField="1"':''} showAll="0" defaultSubtotal="0" compact="0" outline="0"${filters.length?' multipleItemSelectionAllowed="1"':''} sortType="${spec.sort==='source'?'manual':spec.sort==='desc'?'descending':'ascending'}">${items}</pivotField>`;
    }).join('');
    const rows=spec.rows.map(fieldIndex),cols=spec.columns.map(fieldIndex);if(spec.values.length>1)cols.push(-2);
    const axis=(name,fields)=>fields.length?`<${name} count="${fields.length}">${fields.map(x=>`<field x="${x}"/>`).join('')}</${name}>`:'';
    const axisItems=(tag,groups,axisFields,isColumn=false)=>{
      if(!axisFields.length)return '';
      const items=[];
      for(const [,values] of groups){for(let m=0;m<(isColumn&&spec.values.length>1?spec.values.length:1);m++)items.push('<i'+(isColumn?' i="'+m+'"':'')+'>'+values.map((v,k)=>'<x v="'+indexes[axisFields[k]].get(scalarKey(v))+'"/>').join('')+(isColumn&&spec.values.length>1?'<x v="'+m+'"/>':'')+'</i>');}
      if(spec.totals&&(isColumn?spec.columns.length:spec.rows.length))for(let m=0;m<(isColumn?spec.values.length:1);m++)items.push('<i t="grand"'+(isColumn?' i="'+m+'"':'')+'><x/></i>');
      return '<'+tag+' count="'+items.length+'">'+items.join('')+'</'+tag+'>';
    };
    const dataFields=spec.values.map((v,j)=>`<dataField name="${esc(`${aggregations[v.aggregate]}: ${v.field} (${j+1})`)}" fld="${fieldIndex(v.field)}" subtotal="${v.aggregate==='countNumbers'?'countNums':v.aggregate}" baseField="0" baseItem="0"/>`).join('');
    zip.file(tablePath,`${DECL}<pivotTableDefinition xmlns="${NS}" name="${esc(pivotName)}" cacheId="${n}" dataCaption="값" rowGrandTotals="${+spec.totals}" colGrandTotals="${+spec.totals}" compact="0" compactData="0" outline="0" outlineData="0" gridDropZones="0" createdVersion="3" updatedVersion="3" minRefreshableVersion="3"><location ref="A${pageFilters.length?pageFilters.length+2:1}:${address(Math.max(1,output.matrix.length-1)+(pageFilters.length?pageFilters.length+1:0),Math.max(1,output.matrix[0].length-1))}" firstHeaderRow="1" firstDataRow="1" firstDataCol="${Math.max(1,rows.length)}"${pageFilters.length?` rowPageCount="${pageFilters.length}" colPageCount="1"`:''}/><pivotFields count="${headers.length}">${fields}</pivotFields>${axis('rowFields',rows)}${axisItems('rowItems',output.rows,rows)}${axis('colFields',cols)}${axisItems('colItems',output.columns,cols,true)}${pageFilters.length?`<pageFields count="${pageFilters.length}">${pageFilters.map(f=>`<pageField fld="${fieldIndex(f.field)}" hier="-1"/>`).join('')}</pageFields>`:''}<dataFields count="${spec.values.length}">${dataFields}</dataFields><pivotTableStyleInfo name="PivotStyleMedium9" showRowHeaders="1" showColHeaders="1" showRowStripes="1" showColStripes="0" showLastColumn="1"/></pivotTableDefinition>`);
    const sheetNumber=workbook.sheets.indexOf(target)+1;
    await addRel(zip,`xl/worksheets/_rels/sheet${sheetNumber}.xml.rels`,rel(`rIdStudioPivot${n}`,'pivotTable',`../pivotTables/pivotTable${n}.xml`));
    await addRel(zip,'xl/_rels/workbook.xml.rels',rel(`rIdStudioCache${n}`,'pivotCacheDefinition',`pivotCache/pivotCacheDefinition${n}.xml`));
    caches+=`<pivotCache cacheId="${n}" r:id="rIdStudioCache${n}"/>`;
    content(cachePath,'application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml');content(recordPath,'application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheRecords+xml');content(tablePath,'application/vnd.openxmlformats-officedocument.spreadsheetml.pivotTable+xml');
    entry.nativeLayout=nativePivotLayout({table:xml(await zip.file(tablePath).async('string')).documentElement,definition:xml(definition).documentElement});
  }
  if(caches){const xml=await zip.file('xl/workbook.xml').async('string');zip.file('xl/workbook.xml',xml.replace('</workbook>',`<pivotCaches>${caches}</pivotCaches></workbook>`));}
  zip.file(META,`${DECL}<pivots xmlns="${META_NS}">${esc(JSON.stringify(definitions))}</pivots>`);
  const rootRels=await zip.file('xl/_rels/workbook.xml.rels').async('string');
  if(!rootRels.includes('customXml/sheetStudioPivots.xml'))await addRel(zip,'xl/_rels/workbook.xml.rels',rel('rIdStudioPivots','customXml','../'+META));
  content(META,'application/xml');zip.file('[Content_Types].xml',types);
  return zip.generateAsync({type:'uint8array',compression:'DEFLATE'});
}

export async function restorePivotParts(bytes,workbook,{refresh=true}={}){
  const zip=await JSZip.loadAsync(bytes),native=await nativePivotIndex(zip,workbook),retained=new Set(native);
  let sourceXML='';
  // Excel renames custom XML parts when saving. Identify our namespace, not its filename.
  for(const part of zip.file(/^customXml\/[^/]+\.xml$/).slice(0,100)){const text=await part.async('string');if(text.includes(META_NS)){sourceXML=text;break;}}
  if(!sourceXML){await preserveNativePivots(zip,workbook,native);return;}
  if(sourceXML.length>2_000_000)throw new Error('피벗 설정의 크기 제한을 초과했습니다.');
  const match=sourceXML.match(/<pivots\b[^>]*>([\s\S]*?)<\/pivots>/);if(!match||!sourceXML.includes(META_NS))throw new Error('피벗 설정을 읽을 수 없습니다.');
  const definitions=JSON.parse(unesc(match[1]),(key,value)=>{if(['__proto__','prototype','constructor'].includes(key))throw new Error('피벗 설정 키가 올바르지 않습니다.');return value;});
  if(!Array.isArray(definitions)||definitions.length>100)throw new Error('피벗 설정 개수가 올바르지 않습니다.');
  const restored=[];
  for(const entry of definitions){
    const target=workbook.sheets.find(s=>s.name===entry.sheet),source=workbook.sheets.find(s=>s.name===entry.source);
    if(!target)continue;
    validatePivot(entry.spec);if(target===source)throw new Error('피벗 원본은 다른 시트여야 합니다.');
    if(entry.rowOffset!==undefined&&(!Number.isInteger(entry.rowOffset)||entry.rowOffset<0||entry.rowOffset>17))throw new Error('피벗 위치 정보가 올바르지 않습니다.');
    const actual=native.filter(n=>n.sheetId===target.id),empty=entry.nativeAbsent===true||!actual.length&&entry.spec.filters.some(f=>!f.values.length);
    const compatible=empty&&!actual.length||actual.length===1&&(typeof entry.nativeLayout==='string'?entry.nativeLayout===nativePivotLayout(actual[0]):legacyPivotMatches(actual[0],entry));
    if(empty&&!actual.length&&!unchangedEmptyReport(target,entry)){target.pivotImportNotice='외부에서 바꾼 정적 피벗 결과와 추가 셀을 유지했습니다. 필요하면 원본 자료에서 새 피벗을 만드세요.';continue;}
    const extras=actual.length===1&&pivotHasExtraCells(target,entry,actual[0]);
    if(!compatible||extras){target.pivotImportNotice=extras?'피벗 바깥에 추가된 셀과 원본 피벗 결과를 유지합니다. 새로 고침은 Excel에서 진행하세요.':'외부에서 변경된 피벗 설정과 결과를 유지합니다. 필드 편집·새로 고침은 Excel에서 진행하세요.';continue;}
    for(const item of actual)retained.delete(item);
    target.pivot={...entry.spec,sourceSheetId:source?.id||'missing-source'};
    if(target.pivot.sourceTableId){try{target.pivot=restorePivotTableBinding(workbook,target.pivot);}catch(error){target.pivotError=error.message;}}
    if(refresh&&entry.rowOffset){
      const offset=entry.rowOffset,cells={};
      for(const ref in target.cells){const p=parseAddress(ref);if(p.r>=offset)cells[address(p.r-offset,p.c)]=target.cells[ref];}
      target.cells=cells;target.rowHeights=Object.fromEntries(Object.entries(target.rowHeights||{}).filter(([r])=>+r>=offset).map(([r,h])=>[+r-offset,h]));
      for(let i=0;i<offset;i++){
        for(const owner of workbook.sheets){
          for(const cell of Object.values(owner.cells))if(cell.raw?.startsWith('='))cell.raw=transformFormulaReferences(cell.raw,{axis:'r',index:0,delta:-1,targetSheet:target.name,ownerSheet:owner.name});
          shiftWorksheetFeatures(owner,'r',0,-1,target.name);
        }
        updateChartReferences(workbook,target.id,'r',0,-1);
      }
    }
    restored.push(target);
  }
  await preserveNativePivots(zip,workbook,[...retained]);
  if(refresh)for(const target of restored)try{applyPivot(workbook,target,target.pivot);}catch(error){target.pivotError=error.message;}
}
