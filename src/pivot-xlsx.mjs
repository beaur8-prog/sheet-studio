import JSZip from 'jszip';
import {buildPivot,applyPivot,aggregations,scalarKey} from './pivot.mjs';
import {validatePivot} from './pivot-schema.mjs';
import {address,transformFormulaReferences} from './engine.mjs';
import {updateChartReferences} from './chart-model.mjs';

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

async function addRel(zip,path,body){const before=zip.file(path)?await zip.file(path).async('string'):rels('');zip.file(path,before.replace('</Relationships>',body+'</Relationships>'));}

/** App definitions are separate custom XML; native Excel PivotTable/cache parts
 * are emitted for XLSX as well. No Excel executable or macro runtime is needed. */
export async function addPivotParts(bytes,workbook,names){
  const pivots=workbook.sheets.filter(s=>s.pivot);
  if(!pivots.length)return bytes;
  const zip=await JSZip.loadAsync(bytes),definitions=[];
  let types=await zip.file('[Content_Types].xml').async('string');
  const content=(path,type)=>{const part=`/${path}`;if(!types.includes(`PartName="${part}"`))types=types.replace('</Types>',`<Override PartName="${part}" ContentType="${type}"/></Types>`);};
  let caches='';
  for(let i=0;i<pivots.length;i++){
    const target=pivots[i],spec=target.pivot,sourceName=names.get(spec.sourceSheetId);
    const pages=spec.filters.filter(f=>!spec.rows.includes(f.field)&&!spec.columns.includes(f.field)).length;
    definitions.push({sheet:names.get(target.id),source:sourceName||null,spec,rowOffset:pages?pages+1:0});
    const output=buildPivot(workbook,spec),{headers,records}=output.data,n=i+1;
    if(!output.matched)continue; // Excel cannot represent an all-items-excluded filter; retain the empty result and app settings.
    const fieldIndex=name=>headers.indexOf(name);
    const shared=headers.map((_,c)=>[...new Map(records.map(row=>[scalarKey(row[c]),row[c]])).values()]);
    const sharedFields=headers.map(name=>spec.rows.includes(name)||spec.columns.includes(name)||spec.filters.some(f=>f.field===name));
    const indexes=shared.map(values=>new Map(values.map((v,index)=>[scalarKey(v),index])));
    const cacheFields=headers.map((name,c)=>{
      const values=shared[c],number=values.some(v=>typeof v==='number'),string=values.some(v=>typeof v==='string'&&v!==''),blank=values.some(v=>v===''||v===null),other=values.some(v=>typeof v==='boolean');
      return `<cacheField name="${esc(name)}" numFmtId="0"><sharedItems count="${sharedFields[c]?values.length:0}" containsString="${+string}" containsNumber="${+number}" containsBlank="${+blank}" containsMixedTypes="${+(number&&(string||other))}" containsSemiMixedTypes="${+(string||blank||other)}">${sharedFields[c]?values.map(typed).join(''):''}</sharedItems></cacheField>`;
    }).join('');
    const definition=`${DECL}<pivotCacheDefinition xmlns="${NS}" xmlns:r="${REL.slice(0,-1)}" r:id="rIdRecords" refreshOnLoad="0" createdVersion="3" refreshedVersion="3" minRefreshableVersion="3" recordCount="${records.length}"><cacheSource type="worksheet"><worksheetSource ref="${esc(spec.sourceRange)}" sheet="${esc(sourceName)}"/></cacheSource><cacheFields count="${headers.length}">${cacheFields}</cacheFields></pivotCacheDefinition>`;
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
    zip.file(tablePath,`${DECL}<pivotTableDefinition xmlns="${NS}" name="SheetStudioPivot${n}" cacheId="${n}" dataCaption="값" rowGrandTotals="${+spec.totals}" colGrandTotals="${+spec.totals}" compact="0" compactData="0" outline="0" outlineData="0" gridDropZones="0" createdVersion="3" updatedVersion="3" minRefreshableVersion="3"><location ref="A${pageFilters.length?pageFilters.length+2:1}:${address(Math.max(1,output.matrix.length-1)+(pageFilters.length?pageFilters.length+1:0),Math.max(1,output.matrix[0].length-1))}" firstHeaderRow="1" firstDataRow="1" firstDataCol="${Math.max(1,rows.length)}"${pageFilters.length?` rowPageCount="${pageFilters.length}" colPageCount="1"`:''}/><pivotFields count="${headers.length}">${fields}</pivotFields>${axis('rowFields',rows)}${axisItems('rowItems',output.rows,rows)}${axis('colFields',cols)}${axisItems('colItems',output.columns,cols,true)}${pageFilters.length?`<pageFields count="${pageFilters.length}">${pageFilters.map(f=>`<pageField fld="${fieldIndex(f.field)}" hier="-1"/>`).join('')}</pageFields>`:''}<dataFields count="${spec.values.length}">${dataFields}</dataFields><pivotTableStyleInfo name="PivotStyleMedium9" showRowHeaders="1" showColHeaders="1" showRowStripes="1" showColStripes="0" showLastColumn="1"/></pivotTableDefinition>`);
    const sheetNumber=workbook.sheets.indexOf(target)+1;
    await addRel(zip,`xl/worksheets/_rels/sheet${sheetNumber}.xml.rels`,rel(`rIdStudioPivot${n}`,'pivotTable',`../pivotTables/pivotTable${n}.xml`));
    await addRel(zip,'xl/_rels/workbook.xml.rels',rel(`rIdStudioCache${n}`,'pivotCacheDefinition',`pivotCache/pivotCacheDefinition${n}.xml`));
    caches+=`<pivotCache cacheId="${n}" r:id="rIdStudioCache${n}"/>`;
    content(cachePath,'application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml');content(recordPath,'application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheRecords+xml');content(tablePath,'application/vnd.openxmlformats-officedocument.spreadsheetml.pivotTable+xml');
  }
  if(caches){const xml=await zip.file('xl/workbook.xml').async('string');zip.file('xl/workbook.xml',xml.replace('</workbook>',`<pivotCaches>${caches}</pivotCaches></workbook>`));}
  zip.file(META,`${DECL}<pivots xmlns="${META_NS}">${esc(JSON.stringify(definitions))}</pivots>`);
  const rootRels=await zip.file('xl/_rels/workbook.xml.rels').async('string');
  if(!rootRels.includes('customXml/sheetStudioPivots.xml'))await addRel(zip,'xl/_rels/workbook.xml.rels',rel('rIdStudioPivots','customXml','../'+META));
  content(META,'application/xml');zip.file('[Content_Types].xml',types);
  return zip.generateAsync({type:'uint8array',compression:'DEFLATE'});
}

export async function restorePivotParts(bytes,workbook,{refresh=true}={}){
  const zip=await JSZip.loadAsync(bytes);
  let xml='';
  // Excel renames custom XML parts when saving. Identify our namespace, not its filename.
  for(const part of zip.file(/^customXml\/[^/]+\.xml$/).slice(0,100)){const text=await part.async('string');if(text.includes(META_NS)){xml=text;break;}}
  if(!xml)return;
  if(xml.length>2_000_000)throw new Error('피벗 설정의 크기 제한을 초과했습니다.');
  const match=xml.match(/<pivots\b[^>]*>([\s\S]*?)<\/pivots>/);if(!match||!xml.includes(META_NS))throw new Error('피벗 설정을 읽을 수 없습니다.');
  const definitions=JSON.parse(unesc(match[1]),(key,value)=>{if(['__proto__','prototype','constructor'].includes(key))throw new Error('피벗 설정 키가 올바르지 않습니다.');return value;});
  if(!Array.isArray(definitions)||definitions.length>100)throw new Error('피벗 설정 개수가 올바르지 않습니다.');
  const restored=[];
  for(const entry of definitions){
    const target=workbook.sheets.find(s=>s.name===entry.sheet),source=workbook.sheets.find(s=>s.name===entry.source);
    if(!target)continue;
    validatePivot(entry.spec);if(target===source)throw new Error('피벗 원본은 다른 시트여야 합니다.');
    target.pivot={...entry.spec,sourceSheetId:source?.id||'missing-source'};
    if(entry.rowOffset!==undefined&&(!Number.isInteger(entry.rowOffset)||entry.rowOffset<0||entry.rowOffset>17))throw new Error('피벗 위치 정보가 올바르지 않습니다.');
    if(refresh&&entry.rowOffset){for(const owner of workbook.sheets)for(const cell of Object.values(owner.cells))if(cell.raw?.startsWith('='))for(let i=0;i<entry.rowOffset;i++)cell.raw=transformFormulaReferences(cell.raw,{axis:'r',index:0,delta:-1,targetSheet:target.name,ownerSheet:owner.name});for(let i=0;i<entry.rowOffset;i++)updateChartReferences(workbook,target.id,'r',0,-1);}
    restored.push(target);
  }
  if(refresh)for(const target of restored)try{applyPivot(workbook,target,target.pivot);}catch(error){target.pivotError=error.message;}
}
