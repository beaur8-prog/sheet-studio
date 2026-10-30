import JSZip from 'jszip';
import {address,formulaForXlsx} from './engine.mjs';
import {tableBounds} from './table-schema.mjs';
import {readAutoFilter,autoFilterXml,createFilterColorWriter} from './filter-xlsx.mjs';
import {applyFilters} from './filters.mjs';
import {createColorResolver} from './spreadsheet-colors.mjs';
import {worksheetParts} from './xlsx-preservation.mjs';
import {worksheetMetadataXML,splitWorksheetXML} from './worksheet-metadata.mjs';
import {capturePartTree,writeStoredParts,contentTypeWriter} from './xlsx-part-store.mjs';
import {mergePreservedDxfs,remapDxfIds} from './xlsx-dxfs.mjs';
import {xml,stringify,child,children,all,attr,readRelationships,resolvePart,relPath,relationship,relationships,R} from './chart-xml.mjs';
const NS='http://schemas.openxmlformats.org/spreadsheetml/2006/main',TABLE_TYPE='application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml';
const bool=(node,key,fallback=false)=>node?.hasAttribute(key)?['1','true'].includes(attr(node,key)):fallback;
export async function restoreTables(zip,book){
  const stylesText=await zip.file('xl/styles.xml')?.async('string'),styles=stylesText?xml(stylesText):null,dxfs=children(child(styles?.documentElement,'dxfs'),'dxf'),resolve=createColorResolver(await zip.file('xl/theme/theme1.xml')?.async('string'),stylesText);
  for(const part of await worksheetParts(zip)){
    const sheet=book.sheets.find(s=>s.name===part.name);if(!sheet||!part.type.endsWith('/worksheet'))continue;
    const doc=xml(worksheetMetadataXML(await zip.file(part.path).async('string'))),nodes=all(doc,'tablePart');if(!nodes.length)continue;
    const rels=await readRelationships(zip,part.path),manualNode=all(doc,'manualHiddenRows')[0];if(manualNode)sheet.tableManualHiddenRows=attr(manualNode,'rows').split(',').filter(Boolean).map(Number);
    for(const node of nodes){
      const rel=rels.find(r=>r.id===attr(node,'r:id'));if(!rel||rel.mode==='External')continue;
      const path=resolvePart(part.path,rel.target),sourceXML=await zip.file(path)?.async('string');if(!sourceXML)throw Error('표 정의 파일이 없습니다.');
      const source=xml(sourceXML).documentElement,style=child(source,'tableStyleInfo'),table={id:crypto.randomUUID(),name:attr(source,'name')||attr(source,'displayName'),range:attr(source,'ref').replace(/\$/g,''),headerRow:attr(source,'headerRowCount','1')!=='0',totalsRow:Number(attr(source,'totalsRowCount','0'))>0,sourceXML,columns:children(child(source,'tableColumns'),'tableColumn').map(column=>({id:Number(attr(column,'id')),name:attr(column,'name'),...(child(column,'calculatedColumnFormula')?{formula:child(column,'calculatedColumnFormula').textContent}:{}),...(child(column,'totalsRowFormula')?{totalsFormula:child(column,'totalsRowFormula').textContent}:{}),...(attr(column,'totalsRowFunction')?{totalsFunction:attr(column,'totalsRowFunction')}:{}),...(attr(column,'totalsRowLabel')?{totalsLabel:attr(column,'totalsRowLabel')}:{} )}))};
      if(!tableBounds(table.range))throw Error('표의 범위가 올바르지 않습니다.');
      if(style)table.style={name:attr(style,'name'),...Object.fromEntries(['showFirstColumn','showLastColumn','showRowStripes','showColumnStripes'].map(key=>[key,bool(style,key)]))};
      const filter=child(source,'autoFilter');if(filter){table.filter=readAutoFilter(filter,sheet,book,dxfs,resolve);table.columns.forEach((column,i)=>{const node=children(filter,'filterColumn').find(n=>Number(attr(n,'colId'))===i);column.filterButton=!node||!bool(node,'hiddenButton');});}
      const tableRels=await readRelationships(zip,path);if(tableRels.length){table.sourceRels=tableRels.map(r=>({...r,target:r.mode==='External'?r.target:resolvePart(path,r.target)}));table.sourceParts=await capturePartTree(zip,table.sourceRels.filter(r=>r.mode!=='External').map(r=>r.target));}
      sheet.tables||=[];sheet.tables.push(table);
    }
  }
  if(book.sheets.some(s=>s.tables?.length)){
    const custom=children(child(styles?.documentElement,'tableStyles'),'tableStyle');if(custom.length)book.xlsxTableStyles=custom.map(stringify);
    if(dxfs.length)book.xlsxDxfs=dxfs.map(stringify);
  }
  const connection=(await readRelationships(zip,'xl/workbook.xml')).find(r=>r.type.endsWith('/connections')&&r.mode!=='External');
  if(connection){const root=resolvePart('xl/workbook.xml',connection.target);book.xlsxConnections={root,type:connection.type,parts:await capturePartTree(zip,[root])};}
}
export function restoreTableFilterVisibility(book){
  for(const sheet of book.sheets){const manual=new Set(sheet.tableManualHiddenRows||[]);delete sheet.tableManualHiddenRows;
    for(const table of sheet.tables||[]){const filter=table.filter;if(!filter)continue;
      const original=Object.entries(sheet.rowHeights||{}).filter(([r,h])=>h===0&&+r>filter.range.r1&&+r<=filter.range.r2).map(([r])=>+r);
      if(filter.unsupported?.length)filter.hiddenRows=original.filter(r=>!manual.has(r));
      else applyFilters(book,{...sheet,autoFilter:filter});
      for(const r of filter.hiddenRows||[])if(!manual.has(r)){const old=sheet.hiddenSizes?.rows?.[r];if(old&&old!==(sheet.defaultRowHeight||25))sheet.rowHeights[r]=old;else delete sheet.rowHeights[r];if(sheet.hiddenSizes?.rows)delete sheet.hiddenSizes.rows[r];}
    }
  }
}
const strip=path=>path.replace(/^(?:xl\/sheetStudioTables\/(?:table\d+|aux)\/)+/,'');
export async function addTableParts(bytes,book,names){
  if(!book.sheets.some(s=>s.tables?.length)&&!book.xlsxConnections)return bytes;
  const zip=await JSZip.loadAsync(bytes),types=xml(await zip.file('[Content_Types].xml').async('string')),content=contentTypeWriter(types),dxfMap=await mergePreservedDxfs(zip,book.xlsxDxfs),styles=xml(await zip.file('xl/styles.xml').async('string')),colorId=createFilterColorWriter(styles);
  if(book.xlsxTableStyles?.length){let container=child(styles.documentElement,'tableStyles');if(!container){container=styles.createElementNS(NS,'tableStyles');styles.documentElement.appendChild(container);}const existing=new Set(children(container,'tableStyle').map(n=>attr(n,'name').toLocaleLowerCase()));for(const source of book.xlsxTableStyles){const node=xml(source).documentElement;if(existing.has(attr(node,'name').toLocaleLowerCase()))continue;remapDxfIds(node,dxfMap);container.appendChild(styles.importNode(node,true));existing.add(attr(node,'name').toLocaleLowerCase());}container.setAttribute('count',String(children(container,'tableStyle').length));}
  let number=0;
  for(const part of await worksheetParts(zip)){
    const sheet=book.sheets.find(s=>(names?.get(s.id)||s.name)===part.name);if(!sheet?.tables?.length)continue;
    const split=splitWorksheetXML(await zip.file(part.path).async('string')),doc=xml(split.metadata),root=doc.documentElement;let tableParts=child(root,'tableParts');if(tableParts)root.removeChild(tableParts);tableParts=doc.createElementNS(NS,'tableParts');
    const sheetRels=zip.file(relPath(part.path))?xml(await zip.file(relPath(part.path)).async('string')):xml(relationships(''));
    for(const table of sheet.tables){
      const n=++number,path=`xl/tables/table${n}.xml`,id='rIdSheetStudioTable'+n,tableDoc=table.sourceXML?xml(table.sourceXML):xml(`<table xmlns="${NS}"/>`),node=tableDoc.documentElement;
      remapDxfIds(node,dxfMap);node.setAttribute('id',String(n));node.setAttribute('name',table.name);node.setAttribute('displayName',table.name);node.setAttribute('ref',table.range);node.setAttribute('headerRowCount',table.headerRow?'1':'0');node.setAttribute('totalsRowCount',table.totalsRow?'1':'0');node.setAttribute('totalsRowShown',table.totalsRow?'1':'0');
      let columns=child(node,'tableColumns');const oldColumns=children(columns,'tableColumn');if(!columns){columns=tableDoc.createElementNS(NS,'tableColumns');node.appendChild(columns);}while(columns.firstChild)columns.removeChild(columns.firstChild);
      table.columns.forEach((column,index)=>{const old=oldColumns.find(c=>Number(attr(c,'id'))===column.id),target=old?old.cloneNode(true):tableDoc.createElementNS(NS,'tableColumn');target.setAttribute('id',String(column.id||index+1));target.setAttribute('name',column.name);
        for(const [property,tag] of [['formula','calculatedColumnFormula'],['totalsFormula','totalsRowFormula']]){let field=child(target,tag);if(column[property]!==undefined){if(!field){field=tableDoc.createElementNS(NS,tag);target.appendChild(field);}field.textContent=formulaForXlsx('='+column[property],field.textContent,table).slice(1);}else if(field)target.removeChild(field);}
        for(const [property,attribute] of [['totalsFunction','totalsRowFunction'],['totalsLabel','totalsRowLabel']])if(column[property]!==undefined)target.setAttribute(attribute,column[property]);else target.removeAttribute(attribute);columns.appendChild(target);
      });columns.setAttribute('count',String(table.columns.length));
      const originalFilter=child(node,'autoFilter');if(table.filter?.unsupported?.length){if(!originalFilter)throw Error('표의 미지원 필터 원본이 없습니다.');}
      else{if(originalFilter)node.removeChild(originalFilter);if(table.filter){const filter=xml(`<table xmlns="${NS}">${autoFilterXml({...sheet,autoFilter:table.filter},colorId,book)}</table>`).documentElement.firstChild;table.columns.forEach((column,i)=>{if(column.filterButton!==false)return;let field=children(filter,'filterColumn').find(n=>Number(attr(n,'colId'))===i);if(!field){field=filter.ownerDocument.createElementNS(NS,'filterColumn');field.setAttribute('colId',String(i));filter.appendChild(field);}field.setAttribute('hiddenButton','1');});node.insertBefore(tableDoc.importNode(filter,true),columns);}}
      if(table.style){let style=child(node,'tableStyleInfo');if(!style){style=tableDoc.createElementNS(NS,'tableStyleInfo');node.insertBefore(style,child(node,'extLst')||null);}if(table.style.name)style.setAttribute('name',table.style.name);else style.removeAttribute('name');for(const key of ['showFirstColumn','showLastColumn','showRowStripes','showColumnStripes'])if(table.style[key]!==undefined)style.setAttribute(key,table.style[key]?'1':'0');}
      zip.file(path,stringify(tableDoc));content(path,TABLE_TYPE);
      if(table.sourceParts){const map=p=>`xl/sheetStudioTables/table${n}/`+strip(p);await writeStoredParts(zip,table.sourceParts,map,content);if(table.sourceRels?.length)zip.file(relPath(path),relationships(table.sourceRels.map(r=>relationship(r.id,r.type,r.mode==='External'?r.target:'/'+map(r.target),r.mode)).join('')));}
      const relNode=xml(relationships(relationship(id,R+'/table','../tables/table'+n+'.xml'))).documentElement.firstChild;sheetRels.documentElement.appendChild(sheetRels.importNode(relNode,true));const tablePart=doc.createElementNS(NS,'tablePart');tablePart.setAttributeNS(R,'r:id',id);tableParts.appendChild(tablePart);
    }
    tableParts.setAttribute('count',String(sheet.tables.length));root.insertBefore(tableParts,child(root,'extLst')||null);zip.file(relPath(part.path),stringify(sheetRels));
    let extLst=child(root,'extLst');if(!extLst){extLst=doc.createElementNS(NS,'extLst');root.appendChild(extLst);}let manual=all(extLst,'manualHiddenRows')[0];if(!manual){const ext=doc.createElementNS(NS,'ext');ext.setAttribute('uri','{84840430-696E-4AD8-A772-03A637580A33}');manual=doc.createElementNS('urn:sheet-studio:filter','ss:manualHiddenRows');ext.appendChild(manual);extLst.appendChild(ext);}manual.setAttribute('rows',Object.entries(sheet.rowHeights||{}).filter(([,h])=>h===0).map(([r])=>r).join(','));
    let properties=child(root,'sheetPr');if(!properties){properties=doc.createElementNS(NS,'sheetPr');root.insertBefore(properties,root.firstChild);}properties.setAttribute('filterMode',sheet.tables.some(t=>Object.keys(t.filter?.columns||{}).length||t.filter?.unsupported?.length)||Object.keys(sheet.autoFilter?.columns||{}).length?'1':'0');zip.file(part.path,split.restore(stringify(doc)));
  }
  if(book.xlsxConnections){const source=book.xlsxConnections,map=p=>'xl/sheetStudioTables/aux/'+strip(p);await writeStoredParts(zip,source.parts,map,content);const rels=xml(await zip.file('xl/_rels/workbook.xml.rels').async('string')),node=xml(relationships(relationship('rIdSheetStudioConnections',source.type,'/'+map(source.root)))).documentElement.firstChild;rels.documentElement.appendChild(rels.importNode(node,true));zip.file('xl/_rels/workbook.xml.rels',stringify(rels));}
  zip.file('xl/styles.xml',stringify(styles));zip.file('[Content_Types].xml',stringify(types));return zip.generateAsync({type:'uint8array',compression:'DEFLATE'});
}
