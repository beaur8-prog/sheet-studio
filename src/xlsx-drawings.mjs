import {drawingAnchorRect,drawingPrintsWithSheet} from './anchor-position.mjs';
export {drawingAnchorRect} from './anchor-position.mjs';
import {capturePartTree,writeStoredParts,validateStoredParts,validateStoredRelationships} from './xlsx-part-store.mjs';
import {xml,stringify,children,child,all,attr,R,resolvePart,relationship} from './chart-xml.mjs';
const cache=new Map();
export function drawingRelationshipIds(node){const ids=new Set();for(const item of [node,...all(node,'*')])for(const a of Array.from(item.attributes||[]))if(a.namespaceURI===R&&['id','embed','link'].includes(a.localName))ids.add(a.value);return ids;}
export function remapDrawingRelationships(node,map){for(const item of [node,...all(node,'*')])for(const a of Array.from(item.attributes||[]))if(a.namespaceURI===R&&['id','embed','link'].includes(a.localName)&&map.has(a.value))item.setAttributeNS(R,a.name,map.get(a.value));}
export function validateDrawings(groups){
  if(!Array.isArray(groups)||groups.length>100)throw Error('보존된 이미지·도형 목록이 올바르지 않습니다.');let total=0;
  for(const group of groups){if(!group||!Array.isArray(group.anchors)||group.anchors.length>2000)throw Error('보존된 그리기 개체가 올바르지 않습니다.');for(const anchor of group.anchors){if(typeof anchor!=='string'||anchor.length>2_000_000||!['absoluteAnchor','oneCellAnchor','twoCellAnchor','extLst'].includes(xml(anchor).documentElement.localName))throw Error('보존된 그리기 위치가 올바르지 않습니다.');total+=anchor.length;}validateStoredParts(group.parts);validateStoredRelationships(group.rels,group.parts);total+=Object.values(group.parts).reduce((n,f)=>n+f.data.length,0);if(total>64_000_000)throw Error('보존된 이미지·도형 크기 제한을 초과했습니다.');}
}
export async function readOpaqueDrawing(zip,path,anchors,drawingRels){
  const used=new Set(anchors.flatMap(anchor=>[...drawingRelationshipIds(anchor)])),rels=drawingRels.filter(r=>used.has(r.id)).map(r=>({...r,target:r.mode==='External'?r.target:resolvePart(path,r.target)}));
  return {anchors:anchors.map(stringify),rels,parts:await capturePartTree(zip,rels.filter(r=>r.mode!=='External').map(r=>r.target))};
}
export async function writeOpaqueDrawings(zip,sheet,index,content){
  let anchors='',rels='',object=1000000;
  for(const [i,group] of (sheet.xlsxDrawings||[]).entries()){
    const prefix=`xl/sheetStudioDrawings/sheet${index+1}/group${i+1}/`,map=path=>prefix+path.replace(/^(?:xl\/sheetStudioDrawings\/sheet\d+\/group\d+\/)+/,'');await writeStoredParts(zip,group.parts,map,content);
    const ids=new Map(group.rels.map(r=>[r.id,`rIdObject${i+1}_${r.id}`]));for(const r of group.rels)rels+=relationship(ids.get(r.id),r.type,r.mode==='External'?r.target:'/'+map(r.target),r.mode);
    for(const source of group.anchors){const doc=xml(source),anchor=doc.documentElement;remapDrawingRelationships(anchor,ids);for(const node of all(anchor,'cNvPr'))node.setAttribute('id',String(++object));anchors+=stringify(anchor);}
  }
  return {anchors,rels};
}
export function drawingPreviews(sheet){
  const images=[];let other=0;
  for(const group of sheet.xlsxDrawings||[])for(const source of group.anchors){
    let info=cache.get(source);if(!info){const anchor=xml(source).documentElement,picture=all(anchor,'pic')[0],blip=all(picture,'blip')[0],crop=all(picture,'srcRect')[0],transform=all(picture,'xfrm')[0],complex=Array.from(crop?.attributes||[]).some(a=>Number(a.value)!==0)||Number(attr(transform,'rot'))!==0||['flipH','flipV'].some(key=>['1','true'].includes(attr(transform,key)))||children(blip).some(n=>n.localName!=='extLst');info={anchor,id:!complex&&blip?.getAttributeNS(R,'embed'),name:attr(all(picture,'cNvPr')[0],'name')};if(cache.size>500)cache.clear();cache.set(source,info);}
    const rel=group.rels.find(r=>r.id===info.id&&r.mode!=='External'),file=rel&&group.parts[rel.target];
    if(file&&/^image\/(?:png|jpeg|gif|webp)$/.test(file.contentType))images.push({...drawingAnchorRect(info.anchor,sheet),src:`data:${file.contentType};base64,${file.data}`,name:info.name||'가져온 이미지',printWithSheet:drawingPrintsWithSheet(info.anchor)});else other++;
  }
  return {images,other};
}
