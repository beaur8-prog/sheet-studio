import {DOMParser,XMLSerializer} from '@xmldom/xmldom';
export const C='http://schemas.openxmlformats.org/drawingml/2006/chart',A='http://schemas.openxmlformats.org/drawingml/2006/main',R='http://schemas.openxmlformats.org/officeDocument/2006/relationships',D='http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing';
export const DECL='<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
export const xmlEscape=v=>String(v??'').replace(/[&<>"'\u0000-\u0008\u000b\u000c\u000e-\u001f]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&apos;'}[c]||''));
export function xml(source){
  if(typeof source!=='string'||source.length>12_000_000||/<!DOCTYPE|<!ENTITY/i.test(source))throw Error('차트 XML이 올바르지 않거나 너무 큽니다.');
  return new DOMParser({onError:(level,message)=>{if(level!=='warning')throw Error(message);}}).parseFromString(source,'application/xml');
}
export const stringify=node=>new XMLSerializer().serializeToString(node);
export const children=(node,name)=>Array.from(node?.childNodes||[]).filter(n=>n.nodeType===1&&(!name||n.localName===name));
export const child=(node,name)=>children(node,name)[0];
export const all=(node,name)=>Array.from(node?.getElementsByTagNameNS('*',name)||[]);
export const first=(node,name)=>all(node,name)[0];
export const attr=(node,name,fallback='')=>node?.getAttribute(name)||fallback;
export const val=(node,name,fallback='')=>attr(child(node,name),'val',fallback);
export function resolvePart(base,target){
  if(typeof target!=='string'||/[\\\u0000]/.test(target)||/^[a-z][a-z0-9+.-]*:/i.test(target))throw Error('차트 내부 경로가 올바르지 않습니다.');
  const parts=(target.startsWith('/')?target.slice(1):base.slice(0,base.lastIndexOf('/')+1)+target).split('/'),out=[];
  for(const p of parts){if(p==='..'){if(!out.length)throw Error('차트 경로가 파일 범위를 벗어납니다.');out.pop();}else if(p&&p!=='.')out.push(p);}
  return out.join('/');
}
export const relPath=part=>part.slice(0,part.lastIndexOf('/')+1)+'_rels/'+part.slice(part.lastIndexOf('/')+1)+'.rels';
export const relationship=(id,type,target,mode)=>`<Relationship Id="${xmlEscape(id)}" Type="${xmlEscape(type)}" Target="${xmlEscape(target)}"${mode?' TargetMode="'+xmlEscape(mode)+'"':''}/>`;
export const relationships=body=>`${DECL}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${body}</Relationships>`;
export async function readRelationships(zip,part){
  const entry=zip.file(relPath(part));if(!entry)return [];
  return all(xml(await entry.async('string')),'Relationship').map(n=>({id:attr(n,'Id'),type:attr(n,'Type'),target:attr(n,'Target'),mode:attr(n,'TargetMode')}));
}
export function replaceChild(parent,name,replacement){const old=child(parent,name);if(old)parent.replaceChild(replacement,old);else parent.appendChild(replacement);}
