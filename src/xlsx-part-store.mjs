import {xml,stringify,all,attr,resolvePart,relPath,readRelationships} from './chart-xml.mjs';
const safePart=path=>typeof path==='string'&&path.length<2000&&!path.startsWith('/')&&!path.split('/').some(p=>!p||p==='.'||p==='..')&&!/[\\:\u0000]/.test(path);
export function validateStoredParts(parts,limit=32_000_000){
  if(!parts||typeof parts!=='object'||Array.isArray(parts)||Object.keys(parts).length>1000)throw Error('보존된 Excel 연결 파일 목록이 올바르지 않습니다.');let total=0;
  for(const [path,file] of Object.entries(parts)){if(!safePart(path)||!file||typeof file.contentType!=='string'||file.contentType.length>255||typeof file.data!=='string'||!/^[A-Za-z0-9+/]*={0,2}$/.test(file.data)||(total+=file.data.length)>limit)throw Error('보존된 Excel 연결 파일이 올바르지 않거나 너무 큽니다.');}
}
export function validateStoredRelationships(rels,parts){
  if(!Array.isArray(rels)||rels.length>2000)throw Error('보존된 Excel 연결 목록이 올바르지 않습니다.');
  for(const rel of rels)if(!rel||typeof rel.id!=='string'||rel.id.length>200||typeof rel.type!=='string'||rel.type.length>4000||typeof rel.target!=='string'||rel.target.length>4000||rel.mode&&rel.mode!=='External'||rel.mode!=='External'&&(!safePart(rel.target)||!parts?.[rel.target]))throw Error('보존된 Excel 연결이 올바르지 않습니다.');
}
export async function capturePartTree(zip,roots,{maxBytes=24_000_000,maxFiles=1000}={}){
  const types=xml(await zip.file('[Content_Types].xml').async('string')),overrides=new Map(all(types,'Override').map(n=>[attr(n,'PartName').replace(/^\//,''),attr(n,'ContentType')])),defaults=new Map(all(types,'Default').map(n=>[attr(n,'Extension'),attr(n,'ContentType')])),parts={},seen=new Set();let total=0;
  async function read(path){
    if(seen.has(path))return;if(!safePart(path)||seen.size>=maxFiles)throw Error('Excel 연결 파일 경로 또는 개수가 올바르지 않습니다.');seen.add(path);
    const file=zip.file(path);if(!file)throw Error('Excel 연결 파일을 찾을 수 없습니다: '+path);
    const bytes=await file.async('uint8array');total+=bytes.length;if(total>maxBytes)throw Error('보존할 Excel 연결 파일 크기 제한을 초과했습니다.');
    parts[path]={data:await file.async('base64'),contentType:overrides.get(path)||defaults.get(path.split('.').at(-1))||'application/octet-stream'};
    const rels=await readRelationships(zip,path),relFile=zip.file(relPath(path));if(relFile){const data=await relFile.async('base64');total+=data.length*3/4;parts[relPath(path)]={data,contentType:'application/vnd.openxmlformats-package.relationships+xml'};if(total>maxBytes)throw Error('보존할 Excel 연결 파일 크기 제한을 초과했습니다.');}
    for(const rel of rels)if(rel.mode!=='External')await read(resolvePart(path,rel.target));
  }
  for(const root of roots)await read(root);return parts;
}
export async function writeStoredParts(zip,parts,map,content){
  for(const [path,file] of Object.entries(parts||{})){
    if(path.endsWith('.rels')){
      const raw=new TextDecoder().decode(Uint8Array.from(atob(file.data),c=>c.charCodeAt(0))),doc=xml(raw),owner=path.replace(/\/_rels\/([^/]+)\.rels$/,'/$1');
      for(const rel of all(doc,'Relationship'))if(attr(rel,'TargetMode')!=='External')rel.setAttribute('Target','/'+map(resolvePart(owner,attr(rel,'Target'))));zip.file(map(path),stringify(doc));
    }else{zip.file(map(path),file.data,{base64:true});content(map(path),file.contentType);}
  }
}
export function contentTypeWriter(doc){return (path,type)=>{const part='/'+path;let node=all(doc,'Override').find(n=>attr(n,'PartName')===part);if(!node){node=doc.createElementNS(doc.documentElement.namespaceURI,'Override');node.setAttribute('PartName',part);doc.documentElement.appendChild(node);}node.setAttribute('ContentType',type);};}
