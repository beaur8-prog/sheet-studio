import {xml,stringify,children,child,attr,D} from './chart-xml.mjs';
const EMU=9525,number=value=>Number.isFinite(Number(value))?Number(value):0;
export function drawingPrintsWithSheet(source){if(!source)return true;const root=typeof source==='string'?xml(source).documentElement:source;return !['0','false'].includes(attr(child(root,'clientData'),'fPrintsWithSheet'));}
export function anchorPlacement(anchor){return anchor.localName==='absoluteAnchor'||attr(anchor,'editAs')==='absolute'?'fixed':anchor.localName==='oneCellAnchor'||attr(anchor,'editAs')==='oneCell'?'move':'move-size';}
const axisInfo=(sheet,axis)=>axis==='r'?{sizes:sheet.rowHeights||{},normal:sheet.defaultRowHeight||25,limit:1048576}:{sizes:sheet.columnWidths||{},normal:sheet.defaultColumnWidth||90,limit:16384};
export function axisOffset(sheet,axis,index){const {sizes,normal}=axisInfo(sheet,axis);let result=index*normal;for(const [key,value] of Object.entries(sizes))if(+key<index)result+=value-normal;return result;}
function axisMarker(sheet,axis,pixel){
 const {sizes,normal,limit}=axisInfo(sheet,axis);let index=0,position=0;pixel=Math.max(0,pixel);
 for(const [key,size] of Object.entries(sizes).map(([k,v])=>[+k,v]).sort((a,b)=>a[0]-b[0])){
  if(key<index||key>=limit)continue;const end=position+(key-index)*normal;if(pixel<end){const step=Math.floor((pixel-position)/normal);return {index:index+step,off:pixel-position-step*normal};}position=end;index=key;
  if(pixel<position+size)return {index,off:pixel-position};position+=size;index++;
 }
 const step=Math.min(limit-1-index,Math.max(0,Math.floor((pixel-position)/normal)));return {index:Math.min(limit-1,index+step),off:Math.max(0,pixel-position-step*normal)};
}
function markerPoint(node,sheet){const col=Math.max(0,Math.min(16384,number(child(node,'col')?.textContent))),row=Math.max(0,Math.min(1048576,number(child(node,'row')?.textContent)));return {x:axisOffset(sheet,'c',col)+number(child(node,'colOff')?.textContent)/EMU,y:axisOffset(sheet,'r',row)+number(child(node,'rowOff')?.textContent)/EMU};}
export function drawingAnchorRect(anchor,sheet){
 const from=child(anchor,'from')?markerPoint(child(anchor,'from'),sheet):{x:number(attr(child(anchor,'pos'),'x'))/EMU,y:number(attr(child(anchor,'pos'),'y'))/EMU},to=child(anchor,'to')?markerPoint(child(anchor,'to'),sheet):null,ext=child(anchor,'ext');
 return {x:Math.max(0,from.x),y:Math.max(0,from.y),width:Math.max(1,Math.min(10000,to?to.x-from.x:number(attr(ext,'cx'))/EMU||500)),height:Math.max(1,Math.min(10000,to?to.y-from.y:number(attr(ext,'cy'))/EMU||300))};
}
function makeNode(doc,prefix,name,attributes={}){const node=doc.createElementNS(D,prefix+':'+name);for(const [key,value] of Object.entries(attributes))node.setAttribute(key,String(value));return node;}
function makeMarker(doc,prefix,name,x,y,sheet){const node=makeNode(doc,prefix,name),col=axisMarker(sheet,'c',x),row=axisMarker(sheet,'r',y);for(const [key,value] of [['col',col.index],['colOff',Math.round(col.off*EMU)],['row',row.index],['rowOff',Math.round(row.off*EMU)]]){const part=makeNode(doc,prefix,key);part.textContent=String(value);node.appendChild(part);}return node;}
export function positionAnchor(source,rect,sheet,placement){
 const doc=typeof source==='string'?xml(source):source.ownerDocument,old=typeof source==='string'?doc.documentElement:source,prefix=old.prefix||'xdr',mode=placement||anchorPlacement(old),kind=mode==='fixed'?'absoluteAnchor':mode==='move'?'oneCellAnchor':'twoCellAnchor',node=makeNode(doc,prefix,kind);
 for(const a of Array.from(old.attributes||[]))if(a.localName!=='editAs')node.setAttributeNS(a.namespaceURI,a.name,a.value);
 if(mode==='fixed'){node.appendChild(makeNode(doc,prefix,'pos',{x:Math.round(rect.x*EMU),y:Math.round(rect.y*EMU)}));node.appendChild(makeNode(doc,prefix,'ext',{cx:Math.round(rect.width*EMU),cy:Math.round(rect.height*EMU)}));}
 else {node.appendChild(makeMarker(doc,prefix,'from',rect.x,rect.y,sheet));if(mode==='move-size')node.appendChild(makeMarker(doc,prefix,'to',rect.x+rect.width,rect.y+rect.height,sheet));else node.appendChild(makeNode(doc,prefix,'ext',{cx:Math.round(rect.width*EMU),cy:Math.round(rect.height*EMU)}));}
 for(const item of children(old))if(!['from','to','pos','ext'].includes(item.localName))node.appendChild(item.cloneNode(true));
 return stringify(node);
}
export function transformAnchor(source,before,after,operation){
 const doc=xml(source),node=doc.documentElement;if(!['absoluteAnchor','oneCellAnchor','twoCellAnchor'].includes(node.localName))return source;
 const mode=anchorPlacement(node),previous=drawingAnchorRect(node,before);
 if(mode==='fixed')return node.localName==='absoluteAnchor'?source:positionAnchor(node,previous,after,mode);
 if(operation){const {axis,index,delta}=operation,name=axis==='r'?'row':'col',limit=axis==='r'?1048576:16384;for(const marker of [child(node,'from'),mode==='move-size'?child(node,'to'):null]){const coordinate=child(marker,name);if(!coordinate)continue;const old=number(coordinate.textContent),removed=delta<0&&old>=index&&old<index-delta,next=old<index?old:removed?index:old+delta;coordinate.textContent=String(Math.max(0,Math.min(limit-1,next)));if(removed){const off=child(marker,name+'Off');if(off)off.textContent='0';}}}
 if(mode==='move'){const from=markerPoint(child(node,'from'),after);return positionAnchor(node,{...previous,...from},after,mode);}
 return stringify(node);
}
