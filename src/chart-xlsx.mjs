import {readAxisOptions,patchAxisOptions} from './chart-axis.mjs';
import {normalDigitWidth} from './xlsx-default-style.mjs';
import {worksheetMetadataXML,splitWorksheetXML} from './worksheet-metadata.mjs';
import JSZip from 'jszip';
import {chartSeries,chartData,parseChartFormula,refFormula,refValues,referenceSize,chartColors,chartNumberFormat} from './chart-model.mjs';
import {chartRect} from './chart-geometry.mjs';
import {numberFormatKind} from './engine.mjs';
import {anchorPlacement,positionAnchor} from './anchor-position.mjs';
import {drawingAnchorRect,readOpaqueDrawing,writeOpaqueDrawings,drawingRelationshipIds,remapDrawingRelationships} from './xlsx-drawings.mjs';
import {C,A,R,D,DECL,xmlEscape as esc,xml,stringify,children,child,all,first,attr,val,resolvePart,relPath,relationship,relationships,readRelationships} from './chart-xml.mjs';

const EMU=9525,CHART_TYPE='application/vnd.openxmlformats-officedocument.drawingml.chart+xml',DRAWING_TYPE='application/vnd.openxmlformats-officedocument.drawing+xml';
const groups={barChart:'bar',bar3DChart:'bar',lineChart:'line',line3DChart:'line',areaChart:'area',area3DChart:'area',pieChart:'pie',pie3DChart:'pie',doughnutChart:'doughnut',scatterChart:'scatter',bubbleChart:'bubble',radarChart:'radar'};
const asNumber=v=>Number.isFinite(Number(v))?Number(v):0;
const from64=value=>Uint8Array.from(atob(value),c=>c.charCodeAt(0));
function to64(bytes){let s='';for(let i=0;i<bytes.length;i+=32768)s+=String.fromCharCode(...bytes.subarray(i,i+32768));return btoa(s);}
function pointCache(node){
  const cache=first(node,'numCache')||first(node,'strCache')||first(node,'numLit')||first(node,'strLit');
  if(!cache){const levels=all(node,'lvl');return levels.length?all(levels.at(-1),'pt').map(p=>first(p,'v')?.textContent||''):[];}
  const count=Math.min(10000,Number(val(cache,'ptCount'))||all(cache,'pt').length),out=Array(count).fill(null),numeric=cache.localName.startsWith('num');
  for(const pt of children(cache,'pt')){const i=Number(attr(pt,'idx'));if(i>=0&&i<count){const v=first(pt,'v')?.textContent;out[i]=numeric?(Number.isFinite(Number(v))?Number(v):null):v||'';}}
  return out;
}
function readRef(node,workbook,ownerSheetId){if(!node)return undefined;const formula=first(node,'f')?.textContent,ref=parseChartFormula(formula,workbook,ownerSheetId);return {...ref,...(ownerSheetId?{ownerSheetId}:{}),...(formula?{formula}:{}),cache:pointCache(node),...(first(node,'formatCode')?{formatCode:first(node,'formatCode').textContent}:{})};}
function anchorRect(anchor,sheet){
  const rect=drawingAnchorRect(anchor,sheet);return {...rect,width:Math.max(100,rect.width),height:Math.max(100,rect.height)};
}
function chartColor(node,palette,fallback){const fill=child(child(node,'spPr'),'solidFill'),rgb=child(fill,'srgbClr'),scheme=child(fill,'schemeClr');return rgb?'#'+attr(rgb,'val'):palette[attr(scheme,'val')]||fallback;}
async function readParts(zip,roots,types){
  const parts={},seen=new Set();let total=0;
  async function visit(part){
    if(seen.has(part))return;seen.add(part);if(seen.size>250)throw Error('차트 연결 개체가 너무 많습니다.');
    const file=zip.file(part);if(!file)throw Error('차트 연결 파일이 없습니다: '+part);
    if(file._data?.uncompressedSize>12_000_000)throw Error('차트 파일 크기 제한을 초과했습니다.');
    const bytes=await file.async('uint8array');total+=bytes.length;if(total>16_000_000)throw Error('차트 원본 크기 제한을 초과했습니다.');
    parts[part]={data:to64(bytes),contentType:types.overrides[part]||types.defaults[part.split('.').at(-1)]||'application/octet-stream'};
    const rels=await readRelationships(zip,part);if(rels.length){const p=relPath(part);parts[p]={data:await zip.file(p).async('base64'),contentType:'application/vnd.openxmlformats-package.relationships+xml'};for(const rel of rels)if(rel.mode!=='External')await visit(resolvePart(part,rel.target));}
  }
  for(const part of roots)await visit(part);return parts;
}
export async function restoreChartParts(bytes,workbook){
  const zip=await JSZip.loadAsync(bytes),bookXML=xml(await zip.file('xl/workbook.xml').async('string')),bookRels=await readRelationships(zip,'xl/workbook.xml');
  workbook.chartNames=all(bookXML,'definedName').map(n=>({name:attr(n,'name'),formula:n.textContent,sourceXML:stringify(n),...(attr(n,'localSheetId')!==''?{sheetId:workbook.sheets[Number(attr(n,'localSheetId'))]?.id}:{}),hidden:attr(n,'hidden')==='1'}));
  workbook.chartNameSheets=workbook.sheets.map(s=>({id:s.id,name:s.name}));
  const typesXML=xml(await zip.file('[Content_Types].xml').async('string')),types={overrides:{},defaults:{}};
  all(typesXML,'Override').forEach(n=>types.overrides[attr(n,'PartName').replace(/^\//,'')]=attr(n,'ContentType'));all(typesXML,'Default').forEach(n=>types.defaults[attr(n,'Extension')]=attr(n,'ContentType'));
  const palette={};const themeRel=bookRels.find(r=>r.type.endsWith('/theme'));
  const styles=zip.file('xl/styles.xml')?xml(await zip.file('xl/styles.xml').async('string')):null,defaultXF=children(child(styles?.documentElement,'cellXfs'),'xf')[0],defaultFont=children(child(styles?.documentElement,'fonts'),'font')[Number(attr(defaultXF,'fontId','0'))]||first(styles,'font'),fontName=val(defaultFont,'name'),fontSize=Number(val(defaultFont,'sz'))||11;
  const digitWidth=normalDigitWidth({name:fontName,size:fontSize});
  if(themeRel){const theme=zip.file(resolvePart('xl/workbook.xml',themeRel.target));if(theme){const str=await theme.async('string');const colors=first(xml(str),'clrScheme');children(colors).forEach(n=>palette[n.localName]='#'+attr(child(n,'srgbClr')||child(n,'sysClr'),child(n,'srgbClr')?'val':'lastClr'));workbook.chartTheme=str;}}
  let number=0;
  for(const entry of all(bookXML,'sheet')){
    const relation=bookRels.find(r=>r.id===attr(entry,'r:id'));if(!relation||relation.mode==='External')continue;
    let sheet=workbook.sheets.find(s=>s.name===attr(entry,'name'));
    if(!sheet)continue;
    const sheetPart=resolvePart('xl/workbook.xml',relation.target),sheetXML=xml(worksheetMetadataXML(await zip.file(sheetPart).async('string'))),sheetRels=await readRelationships(zip,sheetPart);
    const format=first(sheetXML,'sheetFormatPr');if(sheet.kind!=='chart'){sheet.defaultColumnWidth=Math.round((Number(attr(format,'defaultColWidth'))||8.43)*digitWidth+5);sheet.defaultRowHeight=(Number(attr(format,'defaultRowHeight'))||15)*4/3;}
    for(const drawing of all(sheetXML,'drawing')){
      const drawingRel=sheetRels.find(r=>r.id===attr(drawing,'r:id'));if(!drawingRel||drawingRel.mode==='External')continue;
      const drawingPart=resolvePart(sheetPart,drawingRel.target),drawingXML=xml(await zip.file(drawingPart).async('string')),drawingRels=await readRelationships(zip,drawingPart);
      const opaque=children(drawingXML.documentElement).filter(anchor=>!all(anchor,'chart').some(n=>n.getAttributeNS(R,'id')));if(opaque.length){sheet.xlsxDrawings||=[];sheet.xlsxDrawings.push(await readOpaqueDrawing(zip,drawingPart,opaque,drawingRels));}
      for(const anchor of children(drawingXML.documentElement)){
        const chartNode=all(anchor,'chart').find(n=>attr(n,'r:id'));if(!chartNode)continue;
        if(++number>200)throw Error('차트는 통합 문서당 200개까지 지원합니다.');
        const rel=drawingRels.find(r=>r.id===attr(chartNode,'r:id'));if(!rel||rel.mode==='External')continue;
        const part=resolvePart(drawingPart,rel.target),sourceXML=xml(await zip.file(part).async('string')),plot=first(sourceXML,'plotArea'),groupNodes=children(plot).filter(n=>/Chart$/.test(n.localName)),known=groupNodes.filter(n=>groups[n.localName]);
        const supported=groupNodes.length>0&&known.length===groupNodes.length,base=known[0],type=supported?(known.length>1?'combo':groups[base.localName]):'unsupported';
        const title=all(child(first(sourceXML,'chart'),'title'),'t').map(n=>n.textContent).join('')||'',legend=!!first(sourceXML,'legend'),series=[];
        const axes=children(plot).filter(n=>['valAx','catAx','dateAx'].includes(n.localName)),primaryAxis=attr(children(base,'axId').at(-1),'val');
        for(const group of known){const kind=groups[group.localName]==='bar'&&val(group,'barDir')==='bar'?'bar-horizontal':groups[group.localName];
          for(const ser of children(group,'ser').slice(0,32)){
            const tx=child(ser,'tx'),nameRef=readRef(tx,workbook,sheet.id);series.push({name:child(tx,'v')?.textContent||`계열 ${series.length+1}`,nameRef,catRef:readRef(child(ser,'cat'),workbook,sheet.id),valRef:readRef(child(ser,'val')||child(ser,'yVal'),workbook,sheet.id),xRef:readRef(child(ser,'xVal'),workbook,sheet.id),sizeRef:readRef(child(ser,'bubbleSize'),workbook,sheet.id),color:chartColor(ser,palette,chartColors[series.length%chartColors.length]),pointColors:Object.fromEntries(children(ser,'dPt').map(n=>[val(n,'idx'),chartColor(n,palette,chartColors[Number(val(n,'idx'))%chartColors.length])])),type:kind,secondary:attr(children(group,'axId').at(-1),'val')!==primaryAxis,lines:val(group,'scatterStyle').includes('line')});
          }
        }
        const usedIds=drawingRelationshipIds(anchor),rels=drawingRels.filter(r=>usedIds.has(r.id)).map(r=>({...r,target:r.mode==='External'?r.target:resolvePart(drawingPart,r.target)}));
        const axisIds=children(base,'axId').map(n=>attr(n,'val')),axisTitle=id=>all(child(axes.find(n=>val(n,'axId')===id),'title'),'t').map(n=>n.textContent).join('');
        const chart={id:crypto.randomUUID(),type:type==='bar'&&val(base,'barDir')==='bar'?'bar-horizontal':type,range:'A1:B2',title,plotVisibleOnly:val(first(sourceXML,'chart'),'plotVisOnly','1')!=='0',nativeName:attr(first(anchor,'cNvPr'),'name'),placement:anchorPlacement(anchor),...anchorRect(anchor,sheet),legend,labels:[...all(sourceXML,'showVal'),...all(sourceXML,'showPercent')].some(n=>attr(n,'val')==='1'),grouping:val(base,'grouping','clustered'),series,xTitle:axisTitle(axisIds[0]),yTitle:axisTitle(axisIds[1])};
        const xNode=axes.find(a=>val(a,'axId')===axisIds[0]),yNode=axes.find(a=>val(a,'axId')===axisIds[1]),secondaryNode=axes.find(a=>a.localName==='valAx'&&val(a,'axId')!==axisIds[0]&&val(a,'axId')!==axisIds[1]);
        if(yNode)chart.valueAxis=readAxisOptions(yNode);if(xNode?.localName==='valAx')chart.xAxis=readAxisOptions(xNode);if(secondaryNode)chart.secondaryValueAxis=readAxisOptions(secondaryNode);
        if(groupNodes.some(n=>/3D/.test(n.localName)))chart.previewNote='3D 차트의 2D 미리 보기 · Excel 원본 보존';
        if(series.some(s=>[s.valRef,s.catRef,s.xRef].some(r=>r?.formula&&!r.sheetId)))chart.previewNote='해석할 수 없는 참조는 저장된 차트 값으로 표시합니다.';
        if(known.some(g=>children(g,'ser').length>32))chart.previewNote='미리 보기: 그룹당 최대 32개 계열 · Excel 원본 보존';
        if(type==='unsupported')chart.unsupportedReason='이 Excel 차트 유형은 원본을 유지하며 저장합니다.';
        chart.source={root:part,rootType:rel.type,anchor:stringify(anchor),rels,parts:await readParts(zip,rels.filter(r=>r.mode!=='External').map(r=>r.target),types),formulas:all(sourceXML,'f').map(n=>({...parseChartFormula(n.textContent,workbook,sheet.id),ownerSheetId:sheet.id,formula:n.textContent})),original:{axes:Object.fromEntries(['valueAxis','xAxis','secondaryValueAxis'].filter(k=>chart[k]).map(k=>[k,structuredClone(chart[k])])),placement:chart.placement,title,legend,plotVisibleOnly:chart.plotVisibleOnly,labels:chart.labels,xTitle:chart.xTitle,yTitle:chart.yTitle,geometry:{x:chart.x,y:chart.y,width:chart.width,height:chart.height}}};
        sheet.charts.push(chart);
      }
    }
  }
  if(!workbook.chartNames.length){delete workbook.chartNames;delete workbook.chartNameSheets;}
}

const cacheXML=(values,numeric,format='General')=>`<c:${numeric?'numCache':'strCache'}>${numeric?'<c:formatCode>'+esc(format)+'</c:formatCode>':''}<c:ptCount val="${values.length}"/>${values.map((v,i)=>v===null||v===undefined||numeric&&!Number.isFinite(Number(v))?'':`<c:pt idx="${i}"><c:v>${esc(numeric?Number(v):v)}</c:v></c:pt>`).join('')}</c:${numeric?'numCache':'strCache'}>`;
function refXML(tag,ref,workbook,names,offsets,numeric=false,fallback=[]){
  const values=ref?refValues(ref,workbook):fallback,formula=ref?.sheetId?refFormula(ref,workbook,names,offsets):ref?.formula;
  const cache=cacheXML(values,numeric,chartNumberFormat(ref,workbook));
  return `<c:${tag}>${formula?`<c:${numeric?'numRef':'strRef'}><c:f>${esc(formula)}</c:f>${cache}</c:${numeric?'numRef':'strRef'}>`:cache.replaceAll(numeric?'numCache':'strCache',numeric?'numLit':'strLit')}</c:${tag}>`;
}
const richTitle=title=>`<c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="ko-KR"/><a:t>${esc(title)}</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/></c:title>`;
const axis=(id,cross,type,position,title,secondary=false,percent=false,format='General',options={})=>{const doc=xml(`<c:${type} xmlns:c="${C}"><c:axId val="${id}"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="${position}"/>${type==='valAx'?'<c:majorGridlines/>':''}${title?richTitle(title):''}<c:numFmt formatCode="${esc(percent?'0%':format)}" sourceLinked="1"/><c:majorTickMark val="none"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/><c:crossAx val="${cross}"/><c:crosses val="${secondary?'max':'autoZero'}"/>${type==='catAx'?'<c:auto val="1"/><c:lblAlgn val="ctr"/><c:lblOffset val="100"/>':type==='dateAx'?'<c:auto val="1"/><c:lblOffset val="100"/><c:baseTimeUnit val="days"/>':'<c:crossBetween val="between"/>'}</c:${type}>`);if(type==='valAx')patchAxisOptions(doc.documentElement,options,format,percent);return stringify(doc.documentElement);};
export function newChartXML(chart,sheet,workbook,names,offsets){
  const series=chartSeries(chart,sheet,workbook).slice(0,chart.type==='pie'?1:128),data=chartData(chart,sheet,workbook),byGroup=new Map();
  const categoryFormat=chartNumberFormat(series[0]?.catRef,workbook),dateCategories=numberFormatKind(categoryFormat)==='date';
  series.forEach((s,i)=>{const type=s.type||chart.type,key=type+':'+!!s.secondary;if(!byGroup.has(key))byGroup.set(key,{type,secondary:!!s.secondary,items:[]});byGroup.get(key).items.push([s,i]);});
  let groupsXML='',needsAxes=false,secondary=false;
  for(const group of byGroup.values()){
    const {type,items}=group,tag={bar:'barChart','bar-horizontal':'barChart',line:'lineChart',area:'areaChart',pie:'pieChart',doughnut:'doughnutChart',scatter:'scatterChart',bubble:'bubbleChart',radar:'radarChart'}[type]||'barChart',isXY=['scatter','bubble'].includes(type),hasAxes=!['pie','doughnut'].includes(type),grouping=['stacked','percentStacked'].includes(chart.grouping)?chart.grouping:['line','area'].includes(type)?'standard':'clustered';
    needsAxes||=hasAxes;secondary||=group.secondary;
    let body=tag==='barChart'?`<c:barDir val="${type==='bar-horizontal'?'bar':'col'}"/><c:grouping val="${grouping}"/>`:['line','area'].includes(type)?`<c:grouping val="${grouping}"/>`:type==='scatter'?'<c:scatterStyle val="marker"/>':type==='radar'?'<c:radarStyle val="marker"/>':'';
    if(['barChart','pieChart','doughnutChart'].includes(tag))body+='<c:varyColors val="'+(['pie','doughnut'].includes(type)?1:0)+'"/>';
    for(const [s,i] of items){const color=(s.color||chartColors[i%chartColors.length]).replace('#','');
      body+=`<c:ser><c:idx val="${i}"/><c:order val="${i}"/>${s.nameRef?refXML('tx',s.nameRef,workbook,names,offsets):`<c:tx><c:v>${esc(s.name||data[i]?.name||'계열 '+(i+1))}</c:v></c:tx>`}<c:spPr><a:solidFill><a:srgbClr val="${color}"/></a:solidFill><a:ln>${type==='scatter'?'<a:noFill/>':`<a:solidFill><a:srgbClr val="${color}"/></a:solidFill>`}</a:ln></c:spPr>${['line','radar','scatter'].includes(type)?'<c:marker><c:symbol val="circle"/><c:size val="5"/></c:marker>':''}${Object.entries(s.pointColors||{}).map(([index,color])=>`<c:dPt><c:idx val="${index}"/><c:spPr><a:solidFill><a:srgbClr val="${esc(color.slice(1))}"/></a:solidFill></c:spPr></c:dPt>`).join('')}${isXY?refXML('xVal',s.xRef||s.catRef,workbook,names,offsets,true):refXML('cat',s.catRef,workbook,names,offsets,dateCategories)}${refXML(isXY?'yVal':'val',s.valRef,workbook,names,offsets,true)}${type==='bubble'?refXML('bubbleSize',s.sizeRef,workbook,names,offsets,true):''}</c:ser>`;
    }
    if(chart.labels){const percent=['pie','doughnut'].includes(type);body+=`<c:dLbls><c:showLegendKey val="0"/><c:showVal val="${percent?0:1}"/><c:showCatName val="0"/><c:showSerName val="0"/><c:showPercent val="${percent?1:0}"/><c:showBubbleSize val="0"/></c:dLbls>`;}
    if(tag==='barChart')body+=`<c:gapWidth val="150"/><c:overlap val="${grouping==='clustered'?0:100}"/>`;
    if(type==='doughnut')body+='<c:firstSliceAng val="0"/><c:holeSize val="50"/>';
    if(type==='bubble')body+='<c:bubbleScale val="100"/><c:showNegBubbles val="0"/><c:sizeRepresents val="area"/>';
    if(hasAxes)body+=`<c:axId val="${group.secondary?300:100}"/><c:axId val="${group.secondary?400:200}"/>`;
    groupsXML+=`<c:${tag}>${body}</c:${tag}>`;
  }
  let axes='';const xy=['scatter','bubble'].includes(chart.type),horizontal=chart.type==='bar-horizontal';
  if(needsAxes){axes=axis(100,200,xy?'valAx':dateCategories?'dateAx':'catAx',horizontal?'l':'b',chart.xTitle,false,false,xy?chartNumberFormat(series[0]?.xRef,workbook):categoryFormat,{gridlines:false,...chart.xAxis})+axis(200,100,'valAx',horizontal?'b':'l',chart.yTitle,false,chart.grouping==='percentStacked',chartNumberFormat(series.find(s=>!s.secondary)?.valRef,workbook),chart.valueAxis);if(secondary)axes+=axis(300,400,dateCategories?'dateAx':'catAx','b','',true,false,categoryFormat)+axis(400,300,'valAx','r','',true,chart.grouping==='percentStacked',chartNumberFormat(series.find(s=>s.secondary)?.valRef,workbook),{gridlines:false,...chart.secondaryValueAxis});}
  return `${DECL}<c:chartSpace xmlns:c="${C}" xmlns:a="${A}" xmlns:r="${R}"><c:date1904 val="${workbook.date1904?1:0}"/><c:lang val="ko-KR"/><c:chart>${chart.title?richTitle(chart.title):''}<c:autoTitleDeleted val="${chart.title?0:1}"/><c:plotArea><c:layout/>${groupsXML}${axes}</c:plotArea>${chart.legend!==false&&(chart.type!=='bar'||series.length>1)?'<c:legend><c:legendPos val="b"/><c:overlay val="0"/></c:legend>':''}<c:plotVisOnly val="${chart.plotVisibleOnly===false?0:1}"/><c:dispBlanksAs val="gap"/><c:showDLblsOverMax val="0"/></c:chart></c:chartSpace>`;
}

function updateOriginalXML(chart,workbook,names,offsets){
  const raw=new TextDecoder().decode(from64(chart.source.parts[chart.source.root].data)),doc=xml(raw),root=first(doc,'chart');
  const fs=all(doc,'f');fs.forEach((node,i)=>{const ref=chart.source.formulas[i];if(!ref?.sheetId)return;node.textContent=refFormula(ref,workbook,names,offsets);const parent=node.parentNode,cache=children(parent).find(n=>['numCache','strCache'].includes(n.localName));if(cache&&referenceSize(ref)<=10000){const fragment=xml(`<c:root xmlns:c="${C}">${cacheXML(refValues(ref,workbook),cache.localName==='numCache',child(cache,'formatCode')?.textContent||chartNumberFormat(ref,workbook))}</c:root>`);parent.replaceChild(doc.importNode(fragment.documentElement.firstChild,true),cache);}});
  if(root?.namespaceURI===C){
    const space=doc.documentElement;let dateSystem=child(space,'date1904');if(!dateSystem){dateSystem=doc.createElementNS(C,'c:date1904');space.insertBefore(dateSystem,space.firstChild);}dateSystem.setAttribute('val',workbook.date1904?'1':'0');
    if(chart.plotVisibleOnly!==chart.source.original.plotVisibleOnly){let node=child(root,'plotVisOnly');if(!node){node=doc.createElementNS(C,'c:plotVisOnly');root.insertBefore(node,child(root,'dispBlanksAs')||null);}node.setAttribute('val',chart.plotVisibleOnly===false?'0':'1');}
    if(chart.title!==chart.source.original.title){const old=child(root,'title');if(old)root.removeChild(old);if(chart.title)root.insertBefore(doc.importNode(xml(`<c:root xmlns:c="${C}" xmlns:a="${A}">${richTitle(chart.title)}</c:root>`).documentElement.firstChild,true),root.firstChild);const auto=child(root,'autoTitleDeleted');if(auto)auto.setAttribute('val',chart.title?'0':'1');}
    if(chart.legend!==chart.source.original.legend){const old=child(root,'legend');if(old)root.removeChild(old);if(chart.legend!==false){const node=xml(`<c:legend xmlns:c="${C}"><c:legendPos val="b"/><c:overlay val="0"/></c:legend>`).documentElement;root.insertBefore(doc.importNode(node,true),child(root,'plotVisOnly')||null);}}
    if(chart.labels!==chart.source.original.labels)for(const group of children(first(root,'plotArea')).filter(n=>/Chart$/.test(n.localName))){
      let labels=child(group,'dLbls');
      if(!labels){labels=doc.createElementNS(C,'c:dLbls');group.insertBefore(labels,children(group).find(n=>!['barDir','grouping','scatterStyle','radarStyle','varyColors','ser'].includes(n.localName))||null);}
      const percent=['pieChart','pie3DChart','doughnutChart'].includes(group.localName),order=['showLegendKey','showVal','showCatName','showSerName','showPercent','showBubbleSize','showLeaderLines','extLst'];
      for(const key of ['showVal','showPercent']){let show=child(labels,key);if(!show){show=doc.createElementNS(C,'c:'+key);labels.insertBefore(show,children(labels).find(n=>order.indexOf(n.localName)>order.indexOf(key))||null);}show.setAttribute('val',chart.labels&&(key==='showPercent'?percent:!percent)?'1':'0');}
    }
    const plot=first(root,'plotArea'),group=children(plot).find(n=>/Chart$/.test(n.localName)),axisIds=children(group,'axId').map(n=>attr(n,'val'));
    const primaryX=children(plot).find(n=>val(n,'axId')===axisIds[0]),primaryY=children(plot).find(n=>val(n,'axId')===axisIds[1]),secondaryY=children(plot).find(n=>n.localName==='valAx'&&![...axisIds].includes(val(n,'axId')));
    for(const [key,node] of [['valueAxis',primaryY],['xAxis',primaryX],['secondaryValueAxis',secondaryY]])if(node?.localName==='valAx'&&JSON.stringify(chart[key])!==JSON.stringify(chart.source.original.axes?.[key]))patchAxisOptions(node,chart[key],child(node,'numFmt')?.getAttribute('formatCode')||'General',chart.grouping==='percentStacked'&&key!=='xAxis');
    for(const [key,index] of [['xTitle',0],['yTitle',1]])if(chart[key]!==chart.source.original[key]){const axis=children(plot).find(n=>['catAx','valAx','dateAx'].includes(n.localName)&&val(n,'axId')===axisIds[index]);if(!axis)continue;const old=child(axis,'title');if(old)axis.removeChild(old);if(chart[key]){const node=xml(`<c:root xmlns:c="${C}" xmlns:a="${A}">${richTitle(chart[key])}</c:root>`).documentElement.firstChild;axis.insertBefore(doc.importNode(node,true),child(axis,'numFmt')||child(axis,'majorTickMark')||null);}}
  }
  return stringify(doc);
}
function absoluteAnchor(chart,graphic){const r=chartRect(chart);return `<xdr:absoluteAnchor xmlns:xdr="${D}" xmlns:a="${A}" xmlns:r="${R}"><xdr:pos x="${Math.round(r.x*EMU)}" y="${Math.round(r.y*EMU)}"/><xdr:ext cx="${Math.round(r.width*EMU)}" cy="${Math.round(r.height*EMU)}"/>${graphic}<xdr:clientData/></xdr:absoluteAnchor>`;}

export async function addChartParts(bytes,workbook,names){
  if(!workbook.sheets.some(s=>s.charts?.length||s.xlsxDrawings?.length)&&!workbook.chartNames?.length&&!workbook.chartTheme)return bytes;
  const zip=await JSZip.loadAsync(bytes),types=xml(await zip.file('[Content_Types].xml').async('string'));
  const offsets=new Map(workbook.sheets.map(s=>{const n=s.pivot?.filters.filter(f=>!s.pivot.rows.includes(f.field)&&!s.pivot.columns.includes(f.field)).length||0;return [s.id,n?n+1:0];}));
  function content(part,type){const name='/'+part;let n=all(types,'Override').find(n=>attr(n,'PartName')===name);if(!n){n=types.createElementNS(types.documentElement.namespaceURI,'Override');n.setAttribute('PartName',name);types.documentElement.appendChild(n);}n.setAttribute('ContentType',type);}
  let number=0;
  for(let index=0;index<workbook.sheets.length;index++){
    const sheet=workbook.sheets[index];if(!sheet.charts?.length&&!sheet.xlsxDrawings?.length)continue;
    let anchors='',drawingRelations='';
    for(const chart of sheet.charts||[]){
      const n=++number,id='rIdChart'+n;let graphic,preservedAnchor;
      if(chart.source&&!chart.rebuilt){
        const source=chart.source,prefix=`xl/sheetStudioCharts/chart${n}/`,map=path=>prefix+path.replace(/^(xl\/sheetStudioCharts\/chart\d+\/)+/,'');
        for(const [part,file] of Object.entries(source.parts)){
          let data=from64(file.data);
          if(part.endsWith('.rels')){const doc=xml(new TextDecoder().decode(data));for(const rel of all(doc,'Relationship'))if(attr(rel,'TargetMode')!=='External'&&attr(rel,'Target').startsWith('/'))rel.setAttribute('Target','/'+map(attr(rel,'Target').slice(1)));data=new TextEncoder().encode(stringify(doc));}
          zip.file(map(part),data);if(!part.endsWith('.rels'))content(map(part),file.contentType);
        }
        zip.file(map(source.root),updateOriginalXML(chart,workbook,names,offsets));
        const anchorDoc=xml(source.anchor),anchor=anchorDoc.documentElement;
        const relIds=new Map();for(const rel of source.rels){const newId=id+'_'+rel.id;relIds.set(rel.id,newId);drawingRelations+=relationship(newId,rel.type,rel.mode==='External'?rel.target:'/'+map(rel.target),rel.mode);}remapDrawingRelationships(anchor,relIds);
        all(anchor,'cNvPr').forEach(node=>node.setAttribute('id',String(n)));
        if(source.original.geometry&&['x','y','width','height'].every(key=>chart[key]===source.original.geometry[key])&&(!chart.placement||chart.placement===(source.original.placement||anchorPlacement(anchor))))preservedAnchor=stringify(anchor);
        else preservedAnchor=positionAnchor(anchor,{x:chart.x,y:chart.y,width:chart.width,height:chart.height},sheet,chart.placement);
        graphic=children(anchor).filter(node=>!['from','to','pos','ext','clientData'].includes(node.localName)).map(stringify).join('');
        // Keep original frame metadata but give each drawing object a unique ID.
        const graphicDoc=xml(`<xdr:root xmlns:xdr="${D}" xmlns:a="${A}" xmlns:r="${R}">${graphic}</xdr:root>`);all(graphicDoc,'cNvPr').forEach(node=>node.setAttribute('id',String(n)));graphic=children(graphicDoc.documentElement).map(stringify).join('');
      }else{
        const part=`xl/charts/sheetStudioChart${n}.xml`;zip.file(part,newChartXML(chart,sheet,workbook,names,offsets));content(part,CHART_TYPE);drawingRelations+=relationship(id,R+'/chart','/'+part);
        graphic=`<xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="${n}" name="Chart ${n}"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr><xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm><a:graphic><a:graphicData uri="${C}"><c:chart xmlns:c="${C}" r:id="${id}"/></a:graphicData></a:graphic></xdr:graphicFrame>`;
      }
      anchors+=preservedAnchor||positionAnchor(absoluteAnchor(chart,graphic),chartRect(chart),sheet,chart.placement||'fixed');
    }
    const opaque=await writeOpaqueDrawings(zip,sheet,index,content);anchors+=opaque.anchors;drawingRelations+=opaque.rels;
    const drawing=`xl/drawings/sheetStudioDrawing${index+1}.xml`,sheetPart=`xl/worksheets/sheet${index+1}.xml`;
    zip.file(drawing,`${DECL}<xdr:wsDr xmlns:xdr="${D}" xmlns:a="${A}" xmlns:r="${R}">${anchors}</xdr:wsDr>`);zip.file(relPath(drawing),relationships(drawingRelations));content(drawing,DRAWING_TYPE);
    const preserved=splitWorksheetXML(await zip.file(sheetPart).async('string')),doc=xml(preserved.metadata),node=doc.createElementNS(doc.documentElement.namespaceURI,'drawing');node.setAttributeNS(R,'r:id','rIdSheetStudioDrawing');const before=children(doc.documentElement).find(n=>['legacyDrawing','legacyDrawingHF','picture','oleObjects','controls','webPublishItems','tableParts','extLst'].includes(n.localName));doc.documentElement.insertBefore(node,before||null);zip.file(sheetPart,preserved.restore(stringify(doc)));
    const rels=zip.file(relPath(sheetPart))?xml(await zip.file(relPath(sheetPart)).async('string')):xml(relationships('')),r=xml(relationships(relationship('rIdSheetStudioDrawing',R+'/drawing','/'+drawing))).documentElement.firstChild;rels.documentElement.appendChild(rels.importNode(r,true));zip.file(relPath(sheetPart),stringify(rels));
  }
  if(workbook.chartTheme)zip.file('xl/theme/theme1.xml',workbook.chartTheme);
  if(workbook.chartNames?.length){
    const doc=xml(await zip.file('xl/workbook.xml').async('string')),root=doc.documentElement,old=child(root,'definedNames');if(old)root.removeChild(old);
    const container=doc.createElementNS(root.namespaceURI,'definedNames');
    for(const named of workbook.chartNames){const local=named.sheetId?workbook.sheets.findIndex(s=>s.id===named.sheetId):-1;if(named.sheetId&&local<0)continue;
      const node=named.sourceXML?doc.importNode(xml(named.sourceXML).documentElement,true):doc.createElementNS(root.namespaceURI,'definedName');node.removeAttribute('localSheetId');node.removeAttribute('hidden');node.setAttribute('name',named.name);if(local>=0)node.setAttribute('localSheetId',String(local));if(named.hidden)node.setAttribute('hidden','1');
      node.textContent=named.formula.replace(/"(?:[^"]|"")*"|('(?:[^']|'')+'|[A-Za-z_가-힣][\w가-힣]*)!/g,(match,name)=>{if(match.startsWith('"'))return match;const original=workbook.chartNameSheets?.find(s=>s.name.toLowerCase()===name.replace(/^'|'$/g,'').replace(/''/g,"'").toLowerCase()),sheet=workbook.sheets.find(s=>s.id===original?.id);return sheet?"'"+(names.get(sheet.id)||sheet.name).replace(/'/g,"''")+"'!":match;});container.appendChild(node);
    }
    root.insertBefore(container,child(root,'calcPr')||child(root,'pivotCaches')||null);zip.file('xl/workbook.xml',stringify(doc));
  }
  zip.file('[Content_Types].xml',stringify(types));return zip.generateAsync({type:'uint8array',compression:'DEFLATE'});
}
