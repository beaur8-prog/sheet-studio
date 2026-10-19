import JSZip from 'jszip';
import {address,parseAddress,createEvaluator,formatValue} from './engine.mjs';
import {filterValue,dateNumber,filterDateParts,applyFilters} from './filters.mjs';
import {validateAutoFilter} from './filter-schema.mjs';
import {createColorResolver,xmlColor,normalizeColor} from './spreadsheet-colors.mjs';
import {xml,stringify,child,children,all,attr,readRelationships,resolvePart,xmlEscape as esc} from './chart-xml.mjs';
const NS='http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const ref=b=>`${address(b.r1,b.c1)}:${address(b.r2,b.c2)}`;
export function xlsxFilterConditions(rule){
  return rule.conditions.map(c=>{
    let value=rule.kind==='date'?String(dateNumber(c.value)):c.value,operator=c.operator;
    if(['contains','notContains','beginsWith','endsWith'].includes(operator)){
      value=operator==='beginsWith'?value+'*':operator==='endsWith'?'*'+value:'*'+value+'*';
      operator=operator==='notContains'?'notEqual':'equal';
    }if(operator==='notEqual'&&value==='')value=' ';return {operator,value};
  });
}
export function autoFilterXml(sheet,colorId){
  const f=sheet.autoFilter;if(!f)return '';validateAutoFilter(f);
  let body='';for(const [c,rule] of Object.entries(f.columns)){
    let content='';
    if(rule.type==='values'){
      const dates=rule.kind==='date'||Object.entries(sheet.cells).some(([a,v])=>parseAddress(a).c===+c&&v.style?.format==='date');
      content=`<filters${rule.values.includes('')?' blank="1"':''}>${rule.values.filter(v=>v!=='').map(v=>{
        if(dates&&Number.isFinite(Number(v))){const d=filterDateParts(Number(v));return `<dateGroupItem year="${d.year}" month="${d.month}" day="${d.day}" dateTimeGrouping="day"/>`;}
        return `<filter val="${esc(v)}"/>`;
      }).join('')}</filters>`;
    }
    if(rule.type==='custom')content=`<customFilters and="${rule.join==='and'?1:0}">${xlsxFilterConditions(rule).map(c=>`<customFilter operator="${c.operator}" val="${esc(c.value)}"/>`).join('')}</customFilters>`;
    if(rule.type==='top')content=`<top10 top="${rule.top?1:0}" percent="${rule.percent?1:0}" val="${rule.count}"/>`;
    if(rule.type==='average')content=`<dynamicFilter type="${rule.above?'aboveAverage':'belowAverage'}"/>`;
    if(rule.type==='color')content=`<colorFilter dxfId="${colorId(rule)}" cellColor="${rule.target==='background'?1:0}"/>`;
    body+=`<filterColumn colId="${+c-f.range.c1}">${content}</filterColumn>`;
  }
  if(f.sort&&f.range.r2>f.range.r1)body+=`<sortState ref="${ref({...f.range,r1:f.range.r1+1})}"><sortCondition descending="${f.sort.desc?1:0}" ref="${address(f.range.r1+1,f.sort.column)}:${address(f.range.r2,f.sort.column)}"/></sortState>`;
  return `<autoFilter ref="${ref(f.range)}">${body}</autoFilter>`;
}
async function worksheets(zip){
  const wb=xml(await zip.file('xl/workbook.xml').async('string')),rels=await readRelationships(zip,'xl/workbook.xml');
  return all(wb,'sheet').map(n=>{const rel=rels.find(r=>r.id===attr(n,'r:id'));return {name:attr(n,'name'),path:rel&&rel.mode!=='External'?resolvePart('xl/workbook.xml',rel.target):null};}).filter(v=>v.path);
}
export async function addFilterParts(bytes,book,names){
  if(!book.sheets.some(s=>s.autoFilter))return bytes;
  const zip=await JSZip.loadAsync(bytes),parts=await worksheets(zip);
  const styles=xml(await zip.file('xl/styles.xml').async('string')),styleRoot=styles.documentElement;let dxfs=child(styleRoot,'dxfs');const colorIds=new Map();
  const colorId=rule=>{
    const color=normalizeColor(rule.color),key=rule.target+':'+color;if(colorIds.has(key))return colorIds.get(key);
    if(!dxfs){dxfs=styles.createElementNS(NS,'dxfs');styleRoot.insertBefore(dxfs,children(styleRoot).find(n=>['tableStyles','colors','extLst'].includes(n.localName))||null);}
    const id=children(dxfs,'dxf').length,paint=color?'FF'+color.slice(1).toUpperCase():null,body=rule.target==='background'?`<fill><patternFill patternType="${paint?'solid':'none'}">${paint?`<fgColor rgb="${paint}"/><bgColor indexed="64"/>`:''}</patternFill></fill>`:`<font><color ${paint?`rgb="${paint}"`:'auto="1"'}/></font>`;
    dxfs.appendChild(styles.importNode(xml(`<dxf xmlns="${NS}">${body}</dxf>`).documentElement,true));dxfs.setAttribute('count',children(dxfs,'dxf').length);colorIds.set(key,id);return id;
  };
  for(const s of book.sheets){if(!s.autoFilter)continue;
    const path=parts.find(p=>p.name===(names?.get(s.id)||s.name))?.path;if(!path)continue;
    if(s.autoFilter.unsupported?.length)throw Error('가져온 미지원 아이콘·색 또는 동적 날짜 필터를 지운 뒤 XLSX로 저장하세요. 원본 파일은 변경되지 않습니다.');
    const doc=xml(await zip.file(path).async('string')),root=doc.documentElement,old=child(root,'autoFilter');if(old)root.removeChild(old);
    const node=xml(`<worksheet xmlns="${NS}">${autoFilterXml(s,colorId)}</worksheet>`).documentElement.firstChild;
    const following=children(root).find(n=>['sortState','dataConsolidate','customSheetViews','mergeCells','phoneticPr','conditionalFormatting','dataValidations','hyperlinks','printOptions','pageMargins','pageSetup','headerFooter','rowBreaks','colBreaks','customProperties','cellWatches','ignoredErrors','smartTags','drawing','legacyDrawing','legacyDrawingHF','picture','oleObjects','controls','webPublishItems','tableParts','extLst'].includes(n.localName));
    root.insertBefore(doc.importNode(node,true),following||null);
    let pr=child(root,'sheetPr');if(!pr){pr=doc.createElementNS(NS,'sheetPr');root.insertBefore(pr,root.firstChild);}pr.setAttribute('filterMode',Object.keys(s.autoFilter.columns).length?'1':'0');
    const meta=doc.createElementNS('urn:sheet-studio:filter','ss:manualHiddenRows');meta.setAttribute('rows',Object.entries(s.rowHeights||{}).filter(([,h])=>h===0).map(([r])=>r).join(','));const extNS='http://schemas.openxmlformats.org/spreadsheetml/2006/main';let extLst=child(root,'extLst');if(!extLst){extLst=doc.createElementNS(extNS,'extLst');root.appendChild(extLst);}const ext=doc.createElementNS(extNS,'ext');ext.setAttribute('uri','{84840430-696E-4AD8-A772-03A637580A33}');ext.appendChild(meta);extLst.appendChild(ext);
    zip.file(path,stringify(doc));
  }if(colorIds.size)zip.file('xl/styles.xml',stringify(styles));return zip.generateAsync({type:'uint8array',compression:'DEFLATE'});
}
export async function restoreFilterParts(bytes,book){
  const zip=await JSZip.loadAsync(bytes),parts=await worksheets(zip);
  const stylesText=await zip.file('xl/styles.xml')?.async('string'),styles=stylesText?xml(stylesText):null,dxfs=children(child(styles?.documentElement,'dxfs'),'dxf'),resolve=createColorResolver(await zip.file('xl/theme/theme1.xml')?.async('string'),stylesText);
  for(const s of book.sheets){const path=parts.find(p=>p.name===s.name)?.path;if(!path)continue;
    const doc=xml(await zip.file(path).async('string')),root=doc.documentElement,f=child(root,'autoFilter');if(!f)continue;
    const refs=attr(f,'ref').split(':').map(parseAddress),a=refs[0],b=refs[1]||a;if(!a||!b)continue;
    const range={r1:a.r,c1:a.c,r2:b.r,c2:b.c},columns={},unsupported=[];
    const evaluate=createEvaluator(book,s.id);
    for(const node of children(f,'filterColumn')){
      const c=range.c1+Number(attr(node,'colId')),values=child(node,'filters'),custom=child(node,'customFilters'),top=child(node,'top10'),dynamic=child(node,'dynamicFilter'),color=child(node,'colorFilter');
      if(values){let list=children(values,'filter').map(n=>attr(n,'val'));if(attr(values,'blank')==='1')list.push('');
        const dateColumn=Object.entries(s.cells).some(([a,v])=>parseAddress(a).c===c&&v.style?.format==='date');
        if(dateColumn){const labels=new Map();for(let r=range.r1+1;r<=range.r2;r++){const v=filterValue(s,evaluate,r,c);if(typeof v==='number')labels.set(formatValue(v,s.cells[address(r,c)]?.style),String(Math.floor(v)));}
          list=list.map(v=>v===''?'':labels.get(v)??(Number.isFinite(dateNumber(v))?String(Math.floor(dateNumber(v))):v));}
        const dates=children(values,'dateGroupItem');if(dates.length)for(let r=range.r1+1;r<=range.r2;r++){
          const v=filterValue(s,evaluate,r,c);if(typeof v!=='number')continue;
          const d=new Date(Date.UTC(1899,11,31)+(v-(v>=60?1:0))*86400000),parts={...filterDateParts(v),hour:d.getUTCHours(),minute:d.getUTCMinutes(),second:d.getUTCSeconds()};
          if(dates.some(n=>Object.entries(parts).every(([key,value])=>!n.hasAttribute(key)||Number(attr(n,key))===value)))list.push(String(Math.floor(v)));
        }columns[c]={type:'values',...(dates.length||dateColumn?{kind:'date'}:{}),values:[...new Set(list)]};
      }else if(custom){
        const conditions=children(custom,'customFilter').map(n=>({operator:attr(n,'operator','equal'),value:attr(n,'operator')==='notEqual'&&attr(n,'val')===' '?'':attr(n,'val')}));
        let kind=conditions.every(v=>v.value.trim()!==''&&Number.isFinite(Number(v.value)))?'number':'text';
        if(kind==='number'&&Object.entries(s.cells).some(([a,v])=>parseAddress(a).c===c&&v.style?.format==='date'))kind='date';
        columns[c]={type:'custom',kind,join:attr(custom,'and')==='1'?'and':'or',conditions};
      }else if(top)columns[c]={type:'top',top:attr(top,'top','1')!=='0',percent:attr(top,'percent')==='1',count:Number(attr(top,'val','10'))};
      else if(dynamic&&['aboveAverage','belowAverage'].includes(attr(dynamic,'type')))columns[c]={type:'average',above:attr(dynamic,'type')==='aboveAverage'};
      else if(color){
        const id=attr(color,'dxfId'),dxf=/^\d+$/.test(id)?dxfs[Number(id)]:null,target=['0','false'].includes(attr(color,'cellColor','1'))?'color':'background',format=child(dxf,target==='background'?'fill':'font'),pattern=child(format,'patternFill'),node=target==='background'?child(pattern,'fgColor')||child(pattern,'bgColor'):child(format,'color');
        const none=target==='background'?format&&(!children(format).length||pattern&&attr(pattern,'patternType','none')==='none'):format&&(!node||['1','true'].includes(attr(node,'auto'))),value=xmlColor(node,resolve);
        if(format&&(none||value)&&(target==='color'||none||attr(pattern,'patternType')==='solid'))columns[c]={type:'color',target,color:none?null:value};else unsupported.push(c);
      }
      else unsupported.push(c);
    }
    const hiddenRows=all(child(root,'sheetData'),'row').filter(n=>attr(n,'hidden')==='1').map(n=>Number(attr(n,'r'))-1).filter(r=>r>range.r1&&r<=range.r2);
    const saved=all(root,'manualHiddenRows')[0],manual=new Set(saved?attr(saved,'rows').split(',').filter(Boolean).map(Number):[]);
    s.autoFilter={range,columns,hiddenRows,...(unsupported.length?{unsupported}:{} )};
    const rowNodes=new Map(all(child(root,'sheetData'),'row').map(n=>[Number(attr(n,'r'))-1,n]));
    if(!unsupported.length){applyFilters(book,s);if(!saved)for(const r of hiddenRows)if(!s.autoFilter.hiddenRows.includes(r))manual.add(r);}
    else s.autoFilter.hiddenRows=hiddenRows.filter(r=>!manual.has(r));
    for(const r of hiddenRows){if(manual.has(r)){s.rowHeights[r]=0;continue;}const height=Number(attr(rowNodes.get(r),'ht'));if(height)s.rowHeights[r]=height*4/3;else delete s.rowHeights[r];}
    const sort=child(child(f,'sortState'),'sortCondition');if(sort){const p=parseAddress(attr(sort,'ref').split(':')[0]);if(p)s.autoFilter.sort={column:p.c,desc:attr(sort,'descending')==='1'};}
    validateAutoFilter(s.autoFilter);
  }
}
