import {sourceData,buildPivot,applyPivot,aggregations,scalarKey,pivotLabel} from './pivot.mjs';

const esc=value=>String(value??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
export function createPivotUI({getWorkbook,getSheet,getRange,checkpoint,changed,switchSheet,modal,toast}){
  function open(){
    const book=getWorkbook(),target=getSheet().pivot?getSheet():null;
    const sources=book.sheets.filter(s=>!s.pivot);
    if(!sources.length)return toast('먼저 원본 데이터 시트를 만드세요.');
    let draft=target?structuredClone(target.pivot):{version:1,sourceSheetId:getSheet().id,sourceRange:getRange(),rows:[],columns:[],values:[],filters:[],totals:true,sort:'asc'};
    let data=null,loadedKey='',timer;
    const el=modal(target?'피벗테이블 편집':'피벗테이블 만들기',`
      <p class="pivot-intro">머리글이 있는 원본에서 필드를 배치하세요. 결과는 별도 시트에 표시됩니다.</p>
      <div class="pivot-source"><label>원본 시트<select id="pivot-source">${sources.map(s=>`<option value="${esc(s.id)}" ${s.id===draft.sourceSheetId?'selected':''}>${esc(s.name)}</option>`).join('')}</select></label><label>원본 범위<input id="pivot-range" value="${esc(draft.sourceRange)}" placeholder="A1:D100" spellcheck="false"></label><button id="pivot-load" type="button">필드 불러오기</button></div>
      <div id="pivot-fields" class="pivot-fields"></div>
      <div class="pivot-options"><label><input type="checkbox" id="pivot-totals" ${draft.totals?'checked':''}>행·열 총합계</label><label>정렬<select id="pivot-sort"><option value="asc">오름차순</option><option value="desc">내림차순</option><option value="source">원본 순서</option></select></label></div>
      <div id="pivot-error" role="alert"></div><div id="pivot-preview" aria-live="polite"></div>
      <p class="pivot-footnote">첫 8행·8열 미리 보기 · 새 데이터 행을 추가한 경우 원본 범위를 확장하세요. 결과 값은 원본 수정 후 새로 고침합니다.</p>`,
      [{label:target?'적용':'만들기',primary:true,action:()=>{
        try{
          if(!data||loadedKey!==sourceKey())throw new Error('변경한 원본의 필드를 먼저 불러오세요.');
          const output=buildPivot(book,draft);
          if(!target&&book.sheets.length>=100)throw new Error('시트는 최대 100개까지 지원합니다.');
          let n=1;while(book.sheets.some(s=>s.name===`피벗 ${n}`))n++;
          const next=target?structuredClone(target):{id:crypto.randomUUID(),name:`피벗 ${n}`,cells:{},columnWidths:{},rowHeights:{},merges:[],charts:[]};
          applyPivot(book,next,draft,output);checkpoint();
          if(target)Object.assign(target,next);else book.sheets.push(next);
          el.remove();clearTimeout(timer);switchSheet(next.id);toast(`${output.matched.toLocaleString()}개 데이터 행을 집계했습니다.`);
        }catch(error){showError(error);}
      }}]);
    el.querySelector('.modal').classList.add('pivot-modal');
    const q=s=>el.querySelector(s),sourceKey=()=>q('#pivot-source').value+'|'+q('#pivot-range').value.toUpperCase().replace(/\$/g,'').trim();
    q('#pivot-sort').value=draft.sort;
    function showError(error){q('#pivot-error').textContent=error.message;q('[data-modal-button]').disabled=true;}
    function preview(){
      try{
        if(!data||sourceKey()!==loadedKey)throw new Error('원본 범위를 확인하고 필드 불러오기를 누르세요.');
        const output=buildPivot(book,draft);q('#pivot-error').textContent='';q('[data-modal-button]').disabled=false;
        q('#pivot-preview').innerHTML=`<div class="pivot-preview-title">미리 보기 <span>${output.matched.toLocaleString()} / ${output.total.toLocaleString()}행 · 결과 ${output.matrix.length}행 × ${output.matrix[0].length}열</span></div><div class="pivot-preview-scroll"><table>${output.matrix.slice(0,8).map((row,r)=>`<tr>${row.slice(0,8).map(v=>`<${r?'td':'th'}>${esc(typeof v==='number'?v.toLocaleString('ko-KR',{maximumFractionDigits:2}):v)}</${r?'td':'th'}>`).join('')}</tr>`).join('')}</table></div>`;
      }catch(error){q('#pivot-preview').textContent='';showError(error);}
    }
    const options=current=>data.headers.map(h=>`<option value="${esc(h)}" ${h===current?'selected':''}>${esc(h)}</option>`).join('');
    function renderFields(){
      q('#pivot-fields').innerHTML=['rows','columns','values','filters'].map((area,ai)=>`<section class="pivot-zone"><h3>${['행','열','값','필터'][ai]}<span>${draft[area].length}</span></h3>${draft[area].map((item,i)=>{
        const field=typeof item==='string'?item:item.field;
        return `<div class="pivot-field" data-area="${area}" data-index="${i}"><select data-field aria-label="${['행','열','값','필터'][ai]} 필드 ${i+1}">${options(field)}</select>${area==='values'?`<select data-aggregate aria-label="집계 방식 ${i+1}">${Object.entries(aggregations).map(([value,label])=>`<option value="${value}" ${item.aggregate===value?'selected':''}>${label}</option>`).join('')}</select>`:''}<button type="button" data-remove aria-label="${['행','열','값','필터'][ai]} 필드 ${i+1} 제거">×</button>${area==='filters'?`<select multiple data-filter-values aria-label="필터 항목 ${i+1}">${[...new Map(data.records.map(row=>{const v=row[data.headers.indexOf(field)];return [scalarKey(v),v];})).values()].map(v=>`<option value="${esc(scalarKey(v))}" ${item.values.some(x=>scalarKey(x)===scalarKey(v))?'selected':''}>${esc(pivotLabel(v))}</option>`).join('')}</select><small>Ctrl / Shift로 여러 항목 선택</small>`:''}</div>`;
      }).join('')}<button class="pivot-add" type="button" data-add="${area}">+ ${['행','열','값','필터'][ai]} 추가</button></section>`).join('');
      preview();
    }
    function load(){
      try{
        const next=sourceData(book,q('#pivot-source').value,q('#pivot-range').value.trim());
        data=next;draft.sourceSheetId=next.source.id;draft.sourceRange=next.range;q('#pivot-range').value=next.range;loadedKey=sourceKey();
        draft.rows=draft.rows.filter(f=>data.headers.includes(f));draft.columns=draft.columns.filter(f=>data.headers.includes(f));draft.values=draft.values.filter(v=>data.headers.includes(v.field));draft.filters=draft.filters.filter(f=>data.headers.includes(f.field));
        if(!target&&!draft.rows.length&&!draft.columns.length)draft.rows=[data.headers[0]];
        if(!draft.values.length){const numeric=data.headers.find((_,i)=>data.records.some(row=>typeof row[i]==='number'));draft.values=[{field:numeric||data.headers.at(-1),aggregate:numeric?'sum':'count'}];}
        renderFields();
      }catch(error){data=null;q('#pivot-fields').innerHTML='';q('#pivot-preview').textContent='';showError(error);}
    }
    q('#pivot-load').addEventListener('click',load);
    q('#pivot-source').addEventListener('change',()=>{showError(new Error('원본 범위를 확인하고 필드 불러오기를 누르세요.'));});
    q('#pivot-range').addEventListener('input',()=>showError(new Error('원본 범위를 확인하고 필드 불러오기를 누르세요.')));
    q('#pivot-fields').addEventListener('click',event=>{
      const add=event.target.closest('[data-add]'),remove=event.target.closest('[data-remove]');
      if(add){const area=add.dataset.add,used=[...draft.rows,...draft.columns],field=data.headers.find(h=>!used.includes(h))||data.headers[0];if(draft[area].length>=(area==='rows'||area==='columns'?8:16))return;draft[area].push(area==='values'?{field,aggregate:'sum'}:area==='filters'?{field,values:[...new Map(data.records.map(row=>{const v=row[data.headers.indexOf(field)];return [scalarKey(v),v];})).values()]}:field);renderFields();}
      if(remove){const row=remove.closest('[data-area]');draft[row.dataset.area].splice(+row.dataset.index,1);renderFields();}
    });
    el.addEventListener('change',event=>{
      const row=event.target.closest('[data-area]');
      if(row){const area=row.dataset.area,i=+row.dataset.index,item=draft[area][i];
        if(event.target.matches('[data-field]')){if(typeof item==='string')draft[area][i]=event.target.value;else{item.field=event.target.value;if(area==='filters')item.values=[...new Map(data.records.map(r=>{const v=r[data.headers.indexOf(item.field)];return [scalarKey(v),v];})).values()];}renderFields();return;}
        if(event.target.matches('[data-aggregate]'))item.aggregate=event.target.value;
        if(event.target.matches('[data-filter-values]'))item.values=[...event.target.selectedOptions].map(o=>JSON.parse(o.value)[1]);
      }
      draft.totals=q('#pivot-totals').checked;draft.sort=q('#pivot-sort').value;clearTimeout(timer);timer=setTimeout(preview,80);
    });
    load();
  }
  function refresh(all=false){
    const book=getWorkbook(),targets=all?book.sheets.filter(s=>s.pivot):[getSheet()].filter(s=>s.pivot);
    if(!targets.length)return toast('새로 고칠 피벗테이블이 없습니다.');
    try{
      const copies=structuredClone(book);
      for(const target of targets)applyPivot(copies,copies.sheets.find(s=>s.id===target.id),target.pivot);
      checkpoint();for(const target of targets)Object.assign(target,copies.sheets.find(s=>s.id===target.id));
      changed({rebuild:true});toast(`${targets.length}개 피벗테이블을 새로 고쳤습니다.`);
    }catch(error){toast(`새로 고침 실패: ${error.message}`);}
  }
  function convert(){if(!getSheet().pivot)return;modal('일반 범위로 변환','<p>현재 결과 값을 유지하고 피벗 설정을 제거합니다. 변환 후 셀을 직접 편집할 수 있습니다.</p>',[{label:'변환',primary:true,action:el=>{checkpoint();delete getSheet().pivot;changed();el.remove();}}]);}
  return {open,refresh,convert,source(){const source=getWorkbook().sheets.find(s=>s.id===getSheet().pivot?.sourceSheetId);if(source)switchSheet(source.id);else toast('원본 시트가 없습니다. 피벗 설정에서 다시 선택하세요.');}};
}
