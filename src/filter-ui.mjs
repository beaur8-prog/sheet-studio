import {cloneData} from './clone.mjs';
import {filterOptions,valueKey,dateNumber,filterPredicates,searchFilterValue} from './filters.mjs';
import {address,createEvaluator,formatValue} from './engine.mjs';
const esc=v=>String(v??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
let closeCurrent=()=>{};
export const closeFilterMenu=()=>closeCurrent();
const operators=[['equal','같음'],['notEqual','같지 않음'],['greaterThan','보다 큼'],['greaterThanOrEqual','크거나 같음'],['lessThan','보다 작음'],['lessThanOrEqual','작거나 같음'],['beginsWith','시작 문자'],['endsWith','끝 문자'],['contains','포함'],['notContains','포함하지 않음']];

export function openFilterMenu({book,sheet,column,anchor,apply,sort,focusGrid,prepared}){
  closeCurrent();const {kind,values,colors}=(prepared?.options||filterOptions(book,sheet,column)),rule=sheet.autoFilter.columns[column],title=sheet.cells[address(sheet.autoFilter.range.r1,column)]?.raw||address(sheet.autoFilter.range.r1,column);
  const own=!prepared&&rule?.type!=='color'&&filterPredicates(book,sheet,createEvaluator(book,sheet.id)).find(p=>p.column===column);
  const original=new Set(prepared?prepared.original:rule?.type==='values'?rule.values.map(valueKey):values.filter(v=>!own||own.test(v.value)).map(v=>v.key));
  let selected=new Set(original),search='',mode=rule?.type==='color'?'color':'list',colorRule=rule?.type==='color'?cloneData(rule):null;
  if(colorRule&&!colors[colorRule.target].some(v=>v.color===colorRule.color))colors[colorRule.target].push({color:colorRule.color,count:0});
  const colorMenu=Object.entries(colors).map(([target,items])=>`<fieldset><legend>${target==='background'?'셀 채우기색':'글자색'}</legend>${items.slice(0,1000).map(({color,count})=>`<label><input type="radio" name="filter-color" data-color-target="${target}" value="${color||''}" ${colorRule?.target===target&&colorRule.color===color?'checked':''}><span class="filter-color-swatch ${target==='background'&&!color?'no-fill':''}" style="${target==='background'&&color?'background:'+color:'color:'+(color||'#222222')}">${target==='color'?'가':''}</span><span>${color?esc(color.toUpperCase()):target==='background'?'채우기 없음':'자동 (기본 글자색)'}</span><small>${count.toLocaleString()}개</small></label>`).join('')}${!items.length?'<p>일치하는 행이 없습니다.</p>':items.length>1000?'<p>처음 1,000개 색 표시 · 다른 열 조건으로 범위를 좁히세요.</p>':''}</fieldset>`).join('');
  const el=document.createElement('div');el.className='filter-popup';el.setAttribute('role','dialog');el.setAttribute('aria-label',`${title} 필터`);
  el.innerHTML=`<div class="filter-heading">${esc(title)}</div><div class="filter-main">
    <button type="button" data-sort="asc">↑ ${kind==='text'?'텍스트 오름차순 정렬':'오름차순 정렬'}</button>
    <button type="button" data-sort="desc">↓ ${kind==='text'?'텍스트 내림차순 정렬':'내림차순 정렬'}</button><hr>
    <button type="button" id="filter-clear" ${rule?'':'disabled'}>“${esc(title)}”에서 필터 해제</button>
    <details class="filter-color-menu" ${mode==='color'?'open':''}><summary>색 기준 필터</summary><div>${colorMenu}</div><p>색을 선택한 뒤 확인을 누르세요.</p></details>
    <details class="filter-condition-menu"><summary>${{text:'텍스트',number:'숫자',date:'날짜'}[kind]} 필터</summary><div>${operators.filter(([op])=>kind==='text'||!['beginsWith','endsWith','contains','notContains'].includes(op)).map(([op,label])=>`<button type="button" data-condition="${op}">${kind==='date'&&op==='greaterThan'?'이후':kind==='date'&&op==='lessThan'?'이전':label}…</button>`).join('')}<button type="button" data-condition="between">사이…</button>${kind==='number'?'<button type="button" data-condition="top">상위 10…</button><button type="button" data-average="above">평균 초과</button><button type="button" data-average="below">평균 미만</button>':''}<button type="button" data-condition="custom">사용자 지정 필터…</button></div></details>
    <label class="filter-search"><span>검색</span><input id="filter-search" type="search" placeholder="검색" aria-label="필터 값 검색" autocomplete="off"></label>
    <label class="filter-add" hidden><input type="checkbox" id="filter-add">현재 선택 영역을 필터에 추가</label>
    <div class="filter-values" role="group" aria-label="필터 값 목록"></div>
    <p class="filter-count"></p></div>
    <div class="filter-custom" hidden></div><p class="filter-error" role="alert"></p>
    <div class="filter-buttons"><button type="button" id="filter-ok" class="primary">확인</button><button type="button" id="filter-cancel">취소</button></div>`;
  document.body.append(el);const $=q=>el.querySelector(q),controller=new AbortController();
  const close=()=>{controller.abort();el.remove();closeCurrent=()=>{};focusGrid();};closeCurrent=close;
  const position=()=>{el.style.left=Math.max(6,Math.min(anchor.left,innerWidth-el.offsetWidth-8))+'px';el.style.top=Math.max(6,Math.min(anchor.bottom,innerHeight-el.offsetHeight-8))+'px';};
  const visible=()=>values.filter(v=>searchFilterValue(v.label,search));
  function drawValues(){
    const list=visible(),count=list.filter(v=>selected.has(v.key)).length;
    // Render a bounded page; search covers every value, including those beyond this page.
    $('.filter-values').innerHTML=`<label><input type="checkbox" id="filter-all" ${count===list.length&&count?'checked':''}>${search?'(모든 검색 결과 선택)':'(모두 선택)'}</label>`+list.slice(0,10000).map((v,i)=>`<label title="${esc(v.label)}"><input type="checkbox" data-value="${i}" ${selected.has(v.key)?'checked':''}><span>${esc(v.label)}</span></label>`).join('')+(list.length?'':'<p>일치하는 항목이 없습니다.</p>');
    $('#filter-all').indeterminate=count>0&&count<list.length;
    $('.filter-count').textContent=list.length>10000?`${list.length.toLocaleString()}개 중 처음 10,000개 표시 · 검색으로 범위를 좁히세요`:'';
    $('#filter-ok').disabled=mode==='color'?!colorRule:count===0&&!($('#filter-add').checked&&original.size);
    $('.filter-add').hidden=!search||!rule;
  }
  function custom(op){
    mode=op==='top'?'top':'custom';$('.filter-main').hidden=true;$('.filter-custom').hidden=false;$('#filter-ok').disabled=false;
    if(mode==='top')$('.filter-custom').innerHTML='<h3>상위 10 자동 필터</h3><label>방향<select id="filter-top"><option value="top">상위</option><option value="bottom">하위</option></select></label><label>개수<input id="filter-top-count" type="number" min="1" max="500" value="10"></label><label>단위<select id="filter-top-unit"><option value="items">항목</option><option value="percent">퍼센트</option></select></label>';
    else {
      const initial=op==='custom'&&rule?.type==='custom'?rule.conditions:[{operator:op==='between'?'greaterThanOrEqual':op==='custom'?'equal':op,value:''},...(op==='between'?[{operator:'lessThanOrEqual',value:''}]:[])];
      $('.filter-custom').innerHTML=`<h3>사용자 지정 자동 필터</h3><p>${esc(title)} · ${kind==='date'?'날짜는 YYYY-MM-DD로 입력하세요.':kind==='text'?'*는 여러 문자, ?는 한 문자, ~는 와일드카드 이스케이프입니다.':'숫자를 입력하세요.'}</p>${[0,1].map(i=>`${i?'<label class="filter-join"><select id="filter-join"><option value="and">그리고 (AND)</option><option value="or">또는 (OR)</option></select></label>':''}<div class="filter-condition"><select id="filter-op-${i}" aria-label="${i+1}번째 조건">${i?'<option value="">(조건 없음)</option>':''}${operators.filter(([v])=>kind==='text'||!['beginsWith','endsWith','contains','notContains'].includes(v)).map(([v,label])=>`<option value="${v}">${label}</option>`).join('')}</select><input id="filter-value-${i}" aria-label="${i+1}번째 조건 값" type="${kind==='date'?'date':'text'}"></div>`).join('')}`;
      initial.forEach((c,i)=>{if(i>1)return;$(`#filter-op-${i}`).value=c.operator;$(`#filter-value-${i}`).value=kind==='date'&&c.value!==''&&Number.isFinite(Number(c.value))?formatValue(Number(c.value),{format:'date'},book):c.value;});
      if(op==='custom'&&rule?.type==='custom')$('#filter-join').value=rule.join;
    }position();$('.filter-custom input').focus();
  }
  function confirm(){
    let next;
    if(mode==='color'){
      if(!colorRule)return;next=colorRule;
    }else if(mode==='list'){
      const list=visible(),set=new Set(list.filter(v=>selected.has(v.key)).map(v=>v.key));
      if(search&&$('#filter-add').checked)for(const v of original)set.add(v);
      if(!set.size)return;
      // Selecting every available item removes this column's condition, even with other filters active.
      next=!search&&set.size===values.length?null:{type:'values',...(kind==='date'?{kind}:{}),values:[...set]};
    }else if(mode==='top'){
      const count=Number($('#filter-top-count').value),percent=$('#filter-top-unit').value==='percent';
      if(!Number.isInteger(count)||count<1||count>(percent?100:500)){$('.filter-error').textContent=`1~${percent?100:500} 사이 정수를 입력하세요.`;return;}
      next={type:'top',top:$('#filter-top').value==='top',percent,count};
    }else{
      const conditions=[0,1].map(i=>({operator:$(`#filter-op-${i}`).value,value:$(`#filter-value-${i}`).value})).filter(c=>c.operator);
      if(kind!=='text'&&conditions.some(c=>c.value.trim()===''||!Number.isFinite(kind==='date'?dateNumber(c.value,book):Number(c.value)))){$('.filter-error').textContent='올바른 조건 값을 입력하세요.';return;}
      next={type:'custom',kind,join:$('#filter-join').value,conditions};
    }
    apply(next);close();
  }
  el.addEventListener('pointerdown',e=>e.stopPropagation());el.addEventListener('dblclick',e=>e.stopPropagation());
  el.addEventListener('click',e=>{e.stopPropagation();const s=e.target.closest('[data-sort]'),c=e.target.closest('[data-condition]'),a=e.target.closest('[data-average]');if(s){sort(s.dataset.sort==='desc');close();}if(c)custom(c.dataset.condition);if(a){apply({type:'average',above:a.dataset.average==='above'});close();}});
  const listMode=()=>{mode='list';el.querySelectorAll('[name="filter-color"]').forEach(input=>input.checked=false);};
  $('.filter-color-menu').addEventListener('change',e=>{if(!e.target.matches('[data-color-target]'))return;mode='color';colorRule={type:'color',target:e.target.dataset.colorTarget,color:e.target.value||null};$('#filter-ok').disabled=false;});
  el.querySelectorAll('details').forEach(details=>details.addEventListener('toggle',()=>{if(details.open)position();}));
  $('#filter-search').addEventListener('input',e=>{listMode();search=e.target.value;selected=new Set(search?visible().map(v=>v.key):original);drawValues();});
  $('.filter-values').addEventListener('change',e=>{listMode();const list=visible();if(e.target.id==='filter-all'){for(const v of list)if(e.target.checked)selected.add(v.key);else selected.delete(v.key);}else if(e.target.dataset.value!==undefined){const key=list[+e.target.dataset.value].key;if(e.target.checked)selected.add(key);else selected.delete(key);}const selector=e.target.id?'#'+e.target.id:'[data-value=\"'+e.target.dataset.value+'\"]';drawValues();$(selector)?.focus();});
  $('#filter-add').addEventListener('change',drawValues);$('#filter-ok').onclick=confirm;$('#filter-cancel').onclick=close;
  $('#filter-clear').onclick=()=>{apply(null);close();};
  el.addEventListener('keydown',e=>{e.stopPropagation();if(e.isComposing)return;if(e.key==='Escape'){e.preventDefault();close();}else if(e.key==='Enter'&&e.target.tagName==='INPUT'){e.preventDefault();if(!$('#filter-ok').disabled)confirm();}else if(e.key==='Tab'){
    const controls=[...el.querySelectorAll('button,input,select,summary')].filter(v=>!v.disabled&&v.getClientRects().length),first=controls[0],last=controls.at(-1);
    if(e.shiftKey&&document.activeElement===first){e.preventDefault();last.focus();}else if(!e.shiftKey&&document.activeElement===last){e.preventDefault();first.focus();}
  }});
  document.addEventListener('pointerdown',e=>{if(!el.contains(e.target))close();},{signal:controller.signal});
  window.addEventListener('resize',close,{signal:controller.signal});
  document.querySelector('#grid').addEventListener('scroll',close,{signal:controller.signal});
  drawValues();position();$('#filter-search').focus();
}
