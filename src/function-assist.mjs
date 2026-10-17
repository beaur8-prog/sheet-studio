import {FUNCTION_CATALOG,FUNCTION_MAP,normalizeFunctionName,searchFunctions} from './function-catalog.mjs';
const escape=value=>String(value).replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));

/** Tolerant scanner for incomplete edits; quoted text, sheets and array constants
 * must not change the active argument number or trigger function completion.
 */
export function formulaAssistance(formula,caret=formula.length) {
  if(!formula.startsWith('='))return {suggestions:[],active:null};
  const stack=[];let quote=null,brace=0;
  for(let i=1;i<caret;i++){
    const c=formula[i];
    if(quote){if(c===quote){if(formula[i+1]===quote)i++;else quote=null;}continue;}
    if(c==='"'||c==="'"){quote=c;continue;}
    if(c==='{'){brace++;continue;}if(c==='}'){brace--;continue;}if(brace)continue;
    if(c==='('){const match=/([A-Za-z_][A-Za-z0-9_.]*)\s*$/.exec(formula.slice(0,i));stack.push({name:match?normalizeFunctionName(match[1]):null,argument:0});}
    else if(c===')')stack.pop();else if((c===','||c===';')&&stack.length)stack.at(-1).argument++;
  }
  const current=[...stack].reverse().find(frame=>frame.name&&FUNCTION_MAP.has(frame.name));
  const active=current?{...FUNCTION_MAP.get(current.name),argument:current.argument}:null;
  if(quote||brace)return {suggestions:[],active};
  const match=/([A-Za-z_][A-Za-z0-9_.]*)$/.exec(formula.slice(0,caret));
  if(!match||/[\w.!'$]/.test(formula[match.index-1]||'')||/^\$?[A-Z]{1,3}\d+$/i.test(match[1]))return {suggestions:[],active};
  const prefix=normalizeFunctionName(match[1]);
  const suggestions=FUNCTION_CATALOG.filter(f=>f.name.startsWith(prefix)).slice(0,10);
  return {suggestions,active,start:match.index,end:caret};
}

export function showFunctionLibrary({onInsert,onClose}) {
  const overlay=document.createElement('div');overlay.className='backdrop function-backdrop';
  overlay.innerHTML=`<section class="modal function-library" role="dialog" aria-modal="true" aria-label="함수 라이브러리">
    <div class="function-heading"><div><h2>함수 라이브러리</h2><p>${FUNCTION_CATALOG.length}개 함수 · 이름이나 설명으로 검색하세요.</p></div><span class="function-badge">ƒx</span></div>
    <div class="function-filters"><label>함수 검색<input id="function-search" placeholder="SUMIFS, 평균, 이자…" autocomplete="off"></label><label>범주<select id="function-category"><option value="">모든 함수</option>${[...new Set(FUNCTION_CATALOG.map(f=>f.category))].map(c=>`<option>${escape(c)}</option>`).join('')}</select></label></div>
    <div class="function-browser"><div><div id="function-count" aria-live="polite"></div><div id="function-list" role="listbox" aria-label="지원 함수 목록" tabindex="0"></div></div><div id="function-detail" aria-live="polite"></div></div>
    <p class="function-limit">셀 하나의 결과를 계산합니다. 동적 배열 펼치기, LET·LAMBDA, 외부 데이터 함수는 아직 지원하지 않습니다.</p>
    <div class="modal-actions"><button data-function-close>취소</button><button class="primary" data-function-insert>수식에 삽입</button></div></section>`;
  let results=searchFunctions(),selected=0;
  const query=selector=>overlay.querySelector(selector),close=()=>{overlay.remove();onClose?.();};
  const insert=()=>{if(!results[selected])return;const fn=results[selected];overlay.remove();onInsert(fn);};
  const render=()=>{
    query('#function-count').textContent=`${results.length}개 함수`;
    query('#function-list').innerHTML=results.map((f,i)=>`<button role="option" aria-selected="${i===selected}" data-function-index="${i}" tabindex="-1"><strong>${f.name}</strong><span>${escape(f.description)}</span></button>`).join('')||'<p class="function-empty">일치하는 함수가 없습니다.</p>';
    const fn=results[selected];query('[data-function-insert]').disabled=!fn;
    query('#function-detail').innerHTML=fn?`<small>${escape(fn.category)}</small><h3>${fn.name}</h3><p>${escape(fn.description)}</p><label>구문</label><code>${escape(fn.syntax)}</code><label>예시</label><code>${escape(fn.example)}</code><p class="function-arguments">[대괄호] 안의 인수는 생략할 수 있습니다.<br>…는 인수 또는 조건 쌍을 반복할 수 있다는 뜻입니다.</p>`:'<p>다른 검색어나 범주를 선택하세요.</p>';
    query(`[data-function-index="${selected}"]`)?.scrollIntoView({block:'nearest'});
  };
  const filter=()=>{results=searchFunctions(query('#function-search').value,query('#function-category').value);selected=0;render();};
  query('#function-search').addEventListener('input',filter);query('#function-category').addEventListener('change',filter);
  overlay.addEventListener('click',e=>{const button=e.target.closest('[data-function-index]');if(button){selected=Number(button.dataset.functionIndex);render();query('#function-list').focus();}else if(e.target.closest('[data-function-close]')||e.target===overlay)close();else if(e.target.closest('[data-function-insert]'))insert();});
  overlay.addEventListener('dblclick',e=>{if(e.target.closest('[data-function-index]'))insert();});
  overlay.addEventListener('keydown',e=>{
    if(e.isComposing)return;
    if(e.key==='Escape'){e.preventDefault();e.stopPropagation();close();}
    else if(e.key==='Enter'&&e.target.tagName!=='SELECT'&&e.target.tagName!=='BUTTON'){e.preventDefault();insert();}
    else if(['ArrowDown','ArrowUp'].includes(e.key)&&e.target.tagName!=='SELECT'){e.preventDefault();selected=Math.max(0,Math.min(results.length-1,selected+(e.key==='ArrowDown'?1:-1)));render();}
    else if(e.key==='Tab'){
      const focusables=[...overlay.querySelectorAll('input,select,[tabindex="0"],.modal-actions button:not(:disabled)')],first=focusables[0],last=focusables.at(-1);
      if(e.shiftKey&&document.activeElement===first){e.preventDefault();last.focus();}else if(!e.shiftKey&&document.activeElement===last){e.preventDefault();first.focus();}
    }
  });
  document.body.append(overlay);render();query('#function-search').focus();return overlay;
}

export function createFormulaAssistant({onChange}) {
  const panel=document.createElement('div');panel.id='function-assist';panel.className='function-assist';panel.hidden=true;document.body.append(panel);
  let input=null,state=null,index=0,dismissed='',composing=false;
  const hide=()=>{panel.hidden=true;if(input){input.removeAttribute('aria-controls');input.removeAttribute('aria-activedescendant');input.setAttribute('aria-expanded','false');}};
  const update=element=>{
    input=element;if(!input||composing||document.querySelector('.backdrop'))return hide();
    state=formulaAssistance(input.value,input.selectionStart);index=0;
    if(input.selectionStart!==input.selectionEnd)state.suggestions=[];
    if(dismissed===input.value)state.suggestions=[];else dismissed='';
    render();
  };
  const render=()=>{
    if(!state?.suggestions.length&&!state?.active)return hide();
    const q=input.getBoundingClientRect();panel.style.left=`${Math.min(q.left,Math.max(8,innerWidth-460))}px`;
    panel.style.top=`${Math.min(q.bottom+4,Math.max(8,innerHeight-340))}px`;
    const active=state.active;
    panel.innerHTML=(state.suggestions.length?`<div class="function-suggestions" role="listbox" id="function-suggestions" aria-label="함수 자동 완성">${state.suggestions.map((f,i)=>`<button id="function-suggestion-${i}" role="option" aria-selected="${i===index}" data-suggestion="${i}" tabindex="-1"><strong>${f.name}</strong><span>${escape(f.description)}</span></button>`).join('')}</div><small class="function-key-hint">↑↓ 선택 · Tab 삽입 · Esc 닫기</small>`:'')+(active?`<div class="function-argument-help"><strong>${active.name}</strong> (${active.parameters.split(', ').map((p,i)=>`<span class="${i===Math.min(active.argument,active.parameters.split(', ').length-1)?'current-argument':''}">${escape(p)}</span>`).join(', ')})<small>${escape(active.description)}</small></div>`:'');
    panel.hidden=false;input.setAttribute('aria-expanded',String(!!state.suggestions.length));
    if(state.suggestions.length){input.setAttribute('aria-controls','function-suggestions');input.setAttribute('aria-activedescendant',`function-suggestion-${index}`);}else {input.removeAttribute('aria-controls');input.removeAttribute('aria-activedescendant');}
    panel.querySelector('[aria-selected="true"]')?.scrollIntoView({block:'nearest'});
  };
  const accept=()=>{
    const f=state?.suggestions[index];if(!f)return false;
    const suffix=input.value.slice(state.end),value=input.value.slice(0,state.start)+f.name+(suffix.startsWith('(')?'':'(')+suffix;
    const caret=state.start+f.name.length+1;input.value=value;input.focus();input.setSelectionRange(caret,caret);onChange(input);update(input);return true;
  };
  panel.addEventListener('pointerdown',e=>{const button=e.target.closest('[data-suggestion]');if(button){e.preventDefault();index=Number(button.dataset.suggestion);accept();}});
  return {
    update,hide,
    composition(value){composing=value;if(value)hide();else if(input)update(input);},
    keydown(e){
      if(composing||e.isComposing||panel.hidden||!state?.suggestions.length)return false;
      if(e.key==='Tab'&&!e.shiftKey){e.preventDefault();return accept();}
      if(e.key==='Escape'){e.preventDefault();dismissed=input.value;hide();return true;}
      if(['ArrowDown','ArrowUp'].includes(e.key)){e.preventDefault();index=(index+(e.key==='ArrowDown'?1:-1)+state.suggestions.length)%state.suggestions.length;render();return true;}
      return false;
    },
  };
}
