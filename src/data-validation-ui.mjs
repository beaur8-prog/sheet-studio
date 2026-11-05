import {VALIDATION_TYPES,VALIDATION_OPERATORS,validationAtSelection,planValidationChange,applyValidationChange,assertValidationEditable,validationBounds} from './data-validation.mjs';
import {validationRules} from './cell-constraints.mjs';
const esc=v=>String(v??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
const options=(values,current)=>Object.entries(values).map(([key,label])=>`<option value="${key}" ${key===current?'selected':''}>${label}</option>`).join('');
export function openValidationUI({book,sheet,range,modal,checkpoint,changed,isCurrent=()=>true,review,toast}){
  try{assertValidationEditable(book,sheet,validationBounds(range));}catch(error){toast(error.message);return;}
  const selected=validationAtSelection(sheet,range),old=selected.rule;
  const r=old&&old.type in VALIDATION_TYPES?old:{type:'any',operator:'between',allowBlank:true,showList:true,showError:true,severity:'stop',formulae:[]};
  const source=r.type==='list'?(r.formulae[0]?.startsWith('"')?r.formulae[0].slice(1,-1).replace(/""/g,'"'):'='+r.formulae[0]):'';
  const el=modal('데이터 유효성 검사',`
    <p class="validation-help">선택한 셀의 입력 조건을 설정합니다. 기존 값은 바꾸지 않으며 새 입력부터 검사합니다.${selected.mixed?' <strong>서로 다른 규칙이 섞여 있습니다. 적용하면 입력한 범위의 규칙을 바꿉니다.</strong>':''}</p>
    <label>적용 범위<input id="dv-range" value="${esc(range)}" autocomplete="off"></label>
    <div class="validation-columns"><label>허용<select id="dv-type">${options(VALIDATION_TYPES,r.type)}</select></label><label id="dv-operator-label">조건<select id="dv-operator">${options(VALIDATION_OPERATORS,r.operator||'between')}</select></label></div>
    <div id="dv-values" class="validation-columns"><label><span id="dv-first-label">최솟값 / 수식</span><input id="dv-value" value="${esc(r.formulae[0]||'')}" autocomplete="off"></label><label id="dv-second-label">최댓값 / 수식<input id="dv-value2" value="${esc(r.formulae[1]||'')}" autocomplete="off"></label></div>
    <label id="dv-source-label">목록 원본<input id="dv-source" value="${esc(source)}" placeholder="예,아니오 또는 =$A$1:$A$5" autocomplete="off"><small>항목은 쉼표로 구분합니다. 셀 범위·이름·수식은 =로 시작하세요.</small></label>
    <div class="validation-checks"><label><input id="dv-blank" type="checkbox" ${r.allowBlank?'checked':''}>빈 값 허용</label><label id="dv-list-label"><input id="dv-list" type="checkbox" ${r.showList?'checked':''}>셀에 목록 표시</label></div>
    <details ${r.showInput?'open':''}><summary>입력 안내</summary><label class="worksheet-check"><input id="dv-show-input" type="checkbox" ${r.showInput?'checked':''}>셀을 선택하면 안내 표시</label><label>안내 제목<input id="dv-prompt-title" maxlength="32" value="${esc(r.promptTitle||'')}"></label><label>안내 내용<textarea id="dv-prompt" maxlength="255" rows="2">${esc(r.prompt||'')}</textarea></label></details>
    <details open><summary>오류 알림</summary><label class="worksheet-check"><input id="dv-show-error" type="checkbox" ${r.showError?'checked':''}>잘못된 값을 입력하면 알림 표시</label><label>알림 종류<select id="dv-severity">${options({stop:'중지 — 수정해야 입력 가능',warning:'경고 — 확인 후 입력 가능',information:'정보 — 확인 후 입력 가능'},r.severity)}</select></label><div class="validation-columns"><label>오류 제목<input id="dv-title" maxlength="32" value="${esc(r.title||'')}"></label><label>오류 내용<input id="dv-message" maxlength="255" value="${esc(r.message||'')}"></label></div></details>
    <p class="validation-help">날짜는 YYYY-MM-DD, 시간은 HH:MM, 조건 수식은 =로 입력할 수 있습니다. 상대 참조는 적용 범위의 왼쪽 위 셀을 기준으로 합니다. 직접 입력할 때 검사하며, 붙여넣기·자동 채우기와 기존 값은 자동 검사하지 않습니다.</p><p id="dv-error" class="worksheet-error" role="alert"></p><div id="dv-review" role="status" aria-live="polite"></div>`,[
      {label:'적용',primary:true,action:()=>apply(false)},
      {label:'규칙 지우기',action:()=>apply(true)},
      {label:'현재 값 검사',action:()=>inspect()}
    ]);
  el.querySelector('.modal').classList.add('validation-modal');const q=id=>el.querySelector('#dv-'+id);
  function apply(clear){
    try{
      if(!isCurrent())throw Error('문서가 변경되었습니다. 창을 닫고 다시 실행하세요.');
      const input=clear?null:{type:q('type').value,operator:q('operator').value,formulae:[q('value').value,q('value2').value],source:q('source').value,allowBlank:q('blank').checked,showList:q('list').checked,showInput:q('show-input').checked,promptTitle:q('prompt-title').value,prompt:q('prompt').value,showError:q('show-error').checked,severity:q('severity').value,title:q('title').value,message:q('message').value};
      const plan=planValidationChange(book,sheet,q('range').value,input);
      if(plan.changed){checkpoint();applyValidationChange(sheet,plan);if(changed()===false)return;}
      el.remove();toast(clear?'선택 범위의 입력 규칙을 지웠습니다.':'입력 규칙을 적용했습니다.');
    }catch(error){q('error').textContent=error.message;}
  }
  async function inspect(){
    if(!review)return;
    try{
      if(!isCurrent())throw Error('문서가 변경되었습니다. 창을 닫고 다시 실행하세요.');
      const range=validationBounds(q('range').value),result=await review(range);if(!result||!el.isConnected||!isCurrent())return;
      q('review').innerHTML=`<p>현재 문서에 적용된 규칙으로 ${result.checked.toLocaleString()}개 입력을 검사했습니다. 조건 불일치 ${result.invalid.toLocaleString()}개 · 확인 불가 ${result.unsupported.toLocaleString()}개${result.truncated?' · 처음 100,000개만 검사했습니다. 범위를 줄여 다시 확인하세요.':''}</p>${result.results.length?'<div class="validation-review-table"><table><thead><tr><th>셀</th><th>현재 값</th><th>결과</th></tr></thead><tbody>'+result.results.map(item=>'<tr><td>'+esc(item.ref)+'</td><td>'+esc(item.value)+'</td><td>'+esc(item.severity==='unsupported'?'규칙 수식 확인 불가':item.message)+'</td></tr>').join('')+'</tbody></table></div>':''}<p>빈 셀과 별도 셀 기록이 없는 배열 결과는 제외합니다. 결과 목록은 처음 100개이며 값을 변경하지 않습니다. 창에서 수정 중인 조건은 적용한 뒤 검사하세요.</p>`;
      q('review').scrollIntoView({block:'nearest'});
    }catch(error){q('error').textContent=error.message;}
  }
  const update=()=>{
    const type=q('type').value,numeric=!['any','list','custom'].includes(type);
    q('operator-label').hidden=!numeric;q('values').hidden=type==='list'||type==='any';q('second-label').hidden=!numeric||!q('operator').value.includes('etween');q('source-label').hidden=type!=='list';q('list-label').hidden=type!=='list';
    q('first-label').textContent=type==='custom'?'조건 수식 (TRUE이면 허용)':q('operator').value.includes('etween')?'최솟값 / 수식':'값 / 수식';
    q('value').placeholder=type==='date'?'2026-01-01 또는 =TODAY()':type==='time'?'09:00 또는 =TIME(9,0,0)':type==='custom'?'=AND(ISNUMBER(A1),A1>0)':'0 또는 =$D$1';
    q('value2').placeholder=type==='date'?'2026-12-31':type==='time'?'18:00':'100';
    for(const key of ['severity','title','message'])q(key).disabled=!q('show-error').checked;
    for(const key of ['prompt-title','prompt'])q(key).disabled=!q('show-input').checked;
  };
  for(const key of ['type','operator','show-error','show-input'])q(key).addEventListener('change',update);update();
}

export function createValidationPrompt({container,getSelection,getSheet,getRect}){
  let element=null,lastKey='',lastRule=null,dismissed=false;
  function hide(){if(element)element.hidden=true;}
  return {dismiss(){if(!element||element.hidden)return false;dismissed=true;hide();return true;},render(){
    const {ref,key}=getSelection(),sheet=getSheet(),rule=validationRules(sheet).find(r=>r.showInput&&(r.promptTitle||r.prompt)&&r.ranges.some(b=>{const p=getSelection().point;return p.r>=b.r1&&p.r<=b.r2&&p.c>=b.c1&&p.c<=b.c2;}));
    if(key!==lastKey||rule!==lastRule){dismissed=false;lastKey=key;lastRule=rule;}
    if(!rule||dismissed){hide();return;}
    const rect=getRect(),host=container(),width=host.clientWidth,height=host.clientHeight;
    if(rect.x<40||rect.y<25||rect.x>=width||rect.y>=height){hide();return;}
    if(!element){element=document.createElement('aside');element.id='validation-prompt';element.className='validation-prompt';element.setAttribute('role','note');host.append(element);}
    element.hidden=false;element.innerHTML=(rule.promptTitle?'<strong>'+esc(rule.promptTitle)+'</strong>':'')+(rule.prompt?'<span>'+esc(rule.prompt)+'</span>':'');element.setAttribute('aria-label',ref+' 입력 안내');
    const x=Math.max(40,Math.min(rect.x+Math.min(20,rect.w),width-element.offsetWidth-8)),below=rect.y+rect.h+6,y=below+element.offsetHeight<=height?below:Math.max(25,rect.y-element.offsetHeight-6);
    Object.assign(element.style,{left:x+'px',top:y+'px'});
  }};
}
