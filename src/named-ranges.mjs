import {absoluteFormula,parseAddress} from './engine.mjs';

export function nameError(book,name,sheetId,original){
  if(typeof name!=='string'||!name||name.length>255||!/^[_\p{L}\\][_\p{L}\p{N}.\\]*$/u.test(name)||parseAddress(name)||/^(?:R|C|R\d+C\d+|TRUE|FALSE)$/i.test(name))return '이름은 문자 또는 밑줄로 시작하고, 공백이나 셀 주소를 포함할 수 없습니다. 최대 255자입니다.';
  if((book.chartNames||[]).some(n=>n!==original&&(n.sheetId||'')===(sheetId||'')&&n.name.toLocaleLowerCase()===name.toLocaleLowerCase()))return '같은 범위에 이 이름이 이미 있습니다.';
  return null;
}
export function nameDefinition(book,sheet,{name,formula,sheetId},original){
  name=String(name||'').trim();formula=String(formula||'').trim();
  const error=nameError(book,name,sheetId,original);if(error)throw Error(error);
  if(sheetId&&!book.sheets.some(s=>s.id===sheetId))throw Error('이름 범위의 시트를 찾을 수 없습니다.');
  if(!formula||formula.length>8192)throw Error('참조 대상은 1~8,192자로 입력하세요.');
  // New definitions bind their cells to the selected sheet instead of changing
  // meaning when the same name is used from another sheet.
  formula=absoluteFormula(formula.replace(/^=/,''),sheet.name);
  return {name,formula,...(sheetId?{sheetId}:{}),hidden:false};
}
const esc=v=>String(v??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
export function openNamesUI({getBook,getSheet,range,modal,checkpoint,changed,navigate,toast}){
  const mutate=operation=>{const book=getBook();if(book.nativeSource){toast('XLSM의 이름 정의 편집은 현재 지원하지 않습니다. Excel에서 편집하거나 XLSX 사본을 사용하세요.');return false;}checkpoint();operation(book);book.referenceCachesValid=false;changed();return true;};
  function show(){
    const book=getBook(),names=(book.chartNames||[]).filter(n=>!n.hidden&&!n.name.startsWith('_xlnm.'));
    const el=modal('이름 관리자',`<p>이름을 수식에 사용하거나 참조 범위로 이동할 수 있습니다. 같은 이름이 있으면 현재 시트의 정의가 우선합니다.</p><div class="named-list">${names.length?names.map((n,i)=>`<div><strong>${esc(n.name)}</strong><span>${esc(book.sheets.find(s=>s.id===n.sheetId)?.name||'통합 문서')}</span><code title="${esc(n.formula)}">${esc(n.formula)}</code><button data-name-go="${i}">이동</button><button data-name-edit="${i}">편집</button><button data-name-delete="${i}">삭제</button></div>`).join(''):'<p>정의된 이름이 없습니다.</p>'}</div><p class="format-help">새 이름의 셀 참조는 선택한 시트의 절대 참조로 저장합니다. 가져온 상대 참조 이름은 원본을 보존하며 계산은 제한될 수 있습니다.</p>`,[{label:'새 이름',primary:true,action:el=>{el.remove();edit();}}]);
    el.addEventListener('click',event=>{const button=event.target.closest('[data-name-go],[data-name-edit],[data-name-delete]');if(!button)return;
      const i=Number(button.dataset.nameGo??button.dataset.nameEdit??button.dataset.nameDelete),named=names[i];if(!named)return;
      if(button.dataset.nameGo!==undefined){el.remove();navigate(named.sheetId?"'"+book.sheets.find(s=>s.id===named.sheetId).name.replace(/'/g,"''")+"'!"+named.name:named.name);}
      else if(button.dataset.nameEdit!==undefined){el.remove();edit(named);}
      else if(mutate(book=>{book.chartNames=book.chartNames.filter(n=>n!==named);})){el.remove();show();}
    });
  }
  function edit(original){
    const book=getBook(),sheet=getSheet(),formula=original?.formula||absoluteFormula(range,sheet.name);
    const el=modal(original?'이름 편집':'새 이름',`<label>이름<input id="name-label" maxlength="255" value="${esc(original?.name||'')}"></label><label>범위<select id="name-scope"><option value="">통합 문서</option>${book.sheets.map(s=>`<option value="${esc(s.id)}" ${original?.sheetId===s.id?'selected':''}>${esc(s.name)}</option>`).join('')}</select></label><label>참조 대상<input id="name-formula" value="${esc(formula)}" maxlength="8192"></label><p id="name-error" class="worksheet-error" role="alert"></p>`,[{label:'확인',primary:true,action:el=>{
      try{const values={name:el.querySelector('#name-label').value,sheetId:el.querySelector('#name-scope').value,formula:el.querySelector('#name-formula').value};
        // Keeping an existing definition verbatim also keeps its original
        // relative reference semantics and future Excel function syntax.
        const definition=nameDefinition(book,sheet,values,original);if(original&&values.formula===original.formula)definition.formula=original.formula;
        if(mutate(book=>{book.chartNames||=[];if(original){delete original.sheetId;Object.assign(original,definition);}else book.chartNames.push(definition);})){el.remove();show();}
      }catch(error){el.querySelector('#name-error').textContent=error.message;}
    }}]);
  }
  show();
}
