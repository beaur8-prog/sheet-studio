const esc=v=>String(v??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
export function createFindUI({getWorkbook,getSheet,getRange,getRevision,getDocumentId,search,replace,navigate,modal,toast,runBusy}){
  let saved={query:'',replacement:'',matchCase:false,wholeCell:false,scope:'sheet',lookIn:'formulas',includeHiddenSheets:false};
  return {open({readOnly=false}={}){
    const documentId=getDocumentId(),sheetId=getSheet().id,range={...getRange()};let preview=null,signature='',previewRevision=-1,index=-1,sequence=0;
    const options=()=>({...saved,query:q('#find-text').value,replacement:q('#replace-text').value,scope:q('#find-scope').value,lookIn:q('#find-look-in').value,matchCase:q('#find-case').checked,wholeCell:q('#find-whole').checked,includeHiddenSheets:!readOnly&&q('#find-hidden').checked,sheetId,range});
    const current=()=>el.isConnected&&getDocumentId()===documentId;
    const el=modal('찾기 및 바꾸기',`
      <div class="find-fields"><label>찾을 내용<input id="find-text" value="${esc(saved.query)}" autocomplete="off"></label><label>바꿀 내용<input id="replace-text" value="${esc(saved.replacement)}" autocomplete="off"></label></div>
      <div class="find-options"><label>검색 범위<select id="find-scope"><option value="sheet">현재 시트</option><option value="selection">선택 영역</option><option value="workbook">통합 문서</option></select></label><label>검색 대상<select id="find-look-in"><option value="formulas">값·수식 원문</option><option value="values">표시 값 (검색 전용)</option></select></label></div>
      <div class="find-checks"><label><input type="checkbox" id="find-case" ${saved.matchCase?'checked':''}>대소문자 구분</label><label><input type="checkbox" id="find-whole" ${saved.wholeCell?'checked':''}>셀 내용 전체 일치</label><label><input type="checkbox" id="find-hidden" ${saved.includeHiddenSheets?'checked':''}>숨겨진 시트 포함</label></div>
      <p class="find-help">입력한 글자를 그대로 찾습니다. 모두 바꾸기는 미리 보기를 확인한 뒤 적용하며 한 번에 실행 취소할 수 있습니다.</p>
      <p id="find-status" role="status" aria-live="polite">범위와 대상을 선택하고 찾기를 누르세요.</p><div id="find-results" class="find-results" aria-label="검색 결과"></div>`,[
      {label:'다음 찾기',primary:true,action:()=>next()},
      {label:'모두 찾기 / 미리 보기',action:()=>scan()},
      {label:'모두 바꾸기',action:()=>apply()}
    ]);
    el.querySelector('.modal').classList.add('find-modal');const q=s=>el.querySelector(s);
    q('#find-scope').value=saved.scope;q('#find-look-in').value=saved.lookIn;
    const invalidate=()=>{++sequence;preview=null;signature='';previewRevision=-1;q('#find-results').textContent='';q('#find-status').textContent='조건이 변경되었습니다. 찾기 또는 미리 보기를 실행하세요.';q('[data-modal-button="2"]').disabled=q('#find-look-in').value==='values';q('#replace-text').disabled=q('#find-look-in').value==='values';q('#find-hidden').disabled=q('#find-scope').value!=='workbook';};
    const writeStatus=()=>{q('#find-status').textContent=`${preview.count.toLocaleString()}개 발견${preview.truncated?' · 처음 1,000개만 표시합니다. 범위를 좁혀 검색하세요.':''}${index>=0?' · '+preview.results[index].sheetName+'!'+preview.results[index].ref:''}`;};
    async function scan(){
      if(!current())return;const o=options(),key=JSON.stringify(o),revision=getRevision(),request=++sequence;
      if(!o.query){q('#find-status').textContent='찾을 내용을 입력하세요.';return;}
      try{
        const result=await runBusy('셀을 찾는 중…',async()=>{try{return await search(o);}catch(error){return {error:error.message};}});if(!result||!current()||request!==sequence||revision!==getRevision()||key!==JSON.stringify(options()))return;
        if(result.error){q('#find-status').textContent=result.error;return;}
        saved={...o};preview=result;signature=key;previewRevision=revision;index=-1;writeStatus();
        q('#find-results').innerHTML=result.results.length?'<table><thead><tr><th>시트·셀</th><th>현재 내용</th>'+(o.lookIn==='formulas'?'<th>바꾸기 미리 보기</th>':'')+'</tr></thead><tbody>'+result.results.slice(0,100).map((item,i)=>'<tr><td><button type="button" data-find-result="'+i+'">'+esc(item.sheetName+'!'+item.ref)+'</button></td><td>'+esc(item.text)+'</td>'+(o.lookIn==='formulas'?'<td>'+esc(item.after)+'</td>':'')+'</tr>').join('')+'</tbody></table>'+(result.results.length>100?'<p>목록은 처음 100개입니다. 다음 찾기로 표시된 1,000개 결과까지 이동할 수 있습니다.</p>':''):'';
        return result;
      }catch(error){if(current())q('#find-status').textContent=error.message;}
    }
    const fresh=()=>preview&&signature===JSON.stringify(options())&&previewRevision===getRevision();
    async function jump(i){
      if(!fresh()||!preview.results[i])return;index=i;const item=preview.results[index];await navigate(item);if(!current())return;previewRevision=getRevision();writeStatus();q('#find-text').focus();
    }
    async function next(){if(!fresh())await scan();if(!fresh()||!preview.results.length)return;await jump((index+1)%preview.results.length);}
    async function apply(){
      if(!fresh()){await scan();if(current()&&preview)q('#find-status').textContent+=' · 미리 보기를 확인한 뒤 모두 바꾸기를 다시 누르세요.';return;}
      const o=options();if(readOnly||o.lookIn==='values')return;if(!preview.count)return;
      const key=signature,request=++sequence;
      try{
        const result=await runBusy('바꾸기 결과를 검증하는 중…',async()=>{try{return await replace(o,()=>current()&&request===sequence&&key===JSON.stringify(options()));}catch(error){return {error:error.message};}});
        if(!result||!el.isConnected)return;if(result.error){q('#find-status').textContent='변경하지 않았습니다: '+result.error;return;}preview=null;signature='';previewRevision=-1;q('#find-results').textContent='';q('#find-status').textContent=`${result.count.toLocaleString()}개 셀을 변경했습니다. 실행 취소로 되돌릴 수 있습니다.`;toast(`${result.count.toLocaleString()}개 셀을 변경했습니다.`);
      }catch(error){if(el.isConnected)q('#find-status').textContent='변경하지 않았습니다: '+error.message;}
    }
    for(const input of el.querySelectorAll('input,select'))input.addEventListener(input.tagName==='SELECT'||input.type==='checkbox'?'change':'input',invalidate);
    el.addEventListener('click',event=>{const button=event.target.closest('[data-find-result]');if(button)void jump(+button.dataset.findResult);});
    el.addEventListener('keydown',event=>{if(event.key==='Tab'){const nodes=[...el.querySelectorAll('button:not([disabled]),input:not([disabled]),select:not([disabled])')],i=nodes.indexOf(document.activeElement);if(event.shiftKey&&i===0||!event.shiftKey&&i===nodes.length-1){event.preventDefault();nodes[event.shiftKey?nodes.length-1:0]?.focus();}}});
    q('[data-modal-button="2"]').disabled=saved.lookIn==='values';q('#replace-text').disabled=saved.lookIn==='values';q('#find-hidden').disabled=saved.scope!=='workbook';
    if(readOnly){
      el.querySelector('.modal').setAttribute('aria-label','찾기');el.querySelector('h2').textContent='찾기';
      q('#replace-text').closest('label').hidden=true;q('[data-modal-button="2"]').hidden=true;q('#find-hidden').closest('label').hidden=true;
      q('.find-help').textContent='현재 시트나 통합 문서에서 찾습니다. 보기 모드에서는 문서를 바꾸지 않습니다.';
    }
    q('#find-text').select();return el;
  }};
}
