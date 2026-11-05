import {compatibilityRisks,compatibilitySummary} from './file-compatibility.mjs';
const esc=v=>String(v??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
const riskHTML=book=>'<ul class="compatibility-risks">'+compatibilityRisks(book).map(item=>'<li><strong>'+esc(item.label)+'</strong><p>'+esc(item.description)+'</p></li>').join('')+'</ul>';
export function openCompatibilityUI({book,modal}){
  const risks=compatibilityRisks(book),rows=compatibilitySummary(book);
  const el=modal('파일 호환성','<p>현재 문서에서 앱이 판별하는 기능과 저장 범위입니다.</p>'+(risks.length?'<div class="compatibility-warning"><strong>XLSX 저장 전 확인할 항목 '+risks.length+'개</strong>'+riskHTML(book)+'</div>':'')+'<table class="compatibility-table"><thead><tr><th>기능</th><th>지원 범위</th><th>설명</th></tr></thead><tbody>'+rows.map(row=>'<tr><td>'+esc(row.label)+'</td><td>'+esc(row.mode)+'</td><td>'+esc(row.detail)+'</td></tr>').join('')+'</tbody></table><p>앱 문서는 현재 편집 상태를 저장합니다. 앱이 보관하지 못하는 삽입 개체 등의 원본 파일은 별도로 유지하세요.</p>');
  el.querySelector('.modal').classList.add('compatibility-modal');return el;
}
export function confirmCompatibilityExport({book,modal,onCopy,onProject}){
  const el=modal('XLSX 사본 저장 확인','<p>다음 기능은 이 앱에서 저장한 XLSX 사본에 유지되지 않을 수 있습니다. 원본은 다른 이름으로 보관하세요.</p>'+riskHTML(book)+'<p>현재 편집 상태는 앱 문서로 저장할 수 있습니다. 기능을 제외한 XLSX는 사본으로 저장하며, 앱에서 연 원본 경로에는 덮어쓰지 않습니다.</p>',[
    {label:'앱 문서 저장',action:el=>{el.remove();onProject();}},
    {label:'기능 제외 XLSX 사본',primary:true,action:el=>{el.remove();onCopy();}}
  ]);el.querySelector('[data-close]').textContent='취소';el.querySelector('.modal').classList.add('compatibility-modal');
  if(book.xlsxPivotArchive?.blockReason){const p=document.createElement('p');p.textContent=book.xlsxPivotArchive.blockReason;el.querySelector('.modal-body').append(p);el.querySelector('[data-modal-button="1"]').disabled=true;}
  return el;
}
