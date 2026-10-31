import './update-ui.css';

const RELEASE_PAGE='https://github.com/beaur8-prog/sheet-studio/releases/latest';
const ACTIVE=new Set(['checking','downloading','verifying','installing']);
const LABELS={idle:'업데이트 확인을 눌러 새 버전을 확인하세요.',checking:'새 버전을 확인하는 중…',available:'새 버전을 다운로드할 수 있습니다.',downloading:'업데이트 다운로드 중…',verifying:'다운로드 파일의 무결성을 확인하는 중…',ready:'다운로드와 무결성 확인이 끝났습니다.',installing:'문서 복구 저장을 마치고 설치 화면을 여는 중…','up-to-date':'설치된 버전보다 새로운 공개 버전이 없습니다.',error:'업데이트를 완료하지 못했습니다.'};
function size(bytes){const n=Math.max(0,Number(bytes)||0),unit=n>=1024*1024?['MB',1024*1024]:n>=1024?['KB',1024]:['B',1];return (n/unit[1]).toLocaleString('ko-KR',{maximumFractionDigits:1})+' '+unit[0];}

export function createUpdateUI({bridge,version,modal,onClose=()=>{}}){
 let state={phase:'idle',percent:0,version:null,currentVersion:version,automatic:false,packaged:false,supportsAutomatic:false,platform:'browser',revision:-1},dialog;
 const notice=document.querySelector('#update-status');
 function status(){return state.phase==='downloading'?`업데이트 다운로드 중 ${state.percent}%`:LABELS[state.phase]||LABELS.idle;}
 function progress(element){
  element.hidden=!['checking','downloading','verifying','ready','installing'].includes(state.phase);
  if(state.phase==='downloading'||state.phase==='ready')element.value=state.percent;
  else element.removeAttribute('value');
 }
 function render(){
  notice.hidden=['idle','up-to-date'].includes(state.phase);
  notice.querySelector('span').textContent=status();progress(notice.querySelector('progress'));
  if(!dialog?.isConnected)return;
  dialog.dataset.phase=state.phase;
  const $=selector=>dialog.querySelector(selector), supported=!!bridge?.updateAction&&state.packaged&&state.supportsAutomatic, active=ACTIVE.has(state.phase);
  $('#update-current').textContent=state.currentVersion||version;
  $('#update-latest').textContent=state.version||'—';
  $('#update-phase').textContent=status();progress($('#update-progress'));
  $('#update-transfer').textContent=state.totalBytes&&['downloading','verifying','ready'].includes(state.phase)?`${size(state.downloadedBytes)} / ${size(state.totalBytes)}`:'';
  $('#update-error').textContent=state.error||'';
  $('#update-error').hidden=!state.error;
  $('#update-checked').textContent=state.checkedAt?'마지막 확인: '+new Date(state.checkedAt).toLocaleString('ko-KR'):'';
  const automatic=$('#update-automatic'),manual=$('#update-manual');
  automatic.checked=!!state.automatic;manual.checked=!state.automatic;
  automatic.disabled=manual.disabled=!supported||state.phase==='installing';
  $('#update-mode').hidden=!state.supportsAutomatic;
  $('#update-check').disabled=!supported||active;
  $('#update-download').hidden=!supported||!['available','downloading','verifying'].includes(state.phase);
  $('#update-download').disabled=active;
  $('#update-install').hidden=!supported||!['ready','installing'].includes(state.phase);
  $('#update-install').disabled=active;
  $('#update-install').textContent=state.platform==='darwin'?'다운로드 파일 보기 / 설치 안내':'설치 화면 열기';
  $('#update-platform').textContent=!bridge?.updateAction?'브라우저에서는 설치 파일을 직접 받아 업데이트하세요.':!state.packaged?'개발 실행에서는 자동 업데이트를 지원하지 않습니다. 설치 파일을 직접 받을 수 있습니다.':state.platform==='darwin'?'Mac은 ZIP을 푼 뒤 문서를 저장하고 앱을 종료하여 응용 프로그램 폴더의 기존 앱을 교체하세요.':state.platform==='linux'?'Linux는 최신 설치 파일을 직접 받아 패키지 관리자로 설치하세요.':'설치를 시작하면 문서 복구 저장 후 앱이 종료되고 Windows 설치 화면이 열립니다. 설치를 마치면 앱이 다시 실행됩니다.';
 }
 function accept(next){if(!next||typeof next!=='object'||Number(next.revision)<state.revision)return;state={...state,...next};render();}
 function failed(error){accept({phase:'error',error:error.message||String(error),revision:state.revision});}
 async function refresh(){if(bridge?.getUpdateState)try{accept(await bridge.getUpdateState());}catch(error){failed(error);}}
 async function run(command,value){
  if(!bridge?.updateAction){if(command==='manual')window.open(RELEASE_PAGE,'_blank','noopener,noreferrer');return;}
  try{accept(await bridge.updateAction(command,value));}catch(error){failed(error);}
 }
 async function open(mode='status'){
  if(!dialog?.isConnected){
   dialog=modal('업데이트',`<div class="update-center">
    <dl class="update-versions"><div><dt>현재 버전</dt><dd id="update-current"></dd></div><div><dt>새 버전</dt><dd id="update-latest"></dd></div></dl>
    <div class="update-state" role="status" aria-live="polite"><strong id="update-phase"></strong><progress id="update-progress" max="100" aria-label="업데이트 진행률"></progress><span id="update-transfer"></span></div>
    <p id="update-error" class="update-error" role="alert" hidden></p><small id="update-checked"></small>
    <fieldset id="update-mode"><legend>업데이트 방식</legend><label><input id="update-automatic" name="update-mode" type="radio" value="automatic"> 자동으로 확인하고 다운로드</label><label><input id="update-manual" name="update-mode" type="radio" value="manual"> 수동으로만 확인</label></fieldset>
    <div class="update-buttons"><button id="update-check" class="primary">업데이트 확인</button><button id="update-download" class="primary" hidden>새 버전 다운로드</button><button id="update-install" class="primary" hidden>설치 화면 열기</button></div>
    <section id="update-manual-help"><h3>수동 업데이트</h3><p id="update-platform"></p><button id="update-manual-download">설치 파일 직접 받기</button></section>
    <p class="update-note">이 창을 닫아도 다운로드는 계속됩니다. 아래 상태 표시줄에서 진행 상황을 다시 열 수 있습니다.</p></div>`);
   dialog.querySelector('.modal').classList.add('update-modal');
   for(const [id,command] of [['check','check'],['download','download'],['install','install'],['manual-download','manual']])dialog.querySelector('#update-'+id).addEventListener('click',()=>{void run(command);});
   for(const id of ['automatic','manual'])dialog.querySelector('#update-'+id).addEventListener('change',()=>{void run('automatic',id==='automatic');});
   dialog.addEventListener('keydown',event=>{if(event.key==='Escape')onClose();});
   dialog.querySelector('[data-close]').addEventListener('click',onClose);
  }
  render();await refresh();render();
  if(!dialog?.isConnected)return;
  if(mode==='manual')dialog.querySelector('#update-manual-download').focus();
  else dialog.querySelector('[data-close]').focus();
  if(mode==='check'&&!ACTIVE.has(state.phase)&&state.packaged&&state.supportsAutomatic)void run('check');
 }
 bridge?.onUpdateState?.(accept);void refresh();
 return {open};
}
