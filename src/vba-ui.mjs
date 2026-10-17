import {buildNativePlan} from './native-plan.mjs';
import {readSpreadsheet} from './io.mjs';

export function createVbaUI({getWorkbook,commitEdit,replaceWorkbook,toast}){
  let panel=null,session=null,stopped=false;
  async function open(mode='interactive'){
    if(panel)return;
    commitEdit();const book=getWorkbook();
    if(!book.nativeSource){toast('먼저 기존 XLSM 통합 문서를 여세요.');return;}
    if(!window.desktop?.vbaStart){toast('VBA 실행과 XLSM 편집 저장은 Windows 데스크톱 앱에서 지원합니다.');return;}
    panel=document.createElement('div');panel.className='vba-overlay';
    panel.innerHTML=`<section class="vba-panel" role="dialog" aria-modal="true" aria-label="Excel VBA"><div class="vba-title"><h2>Excel VBA</h2><span>실제 Excel 엔진</span></div><p class="vba-description">기존 VBA 프로젝트와 Excel 개체를 유지한 작업 사본을 사용합니다. Excel의 보안 설정을 따르며 설정을 자동 변경하지 않습니다.</p><p class="vba-description">사본 폴더에서 실행되므로 <code>ThisWorkbook.Path</code>와 상대 경로가 달라집니다. 외부 파일·추가 기능이 필요한 매크로는 Excel에서 확인하세요.</p><label>매크로 이름 <span>(선택)</span><input id="vba-name" placeholder="Module1.MacroName" spellcheck="false"></label><div class="vba-status" role="status">Excel 설치를 확인하고 있습니다.</div><div class="vba-actions"><button data-vba="connect" class="primary" disabled>Excel 연결</button><button data-vba="run" disabled>이름으로 실행</button><button data-vba="dialog" disabled>Excel 매크로 목록</button><button data-vba="finish" disabled>결과 가져오기</button><button data-vba="cancel">닫기</button></div><p class="vba-footnote">연결 후 Excel에서 직접 실행하거나 Alt+F8을 사용할 수 있습니다. 실행을 허용한 통합 문서의 열기 이벤트도 Excel에서 동작할 수 있습니다.</p></section>`;
    document.body.append(panel);stopped=false;session=null;const container=panel;let commandError='';const active=()=>!stopped&&panel===container;
    const state=panel.querySelector('.vba-status'),buttons=Object.fromEntries([...panel.querySelectorAll('[data-vba]')].map(b=>[b.dataset.vba,b]));
    const set=(status,message)=>{state.textContent=message;for(const key of ['run','dialog','finish'])buttons[key].disabled=status!=='ready'||mode==='save';buttons.connect.disabled=status!=='idle';buttons.cancel.textContent=session?'중단·닫기':'닫기';};
    const close=()=>{if(panel!==container)return;panel?.remove();panel=null;session=null;stopped=true;};
    async function finish(){const file=await window.desktop.vbaResult(session);if(!active())return;const loaded=await readSpreadsheet(new Uint8Array(file.data),file.name);if(!active())return;loaded.nativeSource.originToken=file.originToken;loaded.name=book.name;
      replaceWorkbook(loaded);
      if(mode==='save'){const filename=book.name.replace(/[<>:"/\\|?*]/g,'_')+'.xlsm';const saved=await window.desktop.saveFile({name:filename,data:file.data,originToken:file.originToken});if(saved)toast('VBA를 유지한 XLSM 파일을 저장했습니다.');}
      else toast('Excel 매크로 실행 결과를 가져왔습니다.');close();}
    async function poll(){while(active()&&session){const s=await window.desktop.vbaStatus(session);if(!active())return;set(s.state,commandError||s.message);if(s.state==='completed'){await finish();return;}if(s.state==='error'||s.state==='cancelled'){buttons.cancel.textContent='닫기';return;}await new Promise(resolve=>setTimeout(resolve,400));}}
    async function start(){set('starting','현재 편집 내용을 Excel 작업 사본에 반영하고 있습니다.');try{const started=await window.desktop.vbaStart({mode,source:{base64:book.nativeSource.base64,filename:book.nativeSource.filename,originToken:book.nativeSource.originToken},plan:buildNativePlan(book)});if(!active()){await window.desktop.vbaCancel(started.id);return;}session=started.id;await poll();}catch(error){if(active())set('error',error.message);}}
    panel.addEventListener('click',async event=>{const button=event.target.closest('[data-vba]');if(!button||button.disabled)return;const command=button.dataset.vba;
      try{if(command==='cancel'){stopped=true;if(session)await window.desktop.vbaCancel(session);close();return;}if(command==='connect'){await start();return;}const name=panel.querySelector('#vba-name').value.trim();if(command==='run'&&!name){state.textContent='매크로 이름을 입력하거나 Excel 매크로 목록을 사용하세요.';return;}commandError='';set('running',command==='finish'?'Excel 결과를 가져오고 있습니다.':'Excel 작업을 요청했습니다.');await window.desktop.vbaCommand(session,{type:command,name});}
      catch(error){commandError=error.message;if(active())set('ready',error.message);}
    });
    try{const installed=await window.desktop.vbaProbe();if(!active())return;if(!installed.available){set('unavailable',installed.reason);return;}set('idle','설치된 Excel에 연결할 수 있습니다.');if(mode==='save'){panel.querySelector('label').hidden=true;panel.querySelector('.vba-footnote').textContent='저장은 매크로와 이벤트를 비활성화한 Excel 사본에서 처리합니다.';await start();}}
    catch(error){if(active())set('error',error.message);}
  }
  return {open:()=>open('interactive'),save:()=>open('save')};
}
