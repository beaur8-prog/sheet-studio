import './android.css';
import {createAndroidFullscreen} from './android-fullscreen.mjs';
import {androidInsets,androidViewportFrame} from './android-viewport.mjs';

const escape = value => String(value ?? '').replace(/[&<>"']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
const button = (id, label, cls = '') => `<button type="button" data-mobile="${id}" class="${cls}">${label}</button>`;

// Mounted only after Android runtime selection; all worksheet mutations stay
// behind the editor's normal commands, validation and undo transactions.
export function installAndroidUI(api) {
  const {modal, toast} = api;
  const root = document.querySelector('#app'), shell = document.querySelector('#grid-shell');
  const header = document.createElement('header'), summary = document.createElement('div'), nav = document.createElement('nav'), tools = document.createElement('div');
  header.className = 'android-header';
  header.innerHTML = button('open','열기') + '<button class="android-title" data-mobile="info"><strong></strong><small></small></button>' + button('address','A1','android-header-address') + button('fullscreen','전체화면') + button('more','더보기');
  summary.className = 'android-selection';
  summary.setAttribute('role','status');summary.setAttribute('aria-live','polite');
  summary.innerHTML = '<button class="android-address" data-mobile="address" aria-label="셀 주소 또는 범위로 이동"></button><button data-mobile="detail" aria-label="선택 셀 상세 보기"></button>';
  tools.className = 'android-tools';tools.setAttribute('aria-label','편집 도구');
  tools.innerHTML = [['aggregate-sum','Σ 합계'],['aggregate-average','평균'],['function','함수'],['edit-cell','셀 입력'],['copy','복사'],['paste','붙여넣기'],['cell-format','셀 서식'],['text-color','글자색'],['fill-color','배경색'],['borders','테두리'],['column-filter','이 열 필터'],['validation-list','입력 목록'],['undo','실행 취소'],['redo','다시 실행'],['save','사본 저장']].map(([id,label])=>button(id,label)).join('');
  nav.className = 'android-nav';nav.setAttribute('aria-label','주요 명령');
  nav.innerHTML = button('find','찾기') + button('sheets','시트') + button('view','보기 설정') + button('mode','편집','android-primary');
  root.prepend(header,summary,tools);root.append(nav);
  const formulaBar=document.querySelector('.formula-bar'),entryTools=document.createElement('div'),entryAddress=document.createElement('span');
  entryTools.className='android-entry-tools';entryTools.setAttribute('aria-label','입력과 수식 바로가기');
  entryTools.innerHTML=button('aggregate-sum','Σ 합계')+button('aggregate-average','평균')+button('function','함수')+button('keyboard','123')+button('next-cell','다음 ↓');
  entryAddress.className='android-entry-address';formulaBar.prepend(entryAddress);formulaBar.append(entryTools);
  const formula=document.querySelector('#formula');formula.setAttribute('enterkeyhint','enter');formula.setAttribute('autocomplete','off');formula.setAttribute('autocorrect','off');formula.setAttribute('autocapitalize','off');
  const done=formulaBar.querySelector('[data-action=commit-edit]'),cancel=formulaBar.querySelector('[data-action=cancel-edit]');
  done.textContent='완료';done.setAttribute('aria-label','입력 완료');cancel.textContent='취소';
  const manifest=document.createElement('link');manifest.rel='manifest';manifest.href=new URL('android.webmanifest',document.baseURI).href;document.head.append(manifest);
  const viewportMeta=document.querySelector('meta[name=viewport]');if(viewportMeta&&!viewportMeta.content.includes('viewport-fit'))viewportMeta.content+=', viewport-fit=cover';
  let hostInsets=null,viewportFrame=0;
  function safeAreas(active){
    for(const edge of ['top','right','bottom','left']){
      const bar=active?0:hostInsets?.bars[edge]||0,cutout=hostInsets?.cutout?.[edge];
      document.body.style.setProperty('--android-safe-'+edge,`max(${bar}px, ${cutout===undefined?'env(safe-area-inset-'+edge+', 0px)':cutout+'px'})`);
    }
  }
  const fullscreen=createAndroidFullscreen({onChange:active=>{
    document.body.classList.toggle('android-fullscreen',active);safeAreas(active);
    header.querySelector('[data-mobile=fullscreen]').textContent=active?'화면 복원':'전체화면';
    header.querySelector('[data-mobile=fullscreen]').setAttribute('aria-pressed',String(active));
    header.querySelector('[data-mobile=fullscreen]').hidden=!fullscreen.canToggle();
    scheduleViewport();
  }});
  const handles = ['start','end'].map(edge=>{
    const el=document.createElement('button');el.type='button';el.className='android-handle';el.dataset.rangeHandle=edge;
    el.setAttribute('aria-label',edge==='start'?'선택 시작 조절 · 주소 입력으로도 선택할 수 있습니다':'선택 끝 조절 · 주소 입력으로도 선택할 수 있습니다');
    shell.append(el);return el;
  });
  let touches=new Map(), candidate=null, range=null, rangeFrame=0, lastTap=null, lastTouch=0, rangeArmed=false, suppressClick=false;
  const run = operation => Promise.resolve().then(operation).catch(error=>toast(error.message));
  function refresh() {
    const state=api.state(),view=state.mode==='view';
    document.body.classList.toggle('android-view',view);document.body.classList.toggle('android-edit',!view);
    document.body.classList.toggle('android-cell-editing',state.cellEditing);
    entryAddress.textContent=state.editAddress||state.rangeLabel;
    header.querySelector('[data-mobile=address]').textContent=state.rangeLabel;
    entryTools.querySelector('[data-mobile=keyboard]').textContent=state.inputMode==='decimal'?'가/ABC':'123';
    entryTools.querySelector('[data-mobile=keyboard]').setAttribute('aria-label',state.inputMode==='decimal'?'문자 키보드로 전환':'숫자 키보드로 전환');
    header.querySelector('strong').textContent=state.name;
    header.querySelector('small').textContent=(view?'보기 모드':'편집 모드')+(state.dirty?' · 저장 전 변경 있음':'');
    nav.querySelector('[data-mobile=mode]').textContent=view?'편집':'보기로';
    tools.hidden=view||state.cellEditing;
    summary.hidden=state.cellEditing;
    const address=summary.querySelector('[data-mobile=address]'),detail=summary.querySelector('[data-mobile=detail]'),text=state.display||'빈 셀 · 상세 보기';
    if(address.textContent!==state.rangeLabel)address.textContent=state.rangeLabel;
    if(detail.textContent!==text)detail.textContent=text;
    summary.querySelector('[data-mobile=detail]').title=state.display;
    tools.querySelector('[data-mobile=undo]').disabled=!state.canUndo;
    tools.querySelector('[data-mobile=redo]').disabled=!state.canRedo;
    const rects=api.selectionRects(),first=rects[0],last=rects.at(-1);
    handles.forEach((el,i)=>{
      const q=i?last:first,selection=api.selection(),multiple=selection.r!==selection.r2||selection.c!==selection.c2;
      el.hidden=!q||state.cellEditing||state.chartSheet||!multiple&&!rangeArmed;
      if(q){el.style.left=Math.max(24,Math.min(shell.clientWidth-24,q.x+(i?q.w:0)))+'px';el.style.top=Math.max(24,Math.min(shell.clientHeight-24,q.y+(i?q.h:0)))+'px';}
    });
  }
  function detail() {
    const s=api.state(),info=api.cellDetail();
    modal(s.rangeLabel+' · 셀 상세',`<div class="android-detail"><p>표시 값</p><div class="android-value">${escape(info.display||'(빈 셀)')}</div>${info.formula?'<p>수식</p><div class="android-value">'+escape(info.formula)+'</div>':''}${info.note?'<p>메모</p><div class="android-value">'+escape(info.note)+'</div>':''}<p class="android-help">${escape(info.notice||'값을 길게 눌러 텍스트를 선택하거나 범위 복사를 사용하세요.')}</p></div>`,[
      {label:'범위 복사',action:el=>{el.remove();run(()=>api.command('copy'));}},
      ...(s.mode==='edit'?[{label:'셀 입력',primary:true,action:el=>{el.remove();api.editCell();}}]:[]),
    ]);
  }
  function address() {
    const el=modal('셀·범위로 이동',`<label>셀 주소·범위 또는 정의된 이름<input id="android-address" value="${escape(api.state().rangeLabel)}" autocomplete="off" autocapitalize="characters" spellcheck="false"></label><p class="android-help">예: B5, A1:D12. 작은 손잡이를 끌지 않고 범위를 선택할 수 있습니다.</p>`,[{label:'이동',primary:true,action:async dialog=>{if(await api.goTo(dialog.querySelector('input').value))dialog.remove();}}]);
    el.querySelector('input').select();
  }
  function sheets() {
    const s=api.state();
    const el=modal('시트',`<div class="android-list">${api.sheets().map(sheet=>`<button data-mobile-sheet="${escape(sheet.id)}" aria-current="${sheet.id===s.sheetId?'true':'false'}">${escape(sheet.name)}${sheet.id===s.sheetId?' · 현재 시트':''}</button>`).join('')}</div>`,s.mode==='edit'?[{label:'새 시트',action:dialog=>{dialog.remove();api.command('add-sheet');}}]:[]);
    el.addEventListener('click',event=>{const b=event.target.closest('[data-mobile-sheet]');if(b){if(api.switchSheet(b.dataset.mobileSheet)!==false)el.remove();}});
  }
  function view() {
    const s=api.state();
    const el=modal('보기 설정',`<div class="android-list">${button('zoom-out','− 축소')}${button('zoom-in','＋ 확대')}${button('zoom-reset','100%')}${button('fit','선택 영역 너비 맞춤')}${button('toggle-grid','눈금선 표시 전환')}${button('show-formulas','수식 표시 전환')}${button('freeze-first-row','첫 행 고정')}${button('freeze-first-col','첫 열 고정')}${button('freeze-none','틀 고정 해제')}${button('theme','화면 테마 전환')}</div><p class="android-help">시트 배율 ${s.zoom}%. 손가락으로 밀어 스크롤하고 확대 버튼으로 글자를 키울 수 있습니다.</p>${s.mode==='edit'?'<label class="worksheet-check"><input id="android-hardware" type="checkbox" '+(s.hardware?'checked':'')+'>외부 키보드 바로 입력</label><p class="android-help">외부 키보드에서 첫 한글부터 받도록 입력 대상을 준비합니다. 터치 입력은 셀 입력 버튼을 사용하세요.</p>':''}`);
    el.querySelector('#android-hardware')?.addEventListener('change',event=>api.hardware(event.target.checked));
  }
  function more() {
    const s=api.state();
    modal('문서 메뉴',`<div class="android-list">${button('new','새 통합 문서')}${button('recover','미저장 문서 복구')}${button('find','찾기')}${button('sheets','시트')}${button('view','보기 설정')}${button('mode',s.mode==='edit'?'읽기 전용 보기로':'편집 모드로')}${button('share','공유')}${button('address','셀·범위로 이동')}${button('detail','선택 셀 상세')}${button('charts','차트 목록·크게 보기')}${s.mode==='edit'?button('save','XLSX 사본 저장')+button('save-json','앱 문서 사본 저장')+button('save-csv','CSV 사본 저장')+button('ribbon','태블릿 리본 펼치기/접기'):''}${button('compatibility','파일 호환성')}${button('shortcuts','키보드 단축키')}${button('info','문서 정보')}</div>`);
  }
  function context() {
    const edit=api.state().mode==='edit';
    modal('선택 영역 · '+escape(api.state().rangeLabel),`<div class="android-list">${button('copy','복사')}${button('detail','셀 상세')}${button('select-range','손잡이로 범위 선택')}${button('address','범위 주소로 선택')}${edit?button('edit-cell','셀 입력')+button('paste','붙여넣기')+button('cell-format','셀 서식')+button('text-color','글자색')+button('borders','테두리'):button('mode','편집 모드로 전환')}</div>`);
  }
  const panels=new Set(['more','view','sheets','detail','address','info']);
  async function mobileCommand(id) {
    if(api.state().busy)return;
    if(!panels.has(id))document.querySelector('.backdrop')?.remove();
    switch(id) {
      case 'detail':return detail();case 'address':return address();case 'sheets':return sheets();case 'view':return view();case 'more':return more();
      case 'charts':{
        const charts=api.charts();if(!charts.length)return toast('현재 시트에 차트가 없습니다.');
        const el=modal('차트 목록','<div class="android-list">'+charts.map((chart,i)=>'<button data-mobile-chart="'+i+'">'+escape(chart.title||'차트 '+(i+1))+'</button>').join('')+'</div>');
        el.addEventListener('click',event=>{const target=event.target.closest('[data-mobile-chart]');if(target){el.remove();api.showChart(charts[+target.dataset.mobileChart].id);}});return;
      }
      case 'info':return modal('문서 정보',`<div class="android-detail">${escape(api.state().name)}<br>Sheet Studio ${escape(api.state().version)}<br>${escape(api.sourceName()||'새 통합 문서')}<p>이 기기의 복구 저장과 원본 파일 저장은 별개입니다. 사본 다운로드·공유 후에도 원본 파일은 바뀌지 않습니다.</p>${api.state().dirty?'<p>저장 전 변경이 있습니다. 사본 저장 또는 편집본 공유를 사용할 수 있습니다.</p>':''}</div>`);
      case 'mode':await api.changeMode();break;
      case 'edit-cell':api.editCell();break;
      case 'aggregate-sum':await api.aggregate('SUM');break;
      case 'aggregate-average':await api.aggregate('AVERAGE');break;
      case 'next-cell':api.nextCell();break;
      case 'keyboard':api.toggleKeyboard();break;
      case 'fullscreen':await fullscreen.toggle();break;
      case 'column-filter':await api.columnFilter();break;
      case 'select-range':rangeArmed=true;break;
      case 'fit':api.fit();break;
      case 'ribbon':document.body.classList.toggle('android-ribbon-open');break;
      case 'share':await api.share();break;
      default:await api.command(id);
    }
    refresh();
  }
  document.addEventListener('click',event=>{
    // Opening a panel on touch-up can retarget its synthetic click to the new
    // backdrop. Consume that one click before it dismisses the panel again.
    if(suppressClick){suppressClick=false;if(event.pointerType==='touch'||event.detail>0){event.preventDefault();event.stopImmediatePropagation();return;}}
    const target=event.target.closest('[data-mobile]');
    if(target){event.preventDefault();event.stopImmediatePropagation();run(()=>mobileCommand(target.dataset.mobile));return;}
    if(Date.now()-lastTouch<650&&event.target.closest('#grid-shell'))event.stopImmediatePropagation();
  },true);
  // Touch compatibility mousedown runs after pointerup. Its default focus can
  // steal the input we just opened, especially after a keyboard-sized resize.
  document.addEventListener('mousedown',event=>{
    if(event.sourceCapabilities?.firesTouchEvents!==false&&Date.now()-lastTouch<650&&event.target.closest('#grid-shell')){event.preventDefault();event.stopImmediatePropagation();}
  },true);
  document.addEventListener('contextmenu',event=>{
    if(event.target.closest('#grid-shell')){
      event.preventDefault();event.stopImmediatePropagation();
      candidate=null;suppressClick=true;
      const cell=event.target.closest('.cell');if(cell){const r=+cell.dataset.r,c=+cell.dataset.c,b=api.selection();if(r<Math.min(b.r,b.r2)||r>Math.max(b.r,b.r2)||c<Math.min(b.c,b.c2)||c>Math.max(b.c,b.c2))if(api.tapCell({r,c})===false)return;}
      context();
    }
  },true);
  function finishRange(cancel) {
    if(!range)return;
    const old=range;range=null;cancelAnimationFrame(rangeFrame);
    if(cancel)api.setSelection(old.original);
    if(shell.hasPointerCapture(old.id))shell.releasePointerCapture(old.id);
    refresh();
  }
  function dragFrame() {
    if(!range)return;
    api.scrollAt(range.point);const p=api.pickCell(range.point),b=range.original;
    api.setSelection(range.edge==='start'?{r:p.r,c:p.c,r2:b.r2,c2:b.c2}:{r:b.r,c:b.c,r2:p.r,c2:p.c});
    rangeFrame=requestAnimationFrame(dragFrame);
  }
  document.addEventListener('pointerdown',event=>{
    suppressClick=false;
    if(api.state().cellEditing&&event.target.closest('.android-entry-tools button'))event.preventDefault();
    const handle=event.target.closest('[data-range-handle]');
    if(handle){event.preventDefault();event.stopImmediatePropagation();if(api.state().busy||api.state().cellEditing)return;range={id:event.pointerId,edge:handle.dataset.rangeHandle,original:api.selection(),point:{clientX:event.clientX,clientY:event.clientY}};shell.setPointerCapture(event.pointerId);dragFrame();return;}
    if(event.pointerType!=='touch'||!event.target.closest('#grid-shell')||event.target.closest('input'))return;
    event.stopPropagation();lastTouch=Date.now();touches.set(event.pointerId,true);
    if(range){candidate=null;finishRange(true);return;}
    if(touches.size>1){candidate=null;finishRange(true);return;}
    const cell=event.target.closest('.cell'),chart=event.target.closest('[data-chart]'),col=event.target.closest('[data-col]'),row=event.target.closest('[data-row]');
    candidate={id:event.pointerId,x:event.clientX,y:event.clientY,time:Date.now(),cell:cell?{r:+cell.dataset.r,c:+cell.dataset.c}:null,chart:chart?.dataset.chart,col:col?.dataset.col,row:row?.dataset.row};
  },true);
  document.addEventListener('pointermove',event=>{
    if(range?.id===event.pointerId){event.preventDefault();event.stopPropagation();range.point={clientX:event.clientX,clientY:event.clientY};return;}
    if(candidate?.id===event.pointerId&&Math.hypot(event.clientX-candidate.x,event.clientY-candidate.y)>10)candidate=null;
  },{capture:true,passive:false});
  document.addEventListener('pointerup',event=>{
    if(range?.id===event.pointerId){finishRange(false);event.stopPropagation();return;}
    const tap=candidate;touches.delete(event.pointerId);candidate=null;
    if(!tap||tap.id!==event.pointerId||Date.now()-tap.time>600||api.state().busy)return;
    lastTouch=Date.now();
    if(tap.chart){suppressClick=true;api.showChart(tap.chart);return;}
    if(tap.cell){const double=lastTap&&Date.now()-lastTap.time<350&&lastTap.cell.r===tap.cell.r&&lastTap.cell.c===tap.cell.c;
      if(api.tapCell(tap.cell,{edit:api.state().mode==='edit'})===false)return;
      if(api.state().mode==='edit'){lastTap=null;rangeArmed=false;suppressClick=true;return;}
      lastTap={cell:tap.cell,time:Date.now()};if(double){lastTap=null;suppressClick=true;if(api.state().mode==='edit')api.editCell();else detail();}
    }else if(tap.col!==undefined)api.selectHeading('c',+tap.col);else if(tap.row!==undefined)api.selectHeading('r',+tap.row);
  },true);
  document.addEventListener('pointercancel',event=>{touches.delete(event.pointerId);candidate=null;if(range?.id===event.pointerId)finishRange(true);},true);
  shell.addEventListener('lostpointercapture',()=>finishRange(true));
  let layoutWidth=window.innerWidth,layoutHeight=window.innerHeight,previousWindowHeight=window.innerHeight,keyboardMayBeOpen=false;
  const textInput=node=>node?.tagName==='TEXTAREA'||node?.tagName==='INPUT'&&!['checkbox','radio','color','range','file','button','submit'].includes(node.type);
  document.addEventListener('focusin',event=>{if(textInput(event.target))keyboardMayBeOpen=true;});
  function viewport() {
    candidate=null;touches.clear();finishRange(true);
    const frame=androidViewportFrame(window);
    // Soft-keyboard height changes must not turn a portrait layout into a
    // landscape layout. A real width change, or idle resize, updates the base.
    const typing=api.state().cellEditing||textInput(document.activeElement)||keyboardMayBeOpen&&window.innerHeight<layoutHeight;
    if(!typing){layoutWidth=window.innerWidth;layoutHeight=window.innerHeight;keyboardMayBeOpen=false;}
    else if(window.innerWidth!==layoutWidth){
      const swapped=Math.abs(window.innerWidth-layoutHeight)<Math.max(32,layoutHeight*.15),obscured=Math.max(0,layoutHeight-previousWindowHeight);
      layoutHeight=swapped?Math.max(layoutWidth,window.innerHeight):window.innerHeight+obscured;layoutWidth=window.innerWidth;
    }
    previousWindowHeight=window.innerHeight;
    document.body.classList.toggle('android-landscape',layoutWidth>layoutHeight);
    for(const [name,value] of Object.entries(frame))document.body.style.setProperty('--android-'+(name==='top'||name==='left'?'viewport-':'')+name,value+'px');
    refresh();
  }
  function scheduleViewport(){cancelAnimationFrame(viewportFrame);viewportFrame=requestAnimationFrame(viewport);}
  window.addEventListener('sheet-studio-android-insets',event=>{
    hostInsets=androidInsets(event.detail);safeAreas(fullscreen.active());
    viewport();
  });
  window.addEventListener('sheet-studio-android-input',event=>{if(typeof event.detail?.hardwareKeyboard==='boolean')api.hardware(event.detail.hardwareKeyboard);});
  window.addEventListener('resize',scheduleViewport);window.visualViewport?.addEventListener('resize',scheduleViewport);
  window.visualViewport?.addEventListener('scroll',scheduleViewport);
  window.addEventListener('pageshow',scheduleViewport);window.screen.orientation?.addEventListener('change',scheduleViewport);
  window.addEventListener('blur',()=>{candidate=null;touches.clear();finishRange(true);});
  document.addEventListener('keydown',event=>{
    if(event.isComposing||event.keyCode===229||event.defaultPrevented)return;
    if(range&&event.key==='Escape'){event.preventDefault();event.stopImmediatePropagation();finishRange(true);return;}
    if(event.key==='Enter'&&event.target.id==='grid'&&api.state().mode==='view'){event.preventDefault();event.stopImmediatePropagation();detail();return;}
    if(['ArrowLeft','ArrowRight','ArrowUp','ArrowDown','Tab','F2'].includes(event.key)&&['grid','editor'].includes(event.target.id)&&api.state().mode==='edit'&&!api.state().cellEditing)api.hardware(true);
  },true);
  window.addEventListener('sheet-studio-android-back',event=>{
    if(api.state().busy){event.preventDefault();api.cancelBusy();return;}
    const el=document.querySelector('.backdrop');
    if(el){el.querySelector('[data-close],[data-function-close]')?.click();event.preventDefault();return;}
    if(document.querySelector('.filter-popup')){api.closeFilter();event.preventDefault();return;}
    if(api.state().cellEditing){api.cancelOrKeepDraft();event.preventDefault();return;}
    if(api.state().mode==='edit'){run(()=>api.changeMode());event.preventDefault();return;}
    if(api.state().dirty||api.sourceName()){run(()=>api.closeDocument());event.preventDefault();}
  });
  viewport();
  // Register inset listeners before a native host can synchronously reply.
  void fullscreen.initialize().catch(error=>toast(error.message));
  return {refresh, detail, context, cancelGesture:()=>{candidate=null;touches.clear();finishRange(true);}};
}
