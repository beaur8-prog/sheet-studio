import './viewer.css';

export const modeIcon=mode=>`<svg viewBox="0 0 24 24" aria-hidden="true"><path d="${mode==='view'?'M2 12s4-7 10-7 10 7 10 7-4 7-10 7S2 12 2 12 M15 12a3 3 0 1 1-6 0 3 3 0 0 1 6 0':'m4 16 12-12 4 4-12 12-5 1z M14 6l4 4'}"/></svg>`;

// Overlay controls occupy no worksheet row. Touch/pointer navigation reveals
// them; keyboard focus keeps them visible even when the idle timer expires.
export function createViewerRail({onEdit,onMenu,mobile=false}){
  const rail=document.createElement('nav');rail.className='viewer-rail';rail.hidden=true;rail.setAttribute('aria-label','뷰어 도구');
  rail.innerHTML=`<button type="button" data-viewer-edit title="편집 모드로 전환 (Esc)">${modeIcon('edit')}<span>편집</span></button><button type="button" data-viewer-menu aria-label="뷰어 메뉴" title="찾기·시트·확대">⋯</button><span class="viewer-escape-hint">Esc · 편집</span>`;
  document.body.append(rail);let timer;
  if(mobile){rail.querySelector('[data-viewer-edit]').dataset.mobile='mode';rail.querySelector('[data-viewer-menu]').dataset.mobile='more';}
  const reveal=()=>{if(rail.hidden)return;rail.classList.add('viewer-controls-visible');clearTimeout(timer);timer=setTimeout(()=>rail.classList.remove('viewer-controls-visible'),2800);};
  rail.querySelector('[data-viewer-edit]').addEventListener('click',()=>void onEdit());
  rail.querySelector('[data-viewer-menu]').addEventListener('click',()=>void onMenu());
  for(const type of ['pointerdown','pointermove','keydown','focusin'])document.addEventListener(type,reveal,{capture:true,passive:true});
  return {refresh(view){const changed=rail.hidden===view;rail.hidden=!view;if(changed&&view)reveal();if(!view)clearTimeout(timer);},reveal,setEditable(value){const b=rail.querySelector('[data-viewer-edit]');b.disabled=!value;b.querySelector('span').textContent=value?'편집':'보기 전용';}};
}

