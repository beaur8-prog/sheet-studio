import {androidHost} from './android-files.mjs';

export function createAndroidFullscreen({document:doc=document,window:win=window,onChange=()=>{}}={}){
  let native=false;
  const installedFullscreen=()=>!!win.matchMedia?.('(display-mode: fullscreen)').matches;
  const active=()=>native||!!doc.fullscreenElement||installedFullscreen();
  const canToggle=()=>!!androidHost(win)?.setFullscreen||!!doc.fullscreenElement||!installedFullscreen();
  const changed=()=>onChange(active());
  doc.addEventListener('fullscreenchange',changed);
  async function setNative(enabled){
    const host=androidHost(win);if(!host?.setFullscreen)return false;
    const result=await host.setFullscreen({enabled,hideStatusBar:enabled,hideNavigationBar:false,transientBarsBySwipe:true});
    if(result?.enabled!==enabled)throw Error('Android 전체화면 적용을 확인하지 못했습니다.');
    native=enabled;changed();return true;
  }
  return {
    active,
    canToggle,
    // Native hosts can hide the status bar on launch. Browser fullscreen must
    // instead be entered from the user's button press, with user activation.
    async initialize(){await setNative(true);changed();},
    async toggle(){
      const next=!active();if(await setNative(next))return;
      if(next){if(!doc.documentElement.requestFullscreen)throw Error('이 환경에서는 시스템 전체화면을 지원하지 않습니다.');await doc.documentElement.requestFullscreen({navigationUI:'hide'});}
      else if(doc.fullscreenElement)await doc.exitFullscreen();
      changed();
    },
  };
}
