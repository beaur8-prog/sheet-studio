import {fitNumberDisplay,createNumberMeasurer,numericDisplayFormat} from './number-display.mjs';
export const MAX_ANDROID_PRINT_BYTES=32*1024*1024;

// Keep the prepared worksheet pages, never the app chrome, in the native job.
// Native scales these page wrappers to the printer's actual paper and margins.
export function androidPrintPayload(document,result){
  const paper=result.paper;
  if(!paper||![paper.width,paper.height].every(n=>Number.isFinite(n)&&n>=50&&n<=1500))throw Error('인쇄 용지 크기를 확인할 수 없습니다.');
  const clone=document.cloneNode(true),pages=[...clone.body.children].filter(el=>el.classList.contains('print-page'));
  if(!pages.length||pages.length!==result.pageCount||pages.length>500)throw Error('인쇄 페이지를 준비하지 못했습니다.');
  clone.querySelectorAll('script,iframe,object,embed,base,link,meta[http-equiv="refresh"]').forEach(el=>el.remove());
  for(const el of clone.querySelectorAll('*')){
    for(const attr of [...el.attributes]){
      if(/^on/i.test(attr.name))el.removeAttribute(attr.name);
      if(['src','href','xlink:href'].includes(attr.name)&&attr.value&&!/^(?:data:image\/|#)/i.test(attr.value)){
        if(el.tagName.toLowerCase()==='a')el.removeAttribute(attr.name);
        else throw Error('인쇄할 이미지를 읽지 못했습니다. 이미지 포함을 끄고 다시 시도하세요.');
      }
    }
  }
  const measure=createNumberMeasurer(document.createElement('canvas').getContext('2d'));
  const originals=[...document.querySelectorAll('[data-print-number]')],copies=[...clone.querySelectorAll('[data-print-number]')];
  for(let i=0;i<originals.length;i++){
    const source=originals[i].querySelector('.print-cell'),target=copies[i].querySelector('.print-cell'),value=Number(originals[i].dataset.printNumber),format=JSON.parse(originals[i].dataset.printNumberFormat);
    if(source&&target&&numericDisplayFormat(value,format)!=='text'){
      const style=document.defaultView.getComputedStyle(source),font=`${style.fontStyle} ${style.fontWeight} ${style.fontSize} ${style.fontFamily}`,width=source.clientWidth-parseFloat(style.paddingLeft)-parseFloat(style.paddingRight);
      target.textContent=fitNumberDisplay(value,source.textContent,format,Math.max(0,width),text=>measure(text,font));target.style.whiteSpace='nowrap';
    }
    copies[i].removeAttribute('data-print-number');copies[i].removeAttribute('data-print-number-format');
  }
  for(const page of pages){const wrapper=clone.createElement('div');wrapper.className='android-sheet-print-page';page.replaceWith(wrapper);wrapper.append(page);}
  clone.documentElement.dataset.sheetPrintDocument='1';
  const html='<!doctype html>'+clone.documentElement.outerHTML;
  if(new TextEncoder().encode(html).byteLength>MAX_ANDROID_PRINT_BYTES)throw Error('인쇄 문서가 32 MiB를 넘습니다. 범위를 줄이거나 이미지 포함을 끄세요.');
  return {html,title:clone.title||'Sheet Studio',paper:{...paper},pageCount:pages.length};
}
