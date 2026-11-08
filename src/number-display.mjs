import {effectiveNumberFormat,selectNumberSection,numberFormatKind} from './number-formats.mjs';

export const numericDisplayFormat=(value,style)=>numberFormatKind(selectNumberSection(effectiveNumberFormat(style),value).section?.code);

// Presentation only: never feed shortened numbers back into calculation, copy or save.
// `measure` uses the same font and zoom as the cell, without reading cell DOM geometry.
export function fitNumberDisplay(value,display,style,width,measure){
  if(typeof value!=='number'||!Number.isFinite(value)||!display)return display;
  const kind=numericDisplayFormat(value,style);
  if(kind==='text')return display;
  // Numeric cells use nowrap; literal newlines/tabs collapse to spaces in CSS.
  const fits=text=>measure(text.replace(/[\t\n\r\f ]+/g,' '))<=width;
  const hashes=()=> '#'.repeat(Math.max(1,Math.min(512,Math.floor(Math.max(0,width)/Math.max(1,measure('#'))))));
  if(/^#+$/.test(display))return hashes();
  if(kind!=='general')return fits(display)?display:hashes();
  const magnitude=Math.abs(value),scientific=magnitude>=1e11||/[eE]/.test(display);
  if(!scientific&&fits(display))return display;
  // General first drops fractional places. Do not round a nonzero fraction to zero.
  if(!scientific&&display.includes('.')){
    const decimals=Math.min(12,display.split('.')[1].length);
    for(let n=decimals-1;n>=0;n--){const rounded=Number(value.toFixed(n)),text=String(rounded);if(rounded!==0&&fits(text))return text;}
  }
  // Keep the sign, leading significant digits AND exponent together; partial
  // scientific notation can misrepresent magnitude just as badly as tail digits.
  for(let n=11;n>=0;n--){
    const [mantissa,exponent]=value.toExponential(n).split('e');
    const text=String(Number(mantissa))+'E'+(Number(exponent)<0?'-':'+')+String(Math.abs(Number(exponent))).padStart(2,'0');
    if(fits(text))return text;
  }
  return hashes();
}

export function createNumberMeasurer(context){
  const cache=new Map();
  return (text,font)=>{
    const key=font+'\0'+text;
    if(cache.has(key))return cache.get(key);
    context.font=font;
    const width=Math.max(0,...String(text).split('\n').map(line=>context.measureText(line).width));
    if(cache.size>=4096)cache.delete(cache.keys().next().value);
    cache.set(key,width);return width;
  };
}
