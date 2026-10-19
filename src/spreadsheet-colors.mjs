import {all,child,children,attr,xml} from './chart-xml.mjs';

export function normalizeColor(value){
 if(typeof value!=='string')return null;
 const hex=value.trim().replace(/^#/,'');
 return /^[\da-f]{3}$/i.test(hex)?'#'+[...hex].map(c=>c+c).join('').toLowerCase():/^[\da-f]{6}$/i.test(hex)?'#'+hex.toLowerCase():null;
}
const defaultTheme=['ffffff','000000','e7e6e6','44546a','4472c4','ed7d31','a5a5a5','ffc000','5b9bd5','70ad47','0563c1','954f72'];
const indexed=('000000 ffffff ff0000 00ff00 0000ff ffff00 ff00ff 00ffff 000000 ffffff ff0000 00ff00 0000ff ffff00 ff00ff 00ffff 800000 008000 000080 808000 800080 008080 c0c0c0 808080 9999ff 993366 ffffcc ccffff 660066 ff8080 0066cc ccccff 000080 ff00ff ffff00 00ffff 800080 800000 008080 0000ff 00ccff ccffff ccffcc ffff99 99ccff ff99cc cc99ff ffcc99 3366ff 33cccc 99cc00 ffcc00 ff9900 ff6600 666699 969696 003366 339966 003300 333300 993300 993366 333399 333333').split(' ');
function tinted(hex,tint){
 if(!Number.isFinite(tint)||!tint||Math.abs(tint)>1)return '#'+hex;
 const rgb=[0,2,4].map(i=>parseInt(hex.slice(i,i+2),16)/255),hi=Math.max(...rgb),lo=Math.min(...rgb),d=hi-lo,l=(hi+lo)/2,s=d/(1-Math.abs(2*l-1))||0;
 let h=d?(hi===rgb[0]?(rgb[1]-rgb[2])/d%6:hi===rgb[1]?(rgb[2]-rgb[0])/d+2:(rgb[0]-rgb[1])/d+4):0;if(h<0)h+=6;
 const light=tint<0?l*(1+tint):l*(1-tint)+tint,c=(1-Math.abs(2*light-1))*s,x=c*(1-Math.abs(h%2-1)),m=light-c/2;
 const lanes=h<1?[c,x,0]:h<2?[x,c,0]:h<3?[0,c,x]:h<4?[0,x,c]:h<5?[x,0,c]:[c,0,x];
 return '#'+lanes.map(v=>Math.round((v+m)*255).toString(16).padStart(2,'0')).join('');
}
export function createColorResolver(themeXML,stylesXML){
 const theme=[...defaultTheme],palette=[...indexed];
 if(themeXML){const scheme=all(xml(themeXML),'clrScheme')[0];['lt1','dk1','lt2','dk2','accent1','accent2','accent3','accent4','accent5','accent6','hlink','folHlink'].forEach((key,i)=>{const color=children(child(scheme,key))[0],value=attr(color,'lastClr')||attr(color,'val');if(/^[\da-f]{6}$/i.test(value))theme[i]=value;});}
 if(stylesXML){const custom=all(xml(stylesXML),'indexedColors')[0];children(custom,'rgbColor').forEach((n,i)=>{const value=attr(n,'rgb');if(/^[\da-f]{8}$/i.test(value))palette[i]=value.slice(-6);});}
 return color=>{
  if(!color||color.auto===true||color.auto===1||color.auto==='1')return undefined;
  const value=color.argb||color.rgb,hex=/^[\da-f]{6}([\da-f]{2})?$/i.test(value||'')?value.slice(-6):Number.isInteger(color.theme)?theme[color.theme]:Number.isInteger(color.indexed)?palette[color.indexed]:undefined;
  return hex?tinted(hex,Number(color.tint)):undefined;
 };
}
export function xmlColor(node,resolve){
 if(!node)return undefined;
 return normalizeColor(resolve({rgb:attr(node,'rgb'),...(node.hasAttribute('theme')?{theme:Number(attr(node,'theme'))}:{}),...(node.hasAttribute('indexed')?{indexed:Number(attr(node,'indexed'))}:{}),auto:['1','true'].includes(attr(node,'auto')),tint:Number(attr(node,'tint'))}))||undefined;
}
