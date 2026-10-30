import {child,children,attr,C} from './chart-xml.mjs';
import {validateNumberFormat} from './number-formats.mjs';
export const AXIS_KEYS=['min','max','majorUnit','logBase','reverse','numberFormat','gridlines'];
export function axisOptionsError(options){
 if(!options||typeof options!=='object'||Array.isArray(options))return '축 설정이 올바르지 않습니다.';
 for(const key of ['min','max','majorUnit','logBase'])if(options[key]!==undefined&&(typeof options[key]!=='number'||!Number.isFinite(options[key])))return '축 범위와 간격은 유한한 숫자로 입력하세요.';
 if(options.min!==undefined&&options.max!==undefined&&options.min>=options.max)return '축 최댓값은 최솟값보다 커야 합니다.';
 if(options.majorUnit!==undefined&&options.majorUnit<=0)return '주 눈금 간격은 0보다 커야 합니다.';
 if(options.logBase!==undefined&&(options.logBase<2||options.logBase>1000||options.min!==undefined&&options.min<=0||options.max!==undefined&&options.max<=0))return '로그 축의 밑은 2~1000이며 범위는 0보다 커야 합니다.';
 for(const key of ['reverse','gridlines'])if(options[key]!==undefined&&typeof options[key]!=='boolean')return '축 표시 설정이 올바르지 않습니다.';
 if(options.numberFormat!==undefined&&(typeof options.numberFormat!=='string'||options.numberFormat.length>512||/[\u0000-\u001f]/.test(options.numberFormat)))return '축 표시 형식은 512자 이하로 입력하세요.';
 if(options.numberFormat){try{validateNumberFormat(options.numberFormat);}catch(error){return error.message;}}
 return null;
}
const round=value=>Number.isFinite(value)?Number(value.toPrecision(14)):value;
function niceUnit(span){const exp=Math.max(Number.MIN_VALUE,10**Math.floor(Math.log10(span)-Math.log10(5))),fraction=span/5/exp;return ([1,2,2.5,5,10].find(n=>n>=fraction)||10)*exp;}
export function axisScale(values,options={}, {percent=false}={}){
 const factor=percent?100:1,finite=values.filter(Number.isFinite),positive=finite.filter(n=>n>0),log=options.logBase,logarithm=v=>Math.log(v)/Math.log(log),low=options.min===undefined?undefined:options.min*factor,high=options.max===undefined?undefined:options.max*factor;
 let min,max,ticks=[],unit;
 if(log){const data=positive.length?positive:[1,log];min=low??log**Math.floor(logarithm(data.reduce((a,b)=>Math.min(a,b),Infinity)));max=high??log**Math.ceil(logarithm(data.reduce((a,b)=>Math.max(a,b),-Infinity)));if(min===max){if(high===undefined)max*=log;else if(low===undefined)min/=log;}if(min>=max){if(high===undefined)max=min*log;else min=max/log;}
  const a=Math.ceil(logarithm(min)-1e-12),b=Math.floor(logarithm(max)+1e-12),stride=Math.max(1,Math.ceil((b-a+1)/20));for(let p=a;p<=b&&ticks.length<21;p+=stride)ticks.push(round(log**p));
 }else{
  min=low??finite.reduce((a,b)=>Math.min(a,b),0);max=high??finite.reduce((a,b)=>Math.max(a,b),0);if(min===max){if(high===undefined)max=min+(Math.abs(min)||1);else if(low===undefined)min=max-(Math.abs(max)||1);}
  if(min>=max){if(high===undefined)max=min+(Math.abs(min)*.1||1);else min=max-(Math.abs(max)*.1||1);}
  const span=max-min;unit=options.majorUnit===undefined?niceUnit(Number.isFinite(span)?span:Math.max(Math.abs(min),Math.abs(max))):options.majorUnit*factor;if(!Number.isFinite(unit)||unit<=0)unit=1;
  if(low===undefined)min=round(Math.floor(min/unit)*unit);if(high===undefined)max=round(Math.ceil(max/unit)*unit);
  const count=Math.floor((max-min)/unit+1e-10),stride=Math.max(1,Math.ceil(count/20)),start=Math.ceil(min/unit-1e-10)*unit;
  if(Number.isFinite(stride)&&Number.isFinite(start))for(let i=0;i<=20;i++){const value=round(start+i*stride*unit);if(value>max+Math.abs(unit)*1e-8)break;if(value>=min-Math.abs(unit)*1e-8)ticks.push(value);}
 }
 if(!Number.isFinite(min)||min<=0&&log)min=positive.reduce((a,b)=>Math.min(a,b),1);if(!Number.isFinite(max))max=positive.reduce((a,b)=>Math.max(a,b),1);if(min===max)max=min===Number.MAX_VALUE?min:Math.min(Number.MAX_VALUE,min*(log||2));
 ticks=ticks.filter(Number.isFinite);if(!ticks.length)ticks=[min,max];const magnitude=Math.max(Math.abs(min),Math.abs(max),1),ratio=value=>log?(logarithm(value)-logarithm(min))/(logarithm(max)-logarithm(min)):(value/magnitude-min/magnitude)/(max/magnitude-min/magnitude);
 return {min,max,ticks,unit,log,ratio:value=>{if(!Number.isFinite(value)||log&&value<=0)return null;const q=min===max ? .5 : Math.max(-1e6,Math.min(1e6,ratio(value)));return options.reverse?1-q:q;},contains:value=>Number.isFinite(value)&&(!log||value>0)&&value>=min&&value<=max};
}
export function axisLabelStyle(options,fallback,percent=false){return options?.numberFormat?{numberFormat:options.numberFormat}:percent?{format:'percent',decimals:0}:fallback;}
export function readAxisOptions(node){
 if(!node)return {};const scaling=child(node,'scaling'),result={reverse:attr(child(scaling,'orientation'),'val')==='maxMin',gridlines:!!child(node,'majorGridlines')};
 for(const key of ['min','max','logBase']){const raw=attr(child(scaling,key),'val');if(raw!==''&&Number.isFinite(Number(raw)))result[key]=Number(raw);}
 const major=attr(child(node,'majorUnit'),'val');if(major!==''&&Number.isFinite(Number(major))&&Number(major)>0)result.majorUnit=Number(major);
 const fmt=child(node,'numFmt');if(fmt&&attr(fmt,'sourceLinked')!=='1'&&attr(fmt,'formatCode'))result.numberFormat=attr(fmt,'formatCode');return result;
}
// Keep all unrelated axis XML (crossings, minor ticks, display units, extensions).
export function patchAxisOptions(node,options={},fallback='General',percent=false){
 const doc=node.ownerDocument,prefix=node.prefix||'c',make=(name,values={})=>{const el=doc.createElementNS(C,prefix+':'+name);for(const [k,v] of Object.entries(values))el.setAttribute(k,String(v));return el;};
 let scaling=child(node,'scaling');if(!scaling){scaling=make('scaling');node.insertBefore(scaling,children(node).find(n=>n.localName!=='axId')||null);}
 const setScale=(name,value)=>{const old=child(scaling,name);if(old)scaling.removeChild(old);if(value!==undefined)scaling.appendChild(make(name,{val:value}));};
 for(const key of ['logBase','orientation','max','min'])setScale(key,key==='orientation'?options.reverse?'maxMin':'minMax':options[key]);
 const scaleOrder=['logBase','orientation','max','min','extLst'];for(const part of children(scaling).sort((a,b)=>scaleOrder.indexOf(a.localName)-scaleOrder.indexOf(b.localName)))scaling.appendChild(part);
 const order=['axId','scaling','delete','axPos','majorGridlines','minorGridlines','title','numFmt','majorTickMark','minorTickMark','tickLblPos','spPr','txPr','crossAx','crosses','crossesAt','crossBetween','majorUnit','minorUnit','dispUnits','extLst'];
 const set=(name,element)=>{const old=child(node,name);if(old)node.removeChild(old);if(element){const index=order.indexOf(name),next=children(node).find(n=>order.indexOf(n.localName)>index);node.insertBefore(element,next||null);}};
 set('majorGridlines',options.gridlines===false?null:make('majorGridlines'));set('majorUnit',options.majorUnit===undefined?null:make('majorUnit',{val:options.majorUnit}));set('numFmt',make('numFmt',{formatCode:options.numberFormat||(percent?'0%':fallback),sourceLinked:options.numberFormat?'0':'1'}));
}
