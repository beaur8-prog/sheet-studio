const color=v=>typeof v==='string'&&/^#[\da-f]{6}$/i.test(v);
export const VISUAL_CONDITIONAL_TYPES=new Set(['colorScale','dataBar']);
export function validateVisualConditional(rule){
  const points=rule.cfvo;
  if(!Array.isArray(points)||![2,3].includes(points.length)||rule.type==='dataBar'&&points.length!==2)throw Error('색조는 2~3개, 데이터 막대는 2개의 기준점이 필요합니다.');
  points.forEach((point,i)=>{
    if(!point||!['min','max','num','percent','percentile'].includes(point.type)||i===0&&point.type==='max'||i===points.length-1&&point.type==='min'||i>0&&i<points.length-1&&['min','max'].includes(point.type))throw Error('시각 규칙의 기준점 종류가 올바르지 않습니다.');
    if(!['min','max'].includes(point.type)&&(typeof point.value!=='number'||!Number.isFinite(point.value)))throw Error('기준점에 유한한 숫자를 입력하세요.');
    if(['percent','percentile'].includes(point.type)&&(point.value<0||point.value>100))throw Error('백분율·백분위수는 0~100 사이입니다.');
    if(i>0&&point.type===points[i-1].type&&point.value<points[i-1].value)throw Error('기준점의 값은 작은 값부터 입력하세요.');
  });
  if(rule.type==='colorScale'&&(!Array.isArray(rule.colors)||rule.colors.length!==points.length||!rule.colors.every(color)))throw Error('각 기준점의 색상을 선택하세요.');
  if(rule.type==='dataBar'&&!color(rule.barColor))throw Error('데이터 막대 색상을 선택하세요.');
  if(rule.stopIfTrue)throw Error('색조·데이터 막대에는 다음 규칙 중지를 사용할 수 없습니다.');
}
const blend=(a,b,ratio)=>'#'+[1,3,5].map(i=>Math.round(parseInt(a.slice(i,i+2),16)*(1-ratio)+parseInt(b.slice(i,i+2),16)*ratio).toString(16).padStart(2,'0')).join('');
function threshold(point,numbers){
  const lo=numbers[0],hi=numbers.at(-1);
  if(point.type==='min')return lo;if(point.type==='max')return hi;if(point.type==='num')return point.value;
  if(point.type==='percent')return lo+(hi-lo)*point.value/100;
  const index=(numbers.length-1)*point.value/100,first=Math.floor(index),last=Math.ceil(index);return numbers[first]+(numbers[last]-numbers[first])*(index-first);
}
export function visualConditionalStyle(rule,numbers,value){
  if(typeof value!=='number'||!Number.isFinite(value)||!numbers.length)return null;
  const stops=rule.cfvo.map(point=>threshold(point,numbers));if(stops.some((n,i)=>!Number.isFinite(n)||i>0&&n<stops[i-1]))return null;
  if(rule.type==='dataBar'){
    const [lo,hi]=stops,ratio=hi===lo?1:Math.max(0,Math.min(1,(value-lo)/(hi-lo)));
    // ISO/IEC 29500 base data bars use 10%..90%, including negative inputs.
    // Modern center axes/negative colors are an x14 rule and remain preserved.
    return {dataBar:{percent:10+ratio*80,color:rule.barColor}};
  }
  if(value<=stops[0]||stops[0]===stops.at(-1))return {background:rule.colors[0]};
  for(let i=1;i<stops.length;i++)if(value<=stops[i])return {background:blend(rule.colors[i-1],rule.colors[i],stops[i]===stops[i-1]?1:(value-stops[i-1])/(stops[i]-stops[i-1]))};
  return {background:rule.colors.at(-1)};
}
export function dataBarCSS(style){
  const bar=style?.dataBar;if(!bar||!color(bar.color)||!Number.isFinite(bar.percent))return '';
  const percent=Math.max(0,Math.min(100,bar.percent));return `background-image:linear-gradient(to right,${bar.color},${bar.color}44);background-size:${percent}% calc(100% - 6px);background-position:2px center;background-repeat:no-repeat;`;
}
