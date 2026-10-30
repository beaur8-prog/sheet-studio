// Excel format codes are data, never JavaScript. Cache bounded, immutable parses
// because every visible cell, chart tick and filter label uses this path.
const cache=new Map(),colors={black:'#000000',green:'#008000',white:'#ffffff',blue:'#0000ff',magenta:'#ff00ff',yellow:'#ffff00',cyan:'#00ffff',red:'#ff0000'};
const placeholder=t=>t?.kind==='digit',symbol=(t,s)=>t?.kind==='code'&&t.text===s;
function parseSection(code){
  const tokens=[];let condition,color;
  for(let i=0;i<code.length;i++){
    const c=code[i];
    if(c==='"'){let text='';while(++i<code.length&&code[i]!=='"')text+=code[i];tokens.push({kind:'literal',text});}
    else if(c==='\\'){tokens.push({kind:'literal',text:code[++i]||''});}
    else if(c==='_'){i++;tokens.push({kind:'literal',text:' '});}
    else if(c==='*'){i++;} // Column-width filling has no stable plain-text width.
    else if(c==='['){const end=code.indexOf(']',i+1);if(end<0){tokens.push({kind:'literal',text:c});continue;}const text=code.slice(i+1,end),test=/^(<=|>=|<>|=|<|>)([+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:E[+-]?\d+)?)$/i.exec(text);if(test)condition={operator:test[1],value:Number(test[2])};else if(colors[text.toLowerCase()])color=colors[text.toLowerCase()];else if(/^\$/.test(text)){tokens.push({kind:'literal',text:text.slice(1).replace(/-[0-9a-f]+$/i,'')});}else if(/^[hms]{1,2}$/i.test(text))tokens.push({kind:'elapsed',text:'['+text+']'});i=end;}
    else tokens.push({kind:'0#?'.includes(c)?'digit':c==='@'?'text':'code',text:c});
  }
  return {code,tokens,condition,color};
}
export function parsedNumberFormat(code){
  code=String(code??'General');if(cache.has(code))return cache.get(code);
  const parts=[];let start=0,quote=false;
  for(let i=0;i<code.length;i++){if(code[i]==='\\'||code[i]==='_'||code[i]==='*'){i++;continue;}if(code[i]==='"')quote=!quote;else if(!quote&&code[i]===';'){parts.push(code.slice(start,i));start=i+1;}}
  parts.push(code.slice(start));const result=parts.slice(0,4).map(parseSection);if(cache.size>=1024)cache.delete(cache.keys().next().value);cache.set(code,result);return result;
}
export function numberFormatKind(code){
  if(!code||/^general$/i.test(code))return 'general';
  const tokens=parsedNumberFormat(code).flatMap(s=>s.tokens),plain=tokens.filter(t=>t.kind==='code').map(t=>t.text).join('');
  if(tokens.some(t=>t.kind==='elapsed')||/[ymdhs]/i.test(plain))return 'date';
  if(tokens.some(t=>symbol(t,'%')))return 'percent';
  if(/\*(?:"[^"]*"|\\.|.)/.test(code))return 'accounting';
  if(tokens.some(t=>/[$₩€£¥]/.test(t.text)))return 'currency';
  if(tokens.some(t=>t.kind==='text')&&!tokens.some(placeholder))return 'text';
  return 'number';
}
const matches=(n,c)=>({'>':n>c.value,'>=':n>=c.value,'<':n<c.value,'<=':n<=c.value,'=':n===c.value,'<>':n!==c.value})[c.operator];
export function selectNumberSection(code,value){
  const sections=parsedNumberFormat(code);
  if(typeof value!=='number')return {section:sections[3]||(sections.length===1&&sections[0].tokens.some(t=>t.kind==='text')?sections[0]:null),value,negative:false};
  const conditional=sections.slice(0,3).some(s=>s.condition);
  if(conditional){const section=sections.slice(0,3).find(s=>s.condition?matches(value,s.condition):true);return {section:section||null,value:Math.abs(value),negative:value<0,unmatched:!section};}
  const index=value<0&&sections.length>1?1:value===0&&sections.length>2?2:0;
  return {section:sections[index],value:Math.abs(value),negative:value<0&&index===0};
}
const literal=tokens=>tokens.map(t=>t.kind==='digit'?'':t.text).join('');
export function formatTextSection(value,section){return section?section.tokens.map(t=>t.kind==='text'?String(value):t.kind==='digit'?'':t.text).join(''):String(value);}
function integerDigits(digits,tokens){
  let at=digits.length-1,result='';
  for(let i=tokens.length-1;i>=0;i--){const t=tokens[i];if(placeholder(t))result=(at>=0?digits[at--]:t.text==='0'?'0':t.text==='?'?' ':'')+result;else if(!symbol(t,','))result=t.text+result;}
  return (at>=0?digits.slice(0,at+1):'')+result;
}
function fixed(value,tokens,decimals){
  const dot=tokens.findIndex(t=>symbol(t,'.')),integer=dot<0?tokens:tokens.slice(0,dot),fraction=dot<0?[]:tokens.slice(dot+1),digits=fraction.filter(placeholder);
  const precision=decimals??Math.min(12,digits.length),rounded=value.toLocaleString('en-US',{useGrouping:false,minimumFractionDigits:precision,maximumFractionDigits:precision}),[whole,tail='']=rounded.split('.');
  const significant=whole==='0'&&!integer.some(t=>placeholder(t)&&t.text==='0')?'':whole;
  let prefix=integerDigits(significant,integer);
  if(integer.some(t=>symbol(t,',')))prefix=prefix.replace(/\d+/g,part=>part.replace(/\B(?=(\d{3})+(?!\d))/g,','));
  const fractionTokens=decimals===null?fraction:Array.from({length:precision},()=>({kind:'digit',text:'0'}));let last=tail.length-1;while(last>=0&&tail[last]==='0'&&fractionTokens[last]?.text!=='0')last--;
  let body='',at=0;for(const t of fractionTokens){if(placeholder(t)){body+=at<=last||t.text==='0'?(tail[at]||'0'):t.text==='?'?' ':'';at++;}else body+=t.text;}
  return prefix+((dot>=0||precision>0)&&(body||!fractionTokens.length)?'.'+body:'');
}
function bestFraction(value,limit){
  let p0=0,q0=1,p1=1,q1=0,x=value;
  for(let steps=0;steps<64;steps++){const a=Math.floor(x),q2=q0+a*q1;if(q2>limit)break;const p2=p0+a*p1;[p0,q0,p1,q1]=[p1,q1,p2,q2];const remainder=x-a;if(remainder<1e-12)return [p1,q1];x=1/remainder;}
  if(!q1)return [Math.round(value),1];const k=Math.floor((limit-q0)/q1),a=[p0+k*p1,q0+k*q1],b=[p1,q1];return Math.abs(a[0]/a[1]-value)<Math.abs(b[0]/b[1]-value)?a:b;
}
function fractionValue(value,tokens){
  const slash=tokens.findIndex(t=>symbol(t,'/'));if(slash<0)return null;
  let numeratorStart=slash;while(placeholder(tokens[numeratorStart-1]))numeratorStart--;if(numeratorStart===slash)return null;
  let denominatorEnd=slash+1;while(placeholder(tokens[denominatorEnd])||tokens[denominatorEnd]?.kind==='code'&&/^[1-9]$/.test(tokens[denominatorEnd].text))denominatorEnd++;if(denominatorEnd===slash+1)return null;
  const before=tokens.slice(0,numeratorStart),numerator=tokens.slice(numeratorStart,slash),denominator=tokens.slice(slash+1,denominatorEnd),mixed=before.some(placeholder),whole=mixed?Math.floor(value):0;
  const fixedDen=denominator.every(t=>t.kind==='code')?Number(literal(denominator)):null;
  let [n,d]=fixedDen?[Math.round((value-whole)*fixedDen),fixedDen]:bestFraction(value-whole,Math.min(9999,10**denominator.length-1)),integer=whole;
  if(mixed&&n>=d){integer+=Math.floor(n/d);n%=d;}
  const wholeText=integerDigits(integer?String(integer):'',before);
  if(mixed&&!n)return wholeText+' '.repeat(numerator.length+1+denominator.length)+literal(tokens.slice(denominatorEnd));
  const numeratorText=integerDigits(n?String(n):'',numerator),denominatorText=fixedDen?String(fixedDen):integerDigits(String(d),denominator);
  return wholeText+numeratorText+'/'+denominatorText+literal(tokens.slice(denominatorEnd));
}
export function formatNumericSection(value,section,{negative=false,decimals=null}={}){
  if(!section)return '########';
  if(!section.code)return '';
  if(/^general$/i.test(section.code))return (negative?'-':'')+String(Number(value.toPrecision(12)));
  if(section.tokens.some(t=>t.kind==='text')&&!section.tokens.some(placeholder))return (negative?'-':'')+String(value);
  const tokens=section.tokens,first=tokens.findIndex(placeholder);if(first<0){let text='';for(let i=0;i<tokens.length;i++){if(tokens.slice(i,i+7).length===7&&tokens.slice(i,i+7).every(t=>t.kind==='code')&&tokens.slice(i,i+7).map(t=>t.text).join('').toLowerCase()==='general'){text+=(negative?'-':'')+String(Number(value.toPrecision(12)));i+=6;}else text+=tokens[i].text;}return text;}
  const percent=tokens.filter(t=>symbol(t,'%')).length;value*=100**percent;
  let end=tokens.length-1;while(end>=first&&!placeholder(tokens[end]))end--;
  const slash=tokens.findIndex(t=>symbol(t,'/'));if(slash>first){let d=slash+1;while(placeholder(tokens[d])||tokens[d]?.kind==='code'&&/^[1-9]$/.test(tokens[d].text))d++;if(d>slash+1)end=Math.max(end,d-1);}
  // Trailing scaling commas may follow the integer or decimal placeholders.
  let commas=0;const trimmed=tokens.map(t=>({...t}));for(let i=end+1;symbol(trimmed[i],',');i++){trimmed[i].text='';commas++;}
  const dot=trimmed.findIndex(t=>symbol(t,'.'));if(dot>=0)for(let i=dot-1;symbol(trimmed[i],',');i--){trimmed[i].text='';commas++;}
  value/=1000**commas;if(!Number.isFinite(value))return '########';
  const prefix=literal(trimmed.slice(0,first)),body=trimmed.slice(first,end+1),suffix=literal(trimmed.slice(end+1));let number;
  const exponent=body.findIndex((t,i)=>t.kind==='code'&&/^[eE]$/.test(t.text)&&['+','-'].includes(body[i+1]?.text));
  if(exponent>=0){
    const mantissa=body.slice(0,exponent),sign=body[exponent+1].text,places=body.slice(exponent+2).filter(placeholder).length,intPlaces=Math.max(1,mantissa.slice(0,mantissa.findIndex(t=>symbol(t,'.'))<0?mantissa.length:mantissa.findIndex(t=>symbol(t,'.'))).filter(placeholder).length);
    let power=value?Math.floor(Math.log10(value)/intPlaces)*intPlaces:0,scaled=value?value/10**power:0;
    const decimal=mantissa.findIndex(t=>symbol(t,'.')),precision=decimals??(decimal<0?0:mantissa.slice(decimal+1).filter(placeholder).length);
    if(Number(scaled.toFixed(Math.min(12,precision)))>=10**intPlaces){power+=intPlaces;scaled/=10**intPlaces;}
    number=fixed(scaled,mantissa,decimals)+body[exponent].text+(power<0?'-':sign==='+'?'+':'')+String(Math.abs(power)).padStart(places,'0');
  }else number=fractionValue(value,body)??fixed(value,body,decimals);
  return (negative?'-':'')+prefix+number+suffix;
}
export function effectiveNumberFormat(style={}){
  if(typeof style==='string')return style;
  if(typeof style.numberFormat==='string'&&(!style.format||numberFormatKind(style.numberFormat)===style.format))return Number.isInteger(style.decimals)&&numberFormatKind(style.numberFormat)!=='date'?formatWithDecimals(style.numberFormat,style.decimals):style.numberFormat;
  const fraction=Number.isInteger(style.decimals)?Math.max(0,Math.min(12,style.decimals)):2,digits=fraction?'.'+'0'.repeat(fraction):'';
  return {general:'General',number:'#,##0'+digits,currency:'₩#,##0'+digits,accounting:'₩* #,##0'+digits+';[Red](₩* #,##0'+digits+');₩* "-";@',percent:'0'+digits+'%',date:'yyyy-mm-dd',text:'@'}[style.format||'general']||style.format;
}
export function numberFormatColor(value,style){return typeof value==='number'?selectNumberSection(effectiveNumberFormat(style),value).section?.color:undefined;}
export function numberFormatStyle(value,style){const color=numberFormatColor(value,style);return color?{...style,color}:style;}
export function formatWithDecimals(code,decimals){
  const places=Math.max(0,Math.min(12,decimals)),replacement=places?'.'+'0'.repeat(places):'';
  return parsedNumberFormat(code).map(section=>{
    const source=section.code,positions=[];let quote=false;
    for(let i=0;i<source.length;i++){const c=source[i];if(c==='\\'||c==='_'||c==='*'){i++;continue;}if(c==='"'){quote=!quote;continue;}if(quote)continue;if(c==='['){const end=source.indexOf(']',i);if(end<0)break;i=end;continue;}positions.push({c,i});}
    if(positions.some(p=>p.c==='/'))return source;
    const exponent=positions.findIndex(p=>/e/i.test(p.c)),mantissa=exponent<0?positions:positions.slice(0,exponent),digits=mantissa.filter(p=>'0#?'.includes(p.c));if(!digits.length)return source;
    const dot=mantissa.find(p=>p.c==='.');
    if(!dot){const at=digits.at(-1).i+1;return source.slice(0,at)+replacement+source.slice(at);}
    let end=dot.i+1;while('0#?'.includes(source[end]||'\0'))end++;return source.slice(0,dot.i)+replacement+source.slice(end);
  }).join(';');
}
export function validateNumberFormat(code){
  if(typeof code!=='string'||code.length>512)throw Error('표시 형식 코드는 512자 이내로 입력하세요.');
  let quoted=false,sections=1;
  for(let i=0;i<code.length;i++){const c=code[i];if(c==='\\'||c==='_'||c==='*'){if(++i>=code.length)throw Error('표시 형식의 이스케이프 뒤에 문자를 입력하세요.');continue;}if(c==='"')quoted=!quoted;else if(!quoted&&c==='['){const end=code.indexOf(']',i+1);if(end<0)throw Error('표시 형식의 대괄호를 닫아 주세요.');i=end;}else if(!quoted&&c===';')sections++;}
  if(quoted)throw Error('표시 형식의 큰따옴표를 닫아 주세요.');if(sections>4)throw Error('표시 형식은 최대 4개 구간까지 입력하세요.');
}
