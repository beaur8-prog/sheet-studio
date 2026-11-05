import {VISUAL_CONDITIONAL_TYPES,visualConditionalStyle} from './conditional-visual.mjs';
import {xml,children,child,attr} from './chart-xml.mjs';
import {transformPreservedFeatures} from './xlsx-feature-transforms.mjs';
import {shiftIndexMap} from './structure-coordinates.mjs';
import {address,parseAddress,evaluateFormula,createEvaluator,shiftFormula,transformFormulaReferences} from './engine.mjs';
import {validateConditionalRules} from './worksheet-schema.mjs';
import {normalizeColor} from './spreadsheet-colors.mjs';
export function ruleBounds(rule) {
  const [a,b]=rule.range.split(':').map(parseAddress);
  if(!a||!b||a.r>b.r||a.c>b.c)throw Error('조건부 서식 범위가 올바르지 않습니다.');
  return {r1:a.r,c1:a.c,r2:b.r,c2:b.c};
}
export function conditionalFormula(rule) {
  const b=ruleBounds(rule),first=address(b.r1,b.c1),range=rule.range.replace(/([A-Z]+)(\d+)/g,'$$$1$$$2');
  const quote=s=>'"'+String(s).replace(/"/g,'""')+'"';
  switch(rule.type){
    case 'expression':return rule.formulae[0].replace(/^=/,'');
    case 'containsText':return `AND(NOT(ISBLANK(${first})),ISNUMBER(SEARCH(${quote(rule.text||'')},${first})))`;
    case 'duplicateValues':return `AND(NOT(ISBLANK(${first})),COUNTIF(${range},${first})>1)`;
    case 'uniqueValues':return `AND(NOT(ISBLANK(${first})),COUNTIF(${range},${first})=1)`;
    case 'aboveAverage':return `AND(ISNUMBER(${first}),${first}${rule.aboveAverage===false?'<':'>'}AVERAGE(${range}))`;
    case 'top10':return `AND(ISNUMBER(${first}),RANK(${first},${range},${rule.bottom?1:0})<=${rule.percent?`MAX(1,INT(COUNT(${range})*${rule.rank}/100))`:rule.rank})`;
  }
  return '';
}
export function createConditionalEvaluator(book,sheet,evaluate=createEvaluator(book,sheet.id)) {
  const rules=(sheet.conditionalRules||[]).map(rule=>{
    const b=ruleBounds(rule),stats={numbers:[],counts:new Map()};
    if(['top10','aboveAverage','duplicateValues','uniqueValues','colorScale','dataBar'].includes(rule.type)){
      for(let r=b.r1;r<=b.r2;r++)for(let c=b.c1;c<=b.c2;c++){
        const key=address(r,c);if(!sheet.cells[key]?.raw&&!evaluate.isComputed?.(key))continue;
        const v=evaluate(key);if(typeof v==='number'&&Number.isFinite(v))stats.numbers.push(v);
        if(['duplicateValues','uniqueValues'].includes(rule.type)){const k=String(v).toLocaleLowerCase();stats.counts.set(k,(stats.counts.get(k)||0)+1);}
      }
      stats.mean=stats.numbers.reduce((a,b)=>a+b,0)/stats.numbers.length;
      stats.numbers.sort((a,b)=>rule.bottom||VISUAL_CONDITIONAL_TYPES.has(rule.type)?a-b:b-a);
      const count=rule.percent?Math.max(1,Math.floor(stats.numbers.length*rule.rank/100)):rule.rank;
      stats.threshold=stats.numbers[Math.min(count,stats.numbers.length)-1];
    }
    return {rule,b,stats};
  });
  return (r,c,value)=>{
    const style={};
    for(const {rule,b,stats} of rules){
      if(r<b.r1||r>b.r2||c<b.c1||c>b.c2)continue;
      let match=false;
      const formula=f=>evaluateFormula(book,sheet.id,shiftFormula('='+f.replace(/^=/,''),r-b.r1,c-b.c1));
      const blank=!sheet.cells[address(r,c)]?.raw&&!evaluate.isComputed?.(address(r,c));
      if(VISUAL_CONDITIONAL_TYPES.has(rule.type)){const visual=visualConditionalStyle(rule,stats.numbers,value);if(visual){for(const [key,v] of Object.entries(visual))if(!(key in style))style[key]=v;}continue;}
      if(rule.type==='expression'){const v=formula(rule.formulae[0]);match=v===true||typeof v==='number'&&v!==0;}
      else if(rule.type==='cellIs'){
        const [a,z]=rule.formulae.map(formula),v=blank?0:value;
        if(!(typeof v==='string'&&v.startsWith('#'))&&typeof v===typeof a){const x=typeof v==='string'?v.toLowerCase():v,y=typeof a==='string'?a.toLowerCase():a;
          match=({equal:()=>x===y,notEqual:()=>x!==y,greaterThan:()=>x>y,greaterThanOrEqual:()=>x>=y,lessThan:()=>x<y,lessThanOrEqual:()=>x<=y,between:()=>x>=a&&x<=z,notBetween:()=>x<a||x>z})[rule.operator]();}
      }else if(rule.type==='containsText')match=!blank&&formula(conditionalFormula(rule))===true;
      else if(rule.type==='duplicateValues'||rule.type==='uniqueValues')match=!blank&&((stats.counts.get(String(value).toLocaleLowerCase())||0)>1)===(rule.type==='duplicateValues');
      else if(rule.type==='aboveAverage')match=typeof value==='number'&&(rule.aboveAverage===false?value<stats.mean:value>stats.mean);
      else if(rule.type==='top10')match=!blank&&typeof value==='number'&&stats.threshold!==undefined&&(rule.bottom?value<=stats.threshold:value>=stats.threshold);
      if(match){for(const [k,v] of Object.entries(rule.style))if(!(k in style))style[k]=v;if(rule.stopIfTrue)break;}
    }
    return style;
  };
}
export function shiftWorksheetFeatures(sheet,axis,index,delta,targetSheet=sheet.name) {
  const local=targetSheet.toLowerCase()===sheet.name.toLowerCase();
  const limit=axis==='r'?1048576:16384,rebaseAmount=n=>local&&delta<0&&n>=index&&n<index-delta?index-delta-n:0;
  if(local&&sheet.hiddenSizes){const key=axis==='r'?'rows':'cols';sheet.hiddenSizes[key]=shiftIndexMap(sheet.hiddenSizes[key],index,delta,limit);}
  if(local&&sheet.xlsxProtectionDefaults){const key=axis==='r'?'rows':'columns';sheet.xlsxProtectionDefaults[key]=shiftIndexMap(sheet.xlsxProtectionDefaults[key],index,delta,limit);}
  if(local&&sheet.xlsxStyleDefaults){const key=axis==='r'?'rows':'columns';sheet.xlsxStyleDefaults[key]=shiftIndexMap(sheet.xlsxStyleDefaults[key],index,delta,limit);}
  if(local&&sheet.freeze){const key=axis==='r'?'rows':'cols';if(index<sheet.freeze[key])sheet.freeze[key]=delta>0?Math.min(limit,sheet.freeze[key]+delta):Math.max(index,sheet.freeze[key]+delta);}
  const transform=f=>transformFormulaReferences(f,{axis,index,delta,targetSheet,ownerSheet:sheet.name});
  transformPreservedFeatures(sheet,{reference:local?transform:f=>f,formula:transform,rebase:(formula,range)=>{const anchor=parseAddress(range?.split(':')[0]),amount=anchor?rebaseAmount(anchor[axis]):0;return amount?shiftFormula('='+formula.replace(/^=/,''),axis==='r'?amount:0,axis==='c'?amount:0).replace(/^=/,formula.startsWith('=')?'=':''):formula;}});
  sheet.conditionalRules=(sheet.conditionalRules||[]).flatMap(rule=>{
    const range=local?transformFormulaReferences(rule.range,{axis,index,delta,targetSheet,ownerSheet:sheet.name}):rule.range;
    if(range.includes('#REF!'))return [];
    const anchor=ruleBounds(rule),amount=rebaseAmount(axis==='r'?anchor.r1:anchor.c1);
    const formulae=rule.formulae.map(f=>{const rebased=amount?shiftFormula('='+f.replace(/^=/,''),axis==='r'?amount:0,axis==='c'?amount:0).replace(/^=/,f.startsWith('=')?'=':''):f;return transformFormulaReferences(rebased,{axis,index,delta,targetSheet,ownerSheet:sheet.name});});
    return [{...rule,range,formulae}];
  });
}
const rgb=c=>c?.argb&&/^([\da-f]{8})$/i.test(c.argb)?'#'+c.argb.slice(-6).toLowerCase():undefined;
export function importConditional(source,sheet,resolveColor=rgb,rawGroups=[]) {
  const color=c=>normalizeColor(resolveColor(c));
  const rawRules=new Map();for(const text of rawGroups){const root=xml(text).documentElement;for(const node of children(root,'cfRule'))rawRules.set(attr(root,'sqref')+'|'+attr(node,'priority'),node);}
  const rules=[];let unsupported=0;
  for(const group of source.conditionalFormattings||[])for(const original of group.rules||[])for(const ref of group.ref.split(/\s+/)){
    const r={id:crypto.randomUUID(),range:ref.includes(':')?ref:ref+':'+ref,type:original.type,formulae:(original.formulae||[]).map(String),style:{},stopIfTrue:!!original.stopIfTrue,_priority:original.priority||1};
    for(const k of ['operator','text','rank','percent','bottom','aboveAverage'])if(original[k]!==undefined)r[k]=original[k];
    if(original.type==='containsText'&&original.operator&&original.operator!=='containsText'){r.type='expression';delete r.operator;}
    const origin=parseAddress(group.ref.split(/\s+/)[0].split(':')[0]),anchor=parseAddress(ref.split(':')[0]);if(origin&&anchor)r.formulae=r.formulae.map(f=>shiftFormula('='+f.replace(/^=/,''),anchor.r-origin.r,anchor.c-origin.c).slice(1));
    if(VISUAL_CONDITIONAL_TYPES.has(r.type)){
      r.cfvo=(original.cfvo||[]).map(point=>({type:point.type,...(!['min','max'].includes(point.type)?{value:point.value}:{})}));
      const raw=rawRules.get(group.ref+'|'+original.priority),bar=child(raw,'dataBar');
      const advanced=original.x14Id||child(raw,'extLst')||r.type==='dataBar'&&(attr(bar,'showValue','1')==='0'||!['10',''].includes(attr(bar,'minLength'))||!['90',''].includes(attr(bar,'maxLength')));
      if(advanced){unsupported++;continue;}
      if(r.type==='colorScale')r.colors=(original.color||[]).map(color);
      else r.barColor=color(original.color);
    }
    const s=original.style||{};if(color(s.fill?.fgColor)||color(s.fill?.bgColor))r.style.background=color(s.fill?.fgColor)||color(s.fill.bgColor);if(color(s.font?.color))r.style.color=color(s.font.color);
    for(const k of ['bold','italic'])if(s.font?.[k]!==undefined)r.style[k]=s.font[k];
    try{validateConditionalRules([r]);ruleBounds(r);rules.push(r);}catch{unsupported++;}
  }
  if(rules.length)sheet.conditionalRules=rules.sort((a,b)=>a._priority-b._priority).map(({_priority,...r})=>({...r,sourcePriority:_priority}));
  if(unsupported)sheet.unsupportedConditionalCount=unsupported;
}
export function exportConditional(target,sheet) {
  if(sheet.unsupportedConditionalCount&&(sheet.xlsxFeatures?.conditionalFormatting?.length||sheet.xlsxFeatures?.conditionalExtensions?.length))return;
  if(sheet.unsupportedConditionalCount)throw Error('이 시트에 지원되지 않는 조건부 서식이 있습니다. 원본을 보존하려면 XLSM을 사용하거나 규칙 관리에서 모두 지우세요.');
  (sheet.conditionalRules||[]).forEach((rule,i)=>{
    const style={};if(rule.style.background){const color={argb:'FF'+rule.style.background.slice(1).toUpperCase()};style.fill={type:'pattern',pattern:'solid',fgColor:color,bgColor:color};}
    if(rule.style.color||rule.style.bold!==undefined||rule.style.italic!==undefined)style.font={...(rule.style.color?{color:{argb:'FF'+rule.style.color.slice(1).toUpperCase()}}:{}),...(rule.style.bold!==undefined?{bold:rule.style.bold}:{}),...(rule.style.italic!==undefined?{italic:rule.style.italic}:{})};
    // Native expression rules make text/duplicate/top/average behavior portable.
    if(rule.type==='colorScale'){target.addConditionalFormatting({ref:rule.range,rules:[{type:'colorScale',priority:i+1,cfvo:rule.cfvo,color:rule.colors.map(c=>({argb:'FF'+c.slice(1).toUpperCase()}))}]});return;}
    // The base bar is written after ExcelJS, which otherwise emits an empty x14:id.
    if(rule.type==='dataBar')return;
    const entry=rule.type==='cellIs'?{type:'cellIs',operator:rule.operator,formulae:rule.formulae.map(f=>f.replace(/^=/,''))}:{type:'expression',formulae:[conditionalFormula(rule)]};
    target.addConditionalFormatting({ref:rule.range,rules:[{...entry,style,priority:i+1,stopIfTrue:!!rule.stopIfTrue}]});
  });
}
