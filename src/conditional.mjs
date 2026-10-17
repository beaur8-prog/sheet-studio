import {address,parseAddress,evaluateFormula,createEvaluator,shiftFormula,transformFormulaReferences} from './engine.mjs';
import {validateConditionalRules} from './worksheet-schema.mjs';
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
    if(['top10','aboveAverage','duplicateValues','uniqueValues'].includes(rule.type)){
      for(const [key,cell] of Object.entries(sheet.cells)){
        const p=parseAddress(key);if(!p||p.r<b.r1||p.r>b.r2||p.c<b.c1||p.c>b.c2||cell.raw==='')continue;
        const v=evaluate(key);if(typeof v==='number')stats.numbers.push(v);
        const k=String(v).toLocaleLowerCase();stats.counts.set(k,(stats.counts.get(k)||0)+1);
      }
      stats.mean=stats.numbers.reduce((a,b)=>a+b,0)/stats.numbers.length;
      stats.numbers.sort((a,b)=>rule.bottom?a-b:b-a);
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
      const blank=!sheet.cells[address(r,c)]?.raw;
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
export function shiftWorksheetFeatures(sheet,axis,index,delta) {
  if(sheet.hiddenSizes){const key=axis==='r'?'rows':'cols',sizes={};for(const [keyIndex,size] of Object.entries(sheet.hiddenSizes[key]||{})){let n=+keyIndex;if(delta<0&&n===index)continue;if(n>=index)n+=delta;if(n>=0&&n<(axis==='r'?1048576:16384))sizes[n]=size;}sheet.hiddenSizes[key]=sizes;}
  if(sheet.freeze){const key=axis==='r'?'rows':'cols';if(index<sheet.freeze[key])sheet.freeze[key]=Math.max(0,sheet.freeze[key]+delta);}
  sheet.conditionalRules=(sheet.conditionalRules||[]).flatMap(rule=>{
    const range=transformFormulaReferences(rule.range,{axis,index,delta,targetSheet:sheet.name,ownerSheet:sheet.name});
    if(range.includes('#REF!'))return [];
    const formulae=rule.formulae.map(f=>transformFormulaReferences(f,{axis,index,delta,targetSheet:sheet.name,ownerSheet:sheet.name}));
    return [{...rule,range,formulae}];
  });
}
const rgb=c=>c?.argb&&/^([\da-f]{8})$/i.test(c.argb)?'#'+c.argb.slice(-6).toLowerCase():undefined;
export function importConditional(source,sheet) {
  const rules=[];let unsupported=0;
  for(const group of source.conditionalFormattings||[])for(const original of group.rules||[])for(const ref of group.ref.split(/\s+/)){
    const r={id:crypto.randomUUID(),range:ref.includes(':')?ref:ref+':'+ref,type:original.type,formulae:(original.formulae||[]).map(String),style:{},stopIfTrue:!!original.stopIfTrue,_priority:original.priority||1};
    for(const k of ['operator','text','rank','percent','bottom','aboveAverage'])if(original[k]!==undefined)r[k]=original[k];
    if(original.type==='containsText'&&original.operator&&original.operator!=='containsText'){r.type='expression';delete r.operator;}
    const s=original.style||{};if(rgb(s.fill?.bgColor)||rgb(s.fill?.fgColor))r.style.background=rgb(s.fill?.bgColor)||rgb(s.fill.fgColor);if(rgb(s.font?.color))r.style.color=rgb(s.font.color);
    for(const k of ['bold','italic'])if(s.font?.[k]!==undefined)r.style[k]=s.font[k];
    try{validateConditionalRules([r]);ruleBounds(r);rules.push(r);}catch{unsupported++;}
  }
  if(rules.length)sheet.conditionalRules=rules.sort((a,b)=>a._priority-b._priority).map(({_priority,...r})=>({...r,sourcePriority:_priority}));
  if(unsupported)sheet.unsupportedConditionalCount=unsupported;
}
export function exportConditional(target,sheet) {
  if(sheet.unsupportedConditionalCount)throw Error('이 시트에 지원되지 않는 조건부 서식이 있습니다. 원본을 보존하려면 XLSM을 사용하거나 규칙 관리에서 모두 지우세요.');
  (sheet.conditionalRules||[]).forEach((rule,i)=>{
    const style={};if(rule.style.background){const color={argb:'FF'+rule.style.background.slice(1).toUpperCase()};style.fill={type:'pattern',pattern:'solid',fgColor:color,bgColor:color};}
    if(rule.style.color||rule.style.bold!==undefined||rule.style.italic!==undefined)style.font={...(rule.style.color?{color:{argb:'FF'+rule.style.color.slice(1).toUpperCase()}}:{}),...(rule.style.bold!==undefined?{bold:rule.style.bold}:{}),...(rule.style.italic!==undefined?{italic:rule.style.italic}:{})};
    // Native expression rules make text/duplicate/top/average behavior portable.
    const entry=rule.type==='cellIs'?{type:'cellIs',operator:rule.operator,formulae:rule.formulae.map(f=>f.replace(/^=/,''))}:{type:'expression',formulae:[conditionalFormula(rule)]};
    target.addConditionalFormatting({ref:rule.range,rules:[{...entry,style,priority:i+1,stopIfTrue:!!rule.stopIfTrue}]});
  });
}
