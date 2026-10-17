const object = v => v && typeof v === 'object' && !Array.isArray(v);
const cell = s => {
  const m = /^([A-Z]{1,3})([1-9]\d*)$/.exec(s || '');
  return m && Number(m[2]) <= 1048576 && [...m[1]].reduce((n,c)=>n*26+c.charCodeAt(0)-64,0) <= 16384;
};
export function validateFreeze(v) {
  if (!object(v) || !Number.isInteger(v.rows) || !Number.isInteger(v.cols) || v.rows < 0 || v.rows >= 1048576 || v.cols < 0 || v.cols >= 16384) throw Error('틀 고정 위치가 올바르지 않습니다.');
}
export function validateConditionalRules(rules) {
  if (!Array.isArray(rules) || rules.length > 200) throw Error('조건부 서식은 시트당 최대 200개까지 지원합니다.');
  for (const r of rules) {
    const parts = typeof r?.range === 'string' ? r.range.split(':') : [];
    if (!object(r) || typeof r.id !== 'string' || r.id.length > 150 || parts.length !== 2 || !parts.every(cell) || !['cellIs','expression','containsText','duplicateValues','uniqueValues','top10','aboveAverage'].includes(r.type)) throw Error('조건부 서식 규칙이 올바르지 않습니다.');
    const coordinate=s=>{const m=/^([A-Z]+)(\d+)$/.exec(s);return [+m[2],[...m[1]].reduce((n,c)=>n*26+c.charCodeAt(0)-64,0)];},a=coordinate(parts[0]),b=coordinate(parts[1]);
    if(a[0]>b[0]||a[1]>b[1])throw Error('조건부 서식 범위가 올바르지 않습니다.');
    if((b[0]-a[0]+1)*(b[1]-a[1]+1)>100000)throw Error('조건부 서식 범위는 최대 100,000셀입니다.');
    if (r.type === 'cellIs' && !['equal','notEqual','greaterThan','greaterThanOrEqual','lessThan','lessThanOrEqual','between','notBetween'].includes(r.operator)) throw Error('조건부 서식 비교가 올바르지 않습니다.');
    if (!Array.isArray(r.formulae) || r.formulae.length > 2 || r.formulae.some(f=>typeof f!=='string'||f.length>8192)) throw Error('조건부 서식 수식이 올바르지 않습니다.');
    if (r.type==='expression' && r.formulae.length!==1 || r.type==='cellIs' && r.formulae.length!==(r.operator.includes('etween')?2:1)) throw Error('조건부 서식에 비교값 또는 수식이 필요합니다.');
    for (const k of ['stopIfTrue','bottom','percent','aboveAverage']) if (r[k]!==undefined && typeof r[k]!=='boolean') throw Error('조건부 서식 옵션이 올바르지 않습니다.');
    if (r.type==='top10' && (!Number.isInteger(r.rank)||r.rank<1||r.rank>(r.percent?100:1000))) throw Error('상위/하위 개수가 올바르지 않습니다.');
    if (r.text!==undefined && (typeof r.text!=='string'||r.text.length>255)) throw Error('조건부 서식 텍스트가 너무 깁니다.');
    if (!object(r.style) || Object.keys(r.style).some(k=>!['background','color','bold','italic'].includes(k))) throw Error('조건부 서식 색상이 올바르지 않습니다.');
    for (const k of ['background','color']) if (r.style[k]!==undefined && !/^#[\da-f]{6}$/i.test(r.style[k])) throw Error('조건부 서식 색상이 올바르지 않습니다.');
    for (const k of ['bold','italic']) if (r.style[k]!==undefined && typeof r.style[k]!=='boolean') throw Error('조건부 서식 글꼴이 올바르지 않습니다.');
  }
}
