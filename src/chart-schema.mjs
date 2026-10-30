import {axisOptionsError} from './chart-axis.mjs';
const object=v=>v&&typeof v==='object'&&!Array.isArray(v);
const color=v=>typeof v==='string'&&/^#[a-f0-9]{6}$/i.test(v);
const part=v=>typeof v==='string'&&v.length<2000&&!v.startsWith('/')&&!v.split('/').some(s=>!s||s==='.'||s==='..')&&!/[\\:\u0000]/.test(v);
const string=(v,max=4000)=>typeof v==='string'&&v.length<=max;
const fail=()=>{throw Error('차트 설정 또는 보존 데이터가 올바르지 않습니다.');};
export function validateChart(chart){
  for(const key of ['valueAxis','xAxis','secondaryValueAxis'])if(chart[key]!==undefined&&axisOptionsError(chart[key]))fail();
  for(const key of ['title','nativeName','previewNote','unsupportedReason','xTitle','yTitle'])if(chart[key]!==undefined&&!string(chart[key]))fail();
  for(const key of ['legend','labels','plotVisibleOnly','secondaryAxis','rebuilt','radarFilled','firstRowAsHeader','firstColumnAsLabels'])if(chart[key]!==undefined&&typeof chart[key]!=='boolean')fail();
  if(chart.placement!==undefined&&!['move-size','move','fixed'].includes(chart.placement))fail();
  if(chart.seriesBy!==undefined&&!['rows','columns'].includes(chart.seriesBy))fail();
  if(chart.grouping!==undefined&&!['standard','clustered','stacked','percentStacked'].includes(chart.grouping))fail();
  const ref=r=>{if(!object(r))fail();for(const key of ['sheetId','ownerSheetId','range','formula','formatCode'])if(r[key]!==undefined&&!string(r[key]))fail();if(r.cache!==undefined&&(!Array.isArray(r.cache)||r.cache.length>10000||r.cache.some(v=>v!==null&&!(typeof v==='number'&&Number.isFinite(v))&&typeof v!=='boolean'&&!string(v,32768))))fail();};
  if(chart.series!==undefined){if(!Array.isArray(chart.series)||chart.series.length>128)fail();for(const s of chart.series){if(!object(s))fail();if(s.name!==undefined&&!string(s.name))fail();if(s.color!==undefined&&!color(s.color))fail();if(s.type!==undefined&&!string(s.type,80))fail();for(const key of ['secondary','lines'])if(s[key]!==undefined&&typeof s[key]!=='boolean')fail();for(const key of ['nameRef','catRef','valRef','xRef','sizeRef'])if(s[key]!==undefined)ref(s[key]);if(s.pointColors!==undefined&&(!object(s.pointColors)||Object.entries(s.pointColors).some(([k,v])=>!/^\d+$/.test(k)||!color(v))))fail();}}
  if(chart.source!==undefined){const s=chart.source;if(!object(s)||!part(s.root)||!object(s.parts)||!s.parts[s.root]||!string(s.rootType)||!string(s.anchor,2_000_000)||!object(s.original)||!Array.isArray(s.rels)||s.rels.length>100||!Array.isArray(s.formulas)||s.formulas.length>2000)fail();s.formulas.forEach(ref);
    let size=0;const entries=Object.entries(s.parts);if(entries.length>500)fail();for(const [path,file] of entries){if(!part(path)||!object(file)||!string(file.contentType,255)||!string(file.data,24_000_000)||!/^[A-Za-z0-9+/]*={0,2}$/.test(file.data))fail();size+=file.data.length;}if(size>24_000_000)fail();
    for(const r of s.rels)if(!object(r)||!string(r.id,200)||!string(r.type)||!string(r.target)||r.mode&&r.mode!=='External'||r.mode!=='External'&&(!part(r.target)||!s.parts[r.target]))fail();
  }
}
