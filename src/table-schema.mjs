import {validateStoredParts,validateStoredRelationships} from './xlsx-part-store.mjs';
import {validateAutoFilter} from './filter-schema.mjs';
import {validatePreservedXML} from './xlsx-preservation-schema.mjs';
const object=v=>v&&typeof v==='object'&&!Array.isArray(v);
export function tableBounds(text){
  const point=value=>{const m=/^\$?([A-Z]{1,3})\$?([1-9]\d*)$/i.exec(value||'');if(!m)return null;let c=0;for(const letter of m[1].toUpperCase())c=c*26+letter.charCodeAt(0)-64;const r=Number(m[2])-1;return c<=16384&&r<1048576?{r,c:c-1}:null;};
  const parts=String(text).split(':'),a=point(parts[0]),b=point(parts[1]||parts[0]);return parts.length<=2&&a&&b&&a.r<=b.r&&a.c<=b.c?{r1:a.r,r2:b.r,c1:a.c,c2:b.c}:null;
}
export const overlaps=(a,b)=>a.r1<=b.r2&&a.r2>=b.r1&&a.c1<=b.c2&&a.c2>=b.c1;
export function validateTables(tables){
  if(!Array.isArray(tables)||tables.length>1000)throw Error('표 목록이 올바르지 않습니다.');const ids=new Set(),names=new Set(),ranges=[];
  for(const table of tables){
    const b=tableBounds(table?.range);
    if(!object(table)||typeof table.id!=='string'||!table.id||table.id.length>150||ids.has(table.id)||typeof table.name!=='string'||!table.name||table.name.length>255||names.has(table.name.toLocaleLowerCase())||!b||!Array.isArray(table.columns)||table.columns.length!==b.c2-b.c1+1)throw Error('표 이름·범위·열 구성이 올바르지 않습니다.');
    if(typeof table.headerRow!=='boolean'||typeof table.totalsRow!=='boolean'||b.r2-b.r1+1<(table.headerRow?1:0)+(table.totalsRow?1:0))throw Error('표의 머리글·요약 행이 올바르지 않습니다.');
    if(table.columns.some(c=>!object(c)||typeof c.name!=='string'||!c.name||c.name.length>32767||c.id!==undefined&&(!Number.isInteger(c.id)||c.id<1)))throw Error('표의 열 이름이 올바르지 않습니다.');
    for(const column of table.columns)if(column.filterButton!==undefined&&typeof column.filterButton!=='boolean')throw Error('표 필터 버튼 설정이 올바르지 않습니다.');
    const columns=new Set();for(const c of table.columns){const key=c.name.toLocaleLowerCase();if(columns.has(key))throw Error('표의 열 이름은 중복될 수 없습니다.');columns.add(key);for(const field of ['formula','totalsFormula','totalsFunction','totalsLabel'])if(c[field]!==undefined&&(typeof c[field]!=='string'||c[field].length>32768))throw Error('표의 계산 열이 올바르지 않습니다.');}
    if(table.style!==undefined){if(!object(table.style))throw Error('표 스타일이 올바르지 않습니다.');if(table.style.name!==undefined&&(typeof table.style.name!=='string'||table.style.name.length>255))throw Error('표 스타일 이름이 올바르지 않습니다.');for(const key of ['showFirstColumn','showLastColumn','showRowStripes','showColumnStripes'])if(table.style[key]!==undefined&&typeof table.style[key]!=='boolean')throw Error('표 스타일 옵션이 올바르지 않습니다.');}
    if(table.filter!==undefined)validateAutoFilter(table.filter);
    if(table.sourceParts!==undefined)validateStoredParts(table.sourceParts);if(table.sourceRels!==undefined)validateStoredRelationships(table.sourceRels,table.sourceParts);
    if(table.sourceXML!==undefined)validatePreservedXML(table.sourceXML,'table');
    if(ranges.some(range=>overlaps(range,b)))throw Error('표의 범위는 서로 겹칠 수 없습니다.');ids.add(table.id);names.add(table.name.toLocaleLowerCase());ranges.push(b);
  }
}

export function filteredRowsForSheet(sheet){return new Set([...(sheet.autoFilter?.hiddenRows||[]),...(sheet.tables||[]).flatMap(t=>t.filter?.hiddenRows||[])]);}

export function tableAt(sheet,r,c){return (sheet.tables||[]).find(table=>{const b=tableBounds(table.range);return r>=b.r1&&r<=b.r2&&c>=b.c1&&c<=b.c2;});}
