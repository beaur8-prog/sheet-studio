import {address,parseAddress} from './engine.mjs';
import {wholeReferences} from './whole-references.mjs';
import {readStructuredToken} from './table-references.mjs';

const LIMIT_ROWS=1048576,LIMIT_COLS=16384;
export const referenceColors=['#2463d3','#c23b33','#8b47ad','#168153','#b16b08','#168494'];

// A tolerant scanner is needed while the user is still typing incomplete formulas.
// Consume strings and identifiers as whole tokens so A1 inside text or LOG10() is not a reference.
export function formulaReferences(formula){
  if(!formula.startsWith('='))return [];
  const tokens=[],pattern=/"(?:[^"]|"")*(?:"|$)|'(?:[^']|'')*(?:'|$)|[\p{L}\p{N}_.$\\]+(?=!)|[\p{L}_$\\][\p{L}\p{N}_.$\\]*|\d+(?:\.\d*)?(?:[eE][+-]?\d+)?|\S/gu;
  let skip=0;for(const m of formula.matchAll(pattern)){if(m.index<skip)continue;let text=m[0];if(text==='['){const structured=readStructuredToken(formula,m.index);skip=structured?.end||formula.length;text=formula.slice(m.index,skip);tokens.push({text,type:'structured',value:text,start:m.index,end:skip});continue;}const type=text.startsWith('"')?'string':text.startsWith("'")?'sheet':/^\d+(?:\.\d*)?(?:[eE][+-]?\d+)?$/.test(text)?'number':/^[\p{L}_$\\]/u.test(text)?'word':'op',value=type==='sheet'?text.slice(1,-1).replaceAll("''","'"):type==='number'?Number(text):text;tokens.push({text,type,value,start:m.index,end:m.index+text.length});}
  const result=[];
  const read=(i,inherited)=>{
    let t=tokens[i],sheet=inherited,start=t?.start;
    if(!t)return null;
    if(tokens[i+1]?.text==='!'){sheet=t.text.startsWith("'")?t.text.slice(1,-1).replaceAll("''","'"):t.text;t=tokens[i+2];i+=2;}
    const pos=parseAddress(t?.text);if(!pos||tokens[i+1]?.text==='(')return null;
    const spill=tokens[i+1]?.text==='#';return {sheet,pos,start,end:spill?tokens[i+1].end:t.end,next:i+(spill?2:1),...(spill?{spill:true}:{})};
  };
  for(let i=1;i<tokens.length&&result.length<128;){
    if(['[',']',':','!'].includes(tokens[i-1]?.text)){i++;continue;}
    const a=read(i);if(!a){i++;continue;}
    const b=tokens[a.next]?.text===':'?read(a.next+1,a.sheet):null;
    if(b&&String(b.sheet||'').toLowerCase()!==String(a.sheet||'').toLowerCase()){i=b.next;continue;}
    result.push({start:a.start,end:b?.end??a.end,sheet:a.sheet,r1:Math.min(a.pos.r,b?.pos.r??a.pos.r),r2:Math.max(a.pos.r,b?.pos.r??a.pos.r),c1:Math.min(a.pos.c,b?.pos.c??a.pos.c),c2:Math.max(a.pos.c,b?.pos.c??a.pos.c),text:formula.slice(a.start,b?.end??a.end),...(a.spill&&!b?{spill:true}:{})});i=b?.next??a.next;
  }
  for(const range of wholeReferences(tokens)){const {first,last,axis,limit}=range;if(first.index>=limit||last.index>=limit||String(first.sheet||'').toLocaleLowerCase()!==String(last.sheet||'').toLocaleLowerCase())continue;const low=Math.min(first.index,last.index),high=Math.max(first.index,last.index);result.push({start:first.start,end:last.token.end,sheet:first.sheet,r1:axis==='c'?0:low,r2:axis==='c'?LIMIT_ROWS-1:high,c1:axis==='c'?low:0,c2:axis==='c'?high:LIMIT_COLS-1,text:formula.slice(first.start,last.token.end),axis});}
  return result.sort((a,b)=>a.start-b.start).slice(0,128);
}

export function referenceSlot(formula,start,end=start){
  if(!formula.startsWith('='))return null;
  const ref=formulaReferences(formula).find(r=>start>=r.start&&end<=r.end);
  if(ref)return {start:ref.start,end:ref.end};
  // Do not insert addresses into quoted strings, including unfinished strings.
  const prefix=formula.slice(0,start).replace(/"(?:[^"]|"")*"/g,'').replace(/'(?:[^']|'')*'/g,'');
  if(prefix.includes('"')||prefix.includes("'"))return null;
  return /[=+\-*/^&<>,;(:]\s*$/.test(prefix)?{start,end}:null;
}

export function cycleReference(formula,start,end=start){
  const slot=referenceSlot(formula,start,end);if(!slot||slot.start===slot.end)return null;
  const whole=formulaReferences(formula).find(r=>r.axis&&r.start===slot.start&&r.end===slot.end);if(whole){const text=whole.text,coordinates=text.replace(/'(?:[^']|'')*'!|[\p{L}\p{N}_.$\\]+!/gu,''),locked=coordinates.trim().startsWith('$'),replacement=text.replace(/'(?:[^']|'')*'(?=!)|[\p{L}\p{N}_.$\\]+(?=!)|\$?([A-Z]{1,3}|[1-9]\d*)/giu,(m,coordinate)=>coordinate?(locked?'':'$')+coordinate.toUpperCase():m);return {formula:formula.slice(0,slot.start)+replacement+formula.slice(slot.end),start:slot.start,end:slot.start+replacement.length};}
  const text=formula.slice(slot.start,slot.end),addresses=text.replace(/'(?:[^']|'')*'!|[\p{L}_][\p{L}\p{N}_.]*!/gu,''),first=/^(\$?)([A-Z]{1,3})(\$?)([1-9]\d*)/i.exec(addresses);
  if(!first)return null;
  const state=(first[1]?2:0)+(first[3]?1:0),next={0:3,3:1,1:2,2:0}[state];
  const replacement=text.replace(/'(?:[^']|'')*'(?=!)|[\p{L}_][\p{L}\p{N}_.]*(?=!)|\$?([A-Z]{1,3})\$?([1-9]\d*)/giu,(m,col,row)=>col?(next&2?'$':'')+col.toUpperCase()+(next&1?'$':'')+row:m);
  return {formula:formula.slice(0,slot.start)+replacement+formula.slice(slot.end),start:slot.start,end:slot.start+replacement.length};
}

export function rangeReference(range,ownerName,viewName){
  const first=address(Math.min(range.r1,range.r2),Math.min(range.c1,range.c2)),last=address(Math.max(range.r1,range.r2),Math.max(range.c1,range.c2));
  return (ownerName===viewName?'':"'"+viewName.replaceAll("'","''")+"'!")+first+(first===last?'':':'+last);
}

export function completeFormula(formula){
  if(!formula.startsWith('='))return formula;
  let quote=null,depth=0;
  for(let i=0;i<formula.length;i++){const c=formula[i];if(quote){if(c===quote){if(formula[i+1]===quote)i++;else quote=null;}continue;}if(c==='"'||c==="'"){quote=c;continue;}if(c==='(')depth++;if(c===')'&&--depth<0)return formula;}
  return !quote&&depth>0&&depth<=32?formula+')'.repeat(depth):formula;
}

export function autoSumPlan(sheet,selection,evaluate){
  const r1=Math.min(selection.r,selection.r2),r2=Math.max(selection.r,selection.r2),c1=Math.min(selection.c,selection.c2),c2=Math.max(selection.c,selection.c2);
  if((r2-r1+1)*(c2-c1+1)>100000)throw Error('자동 합계는 한 번에 최대 100,000개 셀을 선택하세요.');
  const occupied=(r,c)=>!!sheet.cells[address(r,c)]?.raw||!!evaluate.isComputed?.(address(r,c)),numeric=(r,c)=>{if(!occupied(r,c))return false;const value=evaluate(address(r,c));return typeof value==='number'&&Number.isFinite(value);};
  const edits=[];
  const add=(r,c,a,b)=>{
    if(r>=LIMIT_ROWS||c>=LIMIT_COLS)throw Error('합계를 넣을 빈 셀이 시트 끝을 벗어납니다.');
    if(occupied(r,c))throw Error('합계 위치에 데이터가 있습니다. 빈 합계 셀을 선택하세요.');
    edits.push({r,c,raw:`=SUM(${a}:${b})`});
  };
  if(r1===r2&&c1===c2){
    for(const [dr,dc] of [[-1,0],[0,-1]]){
      let r=r1+dr,c=c1+dc;if(r<0||c<0||!numeric(r,c))continue;
      const last=address(r,c);let first=last,steps=0;
      while(r>=0&&c>=0&&numeric(r,c)){if(steps++>=100000)throw Error('자동 합계 범위는 최대 100,000개 셀입니다.');first=address(r,c);r+=dr;c+=dc;}
      return [{r:r1,c:c1,raw:`=SUM(${first}:${last})`}];
    }
    return [{r:r1,c:c1,raw:'=SUM()'}];
  }
  if(r1===r2){const target=!occupied(r1,c2)?c2:c2+1;add(r1,target,address(r1,c1),address(r1,target-1));return edits;}
  const bottomBlank=Array.from({length:c2-c1+1},(_,i)=>!occupied(r2,c1+i)).every(Boolean);
  const rightBlank=c2>c1&&Array.from({length:r2-r1+1},(_,i)=>!occupied(r1+i,c2)).every(Boolean);
  if(bottomBlank||!rightBlank){const target=bottomBlank?r2:r2+1,lastCol=rightBlank?c2-1:c2;for(let c=c1;c<=lastCol;c++)add(target,c,address(r1,c),address(target-1,c));}
  if(rightBlank)for(let r=r1;r<=r2;r++)add(r,c2,address(r,c1),address(r,c2-1));
  return edits;
}
