import {address,parseAddress} from './engine.mjs';

const LIMIT_ROWS=1048576,LIMIT_COLS=16384;
export const referenceColors=['#2463d3','#c23b33','#8b47ad','#168153','#b16b08','#168494'];

// A tolerant scanner is needed while the user is still typing incomplete formulas.
// Consume strings and identifiers as whole tokens so A1 inside text or LOG10() is not a reference.
export function formulaReferences(formula){
  if(!formula.startsWith('='))return [];
  const tokens=[],pattern=/"(?:[^"]|"")*(?:"|$)|'(?:[^']|'')*(?:'|$)|[\p{L}_$\\][\p{L}\p{N}_.$\\]*|\d+(?:\.\d*)?(?:[eE][+-]?\d+)?|\S/gu;
  for(const m of formula.matchAll(pattern))tokens.push({text:m[0],start:m.index,end:m.index+m[0].length});
  const result=[];
  const read=(i,inherited)=>{
    let t=tokens[i],sheet=inherited,start=t?.start;
    if(!t)return null;
    if(tokens[i+1]?.text==='!'){sheet=t.text.startsWith("'")?t.text.slice(1,-1).replaceAll("''","'"):t.text;t=tokens[i+2];i+=2;}
    const pos=parseAddress(t?.text);if(!pos||tokens[i+1]?.text==='(')return null;
    return {sheet,pos,start,end:t.end,next:i+1};
  };
  for(let i=1;i<tokens.length&&result.length<128;){
    if(['[',']',':','!'].includes(tokens[i-1]?.text)){i++;continue;}
    const a=read(i);if(!a){i++;continue;}
    const b=tokens[a.next]?.text===':'?read(a.next+1,a.sheet):null;
    if(b&&b.sheet!==a.sheet){i=b.next;continue;}
    result.push({start:a.start,end:b?.end??a.end,sheet:a.sheet,r1:Math.min(a.pos.r,b?.pos.r??a.pos.r),r2:Math.max(a.pos.r,b?.pos.r??a.pos.r),c1:Math.min(a.pos.c,b?.pos.c??a.pos.c),c2:Math.max(a.pos.c,b?.pos.c??a.pos.c),text:formula.slice(a.start,b?.end??a.end)});i=b?.next??a.next;
  }
  return result;
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
  const raw=(r,c)=>sheet.cells[address(r,c)]?.raw||'',numeric=(r,c)=>raw(r,c)!==''&&typeof evaluate(address(r,c))==='number'&&Number.isFinite(evaluate(address(r,c)));
  const edits=[];
  const add=(r,c,a,b)=>{
    if(r>=LIMIT_ROWS||c>=LIMIT_COLS)throw Error('합계를 넣을 빈 셀이 시트 끝을 벗어납니다.');
    if(raw(r,c))throw Error('합계 위치에 데이터가 있습니다. 빈 합계 셀을 선택하세요.');
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
  if(r1===r2){const target=raw(r1,c2)===''?c2:c2+1;add(r1,target,address(r1,c1),address(r1,target-1));return edits;}
  const bottomBlank=Array.from({length:c2-c1+1},(_,i)=>raw(r2,c1+i)==='').every(Boolean);
  const rightBlank=c2>c1&&Array.from({length:r2-r1+1},(_,i)=>raw(r1+i,c2)==='').every(Boolean);
  if(bottomBlank||!rightBlank){const target=bottomBlank?r2:r2+1,lastCol=rightBlank?c2-1:c2;for(let c=c1;c<=lastCol;c++)add(target,c,address(r1,c),address(target-1,c));}
  if(rightBlank)for(let r=r1;r<=r2;r++)add(r,c2,address(r,c1),address(r,c2-1));
  return edits;
}
