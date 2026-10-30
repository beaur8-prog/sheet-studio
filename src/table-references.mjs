import {tableBounds} from './table-schema.mjs';
const escaped=new Set(["'",'[',']','#','@']);
export function readStructuredToken(input,start){
  let depth=0;
  for(let i=start;i<input.length;i++){
    if(input[i]==="'"&&escaped.has(input[i+1])){i++;continue;}
    if(input[i]==='[')depth++;
    else if(input[i]===']'&&!--depth)return {value:input.slice(start,i+1),end:i+1};
    if(depth<0)return null;
  }return null;
}
export function structuredParts(selector){
  const stack=[],leaves=[];let span=false,thisRow=false;
  for(let i=0;i<selector.length;i++){
    const char=selector[i],node=stack.at(-1);
    if(char==="'"&&escaped.has(selector[i+1])){if(!node)return null;node.escapedFirst||=node.text.trim().length===0;node.text+=selector[++i];continue;}
    if(char==='['){if(node)node.children++;stack.push({start:i,text:'',children:0,escapedFirst:false});}
    else if(char===']'){
      const done=stack.pop();if(!done)return null;done.end=i+1;
      if(!done.children)leaves.push(done);else if(done.text.includes('@'))thisRow=true;
    }else if(node){if(node.children&&char===':')span=true;else node.text+=char;}
    else if(!/\s/.test(char))return null;
  }
  if(stack.length||!leaves.length)return null;
  const flags=new Set(),columns=[];
  for(const leaf of leaves){let text=leaf.text.trim();if(!text)return null;
    if(!leaf.escapedFirst&&text.startsWith('@')){thisRow=true;text=text.slice(1).trim();if(!text)continue;}
    const flag=text.toLowerCase().replace(/\s+/g,'');
    if(!leaf.escapedFirst&&['#all','#data','#headers','#totals','#thisrow'].includes(flag))flags.add(flag);else columns.push({name:text,...leaf});
  }
  if(thisRow)flags.add('#thisrow');return {flags,columns,span,leaves};
}
const includes=(b,p)=>p.r>=b.r1&&p.r<=b.r2&&p.c>=b.c1&&p.c<=b.c2;
const segments=values=>{const result=[];for(const value of [...new Set(values)].sort((a,b)=>a-b)){const last=result.at(-1);if(last&&value===last[1]+1)last[1]=value;else result.push([value,value]);}return result;};
export function resolveStructuredReference(book,owner,position,tableName,selector,explicitSheet){
  const entries=book.sheets.flatMap(sheet=>(sheet.tables||[]).map(table=>({sheet,table}))),entry=tableName?entries.find(({table})=>table.name.toLocaleLowerCase()===tableName.toLocaleLowerCase()):entries.find(({sheet,table})=>sheet.id===owner.id&&includes(tableBounds(table.range),position));
  if(!entry)return {error:tableName?'#NAME?':'#VALUE!'};
  const {sheet,table}=entry,b=tableBounds(table.range),parts=structuredParts(selector||'[#Data]');if(!parts)return {error:'#REF!'};
  if(explicitSheet&&sheet.name.toLocaleLowerCase()!==explicitSheet.toLocaleLowerCase())return {error:'#REF!'};
  const {flags,columns,span}=parts,first=b.r1+(table.headerRow?1:0),last=b.r2-(table.totalsRow?1:0);let rowSegments=[];
  if(flags.has('#thisrow')){if(flags.size>1||position.r<first||position.r>last||owner.id!==sheet.id)return {error:'#VALUE!'};rowSegments=[[position.r,position.r]];}
  else if(flags.has('#all'))rowSegments=[[b.r1,b.r2]];
  else {
    if(flags.has('#headers')){if(!table.headerRow)return {error:'#REF!'};rowSegments.push([b.r1,b.r1]);}
    if(!flags.size||flags.has('#data'))if(last>=first)rowSegments.push([first,last]);
    if(flags.has('#totals')){if(!table.totalsRow)return {error:'#REF!'};rowSegments.push([b.r2,b.r2]);}
    rowSegments.sort((a,b)=>a[0]-b[0]);for(let i=rowSegments.length-1;i>0;i--)if(rowSegments[i-1][1]+1>=rowSegments[i][0]){rowSegments[i-1][1]=Math.max(rowSegments[i-1][1],rowSegments[i][1]);rowSegments.splice(i,1);}
  }
  let columnSegments;
  if(columns.length){const indexes=columns.map(column=>table.columns.findIndex(c=>c.name.toLocaleLowerCase()===column.name.toLocaleLowerCase()));if(indexes.some(n=>n<0))return {error:'#REF!'};columnSegments=span?[[Math.min(...indexes),Math.max(...indexes)]]:segments(indexes);}
  else columnSegments=[[0,table.columns.length-1]];
  const areas=rowSegments.flatMap(([top,bottom])=>columnSegments.map(([left,right])=>({sheet,top,left:b.c1+left,rows:bottom-top+1,cols:right-left+1})));
  if(!areas.length)return {sheet,top:first,left:b.c1+columnSegments[0][0],rows:0,cols:columnSegments[0][1]-columnSegments[0][0]+1,empty:true};
  return areas.length===1?areas[0]:{areas};
}
export const escapeColumn=name=>String(name).replace(/['\[\]#@]/g,char=>"'"+char);

export function rewriteStructuredSelector(selector,renames){
  const parts=structuredParts(selector);if(!parts)return selector;const patches=[];
  for(const column of parts.columns){const match=Object.entries(renames||{}).find(([name])=>name.toLocaleLowerCase()===column.name.toLocaleLowerCase());if(!match)continue;if(match[1]===null)return null;const thisRow=!column.escapedFirst&&column.text.trim().startsWith('@');patches.push({start:column.start,end:column.end,value:thisRow?'[@['+escapeColumn(match[1])+']]':'['+escapeColumn(match[1])+']'});}
  for(const p of patches.reverse())selector=selector.slice(0,p.start)+p.value+selector.slice(p.end);return selector;
}
