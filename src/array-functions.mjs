import {MAX_FORMULA_CELLS} from './limits.mjs';

const functions=new Set(['FILTER','SORT','UNIQUE','SEQUENCE']);

/** Dynamic-array functions return values, not worksheet cells. The evaluator
 * owns spill placement and keeps these results out of the editable model. */
export function arrayFunction(name,nodes,evaluate,h){
  if(!functions.has(name))return null;
  const {BLANK,FormulaError,fail,asRange,checked,numeric,logical,finite,compare}=h;
  const arity=(min,max)=>{if(nodes.length<min||nodes.length>max)fail('#VALUE!');};
  const optional=(i,fallback)=>!nodes[i]||nodes[i].omitted?fallback:evaluate(nodes[i]);
  const result=(rows,cols,values)=>{
    if(!Number.isInteger(rows)||!Number.isInteger(cols)||rows<1||cols<1)fail('#CALC!');
    if(rows*cols>MAX_FORMULA_CELLS)fail('#NUM!');
    return {value:{kind:'range',dynamic:true,rows,cols,values:values.map(v=>v===BLANK?0:v)}};
  };
  if(name==='SEQUENCE'){
    arity(1,4);
    const rows=Math.trunc(numeric(optional(0,1))),cols=Math.trunc(numeric(optional(1,1)));
    if(rows<1||cols<1)fail('#VALUE!');
    if(rows*cols>MAX_FORMULA_CELLS||rows>1048576||cols>16384)fail('#NUM!');
    const start=numeric(optional(2,1)),step=numeric(optional(3,1));
    const values=Array.from({length:rows*cols},(_,i)=>finite(start+i*step));
    return result(rows,cols,values);
  }
  arity(name==='FILTER'?2:1,name==='FILTER'?3:name==='SORT'?4:3);
  const source=asRange(evaluate(nodes[0]));
  if(source.rows*source.cols>MAX_FORMULA_CELLS)fail('#NUM!');
  if(name==='FILTER'){
    arity(2,3);
    const include=asRange(evaluate(nodes[1]));
    const byRow=include.rows===source.rows&&include.cols===1;
    const byCol=include.rows===1&&include.cols===source.cols;
    if(!byRow&&!byCol)fail('#VALUE!');
    const selected=[];
    include.values.forEach((value,i)=>{if(logical(value))selected.push(i);});
    if(!selected.length){
      if(nodes.length<3)fail('#CALC!');
      const empty=asRange(evaluate(nodes[2]));return result(empty.rows,empty.cols,empty.values);
    }
    const rows=byRow?selected.length:source.rows,cols=byRow?source.cols:selected.length,values=[];
    for(let r=0;r<rows;r++)for(let c=0;c<cols;c++)values.push(source.values[(byRow?selected[r]:r)*source.cols+(byRow?c:selected[c])]);
    return result(rows,cols,values);
  }
  if(name==='SORT'){
    arity(1,4);
    const byCol=logical(optional(3,false)),index=Math.trunc(numeric(optional(1,1)))-1,order=numeric(optional(2,1));
    if(index<0||index>=(byCol?source.rows:source.cols)||![1,-1].includes(order))fail('#VALUE!');
    const count=byCol?source.cols:source.rows,indices=Array.from({length:count},(_,i)=>i);
    const key=i=>source.values[byCol?index*source.cols+i:i*source.cols+index];
    indices.sort((a,b)=>{
      const x=key(a),y=key(b);
      // Physical blanks sort after populated values in either direction. They
      // become zero only in the returned array, after the sort is complete.
      if(x===BLANK||y===BLANK)return x===y?a-b:x===BLANK?1:-1;
      let delta;
      if(x instanceof FormulaError||y instanceof FormulaError)delta=x instanceof FormulaError&&y instanceof FormulaError?x.code.localeCompare(y.code):x instanceof FormulaError?1:-1;
      else delta=compare(x,y);
      return delta*order||a-b;
    });
    const values=[];
    for(let r=0;r<source.rows;r++)for(let c=0;c<source.cols;c++)values.push(source.values[(byCol?r:indices[r])*source.cols+(byCol?indices[c]:c)]);
    return result(source.rows,source.cols,values);
  }
  arity(1,3);
  const byCol=logical(optional(1,false)),once=logical(optional(2,false));
  const count=byCol?source.cols:source.rows,length=byCol?source.rows:source.cols,keys=new Map();
  const identity=value=>{
    if(value===BLANK)return ['blank'];
    if(value instanceof FormulaError)return ['error',value.code];
    return [typeof value,typeof value==='string'?value.toLocaleLowerCase():value];
  };
  for(let i=0;i<count;i++){
    const values=Array.from({length},(_,j)=>identity(source.values[byCol?j*source.cols+i:i*source.cols+j]));
    const key=JSON.stringify(values),entry=keys.get(key);
    if(entry)entry.count++;else keys.set(key,{index:i,count:1});
  }
  const indices=[...keys.values()].filter(entry=>!once||entry.count===1).map(entry=>entry.index);
  if(!indices.length)fail('#CALC!');
  const rows=byCol?source.rows:indices.length,cols=byCol?indices.length:source.cols,values=[];
  for(let r=0;r<rows;r++)for(let c=0;c<cols;c++)values.push(source.values[(byCol?r:indices[r])*source.cols+(byCol?indices[c]:c)]);
  return result(rows,cols,values);
}
