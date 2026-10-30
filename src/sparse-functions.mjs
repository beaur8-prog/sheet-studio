/** Whole-row/column aggregations retain their logical extent while visiting
 * only populated cells. They never materialize millions of trailing blanks.
 */
export function sparseFunction(name,nodes,evaluate,h){
  const supported=new Set(['SUM','AVERAGE','MIN','MAX','COUNT','COUNTA','COUNTBLANK','COUNTIF','SUMIF','AVERAGEIF','COUNTIFS','SUMIFS','AVERAGEIFS','MINIFS','MAXIFS','VLOOKUP','XLOOKUP','MATCH','XMATCH','SUBTOTAL']);
  if(!supported.has(name))return null;
  const {BLANK,FormulaError,fail,checked,numeric,logical,finite,asRange,numericItems,criterion,compare,wildcardRegex,metadata,readRange,readCell,maxCells,aggregate,address}=h;
  const rangeArgument=i=>['SUM','AVERAGE','MIN','MAX','COUNT','COUNTA'].includes(name)||name==='SUBTOTAL'&&i>0||name==='COUNTBLANK'&&i===0||name==='COUNTIF'&&i===0||['SUMIF','AVERAGEIF'].includes(name)&&[0,2].includes(i)||name==='COUNTIFS'&&i%2===0||['SUMIFS','AVERAGEIFS','MINIFS','MAXIFS'].includes(name)&&(i===0||i%2===1)||['VLOOKUP','MATCH','XMATCH'].includes(name)&&i===1||name==='XLOOKUP'&&[1,2].includes(i);
  const refs=nodes.map((node,i)=>{if(!rangeArgument(i)||!['ref','range','whole','referenceOp','name','call'].includes(node.type))return null;try{return metadata(node);}catch(error){if(error instanceof FormulaError)return null;throw error;}});
  if(!refs.some(ref=>ref?.sheet&&(ref.entire||ref.rows*ref.cols>maxCells)))return null;
  const arity=(min,max=min)=>{if(nodes.length<min||nodes.length>max)fail('#VALUE!');},integer=index=>Math.trunc(numeric(evaluate(nodes[index]))),value=index=>evaluate(nodes[index]);
  const range=index=>{if(refs[index]?.sheet)return readRange(refs[index]);const r=asRange(value(index));return {...r,indices:r.values.map((_,i)=>i)};};
  const map=range=>range.entries??=new Map(range.indices.map((index,i)=>[index,range.values[i]]));
  const read=(range,index)=>map(range).get(index)??BLANK;
  const result=value=>({value});
  if(['SUM','AVERAGE','MIN','MAX','COUNT','COUNTA'].includes(name)){
    arity(1,255);let sum=0,count=0,nonempty=0,minimum=Infinity,maximum=-Infinity;const regular=[];
    for(let i=0;i<nodes.length;i++){
      if(!refs[i]?.sheet){try{regular.push(value(i));}catch(error){if(['COUNT','COUNTA'].includes(name)&&error instanceof FormulaError)regular.push(error);else throw error;}continue;}
      const r=range(i);if(!r.summary){const summary={sum:0,count:0,nonempty:0,minimum:Infinity,maximum:-Infinity};for(const v of r.values){if(v!==BLANK)summary.nonempty++;if(v instanceof FormulaError){summary.error??=v;continue;}if(typeof v==='number'){summary.sum+=v;summary.count++;summary.minimum=Math.min(summary.minimum,v);summary.maximum=Math.max(summary.maximum,v);}}r.summary=summary;}
      const s=r.summary;if(s.error&&!['COUNT','COUNTA'].includes(name))checked(s.error);sum+=s.sum;count+=s.count;nonempty+=s.nonempty;minimum=Math.min(minimum,s.minimum);maximum=Math.max(maximum,s.maximum);
    }
    if(name==='COUNTA')return result(nonempty+(regular.length?aggregate(name,regular):0));if(name==='COUNT')return result(count+(regular.length?aggregate(name,regular):0));
    const numbers=numericItems(regular);for(const n of numbers){sum+=n;count++;minimum=Math.min(minimum,n);maximum=Math.max(maximum,n);}
    return result(name==='MIN'?count?minimum:0:name==='MAX'?count?maximum:0:name==='AVERAGE'?count?finite(sum/count):fail('#DIV/0!'):finite(sum));
  }
  if(name==='COUNTBLANK'){arity(1);const r=range(0);return result(r.rows*r.cols-r.values.filter(v=>v!==BLANK&&v!=='').length);}
  if(['COUNTIF','SUMIF','AVERAGEIF','COUNTIFS','SUMIFS','AVERAGEIFS','MINIFS','MAXIFS'].includes(name)){
    const single=['COUNTIF','SUMIF','AVERAGEIF'].includes(name),count=name.startsWith('COUNT'),average=name.startsWith('AVERAGE'),minimum=name==='MINIFS',maximum=name==='MAXIFS';
    if(single)arity(2,count?2:3);else{arity(count?2:3,255);if((nodes.length-(count?0:1))%2)fail('#VALUE!');}
    const checks=[],options=name==='AVERAGEIFS'?{blankAsZero:true,booleanAsNumber:true}:{};
    if(single)checks.push([range(0),criterion(value(1),options)]);else for(let i=count?0:1;i<nodes.length;i+=2)checks.push([range(i),criterion(value(i+1),options)]);
    let target=count?checks[0][0]:single?range(nodes.length>2?2:0):range(0);
    if(single&&!count&&refs[nodes.length>2?2:0]?.sheet){const ref=refs[nodes.length>2?2:0];target=readRange({...ref,rows:checks[0][0].rows,cols:checks[0][0].cols});}
    if(checks.some(([r])=>r.rows!==target.rows||r.cols!==target.cols))fail('#VALUE!');
    const indexes=new Set(checks.flatMap(([r])=>r.indices));if(!count)for(const i of target.indices)indexes.add(i);
    const matchesBlank=checks.every(([,test])=>test(BLANK)),implicit=matchesBlank?target.rows*target.cols-indexes.size:0;let matched=count?implicit:0,total=0,low=Infinity,high=-Infinity;
    for(const index of indexes){if(!checks.every(([r,test])=>test(read(r,index))))continue;if(count){matched++;continue;}let v=checked(read(target,index));if(['SUMIFS','AVERAGEIFS'].includes(name)&&typeof v==='boolean')v=Number(v);if(typeof v!=='number')continue;total+=v;matched++;low=Math.min(low,v);high=Math.max(high,v);}
    return result(count?matched:average?matched?finite(total/matched):fail('#DIV/0!'):minimum?matched?low:0:maximum?matched?high:0:finite(total));
  }
  if(name==='VLOOKUP'){
    arity(3,4);const lookup=checked(value(0)),r=range(1),column=integer(2),approximate=nodes.length<4||logical(value(3));if(column<1)fail('#VALUE!');if(column>r.cols)fail('#REF!');const pattern=typeof lookup==='string'?wildcardRegex(lookup):null;let selected=-1;
    for(let i=0;i<r.values.length;i++){if(r.indices[i]%r.cols)continue;const v=r.values[i];if(v instanceof FormulaError||v===BLANK||typeof v!==typeof lookup)continue;if(!approximate&&(pattern?pattern.test(v):compare(v,lookup)===0)){selected=Math.floor(r.indices[i]/r.cols);break;}if(approximate&&compare(v,lookup)<=0)selected=Math.floor(r.indices[i]/r.cols);}
    if(selected<0)fail('#N/A');const v=checked(readCell(r.sheet,address(r.top+selected,r.left+column-1)));return result(v===BLANK?0:v);
  }
  if(['XLOOKUP','MATCH','XMATCH'].includes(name)){
    arity(2,name==='XLOOKUP'?6:name==='XMATCH'?4:3);if(name==='XLOOKUP'&&nodes.length<3)fail('#VALUE!');const lookup=checked(value(0)),r=range(1);if(r.rows>1&&r.cols>1)fail(name==='XLOOKUP'?'#VALUE!':'#N/A');let mode,search,wildcard;
    if(name==='MATCH'){const type=nodes.length>2?integer(2):1;if(![-1,0,1].includes(type))fail('#N/A');mode=-type;search=type===0?1:-1;wildcard=type===0;}
    else{mode=nodes.length>(name==='XLOOKUP'?4:2)?integer(name==='XLOOKUP'?4:2):0;search=nodes.length>(name==='XLOOKUP'?5:3)?integer(name==='XLOOKUP'?5:3):1;wildcard=mode===2;}
    if(![0,-1,1,2].includes(mode)||![1,-1,2,-2].includes(search))fail('#VALUE!');const pattern=typeof lookup==='string'&&wildcard?wildcardRegex(lookup):null;let selected=-1,best;
    for(let j=0;j<r.values.length;j++){const i=search<0?r.values.length-1-j:j,v=r.values[i];if(v instanceof FormulaError||v===BLANK||typeof v!==typeof lookup)continue;const order=compare(v,lookup);if(pattern?pattern.test(v):order===0){selected=r.indices[i];break;}if(mode===-1&&order<0||mode===1&&order>0)if(best===undefined||(mode===-1?compare(v,best)>0:compare(v,best)<0)){best=v;selected=r.indices[i];}}
    if(name!=='XLOOKUP')return result(selected<0?fail('#N/A'):selected+1);
    const out=refs[2];if(!out?.sheet){const target=range(2);if(target.rows*target.cols!==r.rows*r.cols)fail('#VALUE!');if(selected<0)return result(nodes.length>3&&!nodes[3].omitted?checked(value(3)):fail('#N/A'));const v=checked(read(target,selected));return result(v===BLANK?0:v);}
    if(out.rows*out.cols!==r.rows*r.cols)fail('#VALUE!');if(selected<0)return result(nodes.length>3&&!nodes[3].omitted?checked(value(3)):fail('#N/A'));const v=checked(readCell(out.sheet,address(out.top+Math.floor(selected/out.cols),out.left+selected%out.cols)));return result(v===BLANK?0:v);
  }
  if(name==='SUBTOTAL'){
    arity(2,255);const code=integer(0),n=code>100?code-100:code,target=['','AVERAGE','COUNT','COUNTA','MAX','MIN','PRODUCT','STDEV.S','STDEV.P','SUM','VAR.S','VAR.P'][n];if(!target||code!==n&&code!==n+100)fail('#VALUE!');const values=[];
    for(let i=1;i<nodes.length;i++){const ref=refs[i];if(!ref?.sheet)fail('#VALUE!');const r=readRange(ref,{subtotalCode:code});for(const v of r.values)values.push(v);}
    return result(aggregate(target,[{kind:'range',values,rows:values.length,cols:1}]));
  }
  return null;
}
