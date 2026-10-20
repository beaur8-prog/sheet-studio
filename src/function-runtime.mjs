import * as library from '@formulajs/formulajs';
import {FUNCTION_MAP, FUNCTION_ALIASES} from './function-catalog.mjs';

/** Bridge Excel reference/error/blank semantics to the pinned MIT Formula.js library.
 * The parser never executes JavaScript, and only catalogued names can enter this bridge.
 */
export function createFunctionRuntime(h) {
  const {BLANK, FormulaError, fail, scalar, checked, numeric, textual, logical, finite,
    isRange, isRef, asRange, flatten, numericItems, criterion, compare, wildcardRegex,
    dateSerial, serialParts, rounded, colName, parseAddress, address, formatValue} = h;
  const range = (values, rows=values.length, cols=1) => ({kind:'range',values,rows,cols});
  const raw = value => value === BLANK ? null : value;
  const sum = values => values.reduce((total,value) => total + value,0);
  const combination=(n,k)=>{
    n=Math.trunc(n);k=Math.trunc(k);if(n<0||k<0||k>n||n>Number.MAX_SAFE_INTEGER)fail('#NUM!');k=Math.min(k,n-k);if(k>10000)fail('#NUM!');
    let result=1;for(let i=1;i<=k;i++){result*=((n-k+i)/i);if(!Number.isFinite(result))fail('#NUM!');}return result;
  };
  const integer = v => Math.trunc(numeric(v));
  const day = v => {const n=Math.floor(numeric(v));serialParts(n);return n;};
  const weekday = serial => ((serial+6)%7+7)%7; // Excel's fictitious 1900-02-29 is intentional.
  const monthDays = (y,m) => y===1900&&m===2?29:new Date(Date.UTC(y,m,0)).getUTCDate();
  const serial = (y,m,d) => y===1900&&m===2&&d===29?60:dateSerial(y,m,d);
  const serialDate = v => {
    v=checked(v);if(typeof v!=='string')return day(v);
    const s=v.trim();let m=/^(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})(?:[ T].*)?$/.exec(s),y,mo,d;
    if(m){[,y,mo,d]=m.map(Number);}else{
      m=/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/.exec(s);
      if(m){mo=+m[1];d=+m[2];y=+m[3];}else return fail('#VALUE!');
    }
    if(y<1900||y>9999||mo<1||mo>12||d<1||d>monthDays(y,mo))fail('#VALUE!');
    return serial(y,mo,d);
  };
  const aNumbers = (args,name) => args.flatMap(arg=>{
    const array=isRange(arg)&&!arg.sheet,reference=isRef(arg)||isRange(arg);
    return (isRange(arg)?arg.values:[scalar(arg)]).flatMap(item=>{
      const value=checked(item);if(value===BLANK)return [];
      if(array&&(typeof value==='boolean'||typeof value==='string'))return [];
      return [typeof value==='number'?value:typeof value==='boolean'?+value:reference?0:numeric(value)];
    });
  });
  const reduceStats = (name,numbers) => {
    const n=numbers.length,mean=n?sum(numbers)/n:0;
    if(name==='PRODUCT')return n?numbers.reduce((a,b)=>a*b,1):0;
    if(name==='SUMSQ')return sum(numbers.map(n=>n*n));
    if(['AVERAGEA','AVEDEV'].includes(name)&&!n)fail('#DIV/0!');
    if(name==='AVERAGEA')return mean;
    if(name==='MAXA')return n?numbers.reduce((a,b)=>Math.max(a,b)):0;
    if(name==='MINA')return n?numbers.reduce((a,b)=>Math.min(a,b)):0;
    if(name==='AVEDEV')return sum(numbers.map(n=>Math.abs(n-mean)))/n;
    const ss=sum(numbers.map(n=>(n-mean)**2));
    if(name==='DEVSQ')return ss;
    const sample=['VAR.S','VARA','STDEV.S','STDEVA'].includes(name);
    if(n<1+Number(sample))fail('#DIV/0!');
    const variance=ss/(n-Number(sample));
    return name.startsWith('STDEV')?Math.sqrt(variance):variance;
  };
  const aggregateNames=new Set(['PRODUCT','SUMSQ','AVEDEV','DEVSQ','MEDIAN','MODE.SNGL','GEOMEAN','HARMEAN','KURT','SKEW','SKEW.P','VAR.S','VAR.P','STDEV.S','STDEV.P']);
  const aNames=new Set(['AVERAGEA','MAXA','MINA','VARA','VARPA','STDEVA','STDEVPA']);
  const pairNames=new Set(['CORREL','PEARSON','COVARIANCE.P','COVARIANCE.S','RSQ','SLOPE','INTERCEPT','STEYX','SUMX2MY2','SUMX2PY2','SUMXMY2']);
  const arrayFirst=new Set(['LARGE','SMALL','PERCENTILE.INC','PERCENTILE.EXC','PERCENTRANK.INC','PERCENTRANK.EXC','QUARTILE.INC','QUARTILE.EXC','TRIMMEAN','Z.TEST']);
  const scalarFn = name => {
    if(name==='FORECAST.LINEAR')return library.FORECAST;
    if(name==='ERF.PRECISE')return library.ERF;
    if(name==='ERFC.PRECISE')return library.ERFC;
    let fn=library;for(const part of name.split('.'))fn=fn?.[part];
    return typeof fn==='function'?fn:null;
  };
  const libraryCall = (name,args) => {
    const fn=scalarFn(name);if(!fn)fail('#NAME?');
    const result=fn(...args);
    if(result instanceof Error)fail(/^#/.test(result.message)?result.message:'#VALUE!');
    if(typeof result==='number')return finite(result);
    if(typeof result==='string')return name==='BASE'||name.endsWith('2HEX')?result.toUpperCase():result;
    if(typeof result==='boolean')return result;
    fail('#VALUE!');
  };
  const capture = (node,evaluate) => {try{return scalar(evaluate(node));}catch(e){if(e instanceof FormulaError)return e;throw e;}};
  function lookupIndex(lookup,values,mode=0,search=1,wildcard=false) {
    if(![0,-1,1,2].includes(mode)||![1,-1,2,-2].includes(search))fail('#VALUE!');
    let selected=-1,best;
    for(let j=0;j<values.length;j++){
      const i=search<0?values.length-1-j:j,v=values[i];if(v instanceof FormulaError||v===BLANK||typeof v!==typeof lookup)continue;
      const order=compare(v,lookup),pattern=typeof lookup==='string'&&(wildcard||mode===2)?wildcardRegex(lookup):null;
      if(pattern?typeof v==='string'&&pattern.test(v):order===0)return i;
      if((mode===-1&&order<0)||(mode===1&&order>0)){
        if(best===undefined||(mode===-1?compare(v,best)>0:compare(v,best)<0)){selected=i;best=v;}
      }
    }
    return selected;
  }
  function businessDays(name,args) {
    const intl=name.endsWith('.INTL'),work=name.startsWith('WORKDAY'),start=serialDate(args[0]);
    let end=work?integer(args[1]):serialDate(args[1]);
    const weekend=intl&&args.length>2?checked(args[2]):1;
    let mask;
    if(typeof weekend==='string'){
      if(!/^[01]{7}$/.test(weekend))fail('#VALUE!');
      mask=[weekend[6],...weekend.slice(0,6)].map(ch=>ch==='1');
    }else{
      const w=integer(weekend);mask=Array(7).fill(false);
      if(w>=1&&w<=7){mask[(w+5)%7]=true;mask[(w+6)%7]=true;}
      else if(w>=11&&w<=17)mask[(w-11)%7]=true;else fail('#NUM!');
    }
    const holidayArg=args[intl?3:2],holidays=new Set(holidayArg===undefined?[]:asRange(holidayArg).values.filter(v=>v!==BLANK).map(serialDate));
    const valid=n=>!mask[weekday(n)]&&!holidays.has(n);
    if(work){
      if(Math.abs(end)>100000||mask.every(Boolean))fail('#NUM!');
      let current=start,remaining=Math.abs(end),direction=Math.sign(end),steps=0;
      while(remaining){if(++steps>700000)fail('#NUM!');current+=direction;if(current<0||current>2958465)fail('#NUM!');if(valid(current))remaining--;}
      return current;
    }
    const direction=end<start?-1:1,lo=Math.min(start,end),hi=Math.max(start,end),length=hi-lo+1;
    let count=Math.floor(length/7)*(7-mask.filter(Boolean).length);
    for(let i=0;i<length%7;i++)if(!mask[weekday(lo+i)])count++;
    for(const d of holidays)if(d>=lo&&d<=hi&&!mask[weekday(d)])count--;
    return count*direction;
  }
  return function call(inputName,nodes,evaluate,ctx) {
    const metadata=FUNCTION_MAP.get(inputName);if(!metadata)fail('#NAME?');
    if(nodes.length<metadata.min||nodes.length>metadata.max)fail('#VALUE!');
    const name=FUNCTION_ALIASES[inputName]||inputName;
    const omitted=i=>!nodes[i]||nodes[i].omitted||nodes[i].type==='literal'&&nodes[i].value===BLANK;
    if(name==='IFNA'){
      try{return checked(evaluate(nodes[0]));}catch(e){if(!(e instanceof FormulaError)||e.code!=='#N/A')throw e;return evaluate(nodes[1]);}
    }
    if(name==='IFS'){
      if(nodes.length%2)fail('#VALUE!');
      for(let i=0;i<nodes.length;i+=2)if(logical(evaluate(nodes[i])))return evaluate(nodes[i+1]);
      fail('#N/A');
    }
    if(name==='SWITCH'){
      const value=checked(evaluate(nodes[0]));
      for(let i=1;i+1<nodes.length;i+=2)if(compare(value,evaluate(nodes[i]))===0)return evaluate(nodes[i+1]);
      return nodes.length%2===0?evaluate(nodes.at(-1)):fail('#N/A');
    }
    if(name==='CHOOSE'){
      const index=integer(evaluate(nodes[0]));if(index<1||index>=nodes.length)fail('#VALUE!');return evaluate(nodes[index]);
    }
    if(['ISERROR','ISERR','ISNA','ISBLANK','ISNUMBER','ISTEXT','ISNONTEXT','ISLOGICAL','ERROR.TYPE','TYPE','N','T'].includes(name)){
      if(name==='TYPE'){try{const v=evaluate(nodes[0]);if(isRange(v))return 64;}catch{}}
      const value=capture(nodes[0],evaluate),error=value instanceof FormulaError;
      if(name==='ISERROR')return error;
      if(name==='ISERR')return error&&value.code!=='#N/A';
      if(name==='ISNA')return error&&value.code==='#N/A';
      if(name==='ISBLANK')return value===BLANK;
      if(name==='ISNUMBER')return typeof value==='number';
      if(name==='ISTEXT')return typeof value==='string';
      if(name==='ISNONTEXT')return typeof value!=='string';
      if(name==='ISLOGICAL')return typeof value==='boolean';
      if(name==='TYPE')return error?16:typeof value==='string'?2:typeof value==='boolean'?4:isRange(value)?64:1;
      if(name==='ERROR.TYPE')return error?({'#NULL!':1,'#DIV/0!':2,'#VALUE!':3,'#REF!':4,'#NAME?':5,'#NUM!':6,'#N/A':7,'#GETTING_DATA':8})[value.code]||fail('#N/A'):fail('#N/A');
      checked(value);return name==='N'?(typeof value==='number'?value:typeof value==='boolean'?+value:0):typeof value==='string'?value:'';
    }
    if(name==='NA')fail('#N/A');
    if(name==='TRUE'||name==='FALSE')return name==='TRUE';
    if(name==='ISREF'){try{const v=evaluate(nodes[0]);return !!(isRef(v)||isRange(v)&&v.sheet);}catch{return false;}}
    if(['ROW','COLUMN','ROWS','COLUMNS','FORMULATEXT','ISFORMULA','OFFSET','INDIRECT','ADDRESS','SHEET','SHEETS','SUBTOTAL'].includes(name))return ctx.referenceFunction(name,nodes,evaluate);
    if(name==='TEXTAFTER'||name==='TEXTBEFORE'){
      const source=textual(evaluate(nodes[0])),delimiter=textual(evaluate(nodes[1]));
      const instance=!omitted(2)?integer(evaluate(nodes[2])):1,mode=!omitted(3)?integer(evaluate(nodes[3])):0,end=!omitted(4)?integer(evaluate(nodes[4])):0;
      if(!instance||![0,1].includes(mode)||![0,1].includes(end))fail('#VALUE!');
      if(!source)return '';if(Math.abs(instance)>source.length)fail('#VALUE!');
      const s=mode?source.toLowerCase():source,d=mode?delimiter.toLowerCase():delimiter;let at=-1;
      if(d==='')at=instance>0?0:source.length;
      else if(instance>0){let from=0;for(let i=0;i<instance;i++){at=s.indexOf(d,from);if(at<0){if(end&&i===instance-1)at=s.length;break;}from=at+d.length;}}
      else {let from=s.length;for(let i=0;i<-instance;i++){at=s.lastIndexOf(d,from-d.length);if(at<0){if(end&&i===-instance-1)at=-d.length;break;}from=at;}}
      if(at<0&&!(end&&at===-d.length))return nodes.length>5?evaluate(nodes[5]):fail('#N/A');
      return name==='TEXTBEFORE'?source.slice(0,Math.max(0,at)):source.slice(Math.max(0,at+delimiter.length));
    }
    const args=nodes.map(evaluate),num=i=>numeric(args[i]),str=i=>textual(args[i]);
    if(name==='XOR'){
      const values=flatten(args).filter(({value,reference})=>{checked(value);return value!==BLANK&&!(reference&&typeof value==='string');});
      if(!values.length)fail('#VALUE!');return values.filter(({value})=>logical(value)).length%2===1;
    }
    if(name==='COUNTBLANK')return asRange(args[0]).values.filter(v=>v===BLANK||v==='').length;
    if(['SUMIFS','COUNTIFS','AVERAGEIFS','MINIFS','MAXIFS','AVERAGEIF'].includes(name)){
      const single=name==='AVERAGEIF',count=name==='COUNTIFS',offset=count?0:1;
      if(!single&&(args.length-offset)%2)fail('#VALUE!');
      const result=asRange(single?(args[2]??args[0]):args[0]);
      const checks=single?[[asRange(args[0]),criterion(args[1])]]:Array.from({length:(args.length-offset)/2},(_,i)=>[asRange(args[offset+2*i]),criterion(args[offset+2*i+1])]);
      for(const [r] of checks)if(r.rows!==result.rows||r.cols!==result.cols)fail('#VALUE!');
      let total=0,matched=0,minimum=Infinity,maximum=-Infinity;
      result.values.forEach((value,i)=>{if(!checks.every(([r,test])=>test(r.values[i])))return;
        if(count){matched++;return;}checked(value);if(typeof value!=='number')return;total+=value;matched++;minimum=Math.min(minimum,value);maximum=Math.max(maximum,value);
      });
      if(count)return matched;
      if(name.startsWith('AVERAGE'))return matched?finite(total/matched):fail('#DIV/0!');
      if(name==='MINIFS')return matched?minimum:0;if(name==='MAXIFS')return matched?maximum:0;return finite(total);
    }
    if(name==='SUMPRODUCT'){
      const ranges=args.map(asRange),first=ranges[0];
      if(ranges.some(r=>r.rows!==first.rows||r.cols!==first.cols))fail('#VALUE!');
      return finite(sum(first.values.map((_,i)=>ranges.reduce((product,r)=>{const v=checked(r.values[i]);return product*(typeof v==='number'?v:0);},1))));
    }
    if(name==='INDEX'){
      const r=asRange(args[0]);let row=integer(args[1]),col=args.length>2?integer(args[2]):1;
      if(args.length===2&&r.rows===1){col=row;row=1;}
      if(row<0||col<0||row>r.rows||col>r.cols)fail('#REF!');
      const rows=row?1:r.rows,cols=col?1:r.cols,top=row?row-1:0,left=col?col-1:0,values=[];
      for(let y=top;y<top+rows;y++)for(let x=left;x<left+cols;x++)values.push(r.values[y*r.cols+x]);
      if(r.sheet)return ctx.getRange(r.sheet,r.top+top,r.left+left,rows,cols);
      if(rows*cols===1)return checked(values[0])===BLANK?0:values[0];return range(values,rows,cols);
    }
    if(['MATCH','XMATCH','HLOOKUP','LOOKUP'].includes(name)){
      const lookup=checked(args[0]),r=asRange(args[1]);let values=r.values,index;
      if(name==='HLOOKUP'){
        const row=integer(args[2]);if(row<1)fail('#VALUE!');if(row>r.rows)fail('#REF!');
        values=values.slice(0,r.cols);const approximate=args.length<4||logical(args[3]);index=lookupIndex(lookup,values,approximate?-1:0,1,!approximate);
        if(index<0)fail('#N/A');const v=checked(r.values[(row-1)*r.cols+index]);return v===BLANK?0:v;
      }
      if(r.rows>1&&r.cols>1&&name!=='LOOKUP')fail('#N/A');
      if(name==='MATCH'){
        const type=args.length>2?integer(args[2]):1;if(![-1,0,1].includes(type))fail('#N/A');
        index=lookupIndex(lookup,values,-type,type===0?1:-1,type===0);return index<0?fail('#N/A'):index+1;
      }
      if(name==='XMATCH'){index=lookupIndex(lookup,values,args.length>2?integer(args[2]):0,args.length>3?integer(args[3]):1);return index<0?fail('#N/A'):index+1;}
      let results;
      if(args.length>2){results=asRange(args[2]).values;if(results.length!==values.length)fail('#N/A');}
      else if(r.rows>=r.cols){values=Array.from({length:r.rows},(_,i)=>r.values[i*r.cols]);results=Array.from({length:r.rows},(_,i)=>r.values[i*r.cols+r.cols-1]);}
      else {values=r.values.slice(0,r.cols);results=r.values.slice(-r.cols);}
      index=lookupIndex(lookup,values,-1,-1);if(index<0)fail('#N/A');const v=checked(results[index]);return v===BLANK?0:v;
    }
    if(aggregateNames.has(name)||aNames.has(name)){
      const numbers=aNames.has(name)?aNumbers(args,name):numericItems(args);
      if(['PRODUCT','SUMSQ','AVEDEV','DEVSQ','VAR.S','VAR.P','STDEV.S','STDEV.P'].includes(name)||aNames.has(name))return finite(reduceStats(name,numbers));
      if(!numbers.length)fail(name==='MEDIAN'||name==='MODE.SNGL'?'#NUM!':'#DIV/0!');
      if(['GEOMEAN','HARMEAN'].includes(name)&&numbers.some(n=>n<=0))fail('#NUM!');
      if(['SKEW','SKEW.P'].includes(name)&&numbers.length<3||name==='KURT'&&numbers.length<4)fail('#DIV/0!');
      if(name==='MODE.SNGL'){
        const counts=new Map();for(const n of numbers)counts.set(n,(counts.get(n)||0)+1);
        let count=1,result;for(const [n,c] of counts)if(c>count||c===count&&n<result){count=c;result=n;}return result===undefined?fail('#N/A'):result;
      }
      return libraryCall(name,[numbers]);
    }
    if(pairNames.has(name)){
      const a=asRange(args[0]),b=asRange(args[1]);if(a.values.length!==b.values.length)fail('#N/A');
      const x=[],y=[];a.values.forEach((v,i)=>{checked(v);checked(b.values[i]);if(typeof v==='number'&&typeof b.values[i]==='number'){x.push(v);y.push(b.values[i]);}});
      if(!x.length||name==='COVARIANCE.S'&&x.length<2||name==='STEYX'&&x.length<3)fail('#DIV/0!');
      const flat=a=>a.every(v=>v===a[0]);
      if(['CORREL','PEARSON','RSQ'].includes(name)&&(flat(x)||flat(y))||['SLOPE','INTERCEPT','STEYX'].includes(name)&&flat(y))fail('#DIV/0!');
      return libraryCall(name,[x,y]);
    }
    if(arrayFirst.has(name)){
      const values=numericItems([args[0]]);if(!values.length)fail('#NUM!');
      if(name==='LARGE'||name==='SMALL'){
        const k=name==='LARGE'?Math.ceil(num(1)):Math.floor(num(1));if(k<1||k>values.length)fail('#NUM!');values.sort((a,b)=>a-b);return values[name==='LARGE'?values.length-k:k-1];
      }
      if(name.startsWith('PERCENTRANK')&&(num(1)<values.reduce((a,b)=>Math.min(a,b),Infinity)||num(1)>values.reduce((a,b)=>Math.max(a,b),-Infinity)))fail('#N/A');
      return libraryCall(name,[values,...args.slice(1).map(v=>numeric(v))]);
    }
    if(name==='RANK.AVG'||name==='RANK.EQ')return libraryCall(name,[num(0),numericItems([args[1]]),args.length>2?num(2):0]);
    if(name==='SERIESSUM')return libraryCall(name,[num(0),num(1),num(2),numericItems([args[3]])]);
    if(name==='TEXTJOIN'){
      const delimiter=str(0),ignore=logical(args[1]),values=flatten(args.slice(2)).map(({value})=>textual(value));
      const result=(ignore?values.filter(v=>v!==''):values).join(delimiter);if(result.length>32767)fail('#VALUE!');return result;
    }
    if(name==='FIND'||name==='SEARCH'){
      const needle=str(0),haystack=str(1),start=args.length>2?integer(args[2]):1;if(start<1||start>Math.max(1,haystack.length))fail('#VALUE!');
      let pos;if(name==='FIND')pos=haystack.indexOf(needle,start-1);else {const regex=wildcardRegex(needle);pos=haystack.slice(start-1).search(new RegExp(regex.source.slice(1,-1),regex.flags));if(pos>=0)pos+=start-1;}
      return pos<0?fail('#VALUE!'):pos+1;
    }
    if(name==='SUBSTITUTE'){
      const text=str(0),old=str(1),replacement=str(2);if(!old)return text;
      if(args.length<4)return text.split(old).join(replacement);
      const n=integer(args[3]);if(n<1)fail('#VALUE!');let from=0,at=-1;
      for(let i=0;i<n;i++){at=text.indexOf(old,from);if(at<0)return text;from=at+old.length;}
      return text.slice(0,at)+replacement+text.slice(at+old.length);
    }
    if(name==='REPLACE'){
      const text=str(0),start=integer(args[1]),length=integer(args[2]);if(start<1||length<0)fail('#VALUE!');return text.slice(0,start-1)+str(3)+text.slice(start-1+length);
    }
    if(name==='REPT'){const text=str(0),n=integer(args[1]);if(n<0||text.length*n>32767)fail('#VALUE!');return text.repeat(Math.min(n,32767));}
    if(name==='CHAR'||name==='UNICHAR'){
      const n=integer(args[0]);if(n<1||n>(name==='CHAR'?255:0x10ffff)||n>=0xd800&&n<=0xdfff)fail('#VALUE!');return String.fromCodePoint(n);
    }
    if(name==='CODE'||name==='UNICODE'){if(!str(0))fail('#VALUE!');return str(0).codePointAt(0);}
    if(name==='VALUE'||name==='NUMBERVALUE'){
      let s=str(0).trim();const decimal=name==='VALUE'?'.':args.length>1?str(1).charAt(0):'.',group=name==='VALUE'?',':args.length>2?str(2).charAt(0):',';
      if(!decimal||decimal===group)fail('#VALUE!');
      if(name==='VALUE')s=s.replace(/^[$₩€£¥]\s*/,'');
      let percent=0;while(s.endsWith('%')){percent++;s=s.slice(0,-1).trim();}
      s=s.replace(/\s/g,'');if(s.includes(decimal)&&group&&s.slice(s.indexOf(decimal)+decimal.length).includes(group))fail('#VALUE!');if(group)s=s.split(group).join('');if(decimal!=='.')s=s.split(decimal).join('.');
      if(!s)return 0;if(!/^[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:e[+-]?\d+)?$/i.test(s))fail('#VALUE!');return finite(Number(s)/100**percent);
    }
    if(name==='TEXT'){
      const value=checked(args[0]),format=str(1);if(typeof value==='string')return value;
      const n=numeric(value);if(/^0+$/.test(format))return (n<0?'-':'')+String(Math.abs(rounded(n,0))).padStart(format.length,'0');
      return formatValue(n,{numberFormat:format});
    }
    if(name==='FIXED'||name==='DOLLAR'){
      const decimals=args.length>1?integer(args[1]):2;if(decimals>99)fail('#VALUE!');
      const n=rounded(num(0),decimals),result=Math.abs(n).toLocaleString('en-US',{useGrouping:name==='DOLLAR'||!(args.length>2&&logical(args[2])),minimumFractionDigits:Math.max(0,decimals),maximumFractionDigits:Math.max(0,decimals)});
      return name==='DOLLAR'?(n<0?'-₩'+result:'₩'+result):(n<0?'-':'')+result;
    }
    if(name==='TRUNC')return rounded(num(0),args.length>1?num(1):0,'down');
    if(name==='QUOTIENT'){if(num(1)===0)fail('#DIV/0!');return finite(Math.trunc(num(0)/num(1)));}
    if(name==='COMBIN'||name==='COMBINA'){
      const n=integer(args[0]),k=integer(args[1]);if(n<0||k<0||name==='COMBIN'&&k>n)fail('#NUM!');if(name==='COMBINA'&&n===0)return k===0?1:fail('#NUM!');
      return combination(name==='COMBINA'?n+k-1:n,k);
    }
    if(name==='PERMUT'||name==='PERMUTATIONA'){
      const n=integer(args[0]),k=integer(args[1]);if(n<0||k<0||name==='PERMUT'&&k>n||k>10000)fail('#NUM!');if(name==='PERMUTATIONA')return finite(n**k);
      let result=1;for(let i=0;i<k;i++){result*=n-i;if(!Number.isFinite(result))fail('#NUM!');}return result;
    }
    if(name==='MULTINOMIAL'){
      const values=numericItems(args).map(Math.trunc);if(values.some(n=>n<0))fail('#NUM!');let total=0,result=1;
      for(const n of values){result*=combination(total+n,n);total+=n;}return finite(result);
    }
    if(name==='NPV'){
      const rate=num(0),values=numericItems(args.slice(1));if(rate===-1)fail('#DIV/0!');return finite(sum(values.map((v,i)=>v/(1+rate)**(i+1))));
    }
    if(name==='XNPV'||name==='XIRR'){
      const offset=name==='XNPV'?1:0,values=asRange(args[offset]).values.map(numeric),dates=asRange(args[offset+1]).values.map(serialDate);
      if(values.length!==dates.length||values.length<2||values.length>10000||dates.some(d=>d<dates[0]))fail('#NUM!');
      const times=dates.map(d=>(d-dates[0])/365);
      const npv=rate=>sum(values.map((v,i)=>v/(1+rate)**times[i]));
      if(name==='XNPV'){if(num(0)<=-1)fail('#NUM!');return finite(npv(num(0)));}
      if(!values.some(v=>v>0)||!values.some(v=>v<0))fail('#NUM!');
      let rate=omitted(2)?0.1:num(2);if(rate<=-1)fail('#NUM!');
      const tolerance=1e-10*values.reduce((n,v)=>Math.max(n,Math.abs(v)),1);
      for(let i=0;i<100;i++){
        const value=npv(rate);if(Math.abs(value)<tolerance)return rate;
        const derivative=-sum(values.map((v,i)=>times[i]*v/(1+rate)**(times[i]+1)));
        if(!Number.isFinite(derivative)||derivative===0)break;
        const next=rate-value/derivative;if(!Number.isFinite(next)||next<=-1)break;rate=next;
      }
      // Bounded fallback near the supplied guess. An unresolved root is an Excel error.
      let lo=-0.9999,hi=1,fl=npv(lo),fh=npv(hi);
      for(let i=0;i<30&&fl*fh>0;i++){hi=hi*2+1;fh=npv(hi);}
      if(!Number.isFinite(fl)||!Number.isFinite(fh)||fl*fh>0)fail('#NUM!');
      for(let i=0;i<150;i++){const mid=(lo+hi)/2,fm=npv(mid);if(Math.abs(fm)<tolerance)return mid;if(fl*fm<=0)hi=mid;else {lo=mid;fl=fm;}}
      fail('#NUM!');
    }
    if(name.startsWith('BIT')){
      const a=num(0),b=num(1),limit=281474976710655;
      if(!Number.isInteger(a)||a<0||a>limit||!Number.isInteger(b))fail('#NUM!');
      if(name==='BITLSHIFT'||name==='BITRSHIFT'){
        if(Math.abs(b)>53)fail('#NUM!');const shift=name==='BITLSHIFT'?b:-b,result=shift>=0?BigInt(a)<<BigInt(shift):BigInt(a)>>BigInt(-shift);return result>BigInt(limit)?fail('#NUM!'):Number(result);
      }
      if(b<0||b>limit)fail('#NUM!');return Number(name==='BITAND'?BigInt(a)&BigInt(b):name==='BITOR'?BigInt(a)|BigInt(b):BigInt(a)^BigInt(b));
    }
    if(name==='TIME'){const values=args.map(integer);if(values.some(n=>n>32767))fail('#NUM!');const seconds=values[0]*3600+values[1]*60+values[2];if(seconds<0)fail('#NUM!');return seconds%86400/86400;}
    if(name==='TIMEVALUE'){
      const match=/^(\d{1,2}):(\d{2})(?::(\d{2}))?\s*(AM|PM)?$/i.exec(str(0).trim());if(!match)fail('#VALUE!');let hour=+match[1],minute=+match[2],second=+(match[3]||0);if(hour>23||minute>59||second>59)fail('#VALUE!');if(match[4]){if(hour<1||hour>12)fail('#VALUE!');hour=hour%12+(match[4].toUpperCase()==='PM'?12:0);}return (hour*3600+minute*60+second)/86400;
    }
    if(['HOUR','MINUTE','SECOND'].includes(name)){let n=typeof checked(args[0])==='string'?call('TIMEVALUE',[nodes[0]],evaluate,ctx):num(0);if(n<0||n>2958465.99999)fail('#NUM!');const seconds=Math.round((n%1)*86400)%86400;return name==='HOUR'?Math.floor(seconds/3600):name==='MINUTE'?Math.floor(seconds/60)%60:seconds%60;}
    if(name==='NOW'){const n=new Date();return dateSerial(n.getFullYear(),n.getMonth()+1,n.getDate())+(n.getHours()*3600+n.getMinutes()*60+n.getSeconds())/86400;}
    if(name==='DATEVALUE')return serialDate(args[0]);
    if(name==='DAYS')return serialDate(args[0])-serialDate(args[1]);
    if(name==='EDATE'||name==='EOMONTH'){
      const p=serialParts(serialDate(args[0])),months=integer(args[1]);if(Math.abs(months)>120000)fail('#NUM!');
      const absolute=p.year*12+p.month-1+months,y=Math.floor(absolute/12),m=((absolute%12)+12)%12+1;
      if(y<1900||y>9999)fail('#NUM!');const last=y===1900&&m===2?28:monthDays(y,m);return serial(y,m,name==='EOMONTH'?last:Math.min(Math.max(1,p.day),last));
    }
    if(name==='WEEKDAY'){
      const w=weekday(serialDate(args[0])),type=args.length>1?integer(args[1]):1;
      if(type===1)return w+1;if(type===2)return (w+6)%7+1;if(type===3)return (w+6)%7;
      if(type>=11&&type<=17)return (w-(type-10)%7+7)%7+1;fail('#NUM!');
    }
    if(name==='WEEKNUM'||name==='ISOWEEKNUM'){
      const n=serialDate(args[0]),p=serialParts(n),type=name==='ISOWEEKNUM'?21:args.length>1?integer(args[1]):1;
      if(type===21){const thursday=n+3-(weekday(n)+6)%7,year=serialParts(thursday).year;return Math.floor((thursday-dateSerial(year,1,1))/7)+1;}
      if(![1,2,11,12,13,14,15,16,17].includes(type))fail('#NUM!');const first=type===1?0:type===2?1:(type-10)%7,jan=dateSerial(p.year,1,1);return Math.floor((n-jan+(weekday(jan)-first+7)%7)/7)+1;
    }
    if(name.startsWith('NETWORKDAYS')||name.startsWith('WORKDAY'))return businessDays(name,args);
    if(name==='DATEDIF'){
      const start=serialDate(args[0]),end=serialDate(args[1]),unit=str(2).toUpperCase();if(end<start)fail('#NUM!');const a=serialParts(start),b=serialParts(end),months=(b.year-a.year)*12+b.month-a.month-(b.day<a.day?1:0);
      if(unit==='D')return end-start;if(unit==='M')return months;if(unit==='Y')return Math.floor(months/12);if(unit==='YM')return months%12;
      if(unit==='YD'){let y=b.year;if(b.month<a.month||b.month===a.month&&b.day<a.day)y--;return end-serial(y,a.month,Math.min(a.day,monthDays(y,a.month)));}
      if(unit==='MD'){if(b.day>=a.day)return b.day-a.day;const m=b.month===1?12:b.month-1,y=b.month===1?b.year-1:b.year;return b.day+monthDays(y,m)-a.day;}fail('#NUM!');
    }
    if(name==='T.DIST'||name==='T.DIST.RT'||name==='T.DIST.2T'){
      const x=num(0),df=integer(args[1]);if(df<1)fail('#NUM!');
      const rt=libraryCall('T.DIST.RT',[Math.abs(x),df]);
      if(name==='T.DIST.RT')return x<0?1-rt:rt;
      if(name==='T.DIST.2T')return x<0?fail('#NUM!'):2*rt;
      if(logical(args[2]))return x<0?rt:1-rt;
      return Math.exp(libraryCall('GAMMALN',[(df+1)/2])-libraryCall('GAMMALN',[df/2]))/Math.sqrt(df*Math.PI)*(1+x*x/df)**(-(df+1)/2);
    }
    if(name==='ERF')return args.length>1?libraryCall('ERF',[num(1)])-libraryCall('ERF',[num(0)]):libraryCall('ERF',[num(0)]);
    // Bound library loops/allocations before calling third-party functions.
    if(name==='FACT'&&(num(0)<0||num(0)>170)||name==='FACTDOUBLE'&&(num(0)<-1||num(0)>300))fail('#NUM!');
    if(['COMBIN','COMBINA','PERMUT','PERMUTATIONA','BASE','DECIMAL','MULTINOMIAL'].includes(name)&&flatten(args).some(({value})=>typeof value==='number'&&Math.abs(value)>1e6))fail('#NUM!');
    if(name==='DB'||name==='DDB'){if(num(2)>10000||num(3)>10000)fail('#NUM!');}
    if(name==='CUMIPMT'||name==='CUMPRINC'){if(num(4)>100000)fail('#NUM!');}
    if(name==='REPT'&&num(1)>32767)fail('#VALUE!');
    if(name==='IRR'||name==='XIRR'){if(asRange(args[0]).values.length>10000)fail('#NUM!');}
    if(['GCD','LCM'].includes(name)){
      const values=numericItems(args).map(Math.trunc);if(values.some(n=>n<0||n>Number.MAX_SAFE_INTEGER))fail('#NUM!');
      const gcd=(a,b)=>{while(b){[a,b]=[b,a%b];}return a;};
      const result=values.reduce((a,b)=>name==='GCD'?gcd(a,b):!a||!b?0:a/gcd(a,b)*b,name==='GCD'?0:1);if(result>Number.MAX_SAFE_INTEGER)fail('#NUM!');return result;
    }
    if(name==='ISEVEN'||name==='ISODD'){if(typeof checked(args[0])==='boolean')fail('#VALUE!');return Math.abs(integer(args[0]))%2===(name==='ISEVEN'?0:1);}
    if(name==='LOG'){if(num(0)<=0||!omitted(1)&&num(1)<=0)fail('#NUM!');if(!omitted(1)&&num(1)===1)fail('#DIV/0!');}
    if(name==='IPMT'||name==='PPMT'){if(num(1)<1||num(1)>num(2))fail('#NUM!');}
    if(name.startsWith('BESSEL')&&(num(1)<0||num(1)>10000))fail('#NUM!');
    if(name.startsWith('BINOM')){
      const trials=name==='BINOM.DIST'?integer(args[1]):integer(args[0]),prob=name==='BINOM.DIST'?num(2):num(1);
      if(trials<0||trials>100000||prob<0||prob>1)fail('#NUM!');
      if(name==='BINOM.INV'&&(num(2)<=0||num(2)>=1))fail('#NUM!');
      if(name==='BINOM.INV'){
        const alpha=num(2);let lo=0,hi=trials;
        while(lo<hi){const mid=Math.floor((lo+hi)/2);if(libraryCall('BINOM.DIST',[mid,trials,prob,true])>=alpha)hi=mid;else lo=mid+1;}
        return lo;
      }
      if(name==='BINOM.DIST.RANGE'){
        const lower=integer(args[2]),upper=omitted(3)?lower:integer(args[3]);if(lower<0||upper<lower||upper>trials)fail('#NUM!');
      }
    }
    if(name==='BASE'&&!omitted(2)&&(num(2)<0||num(2)>255))fail('#VALUE!');
    if(name.startsWith('BETA.')){if(num(1)<=0||num(2)<=0)fail('#NUM!');}
    if(name==='GAMMA.DIST'||name==='GAMMA.INV'){if(num(1)<=0||num(2)<=0||num(0)<0)fail('#NUM!');}
    if(name==='POISSON.DIST'&&(num(0)<0||num(1)<=0))fail('#NUM!');
    if(name==='WEIBULL.DIST'&&(num(0)<0||num(1)<=0||num(2)<=0))fail('#NUM!');
    // Preserve errors before the library can accidentally replace their code or ignore them.
    const parameterParts=metadata.parameters.split(', ');
    const converted=args.map((arg,i)=>{
      if(omitted(i)&&parameterParts[i]?.startsWith('['))return undefined;
      const safe=value=>{value=checked(value);return typeof value==='number'?finite(value):raw(value);};
      if(isRange(arg))return Array.from({length:arg.rows},(_,r)=>arg.values.slice(r*arg.cols,(r+1)*arg.cols).map(safe));
      return safe(arg);
    });
    const cumulative={'BETA.DIST':3,'BINOM.DIST':3,'CHISQ.DIST':2,'EXPON.DIST':2,'F.DIST':3,'GAMMA.DIST':3,'HYPGEOM.DIST':4,'LOGNORM.DIST':3,'NEGBINOM.DIST':3,'NORM.DIST':3,'NORM.S.DIST':1,'POISSON.DIST':2,'WEIBULL.DIST':3}[name];
    if(cumulative!==undefined)converted[cumulative]=logical(args[cumulative]);
    return libraryCall(name,converted);
  };
}
