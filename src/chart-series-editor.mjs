import {chartSeries,chartColors,parseChartFormula,refFormula,referenceSize} from './chart-model.mjs';
import {tableBounds} from './table-schema.mjs';
import {address} from './engine.mjs';
const esc=v=>String(v??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
const refText=(ref,book)=>ref?.formula||ref?.sheetId&&ref?.range&&refFormula(ref,book)||'';
const literalName=value=>/^[=']/.test(value)?"'"+value:value;
export function chartSeriesRows(chart,sheet,book){
  return chartSeries(chart,sheet,book).map((series,i)=>({name:series.nameRef&&(series.nameRef.formula||series.nameRef.sheetId)?'='+refText(series.nameRef,book):literalName(String(series.nameRef?.cache?.[0]??series.name??`계열 ${i+1}`)),values:refText(series.valRef,book),categories:refText(series.xRef||series.catRef,book),size:refText(series.sizeRef,book),color:series.color||chartColors[i%chartColors.length],type:series.type==='line'?'line':'bar',secondary:!!series.secondary,lines:!!series.lines}));
}
function readReference(value,book,sheetId,{optional=false,single=false}={}){
  const input=String(value||'').trim().replace(/^=/,'');if(!input&&optional)return undefined;
  const ref=parseChartFormula(input,book,sheetId),b=ref&&tableBounds(ref.range);
  if(!ref||!b||b.r1!==b.r2&&b.c1!==b.c2)throw Error('계열 참조는 한 행 또는 한 열의 셀 범위·이름이어야 합니다. 예: B2:B10');
  const count=referenceSize(ref);if(count<1||count>10000)throw Error('한 계열은 최대 10,000개 값까지 지정하세요.');if(single&&count!==1)throw Error('계열 이름 참조는 셀 하나를 지정하세요.');
  const direct=/^\$?[A-Z]{1,3}\$?[1-9]\d*(?::\$?[A-Z]{1,3}\$?[1-9]\d*)?$/i.test(input)||/!\$?[A-Z]{1,3}\$?[1-9]\d*(?::\$?[A-Z]{1,3}\$?[1-9]\d*)?$/i.test(input);
  const local=input.slice(input.lastIndexOf('!')+1);if(!direct&&(/^[^\[]*[()]/.test(local)||/\d+#$/.test(local)))throw Error('셀 범위 또는 정의된 이름을 사용하세요. 함수·스필 수식은 이름으로 등록한 뒤 지정하세요.');
  return {...ref,...(!direct?{formula:input,ownerSheetId:sheetId}:{})};
}
export function chartSeriesPlan(rows,type,book,sheetId){
  if(!Array.isArray(rows)||rows.length<1||rows.length>32)throw Error('차트에는 1~32개의 계열을 지정하세요.');
  let categoryKey,commonCount;
  const series=rows.map((row,i)=>{
    const name=String(row.name||'').trim()||`계열 ${i+1}`;if(name.length>4000)throw Error('계열 이름은 4,000자 이내로 입력하세요.');
    const valRef=readReference(row.values,book,sheetId),catRef=readReference(row.categories,book,sheetId,{optional:type!=='bubble'}),sizeRef=type==='bubble'?readReference(row.size,book,sheetId):undefined,xy=['scatter','bubble'].includes(type);
    for(const [label,ref] of [['항목 / X값',catRef],['거품 크기',sizeRef]])if(ref&&referenceSize(ref)!==referenceSize(valRef))throw Error(`계열 ${i+1}: ${label}과 값 범위의 셀 수가 같아야 합니다.`);
    if(!xy){const n=referenceSize(valRef);if(commonCount!==undefined&&commonCount!==n)throw Error('범주형 차트의 모든 계열은 같은 수의 값을 지정하세요.');commonCount=n;const key=catRef?(catRef.formula?.toLocaleLowerCase()||catRef.sheetId+'!'+catRef.range):'';if(categoryKey!==undefined&&key!==categoryKey)throw Error('범주형 차트의 계열에는 같은 항목 범위를 지정하세요.');categoryKey=key;}
    if(!/^#[\da-f]{6}$/i.test(row.color||''))throw Error('계열 색을 선택하세요.');
    const categories=catRef||{cache:Array.from({length:referenceSize(valRef)},(_,i)=>xy?i+1:String(i+1))};
    const entry={...(name.startsWith('=')?{nameRef:readReference(name,book,sheetId,{single:true})}:{name:name.startsWith("'")?name.slice(1):name}),valRef,catRef:categories,...(xy?{xRef:categories}:{}),...(sizeRef?{sizeRef}:{}),color:row.color,type:type==='combo'?(row.type==='line'?'line':'bar'):type,secondary:type==='combo'&&!!row.secondary,...(type==='scatter'?{lines:!!row.lines}:{})};return entry;
  });
  const ownerRefs=series.flatMap(s=>[s.valRef,s.catRef,s.sizeRef]).filter(ref=>ref?.sheetId===sheetId),bounds=ownerRefs.map(ref=>tableBounds(ref.range));
  const range=bounds.length?address(Math.min(...bounds.map(b=>b.r1)),Math.min(...bounds.map(b=>b.c1)))+':'+address(Math.max(...bounds.map(b=>b.r2)),Math.max(...bounds.map(b=>b.c2))):'A1:A1';
  return {series,range};
}
export function installSeriesEditor({el,current,sheet,book,getRangeChart,onChange}){
  const host=el.querySelector('#chart-series-editor');let rows=current.series?.length?chartSeriesRows(current,sheet,book):null,seedError='';if(rows?.length>32){seedError='원본에 32개를 넘는 계열이 있습니다. 자동으로 일부만 복제하지 않습니다. 사용할 범위를 다시 구성하세요.';rows=[];}
  const readRows=()=>{for(const field of host.querySelectorAll('[data-series-field]')){const row=rows[+field.closest('[data-series-row]').dataset.seriesRow];row[field.dataset.seriesField]=field.type==='checkbox'?field.checked:field.value;}return rows||[];};
  const render=()=>{
    const type=el.querySelector('#chart-type').value,xy=['scatter','bubble'].includes(type);
    host.innerHTML=`${seedError?'<p class="worksheet-error">'+esc(seedError)+'</p>':''}<p class="chart-help">범주형 차트는 모든 계열에 같은 항목 범위를 지정합니다. 이름은 직접 쓰거나 =셀로 지정하세요. =로 시작하는 글자 이름은 앞에 작은따옴표를 붙이세요. 범위는 다른 시트·정의된 이름도 사용할 수 있습니다.</p><div class="chart-series-list">${(rows||[]).map((row,i)=>`<fieldset data-series-row="${i}"><legend>계열 ${i+1}</legend><div class="chart-series-fields"><label>이름 또는 =셀<input data-series-field="name" aria-label="계열 ${i+1} 이름" value="${esc(row.name)}"></label><label>${xy?'Y값':'값'} 범위<input data-series-field="values" aria-label="계열 ${i+1} 값 범위" value="${esc(row.values)}" placeholder="B2:B10"></label><label>${xy?'X값':'항목'} 범위<input data-series-field="categories" aria-label="계열 ${i+1} 항목 범위" value="${esc(row.categories)}" placeholder="${xy?'A2:A10':'빈칸은 1, 2, 3…'}"></label>${type==='bubble'?`<label>거품 크기 범위<input data-series-field="size" aria-label="계열 ${i+1} 크기 범위" value="${esc(row.size)}" placeholder="C2:C10"></label>`:''}<label>색<input type="color" data-series-field="color" aria-label="계열 ${i+1} 색" value="${row.color}"></label>${type==='combo'?`<label>유형<select data-series-field="type" aria-label="계열 ${i+1} 유형"><option value="bar">막대</option><option value="line" ${row.type==='line'?'selected':''}>꺾은선</option></select></label><label class="chart-check"><input type="checkbox" data-series-field="secondary" aria-label="계열 ${i+1} 보조축" ${row.secondary?'checked':''}>보조축</label>`:''}${type==='scatter'?`<label class="chart-check"><input type="checkbox" data-series-field="lines" aria-label="계열 ${i+1} 점 연결" ${row.lines?'checked':''}>점을 선으로 연결</label>`:''}</div><div class="chart-series-actions"><button type="button" data-series-up="${i}" ${i===0?'disabled':''} aria-label="계열 ${i+1} 위로">↑</button><button type="button" data-series-down="${i}" ${i===(rows.length-1)?'disabled':''} aria-label="계열 ${i+1} 아래로">↓</button><button type="button" data-series-remove="${i}" aria-label="계열 ${i+1} 삭제">삭제</button></div></fieldset>`).join('')}</div><div class="chart-series-actions"><button type="button" data-series-add ${rows?.length>=32?'disabled':''}>계열 추가</button><button type="button" data-series-reset>현재 셀 범위에서 다시 구성</button></div>`;
  };
  const initialize=()=>{if(!rows)rows=chartSeriesRows(getRangeChart(),sheet,book);if(!rows.length&&!seedError)rows=[{name:'계열 1',values:'',categories:'',size:'',color:chartColors[0],type:'bar',secondary:false,lines:false}];render();};
  host.addEventListener('input',()=>{readRows();onChange();});host.addEventListener('change',()=>{readRows();onChange();});
  host.addEventListener('click',event=>{
    const button=event.target.closest('[data-series-add],[data-series-reset],[data-series-up],[data-series-down],[data-series-remove]');if(!button)return;readRows();
    let focusRow=0,focusAction;
    if(button.hasAttribute('data-series-add')){focusRow=rows.length;seedError='';rows.push({name:'계열 '+(rows.length+1),values:'',categories:rows[0]?.categories||'',size:'',color:chartColors[rows.length%chartColors.length],type:'bar',secondary:false,lines:false});}
    else if(button.hasAttribute('data-series-reset')){rows=chartSeriesRows(getRangeChart(),sheet,book);seedError='';}
    else if(button.dataset.seriesRemove!==undefined){focusRow=+button.dataset.seriesRemove;rows.splice(focusRow,1);focusRow=Math.min(focusRow,rows.length-1);}
    else{const i=+(button.dataset.seriesUp??button.dataset.seriesDown),next=i+(button.dataset.seriesUp!==undefined?-1:1);if(next>=0&&next<rows.length)[rows[i],rows[next]]=[rows[next],rows[i]];focusRow=next;focusAction=button.dataset.seriesUp!==undefined?'up':'down';}
    render();onChange();const group=host.querySelector(`[data-series-row="${focusRow}"]`),action=group?.querySelector(`[data-series-${focusAction}]:not([disabled])`);(action||group?.querySelector('[data-series-field="name"]')||host.querySelector('[data-series-add]')).focus();
  });
  return {initialize,refresh(){readRows();if(rows)render();},plan(type){return chartSeriesPlan(readRows(),type,book,sheet.id);},count:()=>rows?.length||0,series:()=>rows||[]};
}
