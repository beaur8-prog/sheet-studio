import {chartTypes,chartLayout,chartRangeError} from './chart-model.mjs';
import {escapeChartHTML as esc} from './chart-render.mjs';

export function chartDialog({chart,range,sheet,workbook,onApply,modal}){
  const normalize=value=>value.includes(':')?value:value+':'+value;
  const current=chart||{type:'bar',grouping:'clustered',title:'차트 제목',range:normalize(range),legend:true,labels:false};
  const initial=chartLayout(current,sheet,workbook);
  const el=modal(chart?'차트 설정':'차트 삽입',`
    <div class="chart-form">
    <label>차트 유형<select id="chart-type">${!chartTypes[current.type]?'<option value="'+esc(current.type)+'">Excel 원본 유형 유지</option>':''}${Object.entries(chartTypes).map(([value,label])=>`<option value="${value}" ${current.type===value?'selected':''}>${label}</option>`).join('')}</select></label>
    <label>배치<select id="chart-grouping">${Object.entries({clustered:'묶은형 / 일반',stacked:'누적',percentStacked:'100% 누적'}).map(([value,label])=>`<option value="${value}" ${current.grouping===value||value==='clustered'&&current.grouping==='standard'?'selected':''}>${label}</option>`).join('')}</select></label>
    <label>제목<input id="chart-title" value="${esc(current.title)}" ${current.type==='unsupported'?'disabled':''}></label>
    <label>데이터 범위<input id="chart-range" value="${esc(current.range||range)}" ${current.source?'disabled':''}></label>
    <label class="chart-check"><input id="chart-first-row" type="checkbox" ${initial.firstRowAsHeader?'checked':''}>첫 행을 계열 이름으로 사용</label>
    <label class="chart-check" id="chart-first-column-option"><input id="chart-first-column" type="checkbox" ${initial.firstColumnAsLabels?'checked':''}>첫 열을 항목 이름으로 사용</label>
    <label class="chart-check"><input id="chart-legend" type="checkbox" ${current.legend!==false?'checked':''}>범례 표시</label>
    <label class="chart-check"><input id="chart-labels" type="checkbox" ${current.labels?'checked':''}>데이터 레이블</label>
    <label>가로축 제목<input id="chart-x-title" value="${esc(current.xTitle||'')}"></label>
    <label>세로축 제목<input id="chart-y-title" value="${esc(current.yTitle||'')}"></label>
    <label class="chart-check"><input id="chart-secondary" type="checkbox" ${current.secondaryAxis!==false?'checked':''}>혼합 차트의 마지막 계열에 보조축 사용</label>
    </div>
    ${current.source?'<p class="chart-help">Excel의 원래 계열과 고급 서식을 유지합니다. 유형·배치·데이터 범위·보조축을 바꾸려면 아래 변환을 선택하세요.</p><label class="chart-check"><input id="chart-convert" type="checkbox">새 차트로 변환 (원본 고급 서식은 제거됨)</label>':''}
    <p class="chart-help">제목·항목 이름은 자동으로 구분하며 위 설정으로 직접 바꿀 수 있습니다. 숫자로 된 첫 행·첫 열은 값으로 포함합니다. 분산형은 첫 열을 X값으로 사용합니다. 거품형은 X값·Y값·크기 순서입니다. 원형은 첫 계열만 사용합니다.</p>
    <div id="chart-error" role="status"></div>`,[{label:chart?'적용':'삽입',primary:true,action:el=>{
      const get=id=>el.querySelector('#'+id),convert=get('chart-convert')?.checked,type=get('chart-type').value,grouping=get('chart-grouping').value;
      const retained=current.source&&!convert;
      if(retained&&(type!==current.type||grouping!==(current.grouping==='standard'?'clustered':current.grouping||'clustered'))){get('chart-error').textContent='유형이나 배치를 바꾸려면 새 차트로 변환을 선택하세요.';return;}
      const selected=normalize(get('chart-range').value.toUpperCase().replace(/\$/g,'').trim());
      const layout={firstRowAsHeader:get('chart-first-row').checked,firstColumnAsLabels:get('chart-first-column').checked};
      const error=!retained&&chartRangeError({type,range:selected,...layout},sheet,workbook);
      if(error){get('chart-error').textContent=error;return;}
      if(!retained&&type==='unsupported'){get('chart-error').textContent='새 차트 유형을 선택하세요.';return;}
      const patch={type,grouping:retained?current.grouping:grouping,title:get('chart-title').value,legend:get('chart-legend').checked,labels:get('chart-labels').checked,xTitle:get('chart-x-title').value,yTitle:get('chart-y-title').value,secondaryAxis:get('chart-secondary').checked};
      if(!retained)Object.assign(patch,{range:selected,...layout,series:undefined,source:undefined,rebuilt:true,previewNote:undefined,unsupportedReason:undefined});
      onApply(patch);el.remove();
    }}]);
  const syncControls=()=>{
    const retained=current.source&&!el.querySelector('#chart-convert')?.checked;
    el.querySelector('#chart-range').disabled=!!retained;
    el.querySelector('#chart-first-row').disabled=!!retained;
    el.querySelector('#chart-first-column').disabled=!!retained||['scatter','bubble'].includes(el.querySelector('#chart-type').value);
    el.querySelector('#chart-first-column-option').hidden=['scatter','bubble'].includes(el.querySelector('#chart-type').value);
    for(const id of ['chart-title','chart-legend','chart-labels','chart-x-title','chart-y-title'])el.querySelector('#'+id).disabled=!!retained&&current.type==='unsupported';
    el.querySelector('#chart-secondary').disabled=!!retained||el.querySelector('#chart-type').value!=='combo';
  };
  const manual={firstRowAsHeader:current.firstRowAsHeader!==undefined,firstColumnAsLabels:current.firstColumnAsLabels!==undefined};
  const infer=()=>{const guessed=chartLayout({type:el.querySelector('#chart-type').value,range:normalize(el.querySelector('#chart-range').value.toUpperCase().replace(/\$/g,'').trim())},sheet,workbook);for(const [key,id] of [['firstRowAsHeader','chart-first-row'],['firstColumnAsLabels','chart-first-column']])if(!manual[key])el.querySelector('#'+id).checked=guessed[key];};
  for(const [key,id] of [['firstRowAsHeader','chart-first-row'],['firstColumnAsLabels','chart-first-column']])el.querySelector('#'+id).addEventListener('change',()=>{manual[key]=true;});
  el.querySelector('#chart-range').addEventListener('input',infer);
  el.querySelector('#chart-convert')?.addEventListener('change',event=>{if(event.target.checked)el.querySelector('#chart-range').value=normalize(range);infer();syncControls();});
  el.querySelector('#chart-type').addEventListener('change',()=>{infer();syncControls();});
  syncControls();
  return el;
}
