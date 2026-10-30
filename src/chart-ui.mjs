import {chartTypes,chartLayout,chartRangeError,chartSeriesBy,chartSeries} from './chart-model.mjs';
import {escapeChartHTML as esc} from './chart-render.mjs';

export function chartDialog({chart,range,sheet,workbook,onApply,modal}){
  const normalize=value=>value.includes(':')?value:value+':'+value;
  const current=chart||{type:'bar',grouping:'clustered',title:'차트 제목',range:normalize(range),legend:true,labels:false};
  const initial=chartLayout(current,sheet,workbook);
  const el=modal(chart?'차트 설정':'차트 삽입',`
    <div class="chart-form">
    <label>차트 유형<select id="chart-type">${!chartTypes[current.type]?'<option value="'+esc(current.type)+'">Excel 원본 유형 유지</option>':''}${Object.entries(chartTypes).map(([value,label])=>`<option value="${value}" ${current.type===value?'selected':''}>${label}</option>`).join('')}</select></label>
    <label>데이터 계열 방향<select id="chart-series-by"><option value="auto">자동 (한 줄은 행 기준)</option><option value="rows">행 기준 (가로로 이어진 값)</option><option value="columns">열 기준 (세로로 이어진 값)</option></select></label>
    <label>배치<select id="chart-grouping">${Object.entries({clustered:'묶은형 / 일반',stacked:'누적',percentStacked:'100% 누적'}).map(([value,label])=>`<option value="${value}" ${current.grouping===value||value==='clustered'&&current.grouping==='standard'?'selected':''}>${label}</option>`).join('')}</select></label>
    <label>제목<input id="chart-title" value="${esc(current.title)}" ${current.type==='unsupported'?'disabled':''}></label>
    <label>데이터 범위<input id="chart-range" value="${esc(current.range||range)}" ${current.source?'disabled':''}></label>
    <label class="chart-check" id="chart-first-row-option"><input id="chart-first-row" type="checkbox" ${initial.firstRowAsHeader?'checked':''}><span id="chart-first-row-label">첫 행을 계열 이름으로 사용</span></label>
    <label class="chart-check" id="chart-first-column-option"><input id="chart-first-column" type="checkbox" ${initial.firstColumnAsLabels?'checked':''}><span id="chart-first-column-label">첫 열을 항목 이름으로 사용</span></label>
    <label class="chart-check"><input id="chart-legend" type="checkbox" ${current.legend!==false?'checked':''}>범례 표시</label>
    <label class="chart-check"><input id="chart-labels" type="checkbox" ${current.labels?'checked':''}>데이터 레이블</label>
    <label class="chart-check"><input id="chart-hidden-data" type="checkbox" ${current.plotVisibleOnly===false?'checked':''}>숨겨진 행·열의 데이터도 표시</label>
    <label>가로축 제목<input id="chart-x-title" value="${esc(current.xTitle||'')}"></label>
    <label>세로축 제목<input id="chart-y-title" value="${esc(current.yTitle||'')}"></label>
    <label class="chart-check"><input id="chart-secondary" type="checkbox" ${current.secondaryAxis!==false?'checked':''}>혼합 차트의 마지막 계열에 보조축 사용</label>
    </div>
    ${current.source?'<p class="chart-help">Excel의 원래 계열과 고급 서식을 유지합니다. 유형·배치·데이터 범위·보조축을 바꾸려면 아래 변환을 선택하세요.</p><label class="chart-check"><input id="chart-convert" type="checkbox">새 차트로 변환 (원본 고급 서식은 제거됨)</label>':''}
    <p class="chart-help" id="chart-data-help"></p>
    <div id="chart-error" role="status"></div>`,[{label:chart?'적용':'삽입',primary:true,action:el=>{
      const get=id=>el.querySelector('#'+id),convert=get('chart-convert')?.checked,type=get('chart-type').value,grouping=get('chart-grouping').value;
      const retained=current.source&&!convert;
      if(retained&&(type!==current.type||grouping!==(current.grouping==='standard'?'clustered':current.grouping||'clustered'))){get('chart-error').textContent='유형이나 배치를 바꾸려면 새 차트로 변환을 선택하세요.';return;}
      const selected=normalize(get('chart-range').value.toUpperCase().replace(/\$/g,'').trim());
      const direction=get('chart-series-by').value,layout={seriesBy:direction==='auto'?chartSeriesBy({type,range:selected},sheet,workbook):direction,firstRowAsHeader:get('chart-first-row').checked,firstColumnAsLabels:get('chart-first-column').checked};
      const error=!retained&&chartRangeError({type,range:selected,...layout},sheet,workbook);
      if(error){get('chart-error').textContent=error;return;}
      if(!retained&&type==='unsupported'){get('chart-error').textContent='새 차트 유형을 선택하세요.';return;}
      const patch={type,grouping:retained?current.grouping:grouping,title:get('chart-title').value,legend:get('chart-legend').checked,labels:get('chart-labels').checked,plotVisibleOnly:!get('chart-hidden-data').checked,xTitle:get('chart-x-title').value,yTitle:get('chart-y-title').value,secondaryAxis:get('chart-secondary').checked};
      if(!retained)Object.assign(patch,{range:selected,...layout,series:undefined,source:undefined,rebuilt:true,previewNote:undefined,unsupportedReason:undefined});
      onApply(patch);el.remove();
    }}]);
  el.querySelector('.modal').classList.add('chart-modal');
  el.querySelector('#chart-series-by').value=current.seriesBy||'auto';
  const candidate=()=>{const type=el.querySelector('#chart-type').value,range=normalize(el.querySelector('#chart-range').value.toUpperCase().replace(/\$/g,'').trim()),direction=el.querySelector('#chart-series-by').value;return {type,range,...(direction==='auto'?{}:{seriesBy:direction})};};
  let requestedLegend=current.legend!==false;
  const syncControls=()=>{
    const retained=current.source&&!el.querySelector('#chart-convert')?.checked;
    el.querySelector('#chart-range').disabled=!!retained;
    el.querySelector('#chart-first-row').disabled=!!retained;
    const source=candidate(),rows=chartSeriesBy(source,sheet,workbook)==='rows',xy=['scatter','bubble'].includes(source.type),parts=source.range.split(':'),single=parts.length===2&&(rows?parts[0].match(/\d+$/)?.[0]===parts[1].match(/\d+$/)?.[0]:parts[0].match(/^[A-Z]+/)?.[0]===parts[1].match(/^[A-Z]+/)?.[0]),implicitX=source.type==='scatter'&&single;
    el.querySelector('#chart-series-by').disabled=!!retained;
    for(const [id,isX] of [['chart-first-row',rows],['chart-first-column',!rows]]){el.querySelector('#'+id).disabled=!!retained||xy&&!implicitX&&isX;el.querySelector('#'+id+'-option').hidden=xy&&!implicitX&&isX;}
    el.querySelector('#chart-first-row-label').textContent=rows?'첫 행을 항목 이름으로 사용':'첫 행을 계열 이름으로 사용';
    el.querySelector('#chart-first-column-label').textContent=rows?'첫 열을 계열 이름으로 사용':'첫 열을 항목 이름으로 사용';
    const resolved={...source,firstRowAsHeader:el.querySelector('#chart-first-row').checked,firstColumnAsLabels:el.querySelector('#chart-first-column').checked},series=retained?current.series:chartSeries(resolved,sheet,workbook),count=series?.length||0;
    el.querySelector('#chart-data-help').textContent=source.type==='bubble'?`거품형은 X값 다음에 Y값·크기 ${rows?'행':'열'} 쌍이 필요합니다. 각 항목의 X·Y·크기 자료를 함께 선택하세요.`:source.type==='scatter'?(implicitX?'한 줄의 값은 Y값으로 사용하고 X값은 1, 2, 3…으로 지정합니다.':`첫 ${rows?'행':'열'}은 X값, 나머지 ${rows?'행':'열'}은 Y값입니다. X·Y가 모두 숫자인 점을 표시합니다.`):source.type==='pie'?(count>1?`원형은 첫 번째 계열의 모든 값을 표시합니다. 현재 ${count}개 계열 중 첫 계열만 사용합니다. 여러 계열은 도넛형을 사용하세요.`:'한 계열의 모든 값을 원의 조각으로 표시합니다. 0은 조각을 만들지 않습니다.'):source.type==='bar'&&count===1?'항목 이름은 각 막대 아래 가로축에 표시합니다. 한 계열은 같은 색을 쓰며 중복 범례는 생략합니다.':source.type==='doughnut'?'각 계열은 하나의 고리이며, 계열의 모든 값을 고리의 조각으로 표시합니다.':`숫자로 된 첫 행·첫 열은 값으로 포함합니다. 현재 ${count}개 계열입니다. 행·열 방향과 제목·항목 이름 설정으로 데이터 구성을 바꿀 수 있습니다.`;
    for(const id of ['chart-title','chart-legend','chart-labels','chart-hidden-data','chart-x-title','chart-y-title'])el.querySelector('#'+id).disabled=!!retained&&current.type==='unsupported';
    const simpleColumn=!retained&&source.type==='bar'&&count===1;el.querySelector('#chart-legend').checked=requestedLegend&&!simpleColumn;el.querySelector('#chart-legend').disabled=!!retained&&current.type==='unsupported'||simpleColumn;
    el.querySelector('#chart-secondary').disabled=!!retained||el.querySelector('#chart-type').value!=='combo';
  };
  const manual={firstRowAsHeader:current.firstRowAsHeader!==undefined,firstColumnAsLabels:current.firstColumnAsLabels!==undefined};
  const infer=()=>{const guessed=chartLayout(candidate(),sheet,workbook);for(const [key,id] of [['firstRowAsHeader','chart-first-row'],['firstColumnAsLabels','chart-first-column']])if(!manual[key])el.querySelector('#'+id).checked=guessed[key];};
  for(const [key,id] of [['firstRowAsHeader','chart-first-row'],['firstColumnAsLabels','chart-first-column']])el.querySelector('#'+id).addEventListener('change',()=>{manual[key]=true;syncControls();});
  el.querySelector('#chart-range').addEventListener('input',()=>{infer();syncControls();});
  el.querySelector('#chart-series-by').addEventListener('change',()=>{manual.firstRowAsHeader=false;manual.firstColumnAsLabels=false;infer();syncControls();});
  el.querySelector('#chart-convert')?.addEventListener('change',event=>{if(event.target.checked)el.querySelector('#chart-range').value=normalize(range);infer();syncControls();});
  el.querySelector('#chart-type').addEventListener('change',()=>{infer();syncControls();});
  el.querySelector('#chart-legend').addEventListener('change',()=>{requestedLegend=el.querySelector('#chart-legend').checked;});
  syncControls();
  return el;
}
