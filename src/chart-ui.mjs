import {chartTypes} from './chart-model.mjs';
import {parseAddress} from './engine.mjs';
import {escapeChartHTML as esc} from './chart-render.mjs';

export function chartDialog({chart,range,onApply,modal}){
  const current=chart||{type:'bar',grouping:'clustered',title:'차트 제목',range,legend:true,labels:false};
  const el=modal(chart?'차트 설정':'차트 삽입',`
    <div class="chart-form">
    <label>차트 유형<select id="chart-type">${!chartTypes[current.type]?'<option value="'+esc(current.type)+'">Excel 원본 유형 유지</option>':''}${Object.entries(chartTypes).map(([value,label])=>`<option value="${value}" ${current.type===value?'selected':''}>${label}</option>`).join('')}</select></label>
    <label>배치<select id="chart-grouping">${Object.entries({clustered:'묶은형 / 일반',stacked:'누적',percentStacked:'100% 누적'}).map(([value,label])=>`<option value="${value}" ${current.grouping===value||value==='clustered'&&current.grouping==='standard'?'selected':''}>${label}</option>`).join('')}</select></label>
    <label>제목<input id="chart-title" value="${esc(current.title)}" ${current.type==='unsupported'?'disabled':''}></label>
    <label>데이터 범위<input id="chart-range" value="${esc(current.range||range)}" ${current.source?'disabled':''}></label>
    <label class="chart-check"><input id="chart-legend" type="checkbox" ${current.legend!==false?'checked':''}>범례 표시</label>
    <label class="chart-check"><input id="chart-labels" type="checkbox" ${current.labels?'checked':''}>데이터 레이블</label>
    <label>가로축 제목<input id="chart-x-title" value="${esc(current.xTitle||'')}"></label>
    <label>세로축 제목<input id="chart-y-title" value="${esc(current.yTitle||'')}"></label>
    <label class="chart-check"><input id="chart-secondary" type="checkbox" ${current.secondaryAxis!==false?'checked':''}>혼합 차트의 마지막 계열에 보조축 사용</label>
    </div>
    ${current.source?'<p class="chart-help">Excel의 원래 계열과 고급 서식을 유지합니다. 유형·배치·데이터 범위·보조축을 바꾸려면 아래 변환을 선택하세요.</p><label class="chart-check"><input id="chart-convert" type="checkbox">새 차트로 변환 (원본 고급 서식은 제거됨)</label>':''}
    <p class="chart-help">첫 행은 계열 이름, 첫 열은 범주, 나머지 열은 값입니다. 분산형은 첫 열을 X값으로 사용합니다. 거품형은 X값·Y값·크기 순서입니다. 원형은 첫 계열만 사용합니다.</p>
    <div id="chart-error" role="status"></div>`,[{label:chart?'적용':'삽입',primary:true,action:el=>{
      const get=id=>el.querySelector('#'+id),convert=get('chart-convert')?.checked,type=get('chart-type').value,grouping=get('chart-grouping').value;
      const retained=current.source&&!convert;
      if(retained&&(type!==current.type||grouping!==(current.grouping==='standard'?'clustered':current.grouping||'clustered'))){get('chart-error').textContent='유형이나 배치를 바꾸려면 새 차트로 변환을 선택하세요.';return;}
      const selected=get('chart-range').value.toUpperCase().replace(/\$/g,'').trim(),[a,b]=selected.split(':').map(parseAddress);
      if(!retained&&(!a||!b||a.r>=b.r||a.c>=b.c||b.r-a.r>10000||b.c-a.c>32)){get('chart-error').textContent='제목 행을 포함한 범위를 입력하세요. 최대 10,000행·32개 값 열을 지원합니다.';return;}
      if(!retained&&type==='bubble'&&b.c-a.c<2){get('chart-error').textContent='거품형에는 X값, Y값, 크기 열이 필요합니다.';return;}
      if(!retained&&type==='unsupported'){get('chart-error').textContent='새 차트 유형을 선택하세요.';return;}
      const patch={type,grouping:retained?current.grouping:grouping,title:get('chart-title').value,legend:get('chart-legend').checked,labels:get('chart-labels').checked,xTitle:get('chart-x-title').value,yTitle:get('chart-y-title').value,secondaryAxis:get('chart-secondary').checked};
      if(!retained)Object.assign(patch,{range:selected,series:undefined,source:undefined,rebuilt:true,previewNote:undefined,unsupportedReason:undefined});
      onApply(patch);el.remove();
    }}]);
  const syncControls=()=>{
    const retained=current.source&&!el.querySelector('#chart-convert')?.checked;
    el.querySelector('#chart-range').disabled=!!retained;
    for(const id of ['chart-title','chart-legend','chart-labels','chart-x-title','chart-y-title'])el.querySelector('#'+id).disabled=!!retained&&current.type==='unsupported';
    el.querySelector('#chart-secondary').disabled=!!retained||el.querySelector('#chart-type').value!=='combo';
  };
  el.querySelector('#chart-convert')?.addEventListener('change',event=>{if(event.target.checked)el.querySelector('#chart-range').value=range;syncControls();});
  el.querySelector('#chart-type').addEventListener('change',syncControls);
  syncControls();
  return el;
}
