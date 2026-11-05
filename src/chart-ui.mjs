import {installSeriesEditor} from './chart-series-editor.mjs';
import {chartAxesHTML,syncChartAxes,readChartAxes,installChartAxes} from './chart-axis-ui.mjs';
import {chartTypes,chartLayout,chartRangeError,chartSeriesBy,chartSeries,chartSourceFormulas} from './chart-model.mjs';
import {escapeChartHTML as esc,layoutChartLegends,layoutChartCategories} from './chart-render.mjs';
import {chartBlankPolicy} from './chart-blanks.mjs';

export function chartDialog({chart,range,sheet,workbook,onApply,modal,preview}){
  const normalize=value=>value.includes(':')?value:value+':'+value;
  const current=chart||{type:'bar',grouping:'clustered',title:'차트 제목',range:normalize(range),legend:true,labels:false,placement:'move-size'};
  const initial=chartLayout(current,sheet,workbook);
  const sourceFormulas=current.source?chartSourceFormulas(current,workbook):[];
  const el=modal(chart?'차트 설정':'차트 삽입',`
    <div class="chart-settings-controls"><div class="chart-form">
    <label>차트 유형<select id="chart-type">${!chartTypes[current.type]?'<option value="'+esc(current.type)+'">Excel 원본 유형 유지</option>':''}${Object.entries(chartTypes).map(([value,label])=>`<option value="${value}" ${current.type===value?'selected':''}>${label}</option>`).join('')}</select></label>
    <label id="chart-data-mode-option">데이터 지정<select id="chart-data-mode"><option value="range">셀 범위에서 자동 구성</option><option value="series" ${current.series?.length&&!current.source?'selected':''} ${workbook?.nativeSource?'disabled':''}>계열별로 직접 지정</option></select></label>
    <label id="chart-series-by-option">데이터 계열 방향<select id="chart-series-by"><option value="auto">자동 (한 줄은 행 기준)</option><option value="rows">행 기준 (가로로 이어진 값)</option><option value="columns">열 기준 (세로로 이어진 값)</option></select></label>
    <label>배치<select id="chart-grouping">${Object.entries({clustered:'묶은형 / 일반',stacked:'누적',percentStacked:'100% 누적'}).map(([value,label])=>`<option value="${value}" ${current.grouping===value||value==='clustered'&&current.grouping==='standard'?'selected':''}>${label}</option>`).join('')}</select></label>
    <label>셀 변경 시 위치<select id="chart-placement" ${workbook?.nativeSource?'disabled':''}>${Object.entries({'move-size':'셀과 함께 이동·크기 조절',move:'셀과 함께 이동 (크기 유지)',fixed:'위치와 크기 고정'}).map(([value,label])=>`<option value="${value}" ${value===(current.placement||'fixed')?'selected':''}>${label}</option>`).join('')}</select></label>
    <label>제목<input id="chart-title" value="${esc(current.title)}" ${current.type==='unsupported'?'disabled':''}></label>
    <label id="chart-range-option" ${current.source?'hidden':''}>데이터 범위<input id="chart-range" value="${esc(current.range||range)}" ${current.source?'disabled':''}></label>
    <label class="chart-check" id="chart-first-row-option"><input id="chart-first-row" type="checkbox" ${initial.firstRowAsHeader?'checked':''}><span id="chart-first-row-label">첫 행을 계열 이름으로 사용</span></label>
    <label class="chart-check" id="chart-first-column-option"><input id="chart-first-column" type="checkbox" ${initial.firstColumnAsLabels?'checked':''}><span id="chart-first-column-label">첫 열을 항목 이름으로 사용</span></label>
    <label class="chart-check"><input id="chart-legend" type="checkbox" ${current.legend!==false?'checked':''}>범례 표시</label>
    <label class="chart-check"><input id="chart-labels" type="checkbox" ${current.labels?'checked':''}>데이터 레이블</label>
    <label class="chart-check"><input id="chart-hidden-data" type="checkbox" ${current.plotVisibleOnly===false?'checked':''}>숨겨진 행·열의 데이터도 표시</label>
    <label id="chart-blank-option">빈 셀 표시<select id="chart-blanks"><option value="gap">간격으로 표시</option><option value="zero">0으로 표시</option><option value="span">선으로 연결</option></select></label>
    <label>가로축 제목<input id="chart-x-title" value="${esc(current.xTitle||'')}"></label>
    <label>세로축 제목<input id="chart-y-title" value="${esc(current.yTitle||'')}"></label>
    <label class="chart-check"><input id="chart-secondary" type="checkbox" ${current.secondaryAxis!==false?'checked':''}>혼합 차트의 마지막 계열에 보조축 사용</label>
    </div>
    ${current.source?'<div id="chart-source-references"><p class="chart-help">원본 데이터 참조</p>'+(sourceFormulas.length?'<ul class="chart-source-list">'+sourceFormulas.slice(0,64).map(formula=>'<li><code>'+esc(formula)+'</code></li>').join('')+'</ul>'+(sourceFormulas.length>64?'<p class="chart-help">전체 '+sourceFormulas.length+'개 참조 중 처음 64개를 표시합니다.</p>':''):'<p class="chart-help">셀 참조 없이 파일에 저장된 계열 값을 사용합니다.</p>')+'</div>':''}
    ${current.source?'<p class="chart-help">Excel의 원래 계열과 고급 서식을 유지합니다. 유형·배치·데이터 범위·보조축을 바꾸려면 아래 변환을 선택하세요.</p><label class="chart-check"><input id="chart-convert" type="checkbox">새 차트로 변환 (원본 고급 서식은 제거됨)</label>':''}
    <section id="chart-series-editor" class="chart-series-editor" hidden aria-label="차트 계열 편집"></section>
    ${chartAxesHTML(current)}
    <p class="chart-help" id="chart-data-help"></p>
    <div id="chart-error" role="status"></div></div>
    <section class="chart-preview-panel" aria-label="차트 미리 보기"><h3>미리 보기와 데이터 구성</h3><p id="chart-preview-status" role="status" aria-live="polite">미리 보기를 준비합니다.</p><div id="chart-preview-image"></div><details open><summary>항목·계열·값 확인</summary><div id="chart-preview-data" class="chart-preview-data"></div><div id="chart-preview-refs" class="chart-preview-refs"></div></details></section>`,[{label:chart?'적용':'삽입',primary:true,action:el=>{
      try{onApply(readDraft().patch);el.remove();}catch(error){el.querySelector('#chart-error').textContent=error.message;}
    }}]);
  function readDraft(){
      const get=id=>el.querySelector('#'+id),convert=get('chart-convert')?.checked,type=get('chart-type').value,grouping=get('chart-grouping').value;
      const retained=current.source&&!convert;
      if(retained&&(type!==current.type||grouping!==(current.grouping==='standard'?'clustered':current.grouping||'clustered'))){throw Error('유형이나 배치를 바꾸려면 새 차트로 변환을 선택하세요.');}
      const selected=normalize(get('chart-range').value.toUpperCase().replace(/\$/g,'').trim());
      const direction=get('chart-series-by').value,layout={seriesBy:direction==='auto'?chartSeriesBy({type,range:selected},sheet,workbook):direction,firstRowAsHeader:get('chart-first-row').checked,firstColumnAsLabels:get('chart-first-column').checked};
      const custom=!retained&&get('chart-data-mode').value==='series';if(custom&&workbook?.nativeSource)throw Error('XLSM의 개별 계열 편집은 Excel에서 진행하거나 XLSX 사본을 사용하세요.');
      const seriesPlan=custom?seriesEditor.plan(type):null;
      const error=!retained&&!custom&&chartRangeError({type,range:selected,...layout},sheet,workbook);
      if(error)throw Error(error);
      if(!retained&&type==='unsupported'){throw Error('새 차트 유형을 선택하세요.');}
      const axes=readChartAxes(el);
      const patch={...axes,type,...(workbook?.nativeSource?{}:{placement:get('chart-placement').value}),grouping:retained?current.grouping:grouping,title:get('chart-title').value,legend:get('chart-legend').checked,labels:get('chart-labels').checked,plotVisibleOnly:!get('chart-hidden-data').checked,...(workbook?.nativeSource?{}:{blankCells:get('chart-blanks').value}),xTitle:get('chart-x-title').value,yTitle:get('chart-y-title').value,secondaryAxis:get('chart-secondary').checked};
      if(!retained)Object.assign(patch,{range:seriesPlan?.range||selected,...layout,series:seriesPlan?.series,source:undefined,rebuilt:true,previewNote:undefined,unsupportedReason:undefined});
      return {patch,retained};
  }
  el.querySelector('.modal').classList.add('chart-modal');if(preview)el.querySelector('.modal').classList.add('chart-preview-modal');installChartAxes(el);
  el.querySelector('#chart-series-by').value=current.seriesBy||'auto';
  el.querySelector('#chart-blanks').value=chartBlankPolicy(current);
  const candidate=()=>{const type=el.querySelector('#chart-type').value,range=normalize(el.querySelector('#chart-range').value.toUpperCase().replace(/\$/g,'').trim()),direction=el.querySelector('#chart-series-by').value;return {type,range,...(direction==='auto'?{}:{seriesBy:direction})};};
  const seriesEditor=installSeriesEditor({el,current,sheet,book:workbook,getRangeChart:()=>({...candidate(),firstRowAsHeader:el.querySelector('#chart-first-row').checked,firstColumnAsLabels:el.querySelector('#chart-first-column').checked}),onChange:()=>{syncControls();schedulePreview();}});
  let previewTimer,previewSequence=0;
  function schedulePreview(){
    clearTimeout(previewTimer);const sequence=++previewSequence;
    if(!preview){el.querySelector('.chart-preview-panel').hidden=true;return;}
    const status=el.querySelector('#chart-preview-status');status.textContent='미리 보기를 계산하는 중…';
    previewTimer=setTimeout(async()=>{
      try{
        const {patch,retained}=readDraft(),result=await preview({chartId:chart?.id,retain:!!retained,patch});if(!el.isConnected||sequence!==previewSequence)return;
        const visible=result.type==='pie'?Math.min(1,result.seriesCount):result.seriesCount;
        status.textContent=`${result.seriesCount}개 계열 · 차트에 ${visible}개 계열 사용`+(result.truncated?' · 화면 차트와 미리 보기는 계열당 처음 1,000개 값입니다. XLSX에는 전체 참조를 저장합니다.':'')+(result.usesCachedValues?' · 미지원 수식의 파일 저장값이 포함되어 있습니다. Excel에서 다시 계산하세요.':'')+(result.note?' · '+result.note:'');
        el.querySelector('#chart-preview-image').innerHTML='<h4>'+esc(result.title)+'</h4><svg role="img" aria-label="차트 미리 보기" viewBox="'+esc(result.viewBox)+'">'+result.svg+'</svg>';
        const xy=['scatter','bubble'].includes(result.type),bubble=result.type==='bubble';
        el.querySelector('#chart-preview-data').innerHTML='<p>처음 8개 계열·계열당 8개 값의 구성입니다. '+result.points.map(s=>esc(s.series)+': '+s.total.toLocaleString()+'개').join(' · ')+'</p><table><thead><tr><th>계열</th><th>'+(xy?'X값':'항목 이름')+'</th><th>'+(xy?'Y값':'값')+'</th>'+(bubble?'<th>크기</th>':'')+'</tr></thead><tbody>'+result.rows.map(row=>'<tr'+(!row.used?' class="chart-unused-series"':'')+'><td>'+esc(row.series)+(!row.used?' (원형에 미사용)':'')+'</td><td>'+esc(xy?row.x:row.item)+'</td><td>'+esc(row.value)+'</td>'+(bubble?'<td>'+esc(row.size)+'</td>':'')+'</tr>').join('')+'</tbody></table>';
        el.querySelector('#chart-preview-refs').innerHTML=result.references.map(ref=>'<p><strong>'+esc(ref.series)+'</strong> · 값 <code>'+esc(ref.values)+'</code> · '+(xy?'X값':'항목')+' <code>'+esc(ref.x||ref.items)+'</code>'+(ref.size?' · 크기 <code>'+esc(ref.size)+'</code>':'')+'</p>').join('');
        layoutChartLegends(el);layoutChartCategories(el);
      }catch(error){if(el.isConnected&&sequence===previewSequence){status.textContent=error.message;el.querySelector('#chart-preview-image').textContent='';el.querySelector('#chart-preview-data').textContent='';el.querySelector('#chart-preview-refs').textContent='';}}
    },180);
  }
  const observer=new MutationObserver(()=>{if(!el.isConnected){clearTimeout(previewTimer);++previewSequence;observer.disconnect();}});observer.observe(el.parentNode,{childList:true});
  let requestedLegend=current.legend!==false;
  const syncControls=()=>{
    const retained=current.source&&!el.querySelector('#chart-convert')?.checked,custom=!retained&&el.querySelector('#chart-data-mode').value==='series';
    el.querySelector('#chart-data-mode-option').hidden=!!retained;el.querySelector('#chart-series-editor').hidden=!custom;el.querySelector('#chart-series-by-option').hidden=custom;
    el.querySelector('#chart-range').disabled=!!retained;
    el.querySelector('#chart-range-option').hidden=!!retained||custom;
    if(current.source)el.querySelector('#chart-source-references').hidden=!retained;
    el.querySelector('#chart-first-row').disabled=!!retained;
    const source=candidate(),rows=chartSeriesBy(source,sheet,workbook)==='rows',xy=['scatter','bubble'].includes(source.type),parts=source.range.split(':'),single=parts.length===2&&(rows?parts[0].match(/\d+$/)?.[0]===parts[1].match(/\d+$/)?.[0]:parts[0].match(/^[A-Z]+/)?.[0]===parts[1].match(/^[A-Z]+/)?.[0]),implicitX=source.type==='scatter'&&single;
    el.querySelector('#chart-series-by').disabled=!!retained;
    for(const [id,isX] of [['chart-first-row',rows],['chart-first-column',!rows]]){el.querySelector('#'+id).disabled=!!retained||xy&&!implicitX&&isX;el.querySelector('#'+id+'-option').hidden=custom||xy&&!implicitX&&isX;}
    el.querySelector('#chart-first-row-label').textContent=rows?'첫 행을 항목 이름으로 사용':'첫 행을 계열 이름으로 사용';
    el.querySelector('#chart-first-column-label').textContent=rows?'첫 열을 계열 이름으로 사용':'첫 열을 항목 이름으로 사용';
    const resolved={...source,firstRowAsHeader:el.querySelector('#chart-first-row').checked,firstColumnAsLabels:el.querySelector('#chart-first-column').checked},series=retained?current.series:custom?seriesEditor.series():chartSeries(resolved,sheet,workbook),count=series?.length||0;
    el.querySelector('#chart-data-help').textContent=custom?'계열별 참조를 사용합니다. 추가·삭제·순서는 적용 전 미리 보기에서 확인하세요. 원형은 첫 계열만, 도넛형은 각 계열을 고리로 표시합니다.':retained?`원본의 ${count}개 계열과 각각의 데이터 참조를 유지합니다. 데이터 구성을 바꾸려면 새 차트로 변환한 뒤 사용할 셀 범위를 지정하세요.`:source.type==='bubble'?`거품형은 X값 다음에 Y값·크기 ${rows?'행':'열'} 쌍이 필요합니다. 각 항목의 X·Y·크기 자료를 함께 선택하세요.`:source.type==='scatter'?(implicitX?'한 줄의 값은 Y값으로 사용하고 X값은 1, 2, 3…으로 지정합니다.':`첫 ${rows?'행':'열'}은 X값, 나머지 ${rows?'행':'열'}은 Y값입니다. X·Y가 모두 숫자인 점을 표시합니다.`):source.type==='pie'?(count>1?`원형은 첫 번째 계열의 모든 값을 표시합니다. 현재 ${count}개 계열 중 첫 계열만 사용합니다. 여러 계열은 도넛형을 사용하세요.`:'한 계열의 모든 값을 원의 조각으로 표시합니다. 0은 조각을 만들지 않습니다.'):source.type==='bar'&&count===1?'항목 이름은 각 막대 아래 가로축에 표시합니다. 한 계열은 같은 색을 쓰며 중복 범례는 생략합니다.':source.type==='doughnut'?'각 계열은 하나의 고리이며, 계열의 모든 값을 고리의 조각으로 표시합니다.':`숫자로 된 첫 행·첫 열은 값으로 포함합니다. 현재 ${count}개 계열입니다. 행·열 방향과 제목·항목 이름 설정으로 데이터 구성을 바꿀 수 있습니다.`;
    syncChartAxes(el,{type:source.type,secondary:source.type==='combo'||series?.some(s=>s.secondary),native:workbook?.nativeSource});
    for(const id of ['chart-title','chart-legend','chart-labels','chart-hidden-data','chart-x-title','chart-y-title'])el.querySelector('#'+id).disabled=!!retained&&current.type==='unsupported';
    const simpleColumn=!retained&&source.type==='bar'&&count===1;el.querySelector('#chart-legend').checked=requestedLegend&&!simpleColumn;el.querySelector('#chart-legend').disabled=!!retained&&current.type==='unsupported'||simpleColumn;
    const blanks=['line','scatter','radar','combo'].includes(source.type);el.querySelector('#chart-blank-option').hidden=!blanks;el.querySelector('#chart-blanks').disabled=!blanks||!!workbook?.nativeSource;el.querySelector('#chart-blanks [value="span"]').disabled=source.type==='scatter'&&!series?.some(s=>s.lines);
    el.querySelector('#chart-secondary').disabled=!!retained||custom||el.querySelector('#chart-type').value!=='combo';
  };
  const manual={firstRowAsHeader:current.firstRowAsHeader!==undefined,firstColumnAsLabels:current.firstColumnAsLabels!==undefined};
  const infer=()=>{const guessed=chartLayout(candidate(),sheet,workbook);for(const [key,id] of [['firstRowAsHeader','chart-first-row'],['firstColumnAsLabels','chart-first-column']])if(!manual[key])el.querySelector('#'+id).checked=guessed[key];};
  for(const [key,id] of [['firstRowAsHeader','chart-first-row'],['firstColumnAsLabels','chart-first-column']])el.querySelector('#'+id).addEventListener('change',()=>{manual[key]=true;syncControls();});
  el.querySelector('#chart-range').addEventListener('input',()=>{infer();syncControls();});
  el.querySelector('#chart-series-by').addEventListener('change',()=>{manual.firstRowAsHeader=false;manual.firstColumnAsLabels=false;infer();syncControls();});
  el.querySelector('#chart-convert')?.addEventListener('change',event=>{if(event.target.checked){el.querySelector('#chart-range').value=normalize(range);if(!workbook?.nativeSource){el.querySelector('#chart-data-mode').value='series';seriesEditor.initialize();}}infer();syncControls();});
  el.querySelector('#chart-data-mode').addEventListener('change',()=>{if(el.querySelector('#chart-data-mode').value==='series')seriesEditor.initialize();syncControls();schedulePreview();});
  el.querySelector('#chart-type').addEventListener('change',()=>{seriesEditor.refresh();infer();syncControls();});
  el.querySelector('#chart-legend').addEventListener('change',()=>{requestedLegend=el.querySelector('#chart-legend').checked;});
  for(const type of ['input','change'])el.addEventListener(type,event=>{if(event.target.matches('input,select')){el.querySelector('#chart-error').textContent='';schedulePreview();}});
  el.addEventListener('click',event=>{if(event.target.closest('[data-axis-reset]'))schedulePreview();});
  if(el.querySelector('#chart-data-mode').value==='series')seriesEditor.initialize();syncControls();schedulePreview();
  return el;
}
