import {chartData,chartColors} from './chart-model.mjs';
import {chartRect} from './chart-geometry.mjs';
import {axisScale,axisLabelStyle} from './chart-axis.mjs';
let plotSequence=0;
import {formatValue} from './engine.mjs';
const esc=value=>String(value??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
const formatChartValue=(n,style={},workbook={})=>style.numberFormat&&style.numberFormat!=='General'||style.format&&style.format!=='general'?formatValue(n,style,workbook):Math.abs(n)>=1e6?(n/1e6).toLocaleString(undefined,{maximumFractionDigits:1})+'M':Number(n.toPrecision(4)).toLocaleString(undefined,{maximumSignificantDigits:4});
const text=(x,y,value,extra='')=>`<text x="${x}" y="${y}" font-size="10" fill="#555" ${extra}>${esc(value)}</text>`;
const line=(x,y,x2,y2,extra='')=>`<line x1="${x}" y1="${y}" x2="${x2}" y2="${y2}" stroke="#e3e6e9" ${extra}/>`;
const poly=(points,color,fill='none',extra='')=>`<polyline points="${points.map(p=>p.join(',')).join(' ')}" stroke="${color}" fill="${fill}" stroke-width="2" ${extra}/>`;

const legendSVG=(items,cols,rows,w,h,marker=9,baseline=8)=>Array.from({length:rows},(_,row)=>`<g class="chart-legend-row" data-legend-width="${w}" transform="translate(0 ${h-baseline-(rows-1-row)*17})">${items.slice(row*cols,(row+1)*cols).map(item=>`<g class="chart-legend-item" aria-label="${esc(item.name)}"><title>${esc(item.name)}</title><rect x="0" y="-8" width="${marker}" height="${marker}" fill="${item.color}"/>${text(14,0,item.name)}</g>`).join('')}</g>`).join('');

// SVG text widths depend on the actual font (including Korean fallback fonts).
// Measure after insertion so each row's painted bounds, rather than fixed slots,
// are centered. Re-run when charts are redrawn after resizing or zooming.
export function layoutChartLegends(root){
 for(const row of root.querySelectorAll('.chart-legend-row')){
  const items=[...row.querySelectorAll('.chart-legend-item')],w=+row.dataset.legendWidth,gap=16,maxWidth=(w-gap*(items.length-1))/items.length;
  const boxes=items.map(item=>{
   const label=item.querySelector('text'),full=item.getAttribute('aria-label');label.textContent=full;
   if(label.getComputedTextLength()>maxWidth-14){const chars=Array.from(full);let low=0,high=chars.length;while(low<high){const mid=Math.ceil((low+high)/2);label.textContent=chars.slice(0,mid).join('')+'…';if(label.getComputedTextLength()<=maxWidth-14)low=mid;else high=mid-1;}label.textContent=chars.slice(0,low).join('')+'…';}
   return item.getBBox();
  });
  let x=(w-boxes.reduce((sum,b)=>sum+b.width,0)-gap*(items.length-1))/2;
  items.forEach((item,i)=>{item.setAttribute('transform',`translate(${x-boxes[i].x} 0)`);x+=boxes[i].width+gap;});
 }
}

export function layoutChartCategories(root){
 for(const group of root.querySelectorAll('.chart-category')){
  const label=group.querySelector('text'),full=group.getAttribute('aria-label').replace(/\s+/g,' ').trim(),width=+group.dataset.labelWidth,x=label.getAttribute('x');
  label.textContent=full;if(label.getComputedTextLength()<=width)continue;
  const chars=Array.from(full),fit=(text,suffix='')=>{let lo=0,hi=text.length;while(lo<hi){const mid=Math.ceil((lo+hi)/2);label.textContent=text.slice(0,mid).join('')+suffix;if(label.getComputedTextLength()<=width)lo=mid;else hi=mid-1;}return lo;};
  const first=fit(chars),rest=chars.slice(first),second=fit(rest),clipped=second<rest.length;
  const lines=[chars.slice(0,first).join(''),clipped?rest.slice(0,fit(rest,'…')).join('')+'…':rest.join('')];
  label.replaceChildren(...lines.map((value,i)=>{const span=document.createElementNS('http://www.w3.org/2000/svg','tspan');span.setAttribute('x',x);span.setAttribute('dy',i?'14':'0');span.textContent=value;return span;}));
 }
}

export function renderChartSVG(chart,sheet,workbook){
  const fmt=(value,style)=>formatChartValue(value,style,workbook);
  const {width,height}=chartRect(chart),w=width-36,h=height-70,data=chartData(chart,sheet,workbook);
  if(chart.type==='unsupported')return {viewBox:`0 0 ${w} ${h}`,svg:text(w/2,h/2,'Excel 원본 차트 보존','text-anchor="middle"')+text(w/2,h/2+20,chart.unsupportedReason||'이 유형의 미리 보기는 아직 지원하지 않습니다.','text-anchor="middle"')};
  if(!data.length||!data.some(s=>s.values.some(Number.isFinite)))return {viewBox:`0 0 ${w} ${h}`,svg:text(w/2,h/2,'표시할 숫자 데이터가 없습니다.','text-anchor="middle"')+text(w/2,h/2+20,'범위와 첫 행·첫 열 설정을 확인하세요.','text-anchor="middle"')};
  const circular=['pie','doughnut'].includes(chart.type),legend=chart.legend!==false&&(chart.type!=='bar'||data.length>1),legendCols=Math.max(1,Math.floor(w/(circular?95:130))),legendCount=circular?data[0].values.length:data.length,legendRows=legend?Math.min(3,Math.ceil(legendCount/legendCols)):0;
  const categorySpace=chart.type==='bar'?(chart.xTitle?64:44):35;
  const left=chart.type==='bar-horizontal'?75:55,right=w-(data.some(s=>s.secondary)?52:20),top=18,bottom=Math.max(top+35,h-categorySpace-legendRows*17),pw=right-left,ph=bottom-top;
  const svg=[],labelSVG=[],plotId='chart-plot-'+(++plotSequence),beginPlot=()=>svg.push(`<defs><clipPath id="${plotId}"><rect x="${left}" y="${top}" width="${Math.max(1,pw)}" height="${Math.max(1,ph)}"/></clipPath></defs><g class="chart-data" clip-path="url(#${plotId})">`),endPlot=()=>svg.push('</g>'),count=Math.max(1,...data.map(s=>s.values.length)),labels=data.find(s=>s.labels.length)?.labels||Array.from({length:count},(_,i)=>String(i+1));
  const pointLabel=(x,y,value,style)=>{if(chart.labels&&Number.isFinite(value))labelSVG.push(text(Math.max(left+4,Math.min(right-4,x)),Math.max(top+10,Math.min(bottom-2,y)),fmt(value,style),'class="chart-data-label" text-anchor="middle"'));};
  if(['pie','doughnut'].includes(chart.type)){
    const series=chart.type==='pie'?data.slice(0,1):data,cx=w/2,cy=(top+bottom)/2,r=Math.max(15,Math.min(pw,ph)/2-10);
    series.forEach((s,k)=>{const values=s.values.map(v=>Math.abs(v||0)),sum=values.reduce((a,b)=>a+b,0);let start=-Math.PI/2;
      values.forEach((v,i)=>{if(!v||!sum)return;const angle=v/sum*Math.PI*2,end=start+angle,color=s.pointColors?.[i]||chartColors[i%chartColors.length],outer=chart.type==='pie'?r:r-k*r/(series.length+1),inner=chart.type==='pie'?0:outer-r/(series.length+1)*.78;
        const p=(rad,a)=>[cx+rad*Math.cos(a),cy+rad*Math.sin(a)],label=esc(s.labels[i]??String(i+1)),attrs=`class="chart-point" data-series="${k}" data-point="${i}"`,title=`<title>${label}: ${v}</title>`;let d;
        if(angle>Math.PI*2-.00001){svg.push(`<circle ${attrs} cx="${cx}" cy="${cy}" r="${(outer+inner)/2}" fill="${inner?'none':color}" stroke="${color}" stroke-width="${inner?outer-inner:0}">${title}</circle>`);}
        else {const [a,b]=p(outer,start),[c,d2]=p(outer,end),[e,f]=p(inner,end),[g,j]=p(inner,start);d=`M ${a} ${b} A ${outer} ${outer} 0 ${angle>Math.PI?1:0} 1 ${c} ${d2} L ${e} ${f} ${inner?`A ${inner} ${inner} 0 ${angle>Math.PI?1:0} 0 ${g} ${j}`:''} Z`;svg.push(`<path ${attrs} d="${d}" fill="${color}" stroke="white" stroke-width="1">${title}</path>`);}
        if(chart.labels){const [x,y]=p((outer+inner)/2,start+angle/2);svg.push(text(x,y,Math.round(v/sum*100)+'%','text-anchor="middle"'));}start=end;
      });
    });
    if(chart.legend!==false){const capacity=legendCols*legendRows,items=labels.slice(0,capacity).map((name,i)=>({name:i===capacity-1&&labels.length>capacity?'외 '+(labels.length-i)+'개 항목':name,color:data[0].pointColors?.[i]||chartColors[i%chartColors.length]}));svg.push(legendSVG(items,legendCols,legendRows,w,h,8,5));}
  }else if(chart.type==='radar'){
    const cx=(left+right)/2,cy=(top+bottom)/2,r=Math.min(pw,ph)/2-16,scale=axisScale(data.flatMap(s=>s.values),chart.valueAxis),radius=v=>scale.ratio(v)??0,n=Math.max(3,count),pos=(i,v)=>[cx+Math.sin(i/n*2*Math.PI)*r*v,cy-Math.cos(i/n*2*Math.PI)*r*v];
    for(const tick of scale.ticks){const ratio=radius(tick);if(chart.valueAxis?.gridlines!==false)svg.push(poly(Array.from({length:n+1},(_,i)=>pos(i%n,ratio)),'#ddd'));svg.push(text(cx+4,cy-r*ratio+4,fmt(tick,axisLabelStyle(chart.valueAxis,data[0].valueFormat)),'class="chart-axis-value"'));}
    labels.forEach((label,i)=>{const [x,y]=pos(i,1.15);svg.push(line(cx,cy,...pos(i,1)));svg.push(text(x,y,label.slice(0,12),'text-anchor="middle"'));});
    svg.push(`<defs><clipPath id="${plotId}"><circle cx="${cx}" cy="${cy}" r="${r}"/></clipPath></defs><g class="chart-data" clip-path="url(#${plotId})">`);data.forEach((s,si)=>{svg.push(poly([...s.values,s.values[0]].map((v,i)=>pos(i%n,radius(v||0))),s.color,chart.radarFilled?s.color:'none',chart.radarFilled?'fill-opacity=".16"':''));s.values.forEach((v,i)=>{if(!Number.isFinite(v)||!scale.contains(v))return;const [x,y]=pos(i%n,radius(v));svg.push(`<circle class="chart-point" data-series="${si}" data-point="${i}" cx="${x}" cy="${y}" r="3" fill="${s.color}"><title>${esc(s.name)} · ${esc(labels[i])}: ${v}</title></circle>`);pointLabel(x,y-7,v,s.valueFormat);});});endPlot();
  }else if(['scatter','bubble'].includes(chart.type)){
    const valueX=(s,i)=>s.xRef?s.xs[i]:i+1,xs=data.flatMap(s=>s.values.map((v,i)=>valueX(s,i))).filter(Number.isFinite),ys=data.flatMap(s=>s.values).filter(Number.isFinite);
    const sx=axisScale(xs,chart.xAxis),sy=axisScale(ys,chart.valueAxis),x=v=>left+sx.ratio(v)*pw,y=v=>bottom-sy.ratio(v)*ph;
    for(const value of sy.ticks){const yy=y(value);if(chart.valueAxis?.gridlines!==false)svg.push(line(left,yy,right,yy));svg.push(text(left-6,yy+3,fmt(value,axisLabelStyle(chart.valueAxis,data[0].valueFormat)),'class="chart-axis-value" text-anchor="end"'));}
    for(const value of sx.ticks){const xx=x(value);if(chart.xAxis?.gridlines)svg.push(line(xx,top,xx,bottom));svg.push(text(xx,bottom+15,fmt(value,axisLabelStyle(chart.xAxis,data[0].xFormat)),'class="chart-axis-x-value" text-anchor="middle"'));}
    beginPlot();
    const maxSize=data.flatMap(s=>s.sizes).reduce((max,value)=>Math.max(max,value),1);
    data.forEach((s,si)=>{const points=s.values.map((v,i)=>Number.isFinite(v)&&Number.isFinite(valueX(s,i))&&sx.ratio(valueX(s,i))!==null&&sy.ratio(v)!==null&&(chart.type!=='bubble'||s.sizes[i]>0)?[x(valueX(s,i)),y(v)]:null);
      if(s.lines){let segment=[];for(const p of [...points,null]){if(p)segment.push(p);else if(segment.length){svg.push(poly(segment,s.color));segment=[];}}}points.forEach((p,i)=>{if(!p)return;const radius=chart.type==='bubble'?Math.max(2,Math.sqrt((s.sizes[i]||0)/maxSize)*Math.min(pw,ph)*.09):4;svg.push(`<circle class="chart-point" data-series="${si}" data-point="${i}" cx="${p[0]}" cy="${p[1]}" r="${radius}" fill="${s.color}" fill-opacity=".65"><title>${esc(s.name)}: ${valueX(s,i)}, ${s.values[i]}</title></circle>`);if(sx.contains(valueX(s,i))&&sy.contains(s.values[i]))pointLabel(p[0],p[1]-radius-3,s.values[i],s.valueFormat);});});endPlot();
  }else{
    const horizontal=chart.type==='bar-horizontal',stack=chart.grouping==='stacked'||chart.grouping==='percentStacked',percent=chart.grouping==='percentStacked';
    // Excel stacks series within the same chart type and axis group. A line
    // on a secondary axis must never inherit a column's accumulated height.
    const groupKey=s=>`${s.type}:${!!s.secondary}`,groups=new Map();
    data.forEach((s,i)=>{const key=groupKey(s);if(!groups.has(key))groups.set(key,[]);groups.get(key).push(i);});
    const totals=new Map([...groups].map(([key,indices])=>[key,Array.from({length:count},(_,k)=>indices.reduce((sum,i)=>sum+Math.abs(data[i].values[k]||0),0))]));
    const values=data.map(s=>s.values.map((v,i)=>v===null?null:percent?(totals.get(groupKey(s))[i]?v/totals.get(groupKey(s))[i]*100:0):v));
    const extents=secondary=>{
      let vals=[];
      for(const indices of groups.values()){
        if(!!data[indices[0]].secondary!==secondary)continue;
        vals.push(...(stack?Array.from({length:count},(_,k)=>[indices.reduce((a,i)=>a+Math.max(0,values[i][k]||0),0),indices.reduce((a,i)=>a+Math.min(0,values[i][k]||0),0)]).flat():indices.flatMap(i=>values[i]).filter(Number.isFinite)));
      }
      return vals;
    };
    const scales=[axisScale(extents(false),chart.valueAxis,{percent}),axisScale(extents(true),chart.secondaryValueAxis,{percent})],coord=(value,secondary=false)=>{const scale=scales[+secondary],ratio=scale.ratio(value);return horizontal?left+ratio*pw:bottom-ratio*ph;};
    for(const val of scales[0].ticks){const p=coord(val),label=fmt(percent?val/100:val,axisLabelStyle(chart.valueAxis,data.find(s=>!s.secondary)?.valueFormat,percent));if(chart.valueAxis?.gridlines!==false)svg.push(horizontal?line(p,top,p,bottom):line(left,p,right,p));svg.push(horizontal?text(p,bottom+16,label,'class="chart-axis-value" text-anchor="middle"'):text(left-6,p+3,label,'class="chart-axis-value" text-anchor="end"'));}
    if(data.some(s=>s.secondary)&&!horizontal)for(const val of scales[1].ticks){const p=coord(val,true);if(chart.secondaryValueAxis?.gridlines)svg.push(line(left,p,right,p));svg.push(text(right+6,p+3,fmt(percent?val/100:val,axisLabelStyle(chart.secondaryValueAxis,data.find(s=>s.secondary)?.valueFormat,percent)),'class="chart-axis-secondary-value"'));}
    const dateValues=data.find(s=>s.dateCategories)?.categories;
    const dates=dateValues?.length===count&&dateValues.every(Number.isFinite)?dateValues:null;
    const unique=dates?[...new Set(dates)].sort((a,b)=>a-b):[],unit=unique.length>1?Math.min(...unique.slice(1).map((v,i)=>v-unique[i])):1;
    const dateMin=dates?Math.min(...dates):0,span=dates?Math.max(...dates)-dateMin+unit:count;
    const step=(horizontal?ph:pw)*(dates?unit:1)/span,categoryMid=i=>(horizontal?top:left)+step*.5+(dates?(dates[i]-dateMin)/unit:i)*step;
    const barIndices=data.map((s,i)=>['bar','bar-horizontal','combo'].includes(s.type)?i:-1).filter(i=>i>=0),positions=new Map([...groups.keys()].map(key=>[key,Array.from({length:count},()=>({positive:0,negative:0}))]));
    const dateLabelWidth=Math.max(60,...labels.map(label=>Math.min(10,label.length)*6))+10,dateLabels=[];
    if(dates){
      labels.forEach((_,i)=>{if(!dateLabels.length||categoryMid(i)-categoryMid(dateLabels.at(-1))>=dateLabelWidth)dateLabels.push(i);});
      const last=labels.length-1;
      if(dateLabels.at(-1)!==last){if(dateLabels.length>1)dateLabels.pop();if(categoryMid(last)-categoryMid(dateLabels.at(-1))>=dateLabelWidth)dateLabels.push(last);}
    }
    labels.forEach((label,i)=>{if(dates?!dateLabels.includes(i):count>16&&i%Math.ceil(count/12)!==0)return;const x=categoryMid(i);svg.push(chart.type==='bar'?`<g class="chart-category" aria-label="${esc(label)}" data-label-width="${Math.max(12,dates?dateLabelWidth-10:step-10)}"><title>${esc(label)}</title><text x="${x}" y="${bottom+18}" text-anchor="middle" font-size="11" fill="#56616f">${esc(label)}</text></g>`:horizontal?text(left-7,categoryMid(i)+3,label.slice(0,10),'class="chart-axis-category" text-anchor="end"'):text(x,bottom+16,label.slice(0,10),'class="chart-axis-category" text-anchor="middle"'));});
    beginPlot();data.forEach((s,si)=>{const points=[],basePoints=[];values[si].forEach((v,i)=>{if(v===null){points.push(null);basePoints.push(null);return;}const side=v>=0?'positive':'negative',position=positions.get(groupKey(s))[i],base=stack?position[side]:0,topValue=base+v;if(stack)position[side]=topValue;
        const scale=scales[+!!s.secondary];if(scale.ratio(topValue)===null){points.push(null);basePoints.push(null);return;}const p=coord(topValue,!!s.secondary),baseP=coord(scale.log&&base<=0?scale.min:base,!!s.secondary),mid=categoryMid(i),visible=scale.contains(topValue);
        if(['bar','bar-horizontal','combo'].includes(s.type)){const lane=stack?0:barIndices.indexOf(si),lanes=stack?1:Math.max(1,barIndices.length),simple=chart.type==='bar'&&data.length===1,thickness=simple?Math.min(56,step*.56):step*.75/lanes,offset=simple?mid-thickness/2:mid-step*.375+lane*thickness;
          svg.push(`<rect class="chart-point" data-series="${si}" data-point="${i}" x="${horizontal?Math.min(p,baseP):offset}" y="${horizontal?offset:Math.min(p,baseP)}" width="${horizontal?Math.max(.5,Math.abs(p-baseP)):thickness}" height="${horizontal?thickness:Math.max(.5,Math.abs(p-baseP))}" ${simple?'rx="3"':''} fill="${s.pointColors?.[i]||s.color}"><title>${esc(s.name)} · ${esc(labels[i])}: ${v}</title></rect>`);if(visible)pointLabel(horizontal?p+(v<0?-10:10):offset+thickness/2,horizontal?offset+thickness/2:p+(v<0?14:-4),percent?v/100:v,percent?{format:'percent'}:s.valueFormat);
        }else{points.push([mid,p]);basePoints.push([mid,baseP]);svg.push(`<circle class="chart-point" data-series="${si}" data-point="${i}" cx="${mid}" cy="${p}" r="3" fill="${s.color}"><title>${esc(s.name)} · ${esc(labels[i])}: ${v}</title></circle>`);if(visible)pointLabel(mid,p-7,percent?v/100:v,percent?{format:'percent'}:s.valueFormat);}
      });
      if(points.length){let segment=[],base=[];const flush=()=>{if(!segment.length)return;if(s.type==='area')svg.push(`<polygon points="${[...segment,...base.reverse()].map(p=>p.join(',')).join(' ')}" fill="${s.color}" fill-opacity=".5"/>`);svg.push(poly(segment,s.color));segment=[];base=[];};points.forEach((p,i)=>{if(!p){flush();return;}segment.push(p);base.push(basePoints[i]);});flush();}
    });endPlot();
  }
  svg.push(...labelSVG);
  if(legend&&!['pie','doughnut'].includes(chart.type)){const cols=Math.max(1,Math.floor(w/130)),items=data.slice(0,cols*legendRows).map((s,i)=>({...s,name:i===cols*legendRows-1&&data.length>i+1?'외 '+(data.length-i)+'개 계열':s.name}));svg.push(legendSVG(items,cols,legendRows,w,h));}
  if([chart.valueAxis,chart.xAxis,chart.secondaryValueAxis].some(a=>a?.logBase)&&data.some(s=>s.values.some(v=>typeof v==='number'&&v<=0)||chart.xAxis?.logBase&&s.xs?.some(v=>typeof v==='number'&&v<=0)))svg.push(text(left,10,'로그 축에서는 0 이하 값을 표시하지 않습니다.'));
  if(data.some(s=>s.truncated))svg.push(text(w-4,10,'미리 보기: 계열당 최대 1,000개','text-anchor="end"'));
  if(data.some(s=>s.usesCachedValues))svg.push(text(left,10,'계산할 수 없는 수식은 파일에 저장된 값으로 표시합니다.'));
  if(chart.xTitle)svg.push(text((left+right)/2,h-legendRows*17-20,chart.xTitle,'text-anchor="middle"'));
  if(chart.yTitle)svg.push(text(10,(top+bottom)/2,chart.yTitle,`text-anchor="middle" transform="rotate(-90 10 ${(top+bottom)/2})"`));
  return {viewBox:`0 0 ${w} ${h}`,svg:svg.join('')};
}
export {esc as escapeChartHTML};
