'use strict';
const fs=require('node:fs/promises'),path=require('node:path'),os=require('node:os');
const {randomUUID,createHash}=require('node:crypto');
const {spawn,execFile}=require('node:child_process');
const {promisify}=require('node:util');
const exec=promisify(execFile);
const MAX_BYTES=15*1024*1024;
const ps=path.join(process.env.SystemRoot||'C:\\Windows','System32/WindowsPowerShell/v1.0/powershell.exe');
const origins=new Map(),sessions=new Map();
let pendingStart=null,closing=false;
const digest=data=>createHash('sha256').update(data).digest('hex');
const unknownZone='[ZoneTransfer]\r\nZoneId=3\r\n';
async function waitForExit(session,timeout=15000){if(session.done)return;await Promise.race([new Promise(resolve=>session.child.once('exit',resolve)),new Promise(resolve=>{const timer=setTimeout(resolve,timeout);timer.unref();})]);}
function registerOrigin(data,zone){const token=randomUUID();origins.set(token,{hash:digest(data),zone});return token;}
async function rememberFile(filename,data){if(process.platform!=='win32')return registerOrigin(data,null);let zone=null;try{zone=await fs.readFile(filename+':Zone.Identifier','utf8');}catch(error){if(!['ENOENT','EINVAL','ENOTSUP'].includes(error.code))zone=unknownZone;}return registerOrigin(data,zone);}
async function markExport(filename,data,token){if(process.platform!=='win32')return;const origin=origins.get(token),zone=origin?.hash===digest(data)?origin.zone:unknownZone;if(zone)await fs.writeFile(filename+':Zone.Identifier',zone);}
const nameOK=value=>typeof value==='string'&&value.length>0&&value.length<=31&&!/[\\/*?:\[\]]/.test(value);
const addrOK=value=>typeof value==='string'&&/^[A-Z]{1,3}[1-9]\d{0,6}$/.test(value)&&Number(value.match(/\d+/)[0])<=1048576&&value.match(/[A-Z]+/)[0].split('').reduce((n,c)=>n*26+c.charCodeAt(0)-64,0)<=16384;
const rangeOK=value=>typeof value==='string'&&value.split(':').length===2&&value.split(':').every(addrOK);
function validatePayload(payload){
  if(!payload||!['interactive','save'].includes(payload.mode))throw Error('잘못된 Excel 작업입니다.');
  const source=payload.source;
  if(!source||typeof source.base64!=='string'||source.base64.length>MAX_BYTES*4/3+4||!/^[A-Za-z0-9+/]*={0,2}$/.test(source.base64))throw Error('XLSM 원본 데이터가 올바르지 않습니다.');
  const data=Buffer.from(source.base64,'base64');if(data.length>MAX_BYTES||data.length<4||data[0]!==80||data[1]!==75)throw Error('XLSM ZIP 파일이 아닙니다.');
  const p=payload.plan;if(!p||!Array.isArray(p.sheets)||!p.sheets.length||p.sheets.length>100||!Array.isArray(p.deleted)||!p.deleted.every(nameOK)||!Array.isArray(p.operations)||p.operations.length>10000)throw Error('변경 계획이 올바르지 않습니다.');
  if(p.pivots!==undefined&&(!Array.isArray(p.pivots)||p.pivots.length>100||JSON.stringify(p.pivots).length>2000000||p.pivots.some(v=>!v||!nameOK(v.sheet)||(v.source!==null&&!nameOK(v.source))||!v.spec||v.spec.version!==1)))throw Error('피벗 저장 설정이 올바르지 않습니다.');
  const ids=new Set(),names=new Set();let cells=0;
  for(const s of p.sheets){
    if(typeof s.id!=='string'||s.id.length>150||ids.has(s.id)||!nameOK(s.name)||names.has(s.name.toLowerCase())||(s.sourceName!==null&&!nameOK(s.sourceName))||!Array.isArray(s.cells))throw Error('시트 변경 계획이 올바르지 않습니다.');ids.add(s.id);names.add(s.name.toLowerCase());
    for(const c of s.cells){if(++cells>200000||!addrOK(c.address)||(c.raw!==undefined&&(typeof c.raw!=='string'||c.raw.length>32768))||(c.note!==undefined&&(typeof c.note!=='string'||c.note.length>32768)))throw Error('셀 변경 계획이 올바르지 않습니다.');if(c.style!==undefined){if(!c.style||typeof c.style!=='object'||Array.isArray(c.style))throw Error('서식이 올바르지 않습니다.');for(const [key,value] of Object.entries(c.style)){if(key==='borders'){
      const sides=['top','right','bottom','left'],styles=['thin','medium','thick','dotted','dashed','double','hair','mediumDashed','dashDot','mediumDashDot','dashDotDot','mediumDashDotDot','slantDashDot'];
      if(value!==null&&(!value||typeof value!=='object'||Array.isArray(value)||Object.entries(value).some(([side,edge])=>!sides.includes(side)||edge!==null&&(!edge||typeof edge!=='object'||Array.isArray(edge)||Object.keys(edge).some(k=>!['style','color'].includes(k))||!styles.includes(edge.style)||!/^#[\da-f]{6}$/i.test(edge.color)))))throw Error('테두리 설정이 올바르지 않습니다.');continue;
    }if(!['bold','italic','underline','wrap','border','fontFamily','fontSize','color','background','align','verticalAlign','format','numberFormat','decimals'].includes(key)||!['string','number','boolean'].includes(typeof value)&&value!==null||typeof value==='string'&&value.length>512||typeof value==='number'&&!Number.isFinite(value))throw Error('서식 속성이 올바르지 않습니다.');}}}
    if(s.manualHiddenRows!==undefined&&(!Array.isArray(s.manualHiddenRows)||s.manualHiddenRows.length>200000||s.manualHiddenRows.some(r=>!Number.isInteger(r)||r<0||r>=1048576)))throw Error('숨긴 행 목록이 올바르지 않습니다.');
    if(s.freeze!==undefined&&(!s.freeze||!Number.isInteger(s.freeze.rows)||!Number.isInteger(s.freeze.cols)||s.freeze.rows<0||s.freeze.rows>=1048576||s.freeze.cols<0||s.freeze.cols>=16384))throw Error('틀 고정 계획이 올바르지 않습니다.');
    if(s.conditionalFormats!==undefined){
      if(!Array.isArray(s.conditionalFormats)||s.conditionalFormats.length>200)throw Error('조건부 서식 계획이 올바르지 않습니다.');
      for(const r of s.conditionalFormats){
        if(!r||!rangeOK(r.range)||![1,2].includes(r.type)||r.type===1&&![1,2,3,4,5,6,7,8].includes(r.operator)||typeof r.formula1!=='string'||r.formula1.length>8192||r.formula2!==null&&(typeof r.formula2!=='string'||r.formula2.length>8192)||typeof r.stopIfTrue!=='boolean'||!r.style||typeof r.style!=='object'||Array.isArray(r.style))throw Error('조건부 서식 조건이 올바르지 않습니다.');
        for(const [key,value] of Object.entries(r.style))if(!(['background','color'].includes(key)&&typeof value==='string'&&/^#[\da-f]{6}$/i.test(value)||['bold','italic'].includes(key)&&typeof value==='boolean'))throw Error('조건부 서식 색상이 올바르지 않습니다.');
      }
    }
    if(s.autoFilter!==undefined&&s.autoFilter!==null){
      const f=s.autoFilter;if(!f||!rangeOK(f.range)||!Array.isArray(f.filters)||f.filters.length>16384||!Array.isArray(f.hiddenRows)||f.hiddenRows.length>200000)throw Error('자동 필터 계획이 올바르지 않습니다.');
      const [start,end]=f.range.split(':'),r1=Number(start.match(/\d+/)[0])-1,r2=Number(end.match(/\d+/)[0])-1;
      const col=a=>a.match(/[A-Z]+/)[0].split('').reduce((n,c)=>n*26+c.charCodeAt(0)-64,0),width=col(end)-col(start)+1;
      if(r2-r1>200000||r2<r1||width<1||f.hiddenRows.some(r=>!Number.isInteger(r)||r<=r1||r>r2))throw Error('자동 필터 범위가 올바르지 않습니다.');
      const criterion=v=>typeof v==='string'&&v.length<=32768||typeof v==='number'&&Number.isFinite(v);
      for(const rule of f.filters)if(!rule||!Number.isInteger(rule.field)||rule.field<1||rule.field>width||rule.operator!==undefined&&![1,2,3,4,5,6,7,11].includes(rule.operator)||!(criterion(rule.criteria1)||Array.isArray(rule.criteria1)&&(rule.criteria1.length>0||Array.isArray(rule.dateValues)&&rule.dateValues.length>0)&&rule.criteria1.length<=200000&&rule.criteria1.every(criterion))||rule.dateValues!==undefined&&(!Array.isArray(rule.dateValues)||rule.dateValues.length>200000||rule.dateValues.some(v=>!Number.isInteger(v)||v<0||v>2958465))||rule.criteria2!==undefined&&!criterion(rule.criteria2))throw Error('자동 필터 조건이 올바르지 않습니다.');
    }
    for(const prop of ['columnWidths','rowHeights']){if(!s[prop]||typeof s[prop]!=='object'||Array.isArray(s[prop]))throw Error('크기가 올바르지 않습니다.');for(const [key,value] of Object.entries(s[prop]))if(!/^\d+$/.test(key)||+key>=(prop==='columnWidths'?16384:1048576)||value!==null&&(!Number.isFinite(value)||value<0||value>10000))throw Error('크기가 올바르지 않습니다.');}
    for(const prop of ['merge','unmerge'])if(!Array.isArray(s[prop])||s[prop].length>10000||!s[prop].every(rangeOK))throw Error('병합 범위가 올바르지 않습니다.');
    if(s.deletedCharts!==undefined&&(!Array.isArray(s.deletedCharts)||s.deletedCharts.length>200||s.deletedCharts.some(n=>typeof n!=='string'||!n||n.length>256)))throw Error('차트 삭제 목록이 올바르지 않습니다.');
    if(s.charts!==undefined){
      if(!Array.isArray(s.charts)||s.charts.length>200)throw Error('차트 변경 계획이 올바르지 않습니다.');
      for(const c of s.charts){
        if(!c||typeof c!=='object'||c.name!==null&&(typeof c.name!=='string'||!c.name||c.name.length>256))throw Error('차트 이름이 올바르지 않습니다.');
        for(const [key,value] of Object.entries(c)){
          if(key==='name')continue;
          if(['x','y','width','height'].includes(key)){if(!Number.isFinite(value)||value<0||value>100000000||['width','height'].includes(key)&&(value<100||value>10000))throw Error('차트 크기가 올바르지 않습니다.');}
          else if(['title','xTitle','yTitle'].includes(key)){if(typeof value!=='string'||value.length>4000)throw Error('차트 제목이 올바르지 않습니다.');}
          else if(['legend','labels','combo','secondaryAxis'].includes(key)){if(typeof value!=='boolean')throw Error('차트 표시 설정이 올바르지 않습니다.');}
          else if(key==='series'){
            if(!Array.isArray(value)||value.length>32||!c.range||c.chartType===undefined)throw Error('차트 계열 계획이 올바르지 않습니다.');
            const size=range=>{if(!rangeOK(range))throw Error('차트 계열 범위가 올바르지 않습니다.');const [a,b]=range.split(':');if(a.match(/[A-Z]+/)[0]!==b.match(/[A-Z]+/)[0])throw Error('차트 계열은 하나의 열이어야 합니다.');const n=Number(b.match(/\d+/)[0])-Number(a.match(/\d+/)[0])+1;if(n<1||n>10000)throw Error('차트 계열 범위가 너무 큽니다.');return n;};
            for(const series of value){
              if(!series||typeof series!=='object'||Array.isArray(series)||Object.keys(series).some(k=>!['valuesRange','nameRange','categoriesRange','sizesRange'].includes(k)))throw Error('차트 계열 계획이 올바르지 않습니다.');
              const count=size(series.valuesRange);
              for(const key of ['categoriesRange','sizesRange'])if(series[key]!==undefined&&size(series[key])!==count)throw Error('차트 계열의 항목 수가 다릅니다.');
              if(series.nameRange!==undefined&&size(series.nameRange)!==1)throw Error('차트 계열 이름은 하나의 셀이어야 합니다.');
            }
          }
          else if(key==='range'){if(!rangeOK(value))throw Error('차트 원본 범위가 올바르지 않습니다.');}
          else if(key==='chartType'){if(![51,52,53,57,58,59,4,63,64,1,76,77,5,-4120,-4169,15,-4151].includes(value))throw Error('지원되지 않는 차트 유형입니다.');}
          else throw Error('허용되지 않는 차트 속성입니다.');
        }
        if(c.name===null&&(!c.range||c.chartType===undefined))throw Error('새 차트의 원본 범위가 필요합니다.');
      }
    }
  }
  for(const op of p.operations)if(typeof op.sheetId!=='string'||!['r','c'].includes(op.axis)||!Number.isInteger(op.index)||op.index<0||op.index>=(op.axis==='r'?1048576:16384)||![-1,1].includes(op.delta))throw Error('행/열 작업이 올바르지 않습니다.');
  if(JSON.stringify(p).length>25*1024*1024)throw Error('변경 계획이 너무 큽니다.');
  return data;
}
async function probe(){if(process.platform!=='win32')return {available:false,reason:'Windows 데스크톱 Excel이 필요합니다.'};try{const {stdout}=await exec(ps,['-NoLogo','-NoProfile','-NonInteractive','-Command',"[bool](Get-Item 'Registry::HKEY_CLASSES_ROOT\\Excel.Application\\CLSID' -ErrorAction SilentlyContinue)"],{windowsHide:true,timeout:10000});return {available:stdout.trim()==='True',reason:'Windows 데스크톱 Excel 설치가 필요합니다.'};}catch{return {available:false,reason:'Excel 등록 정보를 읽을 수 없습니다.'};}}
async function startSession(payload){
  const data=validatePayload(payload);if(!(await probe()).available)throw Error('설치된 Windows 데스크톱 Excel을 찾을 수 없습니다.');
  for(const s of sessions.values())if(!s.done){if(['completed','cancelled','error'].includes((await status(s.id)).state))await waitForExit(s);if(!s.done)throw Error('진행 중인 Excel 작업을 먼저 종료하세요.');}
  const id=randomUUID(),dir=await fs.mkdtemp(path.join(os.tmpdir(),'sheet-studio-vba-'));
  const filename=path.basename(String(payload.source.filename||'Workbook.xlsm')).replace(/[<>:"/\\|?*\u0000-\u001f]/g,'_').replace(/\.xlsm$/i,'').slice(0,120)+'.xlsm';
  const origin=origins.get(payload.source.originToken),zone=origin?.hash===digest(data)?origin.zone:unknownZone;
  await fs.writeFile(path.join(dir,filename),data);if(zone)await fs.writeFile(path.join(dir,filename)+':Zone.Identifier',zone);
  const config={mode:payload.mode,filename,plan:payload.plan,zone};
  await fs.writeFile(path.join(dir,'request.json'),JSON.stringify(config));
  await fs.writeFile(path.join(dir,'worker.ps1'),Buffer.concat([Buffer.from([0xef,0xbb,0xbf]),await fs.readFile(path.join(__dirname,'excel-worker.ps1'))]));
  const child=spawn(ps,['-NoLogo','-NoProfile','-NonInteractive','-STA','-File',path.join(dir,'worker.ps1'),'-JobDir',dir],{windowsHide:true,stdio:['ignore','ignore','pipe']});
  const session={id,dir,child,zone,filename,done:false,stderr:'',last:{state:'starting',message:'Excel 작업 사본을 준비하고 있습니다.'}};sessions.set(id,session);
  child.stderr.on('data',data=>{session.stderr=(session.stderr+data.toString()).slice(-8000);});child.on('error',e=>{session.done=true;session.last={state:'error',message:e.message};});child.on('exit',code=>{session.done=true;session.exitCode=code;});
  return {id};
}
function get(id){const s=sessions.get(id);if(!s)throw Error('Excel 작업을 찾을 수 없습니다.');return s;}
async function status(id){const s=get(id);if(s.cancelled)return {id,state:'cancelled',message:'Excel 작업 사본의 실행을 중단했습니다.'};try{s.last=JSON.parse(await fs.readFile(path.join(s.dir,'status.json'),'utf8'));}catch{}if(s.done&&!['completed','cancelled','error'].includes(s.last.state))s.last={state:'error',message:s.stderr||'Excel 연결이 종료되었습니다.'};return {id,...s.last};}
async function command(id,command){const s=get(id),state=await status(id);if(state.state!=='ready')throw Error('Excel 작업이 준비될 때까지 기다려 주세요.');if(!command||!['run','dialog','finish','cancel'].includes(command.type))throw Error('허용되지 않는 작업입니다.');if(command.type==='run'&&(typeof command.name!=='string'||command.name.length>200||!/^[_\p{L}][_\p{L}\p{N}]*(?:\.[_\p{L}][_\p{L}\p{N}]*)?$/u.test(command.name)))throw Error('매크로 이름은 Module1.MacroName 형식으로 입력하세요.');await fs.writeFile(path.join(s.dir,'command.json'),JSON.stringify({type:command.type,name:command.name||''}),{flag:'wx'});return {accepted:true};}
async function result(id){const s=get(id);if((await status(id)).state!=='completed')throw Error('Excel 작업이 완료되지 않았습니다.');const data=await fs.readFile(path.join(s.dir,'result.xlsm'));if(data.length>MAX_BYTES)throw Error('매크로 실행 결과가 15 MB를 초과했습니다. Excel 작업 사본을 확인하세요.');const originToken=registerOrigin(data,s.zone);return {name:s.filename,data:Array.from(data),originToken};}
async function cancel(id){
  const s=get(id);if(s.done)return;s.cancelled=true;
  await fs.writeFile(path.join(s.dir,'cancel.flag'),'cancel');
  const owner=path.join(s.dir,'owner.json'),deadline=Date.now()+10000;
  let hasOwner=false;
  while(!s.done&&Date.now()<deadline){try{await fs.access(owner);hasOwner=true;break;}catch{}await new Promise(resolve=>setTimeout(resolve,100));}
  // A worker still creating COM reads cancel.flag immediately after ownership is
  // recorded. Do not kill it before then, which could strand an Excel instance.
  if(!hasOwner||s.done)return;
  await exec(ps,['-NoLogo','-NoProfile','-NonInteractive','-File',path.join(s.dir,'worker.ps1'),'-JobDir',s.dir,'-CancelOwned'],{windowsHide:true,timeout:15000});
  s.child.kill();s.done=true;
}
async function start(payload){
  if(closing||pendingStart)throw Error('Excel 연결이 시작되거나 종료되는 중입니다.');
  pendingStart=startSession(payload);
  try{return await pendingStart;}finally{pendingStart=null;}
}
async function closeAll(){closing=true;if(pendingStart)await pendingStart.catch(()=>{});for(const s of sessions.values())if(!s.done)await cancel(s.id).catch(()=>{});}
module.exports={probe,start,status,command,result,cancel,closeAll,rememberFile,markExport,validatePayload};
