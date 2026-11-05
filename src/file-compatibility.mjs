import {xml,all,attr} from './chart-xml.mjs';
const RULES=Object.freeze({
  ole:{label:'첨부·삽입 개체',description:'삽입된 문서·파일(OLE)을 XLSX 사본에 유지하지 못합니다.',path:/^xl\/embeddings\//i,type:/oleObject/i},
  activex:{label:'ActiveX·양식 컨트롤',description:'컨트롤 개체와 설정은 XLSX 사본에서 제거될 수 있습니다.',path:/^xl\/(?:activeX|ctrlProps)\//i,type:/activeX|controlproperties/i},
  threaded:{label:'대화형 댓글',description:'댓글 대화와 작성자 정보는 현재 지원하지 않습니다. 일반 셀 메모와 다른 기능입니다.',path:/^xl\/(?:threadedComments|persons)\//i,type:/threadedcomments|person\+xml/i},
  model:{label:'데이터 모델',description:'Power Pivot·Data Model을 XLSX 사본에 유지하지 못합니다.',path:/^xl\/model\//i,type:/model\+data/i},
  slicer:{label:'슬라이서·타임라인',description:'연결된 필터 UI와 캐시를 XLSX 사본에 유지하지 못합니다.',path:/^xl\/(?:slicers|slicerCaches|timelines|timelineCaches)\//i,type:/slicer|timeline/i},
  web:{label:'웹 추가 기능',description:'통합 문서에 포함된 웹 추가 기능은 실행하거나 저장하지 않습니다.',path:/^(?:xl\/)?webextensions\//i,type:/webextension/i},
  custom:{label:'다른 프로그램의 추가 문서 정보',description:'앱이 해석하지 못하는 Custom XML 정보는 XLSX 사본에 유지하지 못합니다.',path:/^customXml\/(?!_rels\/|.*Props).*\.xml$/i,type:null},
  macro:{label:'VBA 프로젝트',description:'XLSX 이름의 파일에 VBA 프로젝트가 포함되어 있습니다. 매크로를 유지하려면 원본을 Excel에서 사용하세요.',path:/^xl\/vbaProject\.bin$/i,type:/vbaProject$/i},
});

export async function readCompatibilityRisks(zip,filename){
  const types=xml(await zip.file('[Content_Types].xml').async('string')),content=new Map(all(types,'Override').map(n=>[attr(n,'PartName').replace(/^\//,''),attr(n,'ContentType')])),risks=[];
  for(const [kind,rule] of Object.entries(RULES)){
    if(kind==='macro'&&/\.xlsm$/i.test(filename))continue;let count=0;
    for(const [path,file] of Object.entries(zip.files)){
      if(file.dir||!rule.path.test(path)&&!rule.type?.test(content.get(path)||''))continue;
      if(kind==='custom'){
        // Only the app's own namespace is reconstructed by addPivotParts.
        const text=await file.async('string');if(text.includes('urn:sheet-studio:pivots:1'))continue;
      }
      count++;
    }
    if(count)risks.push({kind,count});
  }
  return risks.length?{version:1,sourceName:String(filename).split(/[\\/]/).at(-1).slice(0,255),risks}:undefined;
}

export function validateCompatibilityReview(value){
  if(!value||value.version!==1||typeof value.sourceName!=='string'||!value.sourceName||value.sourceName.length>255||/[\\/\u0000-\u001f]/.test(value.sourceName)||!Array.isArray(value.risks)||value.risks.length>Object.keys(RULES).length)throw Error('파일 호환성 정보가 올바르지 않습니다.');
  const kinds=new Set();for(const risk of value.risks)if(!risk||!Object.hasOwn(RULES,risk.kind)||kinds.has(risk.kind)||!Number.isInteger(risk.count)||risk.count<1||risk.count>1000000)throw Error('파일 호환성 항목이 올바르지 않습니다.');else kinds.add(risk.kind);
}
export function compatibilityRisks(book){return (book.xlsxCompatibility?.risks||[]).map(item=>({...item,...RULES[item.kind]}));}
export function assertCompatibleExport(book,kind,options={}){
  if(kind==='xlsx'&&book.xlsxCompatibility?.risks.length&&options.allowFeatureLoss!==true)throw Error('이 문서에는 XLSX에 유지하지 못하는 기능이 있습니다. 파일 호환성에서 확인한 뒤 기능 제외 사본을 저장하세요.');
}

export function compatibilitySummary(book){
  const count=fn=>book.sheets.reduce((n,s)=>n+fn(s),0),rows=[];
  rows.push({label:'셀·수식',mode:'편집 가능',detail:'지원하는 수식은 앱에서 계산합니다. 미지원 수식의 원본 저장값은 계산 결과와 구분합니다.'});
  const tables=count(s=>s.tables?.length||0);if(tables)rows.push({label:`Excel 표 ${tables}개`,mode:book.nativeSource||book.xlsxPivotArchive?'편집 제한':'편집 가능',detail:'표 범위·계산 열·요약·필터와 지원하는 구조화 참조를 유지합니다.'});
  const charts=count(s=>s.charts?.length||0);if(charts)rows.push({label:`차트 ${charts}개`,mode:'유형별 지원',detail:'지원 유형은 편집하며 다른 유형은 원본 보존과 미리 보기의 범위가 다릅니다.'});
  const pivots=count(s=>s.pivot?1:0);if(pivots)rows.push({label:`앱 피벗 ${pivots}개`,mode:'편집 가능',detail:'필드 편집과 새로 고침을 지원합니다.'});
  const native=book.xlsxPivotArchive?.sheets.length||0;if(native)rows.push({label:`원본 피벗 ${native}개`,mode:book.xlsxPivotArchive.blockReason?'저장 제한':'원본 보존',detail:book.xlsxPivotArchive.blockReason||'저장된 결과와 기본 피벗 정의·캐시를 유지합니다. 새로 고침은 Excel에서 진행하세요.'});
  if(book.xlsxExternalLinks)rows.push({label:'외부 통합 문서 연결',mode:'원본 보존',detail:'연결 정의와 저장값을 보관하며 외부 파일을 열거나 갱신하지 않습니다.'});
  if(book.nativeSource)rows.push({label:'VBA·XLSM',mode:'Excel 연결',detail:'VBA 실행과 매크로 보존 저장에는 Windows 데스크톱 Excel이 필요합니다.'});
  if(count(s=>s.xlsxFeatures?.dataValidations?1:0))rows.push({label:'데이터 유효성 검사',mode:'규칙별 지원',detail:'지원하는 입력 제한·목록을 적용하고 원본 규칙을 보존합니다.'});
  if(count(s=>s.unsupportedConditionalCount||0))rows.push({label:'고급 조건부 서식',mode:'원본 보존',detail:'일부 색조·막대·아이콘 규칙은 화면 편집을 지원하지 않으며 원본 XML을 유지합니다.'});
  return rows;
}
