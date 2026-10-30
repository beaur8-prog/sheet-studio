import {xml,child,attr} from './chart-xml.mjs';
import {createColorResolver} from './spreadsheet-colors.mjs';

const esc=value=>String(value??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
const truth=value=>value==='1'||value==='true';
export const HEADER_FIELDS=['oddHeader','oddFooter','evenHeader','evenFooter','firstHeader','firstFooter'];
export const HEADER_PLACEHOLDERS={P:'페이지',N:'전체 페이지',D:'날짜',T:'시간',F:'파일 이름',A:'시트 이름',Z:'경로',G:'그림'};
const aliases={Page:'P',Pages:'N',Date:'D',Time:'T',File:'F',Tab:'A',Path:'Z',Picture:'G'};

// Tokenize before splitting sections: an escaped &&L is literal text, and a
// quoted font name can itself contain an ampersand or a section letter.
function tokens(source){
  const result=[];source=String(source||'').replace(/_x([0-9a-f]{4})_/gi,(_match,hex)=>String.fromCharCode(parseInt(hex,16)));
  for(let i=0;i<source.length;){
    if(source[i]!=='&'){let end=source.indexOf('&',i);if(end<0)end=source.length;result.push({text:source.slice(i,end)});i=end;continue;}
    const rest=source.slice(i),match=/^&(?:&|"[^"]*"|\[[^\]]+\]|K(?:\d{2}[+-]\d{3}|[a-f\d]{6})|\d{1,3}|[PN][+-]\d+|[A-Z])/i.exec(rest);
    if(!match){result.push({text:'&'});i++;continue;}
    const raw=match[0],body=raw.slice(1),alias=body.startsWith('[')?aliases[body.slice(1,-1)]:null;
    result.push(body==='&'?{text:'&',raw}:{code:alias||body[0].toUpperCase(),argument:alias?'':body.slice(1),raw});i+=raw.length;
  }
  return result;
}
export function splitHeaderSections(source){
  const parts={left:'',center:'',right:''};let section='center';
  for(const token of tokens(source)){
    if(['L','C','R'].includes(token.code)&&!token.argument)section={L:'left',C:'center',R:'right'}[token.code];
    else parts[section]+=token.raw??token.text;
  }
  return parts;
}
export function headerEditorText(source){
  return tokens(source).map(token=>token.text??(HEADER_PLACEHOLDERS[token.code]?'{'+HEADER_PLACEHOLDERS[token.code]+'}'+(token.argument||''):'')).join('');
}
export function headerTextFromEditor(text){
  return String(text).replace(/&/g,'&&').replace(/\{([^}]+)\}/g,(whole,name)=>{const code=Object.keys(HEADER_PLACEHOLDERS).find(key=>HEADER_PLACEHOLDERS[key]===name);return code?'&'+code:whole;});
}
export function readHeaderFooter(sheet){
  const source=sheet.xlsxFeatures?.headerFooter,root=source?xml(source).documentElement:sheet.kind==='chart'&&sheet.xlsxChartSheet?.xml?child(xml(sheet.xlsxChartSheet.xml).documentElement,'headerFooter'):null;
  return {...Object.fromEntries(HEADER_FIELDS.map(name=>[name,child(root,name)?.textContent??''])),present:HEADER_FIELDS.filter(name=>!!child(root,name)),differentFirst:truth(attr(root,'differentFirst')),differentOddEven:truth(attr(root,'differentOddEven')),scaleWithDoc:!['0','false'].includes(attr(root,'scaleWithDoc')),alignWithMargins:!['0','false'].includes(attr(root,'alignWithMargins'))};
}
export function headerVariant(settings,index,pageNumber,kind){
  const prefix=settings.differentFirst&&index===0?'first':settings.differentOddEven&&pageNumber%2===0?'even':'odd';
  const field=prefix+kind,exists=settings.present?settings.present.includes(field):Object.hasOwn(settings,field);return prefix==='even'&&!exists?settings['odd'+kind]||'':settings[field]||'';
}
function renderSection(source,context,notes){
  const resolve=context.resolveColor||createColorResolver(context.book.chartTheme||''),scale=context.scale||1;
  const style={fontFamily:'Arial',fontSize:11,bold:false,italic:false,underline:false,doubleUnderline:false,strike:false,super:false,sub:false,color:'#222222'};let text='',html='';
  const flush=()=>{if(!text)return;const decorations=[style.underline||style.doubleUnderline?'underline':'',style.strike?'line-through':''].filter(Boolean).join(' '),font=style.fontFamily.replace(/[\\'\r\n]/g,'');
    const css=`font-family:'${font}';font-size:${style.fontSize*scale}pt;color:${style.color};font-weight:${style.bold?'bold':'normal'};font-style:${style.italic?'italic':'normal'};text-decoration:${decorations||'none'};${style.doubleUnderline?'text-decoration-style:double;':''}vertical-align:${style.super?'super':style.sub?'sub':'baseline'};`;
    html+='<span style="'+esc(css)+'">'+esc(text)+'</span>';text='';};
  for(const token of tokens(source)){
    if(token.text!==undefined){text+=token.text;continue;}
    const code=token.code,arg=token.argument;
    if(['P','N','D','T','F','A','Z','G'].includes(code)){
      const offset=/^[+-]\d+$/.test(arg)?Number(arg):0;
      text+=code==='P'?context.pageNumber+offset:code==='N'?context.totalPages+offset:code==='D'?context.now.toLocaleDateString('ko-KR'):code==='T'?context.now.toLocaleTimeString('ko-KR',{hour:'2-digit',minute:'2-digit'}):code==='F'?context.book.nativeSource?.filename||context.book.name:code==='A'?context.sheet.name:code==='Z'?context.filePath||'':'';
      if(code==='G')notes.add('머리글·바닥글의 그림은 인쇄 미리 보기를 지원하지 않습니다. 원본은 XLSX에 유지됩니다.');
      if(code==='Z'&&!context.filePath)notes.add('파일 경로는 앱에서 보관하지 않아 머리글·바닥글에 표시하지 않습니다.');
      continue;
    }
    flush();
    if(code==='"'){
      const [font,...rest]=arg.replace(/"$/,'').split(',');if(font&&!['-','+'].includes(font))style.fontFamily=font;const name=rest.join(',');if(name){style.bold=/bold/i.test(name);style.italic=/italic/i.test(name);}
    }else if(/^\d$/.test(code)){style.fontSize=Math.max(1,Math.min(409,Number(code+arg)));}
    else if(code==='K'){
      const theme=/^(\d{2})([+-])(\d{3})$/.exec(arg);if(theme)style.color=resolve({theme:Math.max(0,Math.min(11,Number(theme[1])-1)),tint:Math.max(-1,Math.min(1,Number(theme[3])/100*(theme[2]==='-'?-1:1)))})||'#222222';
      else if(/^[a-f\d]{6}$/i.test(arg))style.color='#'+arg;
    }else if(['B','I','U','E','S','X','Y'].includes(code)){
      const property={B:'bold',I:'italic',U:'underline',E:'doubleUnderline',S:'strike',X:'super',Y:'sub'}[code];if(code==='X'&&style.sub||code==='Y'&&style.super)continue;style[property]=!style[property];
    }else if(!['L','C','R'].includes(code))notes.add('일부 머리글·바닥글 효과는 인쇄에서 생략됩니다. 원본 코드는 XLSX에 유지됩니다.');
  }
  flush();return html;
}
export function renderPageHeaders(settings,context){
  const notes=new Set(),result={header:'',footer:'',notes};
  for(const [kind,key] of [['Header','header'],['Footer','footer']]){
    const source=headerVariant(settings,context.index,context.pageNumber,kind);if(!source)continue;
    const sections=splitHeaderSections(source);result[key]='<div class="print-'+key+'">'+['left','center','right'].map(side=>'<div class="print-header-cell '+side+'">'+renderSection(sections[side],context,notes)+'</div>').join('')+'</div>';
  }
  return result;
}
