import fs from 'node:fs/promises';
import {FUNCTION_CATALOG} from '../src/function-catalog.mjs';

const check=process.argv.includes('--check'),files=['README.md','docs/FUNCTIONS.md'],root=new URL('../',import.meta.url);
let stale=false;
for(const file of files){
  const path=new URL(file,root),text=await fs.readFile(path,'utf8'),pattern=/\*\*\d+개 함수 이름\*\*/g,matches=text.match(pattern);
  if(matches?.length!==1)throw Error(file+': 함수 개수 표기를 한 곳에서 찾을 수 없습니다.');
  const next=text.replace(pattern,`**${FUNCTION_CATALOG.length}개 함수 이름**`);
  if(next===text)continue;
  if(check){stale=true;console.error(file+': 함수 개수가 카탈로그와 다릅니다. npm run docs:functions를 실행하세요.');}
  else await fs.writeFile(path,next);
}
if(stale)process.exitCode=1;else console.log(`Function documentation: ${FUNCTION_CATALOG.length} names match the source catalogue.`);
