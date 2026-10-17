import {createWorkbook,address} from '../../src/engine.mjs';
export const rows=[['지역','매출','상품'],['서울',10,'사과'],['부산',20,'배'],['서울',30,'바나나'],['대전',40,'사과'],['부산',50,'수박'],['서울',60,'A*B'],['서울',null,''],['부산',80,'배'],[null,90,'포도']];
const custom=(kind,operator,value,second,join='and')=>({type:'custom',kind,join,conditions:[{operator,value},...(second?[second]:[])]});
export const cases=[
  {name:'arrows only',filters:[],columns:{}},
  {name:'one value',filters:[{field:1,criteria1:'서울'}],columns:{1:{type:'values',values:['서울']}}},
  {name:'multiple values',filters:[{field:1,criteria1:['서울','부산'],operator:7}],columns:{1:{type:'values',values:['서울','부산']}}},
  {name:'blank',filters:[{field:1,criteria1:'='}],columns:{1:{type:'values',values:['']}}},
  {name:'not blank',filters:[{field:1,criteria1:'<>'}],columns:{1:custom('text','notEqual','')}},
  {name:'number greater',filters:[{field:2,criteria1:'>30'}],columns:{2:custom('number','greaterThan','30')}},
  {name:'between inclusive',filters:[{field:2,criteria1:'>=20',operator:1,criteria2:'<=60'}],columns:{2:custom('number','greaterThanOrEqual','20',{operator:'lessThanOrEqual',value:'60'})}},
  {name:'or conditions',filters:[{field:2,criteria1:'<20',operator:2,criteria2:'>70'}],columns:{2:custom('number','lessThan','20',{operator:'greaterThan',value:'70'},'or')}},
  {name:'combined columns',filters:[{field:1,criteria1:'서울'},{field:2,criteria1:'>20'}],columns:{1:{type:'values',values:['서울']},2:custom('number','greaterThan','20')}},
  {name:'contains',filters:[{field:3,criteria1:'=*과*'}],columns:{3:custom('text','contains','과')}},
  {name:'not contains',filters:[{field:3,criteria1:'<>*과*'}],columns:{3:custom('text','notContains','과')}},
  {name:'wildcard escaped',filters:[{field:3,criteria1:'=A~*B'}],columns:{3:custom('text','equal','A~*B')}},
  {name:'top items',filters:[{field:2,criteria1:'3',operator:3}],columns:{2:{type:'top',top:true,percent:false,count:3}}},
  {name:'bottom percent',filters:[{field:2,criteria1:'30',operator:6}],columns:{2:{type:'top',top:false,percent:true,count:30}}},
  {name:'above average',filters:[{field:2,criteria1:33,operator:11}],columns:{2:{type:'average',above:true}}},
  {name:'no matches',filters:[{field:1,criteria1:'인천'}],columns:{1:{type:'values',values:['인천']}}}
];
export function fixture(){
  const book=createWorkbook(),sheet=book.sheets[0];sheet.name='Sheet1';sheet.cells={};sheet.charts=[];sheet.merges=[];
  rows.forEach((row,r)=>row.forEach((v,c)=>{if(v!==null)sheet.cells[address(r+2,c+1)]={raw:String(v)};}));
  sheet.cells.A1={raw:'범위 밖 제목'};sheet.cells.D15={raw:'=SUBTOTAL(9,C4:C12)'};sheet.cells.D16={raw:'=SUBTOTAL(109,C4:C12)'};
  sheet.autoFilter={range:{r1:2,r2:11,c1:1,c2:3},columns:{},hiddenRows:[]};return book;
}
