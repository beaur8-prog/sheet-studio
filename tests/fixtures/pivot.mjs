import {address} from '../../src/engine.mjs';
export function fixture(){
  const rows=[['지역','상품','매출','수량'],['동부','A',10,1],['동부','A',20,2],['동부','B',30,3],['서부','A',100,4],['서부','B','',5]];
  const source={id:'source',name:'원본',cells:{},merges:[],charts:[],columnWidths:{},rowHeights:{}};
  rows.forEach((row,r)=>row.forEach((v,c)=>source.cells[address(r,c)]={raw:String(v),style:{}}));
  return {name:'피벗 테스트',activeSheetId:'source',sheets:[source]};
}
export const spec=()=>({version:1,sourceSheetId:'source',sourceRange:'A1:D6',rows:['지역'],columns:['상품'],values:[{field:'매출',aggregate:'sum'}],filters:[],totals:true,sort:'asc'});
