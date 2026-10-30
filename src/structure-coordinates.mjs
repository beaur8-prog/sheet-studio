export function structureIndex(value,index,delta){
  if(value<index)return value;
  if(delta<0&&value<index-delta)return null;
  return value+delta;
}
export function shiftIndexMap(values,index,delta,limit){
  const result={};for(const [key,value] of Object.entries(values||{})){const next=structureIndex(+key,index,delta);if(next!==null&&next>=0&&next<limit)result[next]=value;}return result;
}
