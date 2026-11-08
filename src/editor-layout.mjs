export function androidEditorLayout({preference='auto',width,height,finePointer=false,keyboard=false}){
  if(width<840||height<600||preference==='touch')return 'touch';
  return preference==='desktop'||finePointer||keyboard?'desktop':'touch';
}
