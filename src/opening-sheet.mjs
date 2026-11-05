// Files and older recovery records can point activeTab at a hidden sheet.
// Display a visible sheet while preserving every sheet and its hidden state.
export function openingSheetId(book){
  const visible=book.sheets.filter(sheet=>!sheet.state||sheet.state==='visible');
  const target=visible.find(sheet=>sheet.id===book.activeSheetId)||visible[0];
  if(!target)throw Error('통합 문서에는 표시되는 시트가 하나 이상 있어야 합니다.');
  return target.id;
}
