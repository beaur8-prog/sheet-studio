'use strict';

function updateMenu(state, open, setAutomatic) {
  const items = [
    { id: 'sheet-updates-check', label: '업데이트 확인…', click: () => open('check') },
    { id: 'sheet-updates-manual', label: '수동 업데이트…', click: () => open('manual') },
    { id: 'sheet-updates-status', label: '업데이트 진행 상황…', click: () => open('status') },
  ];
  if (state.supportsAutomatic) items.push({ type: 'separator' }, {
    label: '업데이트 자동 확인·다운로드', type: 'checkbox', checked: state.automatic,
    click: item => { void setAutomatic(item.checked); },
  });
  return items;
}
module.exports = { updateMenu };
