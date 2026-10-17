'use strict';

function applicationMenu(platform, send) {
  if (platform !== 'darwin') return null;
  const item = (id, label, accelerator) => ({ id, label, accelerator, click: () => send(id) });
  return [
    { role: 'appMenu' },
    { label: '파일', submenu: [
      item('new', '새 통합 문서', 'CmdOrCtrl+N'),
      item('open', '열기…', 'CmdOrCtrl+O'),
      item('save', '저장', 'CmdOrCtrl+S'),
      { type: 'separator' }, { role: 'close' },
    ] },
    { label: '편집', submenu: [
      item('undo', '실행 취소', 'CmdOrCtrl+Z'),
      item('redo', '다시 실행', 'CmdOrCtrl+Shift+Z'),
      { type: 'separator' },
      item('cut', '잘라내기', 'CmdOrCtrl+X'),
      item('copy', '복사', 'CmdOrCtrl+C'),
      item('paste', '붙여넣기', 'CmdOrCtrl+V'),
      item('select-all', '모두 선택', 'CmdOrCtrl+A'),
      item('find', '찾기 및 바꾸기', 'CmdOrCtrl+F'),
    ] },
    { label: '보기', submenu: [{ role: 'togglefullscreen' }] },
    { role: 'windowMenu' },
  ];
}
module.exports = { applicationMenu };
