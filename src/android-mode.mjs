// Runtime selection is independent of window width. Desktop Electron always
// wins, including a resized window or an Android user-agent used in a test.
export function isAndroidRuntime({desktop, capacitor, navigator: nav = {}} = {}) {
  if (desktop) return false;
  const native = capacitor?.getPlatform?.();
  if (native && native !== 'web') return native === 'android';
  return nav.userAgentData?.platform === 'Android' || /\bAndroid\b/i.test(nav.userAgent || '');
}

const viewCommands = new Set([
  'open', 'new', 'recover', 'copy', 'find', 'select-all', 'sheet-prev', 'sheet-next',
  'theme', 'toggle-grid', 'show-formulas', 'zoom-in', 'zoom-out', 'zoom-reset',
  'freeze-menu', 'freeze-selection', 'freeze-first-row', 'freeze-first-col', 'freeze-none',
  'compatibility', 'about', 'shortcuts', 'cancel-edit', 'print',
]);
const unavailable = new Set(['vba', 'save-xlsm', 'page-setup', 'updates-check', 'updates-manual', 'updates-status']);
export function androidCommandAllowed(name, mode) {
  if (unavailable.has(name)) return false;
  return mode === 'edit' || viewCommands.has(name);
}

export function createAndroidSession() {
  let mode = 'view', activeSheetId = null, epoch = 0, edited = false;
  const views = new Map();
  return {
    get mode() { return mode; }, get activeSheetId() { return activeSheetId; },
    get epoch() { return epoch; }, get edited() { return edited; },
    setMode(next) { if (!['view', 'edit'].includes(next)) throw Error('잘못된 편집 모드입니다.'); if (mode !== next) { mode = next; epoch++; } },
    markEdited() { edited = true; },
    open(id, {editing = false, restored = false} = {}) { views.clear(); activeSheetId = id; mode = editing ? 'edit' : 'view'; edited = restored; epoch++; },
    activate(id, view) { if (activeSheetId && view) views.set(activeSheetId, structuredClone(view)); activeSheetId = id; return structuredClone(views.get(id) || {selection:{r:0,c:0,r2:0,c2:0},left:0,top:0,zoom:100}); },
    getFreeze(id) { return views.get(id)?.freeze; },
    setFreeze(id, freeze) { views.set(id, {...views.get(id), freeze}); },
    ticket(documentId, revision) { return {documentId, revision, epoch}; },
    current(ticket, documentId, revision) { return ticket.documentId === documentId && ticket.revision === revision && ticket.epoch === epoch; },
  };
}
