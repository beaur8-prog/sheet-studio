export const RECOVERY_DB = 'sheet-studio-recovery';
const BUCKETS = 1024;
export function cellBucket(ref) {
  let hash = 2166136261;
  for (let i = 0; i < ref.length; i++) hash = Math.imul(hash ^ ref.charCodeAt(i), 16777619);
  return (hash >>> 0) % BUCKETS;
}
export const workbookMetadata = book => ({...book, sheets: book.sheets.map(({cells, ...sheet}) => sheet)});
const request = req => new Promise((resolve, reject) => { req.onsuccess = () => resolve(req.result); req.onerror = () => reject(req.error); });
const completed = tx => new Promise((resolve, reject) => { tx.oncomplete = resolve; tx.onabort = () => reject(tx.error || Error('자동 복구 저장이 취소되었습니다.')); tx.onerror = () => {}; });

export async function openRecoveryStore() {
  const opening = indexedDB.open(RECOVERY_DB, 1);
  opening.onupgradeneeded = () => {
    opening.result.createObjectStore('state');
    opening.result.createObjectStore('chunks');
  };
  const db = await request(opening);
  db.onversionchange = () => db.close();
  let saved = null, previous = null, fresh = false;
  async function readState() {
    const tx = db.transaction('state', 'readonly'), done = completed(tx), state = tx.objectStore('state');
    [saved, previous] = await Promise.all([request(state.get('current')), request(state.get('previous'))]);
    await done;
  }
  async function read(snapshot) {
    if (!snapshot) return null;
    const tx = db.transaction('chunks', 'readonly'), done = completed(tx);
    const book = {...snapshot.metadata, sheets: snapshot.metadata.sheets.map(sheet => ({...sheet, cells: {}}))};
    const jobs = book.sheets.map(async sheet => {
      const generation = snapshot.generations[sheet.id];
      if (!generation) throw Error('자동 복구 시트 정보가 올바르지 않습니다.');
      const range = IDBKeyRange.bound([generation, 0], [generation, BUCKETS - 1]);
      await new Promise((resolve, reject) => {
        const cursor = tx.objectStore('chunks').openCursor(range);
        cursor.onerror = () => reject(cursor.error);
        cursor.onsuccess = () => {
          const current = cursor.result;
          if (!current) return resolve();
          Object.assign(sheet.cells, current.value); current.continue();
        };
      });
    });
    await Promise.all(jobs); await done; return book;
  }
  async function load() { await readState(); fresh = false; return read(saved); }
  // A cold launch reads only metadata. Keep the old chunks untouched until the
  // first edit is saved, then retain their snapshot as the previous document.
  async function startBlank(documentId) {
    await readState(); fresh = true;
    if (saved && !saved.documentId && documentId) {
      const next = {...saved, documentId}, tx = db.transaction('state', 'readwrite'), done = completed(tx);
      tx.objectStore('state').put(next, 'current'); await done; saved = next;
    }
    return !!(saved || previous);
  }
  async function recover() {
    const snapshot = fresh ? saved || previous : previous || saved;
    return snapshot ? {book: await read(snapshot), documentId: snapshot.documentId} : null;
  }
  async function save(book, dirty, replace, metadataChanged = true, documentId) {
    const retained = fresh ? saved || previous : previous;
    const base = fresh ? null : saved;
    const publish = !base || metadataChanged || replace.size > 0 || base.documentId !== documentId;
    const metadata = publish ? workbookMetadata(book) : base.metadata, generations = {}, groups = [];
    const obsolete = new Set([...Object.values(saved?.generations || {}), ...Object.values(previous?.generations || {})]);
    for (const generation of Object.values(retained?.generations || {})) obsolete.delete(generation);
    for (const sheet of book.sheets) {
      const full = !base?.generations[sheet.id] || replace.has(sheet.id);
      const generation = full ? crypto.randomUUID() : base.generations[sheet.id];
      generations[sheet.id] = generation; obsolete.delete(generation);
      const chunks = new Map();
      if (full) {
        for (const ref in sheet.cells) {
          const bucket = cellBucket(ref);
          if (!chunks.has(bucket)) chunks.set(bucket, {});
          chunks.get(bucket)[ref] = sheet.cells[ref];
        }
      } else {
        for (const ref of dirty.get(sheet.id) || []) {
          const bucket = cellBucket(ref);
          if (!chunks.has(bucket)) chunks.set(bucket, []);
          chunks.get(bucket).push(ref);
        }
      }
      groups.push({sheet, generation, full, chunks});
    }
    const tx = db.transaction(['state', 'chunks'], 'readwrite'), done = completed(tx), store = tx.objectStore('chunks');
    try {
      const patches = [];
      for (const group of groups) for (const [bucket, value] of group.chunks) {
        const key = [group.generation, bucket];
        if (group.full) store.put(value, key);
        else patches.push(request(store.get(key)).then(chunk => {
          chunk ||= {};
          for (const ref of value) {
            if (Object.hasOwn(group.sheet.cells, ref)) chunk[ref] = group.sheet.cells[ref]; else delete chunk[ref];
          }
          if (Object.keys(chunk).length) store.put(chunk, key); else store.delete(key);
        }));
      }
      await Promise.all(patches);
      for (const generation of obsolete) store.delete(IDBKeyRange.bound([generation, 0], [generation, BUCKETS - 1]));
      const next = {metadata, generations, ...(documentId ? {documentId} : {})};
      if (publish) tx.objectStore('state').put(next, 'current');
      if (fresh && retained) tx.objectStore('state').put(retained, 'previous');
      await done; saved = next; previous = retained; fresh = false;
      return {chunksWritten: groups.reduce((n, group) => n + group.chunks.size, 0), fullSheets: groups.filter(group => group.full).length, metadataWritten: publish};
    } catch (error) {
      try { tx.abort(); } catch {}
      await done.catch(() => {}); throw error;
    }
  }
  return {load, startBlank, recover, save, get documentId() { return saved?.documentId; }, close: () => db.close()};
}
