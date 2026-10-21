// Associate location IDs with document identity, including undo/redo replacements.
// Only the ID is persisted here; actual paths are owned by the desktop process.
export function createDocumentLocation(getBook, storage) {
  const ids = new WeakMap(), key = 'sheet-studio.document-location';
  let recovered;
  try { recovered = storage.getItem(key); } catch {}
  if (typeof recovered === 'string' && /^[0-9a-f]{8}(?:-[0-9a-f]{4}){3}-[0-9a-f]{12}$/i.test(recovered)) ids.set(getBook(), recovered);
  function id() {
    const book = getBook();
    if (!ids.has(book)) ids.set(book, crypto.randomUUID());
    const value = ids.get(book);
    if (value !== recovered) {
      try { storage.setItem(key, value); recovered = value; } catch {}
    }
    return value;
  }
  return { id, carry: previous => ids.set(getBook(), previous) };
}
