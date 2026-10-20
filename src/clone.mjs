const originals = new WeakMap();
const trusted = new WeakSet();
// Only use for data cloned by the document worker, which cannot contain proxies.
export function trustData(value) { if (value && typeof value === 'object') trusted.add(value); return value; }
export const unwrap = value => value && typeof value === 'object' ? originals.get(value) || value : value;
export function registerProxy(proxy, original) { originals.set(proxy, original); return proxy; }

// Values assembled with object spreads can contain a tracked child object.
export function plain(value, seen = new WeakSet()) {
  const original = unwrap(value);
  if (original !== value || !value || typeof value !== 'object' || trusted.has(value) || seen.has(value)) return original;
  seen.add(value);
  for (const key of Object.keys(value)) {
    const child = plain(value[key], seen);
    if (child !== value[key]) value[key] = child;
  }
  return value;
}
export const cloneData = value => structuredClone(plain(unwrap(value)));
