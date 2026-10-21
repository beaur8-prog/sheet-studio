'use strict';
const fs = require('node:fs/promises');
const path = require('node:path');
const { randomUUID } = require('node:crypto');

const validId = value => typeof value === 'string' && /^[0-9a-f]{8}(?:-[0-9a-f]{4}){3}-[0-9a-f]{12}$/i.test(value);
const validDirectory = value => typeof value === 'string' && value.length <= 32768 && !value.includes('\0') && path.isAbsolute(value);

// Paths stay in the main process and user preferences, never in exported books.
function createFileLocations({ stateFile, documents }) {
  let lastDirectory, loaded, writing = Promise.resolve();
  const directories = new Map(), pending = new Map();
  function load() {
    return loaded ||= (async () => {
      try {
        if ((await fs.stat(stateFile)).size > 131072) return;
        const state = JSON.parse(await fs.readFile(stateFile, 'utf8'));
        if (state.version !== 1) return;
        if (validDirectory(state.lastDirectory)) lastDirectory = state.lastDirectory;
        for (const entry of Array.isArray(state.directories) ? state.directories.slice(-64) : []) {
          if (Array.isArray(entry) && validId(entry[0]) && validDirectory(entry[1])) directories.set(entry[0], entry[1]);
        }
      } catch { /* Missing or damaged preferences must not prevent opening files. */ }
    })();
  }
  async function directory(documentId) {
    await load();
    for (const candidate of new Set([directories.get(documentId), lastDirectory, documents])) {
      if (!candidate) continue;
      try { if ((await fs.stat(candidate)).isDirectory()) return candidate; } catch {}
    }
    return documents;
  }
  async function remember(documentId, file) {
    await load();
    lastDirectory = path.dirname(file);
    if (validId(documentId)) {
      directories.delete(documentId); directories.set(documentId, lastDirectory);
      while (directories.size > 64) directories.delete(directories.keys().next().value);
    }
    const data = JSON.stringify({ version: 1, lastDirectory, directories: [...directories] });
    writing = writing.then(async () => {
      const temp = stateFile + '.' + randomUUID() + '.tmp';
      try {
        await fs.mkdir(path.dirname(stateFile), { recursive: true });
        await fs.writeFile(temp, data, { flag: 'wx', mode: 0o600 });
        await fs.rename(temp, stateFile);
      } finally { await fs.unlink(temp).catch(() => {}); }
    }).catch(error => console.warn('파일 대화상자 위치를 기록하지 못했습니다:', error.code));
    await writing;
  }
  function stage(file) {
    const token = randomUUID(); pending.set(token, file);
    while (pending.size > 16) pending.delete(pending.keys().next().value);
    return token;
  }
  async function confirm({ token, documentId } = {}) {
    if (!pending.has(token) || !validId(documentId)) throw Error('열린 파일의 위치 정보가 올바르지 않습니다.');
    const file = pending.get(token); pending.delete(token);
    await remember(documentId, file);
  }
  return { directory, remember, stage, confirm };
}
module.exports = { createFileLocations };
