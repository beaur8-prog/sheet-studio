import test from 'node:test';
import assert from 'node:assert/strict';
import path from 'node:path';
import { pathToFileURL } from 'node:url';
import { FileOpenRequests, fileAssociations } from '../electron/file-open-requests.cjs';
import { installRequestedFileOpener } from '../src/file-open-requests.mjs';
import { windowsFileAssociations } from '../scripts/windows-file-associations.cjs';

test('file-association cleanup never recursively deletes a registry key', () => {
  const script = windowsFileAssociations();
  const deletions = script.split('\n').map(line => line.trim()).filter(line => line.startsWith('DeleteRegKey'));
  assert.ok(deletions.length);
  assert.ok(deletions.every(line => line.startsWith('DeleteRegKey /ifempty ')));
  assert.ok(!script.includes('UserChoice'));
  for (const ext of ['xlsx', 'xlsm', 'csv', 'tsv']) {
    assert.ok(script.includes(`DeleteRegValue SHELL_CONTEXT "Software\\Classes\\.${ext}\\OpenWithProgids" "SheetStudio.${ext}"`));
  }
});

test('OS paths and launch arguments preserve order and coalesce pending duplicates', () => {
  const queue = new FileOpenRequests(), directory = path.resolve('한글 폴더');
  const first = path.join(directory, '첫 문서.XLSX'), second = path.join(directory, '다음 문서.csv');
  queue.addArguments(['sheet-studio', '--inspect', '첫 문서.XLSX', pathToFileURL(first).href, '다음 문서.csv'], directory);
  assert.equal(queue.size, 2);
  assert.equal(queue.take(), first); assert.equal(queue.take(), second);
  assert.equal(queue.take(), undefined);
  assert.equal(queue.add(first), true, 'an explicit later open can reopen the same file');
  for (const extension of ['xlsm', 'tsv', 'sheet.json', 'json']) assert.equal(queue.add(path.join(directory, '문서.' + extension)), true);
});

test('OS request queue rejects switches, remote URLs, invalid paths and unsupported formats', () => {
  const queue = new FileOpenRequests();
  for (const invalid of [null, '', 12, '--export=file.xlsx', 'https://example.com/file.xlsx', 'file://host/%ZZ', 'bad\0.xlsx', 'deck.pptx', 'script.js'])
    assert.equal(queue.add(invalid), false, String(invalid));
  for (let i = 0; i < 64; i++) assert.equal(queue.add(`file-${i}.xlsx`), true);
  assert.equal(queue.add('overflow.xlsx'), false);
  assert.equal(queue.size, 64);
});

test('installation offers supported spreadsheet types without claiming all JSON documents', () => {
  assert.deepEqual(fileAssociations.map(type => type.ext), ['xlsx', 'xlsm', 'csv', 'tsv']);
  assert.ok(fileAssociations.every(type => type.rank === 'Alternate' && type.role === 'Editor' && type.mimeType));
});

async function until(predicate) {
  const deadline = Date.now() + 2000;
  while (!predicate()) {
    if (Date.now() > deadline) throw Error('File-open controller timed out');
    await new Promise(resolve => setTimeout(resolve, 10));
  }
}
function bridge(files) {
  let listener;
  return {
    takeRequestedFile: async () => files.shift() || null,
    onOpenRequested: callback => { listener = callback; return () => { listener = undefined; }; },
    notify: () => listener?.(),
  };
}

test('editor delays startup and running requests while a dialog or workbook operation is active', async () => {
  let ready = false;
  const files = [{ name: 'one.xlsx' }, { name: 'two.csv' }], opened = [], errors = [];
  const desktop = bridge(files);
  const dispose = installRequestedFileOpener({ desktop, canOpen: () => ready, openFile: async file => { opened.push(file.name); ready = false; }, onError: error => errors.push(error) });
  try {
    await new Promise(resolve => setTimeout(resolve, 100)); assert.deepEqual(opened, []);
    ready = true; await until(() => opened.length === 1);
    desktop.notify(); await new Promise(resolve => setTimeout(resolve, 100)); assert.deepEqual(opened, ['one.xlsx']);
    ready = true; await until(() => opened.length === 2);
    assert.deepEqual(opened, ['one.xlsx', 'two.csv']); assert.deepEqual(errors, []);
  } finally { dispose(); }
});

test('a dialog opened while reading file bytes retains that file until dismissal', async () => {
  let ready = true, finishRead;
  const desktop = bridge([]), opened = [];
  desktop.takeRequestedFile = () => new Promise(resolve => { finishRead = resolve; });
  const dispose = installRequestedFileOpener({ desktop, canOpen: () => ready, openFile: async file => { opened.push(file.name); }, onError: error => { throw error; } });
  try {
    await until(() => finishRead);
    ready = false; finishRead({ name: 'kept.xlsx' });
    await new Promise(resolve => setTimeout(resolve, 100)); assert.deepEqual(opened, []);
    desktop.takeRequestedFile = async () => null;
    ready = true; await until(() => opened.length);
    assert.deepEqual(opened, ['kept.xlsx']);
  } finally { dispose(); }
});

test('unreadable and invalid documents do not strand the next file or a notification during a read', async () => {
  const files = [new Error('missing'), { name: 'broken.xlsx' }, { name: 'valid.xlsx' }], opened = [], errors = [];
  const desktop = bridge(files);
  desktop.takeRequestedFile = async () => {
    const file = files.shift();
    if (files.length) desktop.notify();
    if (file instanceof Error) throw file;
    return file || null;
  };
  const dispose = installRequestedFileOpener({ desktop, canOpen: () => true, openFile: async file => {
    if (file.name === 'broken.xlsx') throw Error('invalid'); opened.push(file.name);
  }, onError: error => errors.push(error.message) });
  try {
    await until(() => opened.length);
    assert.deepEqual(opened, ['valid.xlsx']); assert.deepEqual(errors, ['missing', 'invalid']);
  } finally { dispose(); }
});


test('a persistent native read failure returns control and waits for a new request', async () => {
  const desktop = bridge([]), errors = [], opened = [];
  let attempts = 0, failing = true;
  desktop.takeRequestedFile = async () => {
    attempts++;
    if (failing) throw Error('native unavailable');
    return attempts === 2 ? { name: 'later.xlsx' } : null;
  };
  const dispose = installRequestedFileOpener({ desktop, canOpen: () => true,
    openFile: async file => opened.push(file.name), onError: error => errors.push(error.message) });
  try {
    await until(() => errors.length);
    await new Promise(resolve => setTimeout(resolve, 200));
    assert.equal(attempts, 1);
    assert.deepEqual(errors, ['native unavailable']);
    failing = false; desktop.notify();
    await until(() => opened.length);
    assert.deepEqual(opened, ['later.xlsx']);
  } finally { dispose(); }
});
