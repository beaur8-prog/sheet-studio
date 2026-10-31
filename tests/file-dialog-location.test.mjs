import test from 'node:test';
import assert from 'node:assert/strict';
import { prepareFileDialogLocation } from '../electron/file-dialog-location.cjs';

test('Mac chooser uses an app-scoped plain directory instead of resolving a saved bookmark', () => {
  const written = [];
  const preferences = { setUserDefault: (...args) => written.push(args) };
  prepareFileDialogLocation(preferences, '/Users/test/한글 폴더', 'darwin');
  assert.deepEqual(written, [['NSOSPLastRootDirectory', 'string', '/Users/test/한글 폴더']]);
  for (const platform of ['win32', 'linux']) prepareFileDialogLocation(preferences, '/documents', platform);
  for (const directory of [undefined, null, '', 'relative']) prepareFileDialogLocation(preferences, directory, 'darwin');
  assert.equal(written.length, 1, 'other platforms and shared recent-folder defaults are untouched');
});
