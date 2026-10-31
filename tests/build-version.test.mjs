import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import buildVersion from '../scripts/build-version.cjs';

test('launch and packaging reject missing, unversioned and stale renderer artifacts', t => {
  const root = fs.mkdtempSync(path.join(os.tmpdir(), 'sheet-build-version-'));
  t.after(() => fs.rmSync(root, { recursive: true, force: true }));
  const check = expected => {
    assert.equal(buildVersion.isCurrentBuild(root, '0.7.32'), expected);
    if (expected) assert.doesNotThrow(() => buildVersion.assertCurrentBuild(root, '0.7.32'));
    else assert.throws(() => buildVersion.assertCurrentBuild(root, '0.7.32'), /npm run build/);
  };
  check(false);
  fs.mkdirSync(path.join(root, 'dist'));
  fs.writeFileSync(path.join(root, 'dist/index.html'), '<html></html>');
  check(false);
  const stamp = path.join(root, 'dist/app-version.json');
  fs.writeFileSync(stamp, JSON.stringify({ version: '0.7.31' }));
  check(false);
  fs.writeFileSync(stamp, 'incomplete build');
  check(false);
  fs.writeFileSync(stamp, JSON.stringify({ version: '0.7.32' }));
  check(true);
  fs.unlinkSync(path.join(root, 'dist/index.html'));
  check(false);
});
