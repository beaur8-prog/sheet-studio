'use strict';
const path = require('node:path');
const { spawnSync } = require('node:child_process');
const root = path.resolve(__dirname, '..');
const script = path.join(path.dirname(require.resolve('vite/package.json', { paths: [root] })), 'bin/vite.js');
const result = spawnSync(process.execPath, [script, ...process.argv.slice(2)], {
  cwd: root, stdio: 'inherit', windowsHide: true,
});
if (result.error) throw result.error;
process.exitCode = result.status ?? 1;
