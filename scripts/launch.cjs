'use strict';
const fs = require('node:fs');
const path = require('node:path');
const { spawn, spawnSync } = require('node:child_process');
const root = path.resolve(__dirname, '..');
const suffix = process.platform === 'win32' ? 'electron.exe'
  : process.platform === 'darwin' ? 'Electron.app/Contents/MacOS/Electron' : 'electron';
const locations = [root, path.dirname(root), path.join(path.dirname(root), 'word-studio'), path.join(path.dirname(root), 'slide-studio')];
const executable = locations.map(folder => path.join(folder, 'node_modules/electron/dist', suffix)).find(fs.existsSync);
if (!executable) throw new Error('Electron runtime not found. Run npm install, or use the packaged Sheet Studio app.');
if (!fs.existsSync(path.join(root, 'dist/index.html'))) {
  const build = spawnSync(process.execPath, [path.join(__dirname, 'vite.cjs'), 'build'], { cwd: root, stdio: 'inherit', windowsHide: true });
  if (build.error) throw build.error;
  if (build.status !== 0) process.exit(build.status ?? 1);
}
const env = { ...process.env };
delete env.ELECTRON_RUN_AS_NODE;
const child = spawn(executable, [root, ...process.argv.slice(2)], { cwd: root, env, stdio: 'inherit', windowsHide: true });
child.on('error', error => { console.error(error.message); process.exitCode = 1; });
child.on('exit', code => { process.exitCode = code ?? 0; });
