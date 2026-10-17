'use strict';
const fs = require('node:fs');
const path = require('node:path');
function buildTarget(args, host = process.platform, hostArch = process.arch) {
  const allowed = new Set(['--mac', '--win', '--dir', '--x64', '--arm64', '--arch=x64', '--arch=arm64']);
  for (const arg of args) if (!allowed.has(arg)) throw new Error('Unknown package option: ' + arg);
  if (args.includes('--mac') && args.includes('--win')) throw new Error('Choose one platform.');
  const platform = args.includes('--mac') ? 'darwin' : args.includes('--win') ? 'win32' : host;
  if (!['win32', 'darwin'].includes(platform)) throw new Error('Packaging supports Windows and macOS.');
  if (platform !== host) throw new Error('Build macOS packages on macOS and Windows packages on Windows.');
  const arches = [...new Set(args.filter(arg => /^(--arch=|--)(x64|arm64)$/.test(arg)).map(arg => arg.replace(/^--(?:arch=)?/, '')))];
  if (arches.length > 1) throw new Error('Choose one architecture per build.');
  const arch = arches[0] || (platform === 'win32' ? 'x64' : hostArch);
  if (!['x64', 'arm64'].includes(arch) || (platform === 'win32' && arch !== 'x64')) throw new Error('Supported targets: Windows x64, macOS x64 and arm64.');
  return { platform, arch, directory: args.includes('--dir') };
}
function runtimeSuffix(platform = process.platform) {
  return platform === 'win32' ? 'electron.exe' : platform === 'darwin' ? 'Electron.app/Contents/MacOS/Electron' : 'electron';
}
function runtimeLocations(root) {
  return [root, path.dirname(root), path.join(path.dirname(root), 'word-studio'), path.join(path.dirname(root), 'slide-studio')];
}
function findRuntime(root, platform = process.platform) {
  const folder = runtimeLocations(root).map(base => path.join(base, 'node_modules/electron')).find(base => fs.existsSync(path.join(base, 'dist', runtimeSuffix(platform))));
  if (!folder) throw new Error('Electron runtime not found. Run npm install first.');
  return folder;
}
function outputDir(root, version, platform = process.platform, arch = process.arch) {
  return path.join(root, 'release', version, (platform === 'darwin' ? 'mac' : 'win') + '-' + arch);
}
function packagedExecutable(root, version, platform = process.platform, arch = process.arch) {
  const output = outputDir(root, version, platform, arch);
  return platform === 'darwin' ? path.join(output, arch === 'arm64' ? 'mac-arm64' : 'mac', 'Sheet Studio.app/Contents/MacOS/Sheet Studio') : path.join(output, 'win-unpacked/Sheet Studio.exe');
}
module.exports = { buildTarget, runtimeSuffix, runtimeLocations, findRuntime, outputDir, packagedExecutable };
