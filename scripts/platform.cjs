'use strict';
const fs = require('node:fs');
const path = require('node:path');
function buildTarget(args, host = process.platform, hostArch = process.arch) {
  const allowed = new Set(['--mac', '--win', '--linux', '--dir', '--x64', '--arm64', '--arch=x64', '--arch=arm64']);
  for (const arg of args) if (!allowed.has(arg)) throw new Error('Unknown package option: ' + arg);
  if (['--mac','--win','--linux'].filter(flag=>args.includes(flag)).length>1) throw new Error('Choose one platform.');
  const platform = args.includes('--mac') ? 'darwin' : args.includes('--win') ? 'win32' : args.includes('--linux') ? 'linux' : host;
  if (!['win32', 'darwin', 'linux'].includes(platform)) throw new Error('Packaging supports Windows, macOS and Linux.');
  if (platform !== host) throw new Error('Build macOS packages on macOS, Windows packages on Windows and Linux packages on Linux.');
  const arches = [...new Set(args.filter(arg => /^(--arch=|--)(x64|arm64)$/.test(arg)).map(arg => arg.replace(/^--(?:arch=)?/, '')))];
  if (arches.length > 1) throw new Error('Choose one architecture per build.');
  const arch = arches[0] || (platform === 'win32' ? 'x64' : hostArch);
  if (!['x64', 'arm64'].includes(arch) || (['win32','linux'].includes(platform) && arch !== 'x64')) throw new Error('Supported targets: Windows x64, Linux x64, macOS x64 and arm64.');
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
  return path.join(root, 'release', version, (platform === 'darwin' ? 'mac' : platform === 'linux' ? 'linux' : 'win') + '-' + arch);
}
function packagedExecutable(root, version, platform = process.platform, arch = process.arch) {
  const output = outputDir(root, version, platform, arch);
  return platform === 'darwin' ? path.join(output, arch === 'arm64' ? 'mac-arm64' : 'mac', 'Sheet Studio.app/Contents/MacOS/Sheet Studio') : platform === 'linux' ? path.join(output, 'linux-unpacked/sheet-studio') : path.join(output, 'win-unpacked/Sheet Studio.exe');
}
module.exports = { buildTarget, runtimeSuffix, runtimeLocations, findRuntime, outputDir, packagedExecutable };
