'use strict';
const fs = require('node:fs/promises');
const nativeFs = require('node:fs');
const path = require('node:path');
const root = path.resolve(__dirname, '..');
const { createIcon, brandExecutable } = require('./branding.cjs');
const { verifyRuntime } = require('./runtime-proof.cjs');
const { generateManifest } = require('./artifacts.cjs');

async function packageApp() {
  const { build, Platform, Arch } = require(require.resolve('electron-builder', { paths: [root] }));
  const manifest = JSON.parse(await fs.readFile(path.join(root, 'package.json'), 'utf8'));
  const locations = [root, path.dirname(root), path.join(path.dirname(root), 'word-studio'), path.join(path.dirname(root), 'slide-studio')];
  const electronRoot = locations.map(folder => path.join(folder, 'node_modules/electron')).find(folder => nativeFs.existsSync(path.join(folder, 'dist/electron.exe')));
  if (!electronRoot) throw new Error('An installed Windows Electron runtime is required. Run npm install first.');
  const runtime = JSON.parse(await fs.readFile(path.join(electronRoot, 'package.json'), 'utf8'));
  const provenance = await verifyRuntime(electronRoot, locations, root);
  // Package only the compiled renderer and its restricted desktop bridge.
  // ExcelJS is already bundled by Vite; no development modules enter the app.
  const stage = path.join(root, '.package');
  await fs.mkdir(stage, { recursive: true });
  await fs.cp(path.join(root, 'dist'), path.join(stage, 'dist'), { recursive: true });
  await fs.cp(path.join(root, 'electron'), path.join(stage, 'electron'), { recursive: true });
  const icon = await createIcon(path.join(stage, 'electron'));
  await fs.writeFile(path.join(stage, 'package.json'), JSON.stringify({
    name: manifest.name, version: manifest.version, description: manifest.description,
    author: manifest.author, main: manifest.main,
  }, null, 2));
  const artifacts = await build({
    projectDir: stage,
    targets: Platform.WINDOWS.createTarget(process.argv.includes('--dir') ? 'dir' : 'portable', Arch.x64),
    config: {
      appId: 'local.lattice.sheetstudio', productName: 'Sheet Studio',
      electronVersion: runtime.version, electronDist: path.join(electronRoot, 'dist'),
      npmRebuild: false, asar: true,
      directories: { output: path.join(root, 'release') },
      files: ['dist/**/*', 'electron/**/*', 'package.json'],
      win: { target: 'portable', signAndEditExecutable: false, icon },
      portable: { artifactName: 'Sheet-Studio-${version}-Windows.exe' },
      afterPack: async context => brandExecutable(path.join(context.appOutDir, 'Sheet Studio.exe'), icon, manifest.version, root),
    },
  });
  await fs.writeFile(path.join(root, 'release/runtime-provenance.json'), JSON.stringify(provenance, null, 2));
  await generateManifest(!process.argv.includes('--dir'));
  console.log('Packaged Sheet Studio:', artifacts.join('\n'));
}

packageApp().catch(error => { console.error(error); process.exitCode = 1; });
