'use strict';
const fs = require('node:fs/promises');
const path = require('node:path');
const root = path.resolve(__dirname, '..');
const { createIcon, brandExecutable } = require('./branding.cjs');
const { verifyRuntime } = require('./runtime-proof.cjs');
const { generateManifest } = require('./artifacts.cjs');
const { buildTarget, findRuntime, runtimeLocations, outputDir } = require('./platform.cjs');

async function packageApp() {
  const target = buildTarget(process.argv.slice(2));
  const { build, Platform, Arch } = require('electron-builder');
  const manifest = JSON.parse(await fs.readFile(path.join(root, 'package.json'), 'utf8'));
  const electronRoot = findRuntime(root);
  const runtime = JSON.parse(await fs.readFile(path.join(electronRoot, 'package.json'), 'utf8'));
  if (runtime.version !== manifest.devDependencies.electron) throw new Error('Installed Electron version differs from package.json. Run npm install.');
  const reuseRuntime = target.arch === process.arch;
  const provenance = reuseRuntime
    ? await verifyRuntime(electronRoot, runtimeLocations(root), root, target)
    : { version: runtime.version, platform: target.platform, architecture: target.arch, verification: 'Target runtime downloaded by electron-builder; local host runtime not reused.' };
  // Isolated staging avoids stale files from earlier builds.
  await fs.mkdir(path.join(root, '.package'), { recursive: true });
  const stage = await fs.mkdtemp(path.join(root, '.package/build-'));
  await fs.cp(path.join(root, 'dist'), path.join(stage, 'dist'), { recursive: true });
  await fs.cp(path.join(root, 'electron'), path.join(stage, 'electron'), { recursive: true });
  const icon = await createIcon(path.join(stage, 'electron'));
  await fs.writeFile(path.join(stage, 'package.json'), JSON.stringify({
    name: manifest.name, version: manifest.version, description: manifest.description,
    author: manifest.author, main: manifest.main,
  }, null, 2));
  const output = process.env.SHEET_STUDIO_OUTPUT_DIR
    ? path.resolve(process.env.SHEET_STUDIO_OUTPUT_DIR)
    : outputDir(root, manifest.version, target.platform, target.arch);
  const mac = target.platform === 'darwin';
  let appOutDir;
  const artifacts = await build({
    projectDir: stage,
    publish: 'never',
    targets: (mac ? Platform.MAC : Platform.WINDOWS).createTarget(target.directory ? 'dir' : mac ? ['dmg', 'zip'] : 'portable', Arch[target.arch]),
    config: {
      appId: 'local.lattice.sheetstudio', productName: 'Sheet Studio', publish: null,
      electronVersion: runtime.version, ...(reuseRuntime ? { electronDist: path.join(electronRoot, 'dist') } : {}),
      npmRebuild: false, asar: true,
      directories: { output },
      files: ['dist/**/*', 'electron/**/*', 'package.json'],
      win: { target: 'portable', signAndEditExecutable: false, icon },
      portable: { artifactName: 'Sheet-Studio-${version}-Windows-${arch}.exe' },
      mac: {
        category: 'public.app-category.productivity', icon: path.join(stage, 'electron/app.png'),
        artifactName: 'Sheet-Studio-${version}-macOS-${arch}.${ext}',
        ...(process.env.SHEET_STUDIO_ADHOC === '1' ? { identity: '-', hardenedRuntime: false, notarize: false } : {}),
      },
      afterPack: async context => {
        appOutDir = context.appOutDir;
        // Clean Finder metadata from the generated bundle before macOS signing.
        // Leave the installed runtime untouched.
        if (mac) await require('node:util').promisify(require('node:child_process').execFile)(
          '/usr/bin/xattr', ['-cr', path.join(context.appOutDir, 'Sheet Studio.app')]);
        if (!mac) await brandExecutable(path.join(context.appOutDir, 'Sheet Studio.exe'), icon, manifest.version, root);
      },
    },
  });
  await fs.writeFile(path.join(output, 'runtime-provenance.json'), JSON.stringify(provenance, null, 2));
  await generateManifest({ output, appOutDir, files: artifacts, platform: target.platform, arch: target.arch });
  console.log('Packaged Sheet Studio:', artifacts.join('\n'));
}
packageApp().catch(error => { console.error(error); process.exitCode = 1; });
