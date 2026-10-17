'use strict';
const fs = require('node:fs/promises');
const nativeFs = require('node:fs');
const path = require('node:path');
const { createHash } = require('node:crypto');
const root = path.resolve(__dirname, '..');
async function generateManifest({ output, appOutDir, files = [], platform, arch }) {
  const { version } = JSON.parse(await fs.readFile(path.join(root, 'package.json'), 'utf8'));
  const appFiles = platform === 'darwin'
    ? ['Sheet Studio.app/Contents/MacOS/Sheet Studio', 'Sheet Studio.app/Contents/Resources/app.asar']
    : ['Sheet Studio.exe', 'resources/app.asar'];
  const artifacts = [];
  for (const absolute of new Set([...files, ...appFiles.map(file => path.join(appOutDir, file))])) {
    const stat = await fs.stat(absolute);
    if (!stat.isFile()) continue;
    const hash = createHash('sha256');
    for await (const chunk of nativeFs.createReadStream(absolute)) hash.update(chunk);
    artifacts.push({ filename: path.relative(output, absolute).split(path.sep).join('/'), bytes: stat.size, sha256: hash.digest('hex') });
  }
  const result = {
    product: 'Sheet Studio', version, platform, architecture: arch,
    createdAt: new Date().toISOString(), artifacts,
    note: platform === 'darwin' ? 'Keep the complete .app bundle. Unsigned builds are for local testing.' : 'The portable EXE is self contained. The unpacked EXE requires the complete win-unpacked directory.',
  };
  await fs.writeFile(path.join(output, 'artifact-manifest.json'), JSON.stringify(result, null, 2));
  return result;
}
module.exports = { generateManifest };
