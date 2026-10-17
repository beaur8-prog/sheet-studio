'use strict';
const fs = require('node:fs/promises');
const nativeFs = require('node:fs');
const path = require('node:path');
const { createHash } = require('node:crypto');
const root = path.resolve(__dirname, '..');

async function generateManifest(includePortable = true) {
  const { version } = JSON.parse(await fs.readFile(path.join(root, 'package.json'), 'utf8'));
  const files = ['win-unpacked/Sheet Studio.exe', 'win-unpacked/resources/app.asar'];
  if (includePortable) files.unshift(`Sheet-Studio-${version}-Windows.exe`);
  const artifacts = [];
  for (const filename of files) {
    const absolute = path.join(root, 'release', filename);
    const hash = createHash('sha256');
    for await (const chunk of nativeFs.createReadStream(absolute)) hash.update(chunk);
    artifacts.push({ filename, bytes: (await fs.stat(absolute)).size, sha256: hash.digest('hex') });
  }
  const result = {
    product: 'Sheet Studio', version, platform: 'win32', architecture: 'x64',
    createdAt: new Date().toISOString(), artifacts,
    note: 'The portable EXE is self contained. The unpacked EXE requires the complete win-unpacked directory.',
  };
  await fs.writeFile(path.join(root, 'release/artifact-manifest.json'), JSON.stringify(result, null, 2));
  return result;
}
module.exports = { generateManifest };
if (require.main === module) generateManifest(!process.argv.includes('--dir')).then(value => console.log(JSON.stringify(value, null, 2))).catch(error => { console.error(error); process.exitCode = 1; });
