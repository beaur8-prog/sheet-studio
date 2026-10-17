'use strict';
const fs = require('node:fs/promises');
const nativeFs = require('node:fs');
const path = require('node:path');
const { createHash } = require('node:crypto');

async function hashFile(file) {
  const hash = createHash('sha256');
  for await (const chunk of nativeFs.createReadStream(file)) hash.update(chunk);
  return hash.digest('hex');
}

async function verifyRuntime(electronRoot, locations, root) {
  const { version } = JSON.parse(await fs.readFile(path.join(electronRoot, 'package.json'), 'utf8'));
  const filename = `electron-v${version}-win32-x64.zip`;
  const archive = locations.map(directory => path.join(directory, '.cache/electron', filename)).find(nativeFs.existsSync);
  const executable = path.join(electronRoot, 'dist/electron.exe');
  const executableSha256 = await hashFile(executable);
  const result = {
    version, source: `https://github.com/electron/electron/releases/download/v${version}/${filename}`,
    installedRuntime: path.join(electronRoot, 'dist'), executableSha256,
    verification: 'installed npm Electron runtime',
  };
  if (archive) {
    const checksums = JSON.parse(await fs.readFile(path.join(electronRoot, 'checksums.json'), 'utf8'));
    const archiveSha256 = await hashFile(archive);
    if (archiveSha256 !== checksums[filename]) throw new Error('Cached Electron archive checksum mismatch.');
    const JSZip = require(require.resolve('jszip', { paths: [root] }));
    const zip = await JSZip.loadAsync(await fs.readFile(archive));
    const officialExe = await zip.file('electron.exe').async('nodebuffer');
    const expected = createHash('sha256').update(officialExe).digest('hex');
    if (expected !== executableSha256) throw new Error('Installed Electron executable differs from its verified release archive.');
    result.archiveSha256 = archiveSha256;
    result.verification = 'release archive SHA-256 matches npm Electron checksums; installed electron.exe matches archive';
  }
  return result;
}

module.exports = { verifyRuntime };
