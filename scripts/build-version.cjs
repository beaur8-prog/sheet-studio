'use strict';
const fs = require('node:fs');
const path = require('node:path');

function isCurrentBuild(root, version) {
  try {
    return fs.statSync(path.join(root, 'dist/index.html')).isFile() &&
      JSON.parse(fs.readFileSync(path.join(root, 'dist/app-version.json'), 'utf8')).version === version;
  } catch { return false; }
}
function assertCurrentBuild(root, version) {
  if (!isCurrentBuild(root, version)) {
    throw Error(`The renderer build is missing or differs from app version ${version}. Run npm run build before packaging.`);
  }
}
module.exports = { isCurrentBuild, assertCurrentBuild };
