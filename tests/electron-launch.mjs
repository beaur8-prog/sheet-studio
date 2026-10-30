import fs from 'node:fs';
import path from 'node:path';
import platform from '../scripts/platform.cjs';
// The same user scenarios can exercise source Electron or the generated app.
// Always require a test profile so a packaged run cannot load the user's work.
export function desktopLaunch(root,environment){
 if(!environment?.SHEET_STUDIO_DATA)throw Error('An isolated SHEET_STUDIO_DATA profile is required.');
 const env={...environment};delete env.ELECTRON_RUN_AS_NODE;
 const packaged=process.argv.includes('--packaged'),version=JSON.parse(fs.readFileSync(path.join(root,'package.json'),'utf8')).version;
 return {executablePath:env.SHEET_STUDIO_EXECUTABLE||(packaged?platform.packagedExecutable(root,version):path.join(platform.findRuntime(root),'dist',platform.runtimeSuffix())),args:packaged?[]:[root],env};
}
