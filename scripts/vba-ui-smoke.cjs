'use strict';
const fs=require('node:fs'),fsp=require('node:fs/promises'),path=require('node:path'),assert=require('node:assert/strict');
const {_electron}=require('playwright');
const root=path.resolve(__dirname,'..');
const folders=[root,path.dirname(root),path.join(path.dirname(root),'word-studio'),path.join(path.dirname(root),'slide-studio')];
const runtime=folders.map(p=>path.join(p,'node_modules/electron/dist/electron.exe')).find(fs.existsSync);
async function run(){
  const packaged=process.argv.includes('--packaged');
  const source=path.resolve(process.argv.slice(2).find(a=>!a.startsWith('--'))||path.join(root,'.verification/native-source.xlsm'));
  const destination=path.join(root,'.verification/ui-saved.xlsm');
  const env={...process.env,SHEET_STUDIO_DATA:path.join(root,'.package/vba-smoke-data')};delete env.ELECTRON_RUN_AS_NODE;
  const application=await _electron.launch({executablePath:packaged?require('./platform.cjs').packagedExecutable(root,require('../package.json').version):runtime,args:packaged?[]:[root],env});
  try{
    const page=await application.firstWindow(),errors=[];page.on('pageerror',e=>errors.push(e.message));
    await page.waitForSelector('.cell');
    await application.evaluate(({dialog},paths)=>{dialog.showOpenDialog=async()=>({canceled:false,filePaths:[paths.source]});dialog.showSaveDialog=async()=>({canceled:false,filePath:paths.destination});},{source,destination});
    await page.keyboard.press('Control+o');
    await page.waitForFunction(()=>document.querySelector('#status')?.textContent.includes('XLSM'));
    await page.locator('[data-tab="개발 도구"]').click();
    await page.locator('[data-action="vba"]').click();
    await page.locator('[data-vba="connect"]:enabled').waitFor();
    await page.screenshot({path:path.join(root,'test-results/vba-panel.png')});
    await page.locator('[data-vba="connect"]').click();
    await page.locator('[data-vba="finish"]:enabled').waitFor({timeout:60000});
    await page.keyboard.press('Control+n');assert.equal(await page.locator('.vba-overlay').count(),1);
    await page.locator('[data-vba="finish"]').click();
    await page.locator('.vba-overlay').waitFor({state:'detached',timeout:60000});
    await page.locator('.cell[aria-label="A1"]').dblclick();await page.locator('#editor').fill('13');await page.locator('#editor').press('Enter');
    await page.locator('[data-action="save-xlsm"]').click();
    await page.locator('.vba-overlay').waitFor({state:'detached',timeout:60000});
    const {readSpreadsheet}=await import('../src/io.mjs');
    const saved=await readSpreadsheet(await fsp.readFile(destination),'saved.xlsm');
    assert.equal(saved.sheets[0].cells.A1.raw,'13');assert.equal(saved.sheets[0].cells.A2.cachedValue,26);
    assert.deepEqual(errors,[]);console.log(JSON.stringify({ok:true,xlsmUI:true,nativeConnect:true,importResult:true,preservingSave:true,rendererErrors:errors}));
  }finally{await application.close();}
}
run().catch(e=>{console.error(e);process.exitCode=1;});
