import {_electron} from 'playwright';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import path from 'node:path';
import {createHash} from 'node:crypto';
import {desktopLaunch} from './electron-launch.mjs';

const root=path.resolve(import.meta.dirname,'..'),profile=await fs.mkdtemp(path.join(root,'.package/update-ui-'));
const environment={...process.env,SHEET_STUDIO_DATA:profile,STUDIO_DISABLE_UPDATES:'1'};
let application,page;
const errors=[];
async function launch(){
 application=await _electron.launch(desktopLaunch(root,environment));
 page=await application.firstWindow();page.setDefaultTimeout(20000);
 page.on('pageerror',error=>errors.push(error.message));await page.locator('.cell').first().waitFor();
 // All test updates are inert byte strings served in memory. They are never
 // executed. The real OS clipboard, browser, installer and user profile are untouched.
 await application.evaluate(({app,dialog,shell},sha256)=>{
  Object.defineProperty(app,'isPackaged',{value:true,configurable:true});
  const bytes=Buffer.from('0123456789abcdef'.repeat(4));
  const fixture=globalThis.updateFixture={bytes,sha256,mode:'available',requests:[],opened:[],folders:[],dialogs:[],feed:null};
  shell.openExternal=async url=>{fixture.opened.push(url);};shell.showItemInFolder=file=>fixture.folders.push(file);
  dialog.showMessageBox=async(...args)=>{fixture.dialogs.push(args.at(-1));return {response:process.platform==='darwin'?0:1};};
  globalThis.fetch=async url=>{
   fixture.requests.push(url);
   if(url.startsWith('https://api.github.com/repos/beaur8-prog/sheet-studio/')){
    if(fixture.mode==='error')throw Error('test network failure <test>');
    const version=fixture.mode==='current'?app.getVersion():'9.9.9';
    const filename=process.platform==='darwin'?`Sheet-Studio-${version}-macOS-${process.arch}.zip`:`Sheet-Studio-${version}-Windows-x64-Setup.exe`;
    const release={tag_name:'v'+version,draft:false,prerelease:false,assets:[{name:filename,state:'uploaded',size:bytes.length,digest:'sha256:'+sha256,browser_download_url:`https://github.com/beaur8-prog/sheet-studio/releases/download/v${version}/${filename}`}]};
    return Response.json(process.platform==='darwin'?[release]:release);
   }
   if(!url.startsWith('https://github.com/beaur8-prog/sheet-studio/releases/download/v9.9.9/'))throw Error('Unexpected test URL');
   return new Response(new ReadableStream({start(controller){fixture.feed=controller;}}));
  };
 },createHash('sha256').update('0123456789abcdef'.repeat(4)).digest('hex'));
}
const center=()=>page.getByRole('dialog',{name:'업데이트',exact:true});
const close=()=>center().getByRole('button',{name:'닫기',exact:true}).click();
const native=mode=>application.evaluate(({Menu},mode)=>Menu.getApplicationMenu().getMenuItemById('sheet-updates-'+mode).click(),mode);
const phase=value=>page.waitForFunction(value=>document.querySelector('.update-modal')?.closest('.backdrop')?.dataset.phase===value,value);
try{
 await fs.mkdir(path.join(root,'test-results'),{recursive:true});
 await launch();
 await page.locator('#name-box').fill('A1');await page.locator('#name-box').press('Enter');
 await page.keyboard.type('update retains document');await page.keyboard.press('Enter');
 await page.locator('[data-tab="도움말"]').click();
 for(const name of ['업데이트 확인','수동 업데이트','업데이트 진행 상황'])assert.equal(await page.getByRole('button',{name,exact:true}).count(),1);
 await native('status');await center().waitFor();
 assert.equal(await center().locator('#update-current').innerText(),await page.evaluate(()=>desktop.getAppVersion()));
 await page.locator('#update-manual-download').click();
 assert.deepEqual(await application.evaluate(()=>updateFixture.opened),['https://github.com/beaur8-prog/sheet-studio/releases/latest']);
 assert.equal(await application.evaluate(()=>updateFixture.requests.length),0);
 if(process.platform!=='linux'){
  await page.locator('#update-manual').check();
  await page.waitForFunction(async()=>!(await desktop.getUpdateState()).automatic);
  assert.equal(JSON.parse(await fs.readFile(path.join(profile,'updates/settings.json'),'utf8')).automatic,false);
  await close();await native('check');await phase('available');
  assert.equal(await application.evaluate(()=>updateFixture.requests.length),1,'check must not download');
  assert.equal(await page.locator('#update-latest').innerText(),'9.9.9');
  await page.locator('#update-download').click();await phase('downloading');
  await application.evaluate(()=>new Promise((resolve,reject)=>{const end=Date.now()+5000;function poll(){if(updateFixture.feed)resolve();else if(Date.now()>end)reject(Error('Missing download stream'));else setTimeout(poll,10);}poll();}));
  await application.evaluate(()=>updateFixture.feed.enqueue(updateFixture.bytes.subarray(0,32)));
  await page.waitForFunction(()=>document.querySelector('#update-progress').value===50);
  await close();assert.equal(await page.locator('#update-status').isVisible(),true);
  await application.evaluate(()=>updateFixture.feed.enqueue(updateFixture.bytes.subarray(32,48)));
  await page.waitForFunction(()=>document.querySelector('#update-status').textContent.includes('75%'));
  await page.locator('#update-status').click();await phase('downloading');
  assert.equal(await page.locator('#update-progress').getAttribute('value'),'75');
  assert.match(await page.locator('#update-transfer').innerText(),/48 B \/ 64 B/);
  await page.screenshot({path:path.join(root,'test-results',`updates-progress-${process.platform}.png`)});
  await application.evaluate(()=>{updateFixture.feed.enqueue(updateFixture.bytes.subarray(48));updateFixture.feed.close();});
  await phase('ready');assert.equal(await page.locator('#update-progress').getAttribute('value'),'100');
  await page.screenshot({path:path.join(root,'test-results',`updates-ready-${process.platform}.png`)});
  const ready=await page.evaluate(()=>desktop.getUpdateState());assert.equal(ready.downloadedBytes,64);assert.equal(ready.totalBytes,64);
  const folder=path.join(profile,'updates',await application.evaluate(()=>updateFixture.sha256));
  const files=await fs.readdir(folder);assert.equal(files.length,1);assert.equal((await fs.stat(path.join(folder,files[0]))).size,64);
  await page.locator('#update-install').click();
  await application.evaluate(()=>new Promise((resolve,reject)=>{const end=Date.now()+5000;function poll(){if(updateFixture.dialogs.length)resolve();else if(Date.now()>end)reject(Error('Missing install instructions'));else setTimeout(poll,10);}poll();}));
  assert.equal(await page.evaluate(()=>sheetStudio.getCell('A1').raw),'update retains document');
  await application.evaluate(()=>{updateFixture.mode='current';});await page.locator('#update-check').click();await phase('up-to-date');
  await application.evaluate(()=>{updateFixture.mode='error';});await page.locator('#update-check').click();await phase('error');
  assert.match(await page.locator('#update-error').innerText(),/test network failure <test>/);
  await fs.mkdir(path.join(root,'test-results'),{recursive:true});
  await page.screenshot({path:path.join(root,'test-results',`updates-error-${process.platform}.png`)});
  await application.evaluate(()=>{updateFixture.mode='current';});await page.locator('#update-check').click();await phase('up-to-date');
  assert.equal(await page.locator('#update-error').isVisible(),false);
  await application.close();application=null;await launch();
  await page.waitForFunction(async()=>!(await desktop.getUpdateState()).automatic);
  await native('manual');await center().waitFor();assert.equal(await page.locator('#update-manual').isChecked(),true);
  assert.equal(await page.evaluate(()=>sheetStudio.getCell('A1').raw),'update retains document');
 }else{
  assert.equal(await page.locator('#update-mode').isVisible(),false);
  assert.equal(await page.locator('#update-check').isEnabled(),false);
  assert.match(await page.locator('#update-platform').innerText(),/Linux/);
 }
 for(const [command,value] of [['executeJavaScript',undefined],['manual','https://example.invalid'],['automatic','false']]){
  assert.equal(await page.evaluate(async([command,value])=>{try{await desktop.updateAction(command,value);return false;}catch{return true;}},[command,value]),true);
 }
 assert.deepEqual(errors,[]);
 await fs.mkdir(path.join(root,'test-results'),{recursive:true});
 await page.screenshot({path:path.join(root,'test-results',`updates-${process.platform}.png`)});
 await fs.writeFile(path.join(root,'test-results',`updates-${process.platform}.json`),JSON.stringify({platform:process.platform,packaged:process.argv.includes('--packaged'),ok:true,scope:process.platform==='linux'?'native/ribbon menus, manual release page, unsupported controls, safe IPC':'native/ribbon menus, manual mode, 50/75/100 percent and byte progress, close/reopen progress, SHA-256 download, explicit install instructions, network retry, preference and document recovery',errors},null,2));
 console.log('PASS: update menus, manual downloads, isolated state/progress, preferences, safe IPC and document recovery ('+process.platform+').');
}finally{if(application)await application.close();await fs.rm(profile,{recursive:true,force:true});}
