// Exercises the packaged Sheet editor's existing recovery/close IPC with a real
// Windows helper and disposable fixture installer, never an installed user app.
import {_electron} from 'playwright';
import fs from 'node:fs/promises';
import path from 'node:path';
import os from 'node:os';
import assert from 'node:assert/strict';
import {createHash} from 'node:crypto';
import {execFileSync} from 'node:child_process';
import {createRequire} from 'node:module';
import {assertBlankStart,recoverPrevious} from './recovery-ui.mjs';
import {powershellCommand} from '../electron/studio-updater.cjs';
import platform from '../scripts/platform.cjs';
if(process.platform!=='win32'||process.env.GITHUB_ACTIONS!=='true')throw Error('Requires a disposable Windows CI runner');
const root=path.resolve(import.meta.dirname,'..'),version=createRequire(import.meta.url)('../package.json').version;
const original=platform.packagedExecutable(root,version),directory=await fs.mkdtemp(path.join(os.tmpdir(),'sheet-update-close-'));
const report=path.join(root,'test-results/windows-update-close');await fs.mkdir(report,{recursive:true});
const ps=path.join(process.env.SystemRoot,'System32/WindowsPowerShell/v1.0/powershell.exe');
const delay=ms=>new Promise(r=>setTimeout(r,ms));
const exists=file=>fs.access(file).then(()=>true,()=>false);
async function until(check){const end=Date.now()+45000;while(Date.now()<end){if(await check())return;await delay(100);}throw Error('Timed out waiting for native update fixture');}
await fs.writeFile(path.join(directory,'payload.cs'),'using System; using System.IO; using System.Reflection; [assembly: AssemblyVersion("9.9.9.0")] [assembly: AssemblyInformationalVersion("9.9.9")] public class NewApp { public static void Main() { File.WriteAllText(Path.Combine(Environment.GetEnvironmentVariable("SHEET_UPDATE_FIXTURE_ROOT"), "reopened"), "yes"); } }');
await fs.writeFile(path.join(directory,'installer.cs'),'using System; using System.IO; using System.Diagnostics; public class Setup { public static void Main() { try { var parent = Process.GetProcessById(int.Parse(Environment.GetEnvironmentVariable("SHEET_UPDATE_FIXTURE_PID"))); if (!parent.HasExited) throw new Exception("Editor still running"); } catch (ArgumentException) {} var root = Environment.GetEnvironmentVariable("SHEET_UPDATE_FIXTURE_ROOT"); File.Copy(Path.Combine(root,"payload.exe"), Environment.GetEnvironmentVariable("SHEET_UPDATE_FIXTURE_EXE"), true); File.WriteAllText(Path.Combine(root,"installed"),"yes"); } }');
execFileSync(ps,['-NoLogo','-NoProfile','-NonInteractive','-EncodedCommand',powershellCommand('param($JobFile)\n$ErrorActionPreference="Stop"\nforeach ($name in @("payload","installer")) { Add-Type -Path (Join-Path $JobFile "$name.cs") -OutputAssembly (Join-Path $JobFile "$name.exe") -OutputType WindowsApplication }',directory)],{stdio:'inherit',timeout:60000});
const results=[];let application,page;
async function launch(executable,env){console.log('Launching packaged fixture:',executable);const environment={...env};delete environment.ELECTRON_RUN_AS_NODE;application=await _electron.launch({executablePath:executable,args:[],env:environment,timeout:45000});page=await application.firstWindow();page.setDefaultTimeout(20000);await page.locator('.cell').first().waitFor();console.log('Packaged fixture editor ready');}
async function edit(ref,text){await page.locator('#name-box').fill(ref);await page.locator('#name-box').press('Enter');await page.keyboard.type(text);await page.keyboard.press('Enter');}
try{
 for(const mode of ['install','save-failure']){
  const folder=path.join(directory,mode),appFolder=path.join(folder,'editor'),profile=path.join(folder,"profile 한글 O'Brien; &");await fs.mkdir(folder);
  await fs.cp(path.dirname(original),appFolder,{recursive:true});
  for(const file of ['payload.exe','installer.exe'])await fs.copyFile(path.join(directory,file),path.join(folder,file));
  const executable=path.join(appFolder,'Sheet Studio.exe');await fs.writeFile(path.join(appFolder,'Uninstall Sheet Studio.exe'),'disposable marker');
  const env={...process.env,SHEET_STUDIO_DATA:profile,STUDIO_DISABLE_UPDATES:'1',SHEET_UPDATE_FIXTURE_ROOT:folder,SHEET_UPDATE_FIXTURE_EXE:executable};
  await launch(executable,env);
  const bytes=await fs.readFile(path.join(folder,'installer.exe')),sha256=createHash('sha256').update(bytes).digest('hex');
  await application.evaluate(({app,dialog,shell,ipcMain},{bytes,sha256,mode,folder})=>{
   const fs=process.getBuiltinModule('fs'),path=process.getBuiltinModule('path');
   process.env.SHEET_UPDATE_FIXTURE_PID=String(process.pid);
   const data=Buffer.from(bytes),item={tag_name:'v9.9.9',draft:false,prerelease:false,assets:[{name:'Sheet-Studio-9.9.9-Windows-x64-Setup.exe',state:'uploaded',size:data.length,digest:'sha256:'+sha256,browser_download_url:'https://github.com/beaur8-prog/sheet-studio/releases/download/v9.9.9/Sheet-Studio-9.9.9-Windows-x64-Setup.exe'}]};
   globalThis.fetch=async url=>url.startsWith('https://api.github.com/')?Response.json(item):new Response(data);
   globalThis.updateCloseFixture={errors:[]};
   dialog.showMessageBox=async(...args)=>({response:args.at(-1).type==='warning'?1:0});
   dialog.showErrorBox=(title,message)=>updateCloseFixture.errors.push({title,message});shell.showItemInFolder=()=>{};
   app.on('before-quit',()=>{const cache=path.join(app.getPath('userData'),'updates',sha256);fs.writeFileSync(path.join(folder,'before-quit.json'),JSON.stringify({helperAcknowledged:fs.existsSync(cache)&&fs.readdirSync(cache).some(n=>n.endsWith('.ready'))}));});
   if(mode==='save-failure'){
    const original=ipcMain.listeners('sheet:close-ready')[0];ipcMain.removeListener('sheet:close-ready',original);let fail=true;
    ipcMain.on('sheet:close-ready',(event,result)=>{if(fail){fail=false;original(event,{...result,error:'fixture recovery save failure'});}else original(event,result);});
   }
  },{bytes:[...bytes],sha256,mode,folder});
  await edit('A1','업데이트 후 복구할 문서');await edit('A2','=1+2');
  console.log('Downloading inert test update:',mode);
  await page.evaluate(()=>desktop.updateAction('download'));
  assert.equal((await page.evaluate(()=>desktop.getUpdateState())).phase,'ready');
  await edit('A3','종료 직전 입력');
  console.log('Requesting native update handoff:',mode);
  const closed=mode==='install'?application.waitForEvent('close',{timeout:45000}):null;
  await page.evaluate(()=>{void desktop.updateAction('install').catch(()=>{});});
  if(mode==='install'){
   await closed;console.log('Real Sheet quit completed');application=null;await until(()=>exists(path.join(folder,'reopened')));
   assert(await exists(path.join(folder,'installed')));
   assert.equal(JSON.parse(await fs.readFile(path.join(folder,'before-quit.json'),'utf8')).helperAcknowledged,true);
   results.push({mode,realSheetClose:true,helperAcknowledgedBeforeQuit:true,installed:true,reopened:true});
  }else{
   await until(async()=>await application.evaluate(()=>updateCloseFixture.errors.length>0));
   assert.match((await application.evaluate(()=>updateCloseFixture.errors[0])).message,/fixture recovery save failure/);
   await page.waitForFunction(async()=>(await desktop.getUpdateState()).phase==='ready');
   assert.equal(await page.evaluate(()=>sheetStudio.getCell('A3').raw),'종료 직전 입력');
   assert.equal(await exists(path.join(folder,'installed')),false);
   await application.close();application=null;await delay(1500);
   assert.equal(await exists(path.join(folder,'installed')),false,'later quit must not install a cancelled update');
   results.push({mode,realSheetCloseGuard:true,editorKeptOpen:true,cancelledHelper:true,noInstallationAfterLaterQuit:true});
  }
  await fs.copyFile(path.join(profile,'updates/last-install.log'),path.join(report,mode+'-install.log'));
  await fs.copyFile(path.join(profile,'updates/helper-startup.log'),path.join(report,mode+'-startup.log'));
  await launch(original,env);await assertBlankStart(page);await recoverPrevious(page);
  assert.equal(await page.evaluate(()=>sheetStudio.getCell('A1').raw),'업데이트 후 복구할 문서');
  assert.equal(await page.evaluate(()=>sheetStudio.evaluateAsync('A2')),3);
  assert.equal(await page.evaluate(()=>sheetStudio.getCell('A3').raw),'종료 직전 입력');
  results.at(-1).documentRecovered=true;results.at(-1).formulaResult=3;
  await application.close();application=null;
  console.log('PASS packaged Sheet update close:',mode);
 }
 await fs.writeFile(path.join(report,'results.json'),JSON.stringify({ok:true,version,results},null,2));
}catch(error){
 console.error('Original native update fixture failure:',error.stack||error);
 if(application){
  try{console.error('Editor state:',await application.evaluate(()=>globalThis.updateCloseFixture));}catch{}
  try{console.error('Updater state:',await page.evaluate(()=>desktop.getUpdateState()));}catch{}
 }
 for(const mode of ['install','save-failure']){
  const updates=path.join(directory,mode,"profile 한글 O'Brien; &",'updates');
  await fs.cp(updates,path.join(report,mode+'-failure-updates'),{recursive:true}).catch(()=>{});
 }
 throw error;
}finally{
 if(application)await application.close().catch(error=>console.error('Fixture close:',error.message));
 await fs.rm(directory,{recursive:true,force:true,maxRetries:10,retryDelay:200}).catch(error=>console.error('Fixture cleanup:',error.message));
}
