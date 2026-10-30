import {_electron} from 'playwright';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import path from 'node:path';
import {createRequire} from 'node:module';
import platform from '../scripts/platform.cjs';

const root=path.resolve(import.meta.dirname,'..'),packaged=process.argv.includes('--packaged');
const version=createRequire(import.meta.url)('../package.json').version;
await fs.mkdir(path.join(root,'.package'),{recursive:true});
const fixture=await fs.mkdtemp(path.join(root,'.package/user-journeys-'));
const env={...process.env,SHEET_STUDIO_DATA:path.join(fixture,'profile')};delete env.ELECTRON_RUN_AS_NODE;
const executablePath=process.env.SHEET_STUDIO_EXECUTABLE||(packaged?platform.packagedExecutable(root,version):path.join(platform.findRuntime(root),'dist',platform.runtimeSuffix()));
let app,page;const errors=[],checks=[];
async function launch(){
 app=await _electron.launch({executablePath,args:packaged?[]:[root],env});page=await app.firstWindow();page.setDefaultTimeout(20000);
 await app.evaluate(({ipcMain})=>{globalThis.journeyClipboard={text:'',html:''};ipcMain.removeHandler('sheet:write-clipboard');ipcMain.handle('sheet:write-clipboard',(_event,payload)=>{globalThis.journeyClipboard=payload;});});
 page.on('pageerror',e=>errors.push(e.message));await page.locator('.cell').first().waitFor();
}
const settle=async()=>{await page.locator('.workbook-busy').waitFor({state:'detached'});await page.evaluate(()=>sheetStudio.whenIdle());};
const go=async ref=>{await page.locator('#name-box').fill(ref);await page.locator('#name-box').press('Enter');await settle();};
const book=()=>page.evaluate(()=>sheetStudio.getWorkbook());
const cells=()=>page.evaluate(()=>sheetStudio.getWorkbook().sheets.find(s=>s.id===sheetStudio.getWorkbook().activeSheetId).cells);
const enter=async(ref,value,commit=true)=>{await go(ref);await page.keyboard.type(value);if(commit){await page.keyboard.press('Enter');await settle();}};
const home=()=>page.getByRole('button',{name:'홈',exact:true}).click();
const menu=action=>app.evaluate(({BrowserWindow},action)=>BrowserWindow.getAllWindows()[0].webContents.send('sheet:menu-action',action),action);
const record=label=>{checks.push(label);console.log('PASS:',label);};
async function fresh(){await page.getByRole('button',{name:'파일',exact:true}).click();await page.locator('[data-action=new]').click();await page.getByRole('button',{name:'새로 만들기',exact:true}).click();await home();await settle();}
async function clipboardPaste(payload){await page.evaluate(payload=>{const data=new DataTransfer();data.setData('text/plain',payload.text);data.setData('text/html',payload.html||'');document.querySelector('#grid').dispatchEvent(new ClipboardEvent('paste',{bubbles:true,cancelable:true,clipboardData:data}));},payload);}
async function externalPaste(ref,text){await go(ref);await app.evaluate(({},text)=>{globalThis.journeyClipboard={text,html:''};},text);await clipboardPaste({text});await settle();}
async function copy(ref,cut=false){await go(ref);await home();await page.getByRole('button',{name:cut?'잘라내기':'복사',exact:true}).click();await settle();}
async function paste(ref){await go(ref);await clipboardPaste(await app.evaluate(()=>globalThis.journeyClipboard));await settle();}
async function reversible(label,operation){await settle();const before=await book();await operation();await settle();const after=await book();assert.notDeepEqual(after,before,label+' changes data');await page.locator('#undo').click();await settle();assert.deepEqual(await book(),before,label+' undo');await page.locator('#redo').click();await settle();assert.deepEqual(await book(),after,label+' redo');record(label);}
async function save(file,format='xlsx'){
 await app.evaluate(({dialog},file)=>{dialog.showSaveDialog=async()=>({canceled:false,filePath:file});},file);
 await page.evaluate(()=>document.querySelector('.toast')?.remove());
 if(format==='xlsx')await page.locator('#save').click();else{await page.getByRole('button',{name:'파일',exact:true}).click();await page.locator('[data-action=save-'+format+']').click();}
 await page.locator('.toast').filter({hasText:'파일을 저장했습니다.'}).waitFor();await settle();assert.ok((await fs.stat(file)).size>0);
}
async function open(file){
 await app.evaluate(({dialog},file)=>{dialog.showOpenDialog=async()=>({canceled:false,filePaths:[file]});},file);
 await page.evaluate(()=>document.querySelector('.toast')?.remove());await menu('open');await page.locator('.toast').filter({hasText:'파일을 열었습니다.'}).waitFor();await settle();
}
async function verifyInvoice(){
 const b=await book(),s=b.sheets.find(s=>s.name==='Sheet1');assert.ok(s);assert.equal(b.sheets.length,2);
 for(const [ref,raw] of Object.entries({A1:'항목',A2:'연필',B2:'2',C2:'1200',D2:'=B2*C2',A3:'공책',B3:'3',C3:'3500',D3:'=B3*C3',D4:'=SUM(D2:D3)',A6:'월간 정산'}))assert.equal(s.cells[ref]?.raw,raw,ref);
 assert.ok(s.cells.A1.style.bold);assert.equal(s.cells.A1.style.borders.top.style,'thin');assert.deepEqual(s.merges,['A6:D6']);
 assert.equal(b.sheets.find(s=>s.name==='Sheet2').cells.A1.raw,'=SUM(Sheet1!D2:D3)');
 assert.equal(await page.evaluate(()=>sheetStudio.evaluateAsync('A1')),12900);
}
try{
 await launch();await fresh();
 await reversible('paste a Korean invoice with numbers and formulas',()=>externalPaste('A1','항목\t수량\t단가\t금액\n연필\t2\t1200\t=B2*C2\n공책\t3\t3500\t=B3*C3'));
 await go('A1:D1');await reversible('format a selected header and undo/redo',()=>page.getByRole('button',{name:'굵게',exact:true}).click());
 await go('A1:D3');await reversible('apply all borders as one undo step',async()=>{await page.getByRole('button',{name:'테두리',exact:true}).click();await page.locator('[data-border-preset=all]').click();await page.getByRole('button',{name:'적용',exact:true}).click();});
 await enter('A6','월간 정산');await go('A6:D6');await reversible('merge and restore a title range',()=>page.getByRole('button',{name:'병합하고 가운데 맞춤',exact:true}).click());
 await copy('D2');await reversible('paste a relative formula with styles',()=>paste('D5'));assert.equal((await cells()).D5.raw,'=B5*C5');
 await go('D5');await reversible('delete and restore a formula cell',()=>page.keyboard.press('Delete'));
 await page.getByRole('button',{name:'새 시트',exact:true}).click();await enter('A1','=SUM(Sheet1!D2:D3)');assert.equal(await page.evaluate(()=>sheetStudio.evaluateAsync('A1')),12900);
 await page.getByRole('button',{name:'Sheet1',exact:true}).click();await enter('D4','=SUM(D2:D3)',false);
 const xlsx=path.join(fixture,'사용자 검수.xlsx');await save(xlsx);assert.equal((await cells()).D4.raw,'=SUM(D2:D3)');record('saving commits the in-progress formula');
 await page.getByRole('button',{name:'Sheet2',exact:true}).click();await save(xlsx);await verifyInvoice();
 await app.close();app=null;await launch();await verifyInvoice();record('process close/restart preserves edits, formulas, borders and merged cells');
 await fresh();await open(xlsx);await verifyInvoice();record('actual XLSX save/reopen preserves two sheets and calculations');
 const json=path.join(fixture,'invoice.sheet.json');await save(json,'json');const beforeJson=await book();await fresh();await open(json);assert.deepEqual(await book(),beforeJson);record('app-document save/reopen preserves the complete model');
 await home();await enter('B10','first');await enter('B10','second');await page.locator('#undo').click();await settle();assert.equal((await cells()).B10.raw,'first');await enter('C10','new branch');assert.ok(await page.locator('#redo').isDisabled());record('editing after undo discards the old redo branch');

 // Default sheet IDs repeat in different documents; a pending cut cannot use
 // an unrelated document's sheet or delete the same address in that document.
 await fresh();await enter('A1','source');await copy('A1',true);await fresh();assert.equal(await page.locator('.copy-outline').count(),0);await enter('A1','destination');
 await reversible('cut followed by a new document becomes a safe copy',()=>paste('B1'));assert.equal((await cells()).A1.raw,'destination');assert.equal((await cells()).B1.raw,'source');
 await fresh();await enter('A1','keep');await enter('A2','move');await copy('A2',true);await go('A1');await page.getByRole('button',{name:'행 삽입',exact:true}).click();
 await reversible('cut followed by inserted rows cannot delete a shifted neighbor',()=>paste('B1'));assert.equal((await cells()).A2.raw,'keep');assert.equal((await cells()).A3.raw,'move');
 await copy('A3',true);await enter('A3','new source');await paste('C1');assert.equal((await cells()).A3.raw,'new source');assert.equal((await cells()).C1.raw,'move');record('editing the cut source preserves the newer value');
 await copy('A3',true);await page.getByRole('button',{name:'새 시트',exact:true}).click();await reversible('cut between sheets in the same document still moves cells',()=>paste('B2'));assert.equal((await book()).sheets[0].cells.A3,undefined);assert.equal((await cells()).B2.raw,'new source');
 await copy('B2',true);await open(xlsx);await home();await paste('C1');assert.equal((await cells()).A1.raw,'=SUM(Sheet1!D2:D3)');assert.equal((await cells()).C1.raw,'new source');record('cut followed by opening a workbook cannot delete its existing cells');

 // Hold only the file chooser response, keeping the real renderer export,
 // IPC and disk writer. Exercise native commands while a save is outstanding.
 await enter('C10','save snapshot');await go('C10');const stable=await book();
 await app.evaluate(({dialog})=>{globalThis.releaseJourneySave=null;dialog.showSaveDialog=()=>new Promise(resolve=>{globalThis.releaseJourneySave=resolve;});});
 await menu('save');await page.locator('.workbook-busy').waitFor();
 for(let i=0;i<100&&!await app.evaluate(()=>!!globalThis.releaseJourneySave);i++)await new Promise(r=>setTimeout(r,20));
 assert.ok(await app.evaluate(()=>!!globalThis.releaseJourneySave));
 for(const command of ['undo','redo','new','cut','bold'])await menu(command);
 await clipboardPaste({text:'must not be pasted'});
 await page.keyboard.press('ControlOrMeta+z');await page.keyboard.type('must not be typed');
 assert.deepEqual(await book(),stable);assert.equal(await page.locator('.backdrop').count(),0);
 const busyFile=path.join(fixture,'save-snapshot.sheet.json');
 await app.evaluate(({ },file)=>globalThis.releaseJourneySave({canceled:false,filePath:file}),busyFile);
 await settle();assert.deepEqual(await book(),stable);
 // The normal Save button exports XLSX, regardless of a mocked chooser suffix.
 assert.equal((await fs.readFile(busyFile)).subarray(0,2).toString(),'PK');
 record('save blocks native undo/redo, paste, typing and document changes');
 await home();await enter('C10','after save');await page.locator('#undo').click();await settle();assert.deepEqual(await book(),stable);record('editing and undo resume normally after save');
 assert.deepEqual(errors,[]);await fs.mkdir(path.join(root,'test-results'),{recursive:true});await page.screenshot({path:path.join(root,'test-results/user-journeys.png')});
 await fs.writeFile(path.join(root,'test-results/user-journeys.json'),JSON.stringify({ok:true,platform:process.platform,packaged,checks},null,2));
}catch(error){await fs.mkdir(path.join(root,'test-results'),{recursive:true});await page?.screenshot({path:path.join(root,'test-results/user-journeys-failure.png')}).catch(()=>{});throw error;}
finally{if(app)await app.close();await fs.rm(fixture,{recursive:true,force:true});}
