import {desktopLaunch} from './electron-launch.mjs';
import {_electron} from 'playwright';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import path from 'node:path';
import {createWorkbook} from '../src/engine.mjs';
const root=path.resolve(import.meta.dirname,'..'),profile=path.join(root,'.package','zoom-'+Date.now()),env={...process.env,SHEET_STUDIO_DATA:profile};delete env.ELECTRON_RUN_AS_NODE;
const app=await _electron.launch(desktopLaunch(root,env)),page=await app.firstWindow(),errors=[];page.on('pageerror',e=>errors.push(e.message));page.setDefaultTimeout(12000);
const idle=()=>page.evaluate(()=>sheetStudio.whenIdle());
const scale=()=>page.locator('#zoom-label').innerText().then(s=>parseInt(s));
const snapshot=()=>page.evaluate(()=>JSON.stringify({book:sheetStudio.getWorkbook(),history:sheetStudio.historyStats(),dirty:sheetStudio.isDirty(),selection:sheetStudio.getSelection()}));
try{
 await page.waitForFunction(()=>!!window.sheetStudio);await idle();
 const b=createWorkbook();b.name='Native desktop zoom';b.sheets[0].cells={A1:{raw:'10'},B1:{raw:'20'},C3:{raw:'=SUM(A1:B1)'}};
 await page.locator('#file-input').setInputFiles({name:'native-zoom.sheet.json',mimeType:'application/json',buffer:Buffer.from(JSON.stringify(b))});await page.waitForFunction(()=>sheetStudio.getWorkbook().name==='Native desktop zoom');await idle();
 const original=await snapshot(),toolbar=await page.locator('.ribbon').boundingBox(),grid=await page.locator('#grid').boundingBox(),point={x:grid.x+350,y:grid.y+215};
 await page.evaluate(()=>{window.nativeWheels=[];window.addEventListener('wheel',e=>nativeWheels.push({ctrl:e.ctrlKey,trusted:e.isTrusted,prevented:e.defaultPrevented}),{passive:true});});
 await page.mouse.move(point.x,point.y);await page.keyboard.down('Control');await page.mouse.wheel(0,-100);await page.keyboard.up('Control');await idle();assert.equal(await scale(),110);
 const cdp=await page.context().newCDPSession(page);await cdp.send('Input.synthesizePinchGesture',{...point,scaleFactor:1.2,relativeSpeed:800,gestureSourceType:'mouse'});await idle();assert.ok(await scale()>110);
 assert.equal(await app.evaluate(({BrowserWindow})=>BrowserWindow.getAllWindows()[0].webContents.getZoomFactor()),1);assert.equal(await page.evaluate(()=>visualViewport.scale),1);assert.equal((await page.locator('.ribbon').boundingBox()).height,toolbar.height);assert.equal(await snapshot(),original);assert.ok(await page.evaluate(()=>nativeWheels.some(e=>e.ctrl&&e.trusted&&e.prevented)));
 await fs.mkdir(path.join(root,'test-results/zoom-gestures'),{recursive:true});await page.screenshot({path:path.join(root,'test-results/zoom-gestures/electron-pinch.png')});assert.deepEqual(errors,[]);
 console.log('PASS actual isolated Electron: trusted Ctrl+wheel and touchpad pinch, sheet scale changes, webContents/page scale stays 1, toolbar/document/selection/history unchanged');
}catch(error){await fs.mkdir(path.join(root,'test-results/zoom-gestures'),{recursive:true});await page.screenshot({path:path.join(root,'test-results/zoom-gestures/electron-failure.png')});throw error;}finally{await app.close();await fs.rm(profile,{recursive:true,force:true});}
