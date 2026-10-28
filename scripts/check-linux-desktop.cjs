'use strict';
// Exercise the installed package on a disposable Linux CI desktop.
const assert=require('node:assert/strict');
const fs=require('node:fs/promises');
const path=require('node:path');
const os=require('node:os');
const {execFileSync}=require('node:child_process');
const {_electron}=require('playwright');
const asar=require('@electron/asar');
const manifest=require('../package.json');
const {fileAssociations}=require('../electron/file-open-requests.cjs');

async function check(){
  assert.equal(process.platform,'linux');
  const executable=process.env.SHEET_STUDIO_EXECUTABLE;
  assert.ok(executable,'Set SHEET_STUDIO_EXECUTABLE to the installed application');
  const profile=await fs.mkdtemp(path.join(os.tmpdir(),'sheet-desktop-'));
  const output=path.resolve('test-results/linux-desktop');await fs.mkdir(output,{recursive:true});
  const env={...process.env,SHEET_STUDIO_DATA:profile};delete env.ELECTRON_RUN_AS_NODE;
  let app;
  try{
    if(process.argv.includes('--wayland')){
      assert.ok(env.WAYLAND_DISPLAY,'A disposable Wayland compositor is required');
      const logFile=path.join(output,'wayland-protocol.log'),wrapper=path.join(profile,'launch-wayland');
      const quote=value=>"'"+value.replaceAll("'","'\\''")+"'";
      // Capture stderr from process start, including messages before Playwright connects.
      await fs.writeFile(wrapper,`#!/bin/bash\nexec ${quote(executable)} "$@" 2> >(tee ${quote(logFile)} >&2)\n`,{mode:0o700});
      app=await _electron.launch({executablePath:wrapper,args:['--ozone-platform=wayland','--disable-gpu'],env:{...env,WAYLAND_DEBUG:'1'},timeout:15000});
      console.log('Wayland: connected to the installed Electron app');
      // A headless compositor can defer the first hidden-window paint. Map the
      // actual packaged window explicitly; do not alter its application ID.
      const createdDeadline=Date.now()+10000;
      while(!await app.evaluate(({BrowserWindow})=>BrowserWindow.getAllWindows().length)&&Date.now()<createdDeadline)await new Promise(r=>setTimeout(r,100));
      await app.evaluate(({BrowserWindow})=>BrowserWindow.getAllWindows()[0].show());
      let log='';const deadline=Date.now()+10000;
      while(!/set_app_id\("[^"]+"\)/.test(log)&&Date.now()<deadline){log=await fs.readFile(logFile,'utf8');if(!/set_app_id\("[^"]+"\)/.test(log))await new Promise(r=>setTimeout(r,100));}
      assert.equal(app.process().exitCode,null,'The Wayland app must remain running');
      const ids=[...log.matchAll(/set_app_id\("([^"]+)"\)/g)].map(m=>m[1]);
      assert.ok(ids.length,'The real Wayland window must announce an app_id');
      assert.ok(ids.every(id=>id===manifest.desktopName),JSON.stringify(ids));
      await fs.writeFile(path.join(output,'wayland.json'),JSON.stringify({version:manifest.version,appIds:ids,desktopFile:manifest.desktopName+'.desktop'},null,2)+'\n');
      console.log('PASS: native Wayland app_id matches the installed desktop file:',ids.join(', '));
      return;
    }
    const desktopFile=`/usr/share/applications/${manifest.desktopName}.desktop`;
    const archive=path.join(path.dirname(executable),'resources/app.asar');
    const packaged=JSON.parse(asar.extractFile(archive,'package.json'));
    assert.equal(packaged.version,manifest.version);assert.equal(packaged.desktopName,manifest.desktopName);
    for(const entry of asar.listPackage(archive)){
      const name=entry.replace(/^\//,'');if(!/^(dist|electron)\//.test(name))continue;
      const source=path.resolve(name),stat=await fs.stat(source).catch(()=>null);if(!stat?.isFile())continue;
      assert.ok((await fs.readFile(source)).equals(asar.extractFile(archive,name)),`Candidate differs from current source: ${name}`);
    }
    execFileSync('desktop-file-validate',[desktopFile]);
    const desktop=await fs.readFile(desktopFile,'utf8');
    const entries=Object.fromEntries(desktop.split('\n').filter(line=>line.includes('=')).map(line=>[line.slice(0,line.indexOf('=')),line.slice(line.indexOf('=')+1)]));
    assert.equal(entries.Name,'Sheet Studio');assert.equal(entries.StartupWMClass,manifest.desktopName);assert.equal(entries.Icon,'sheet-studio');
    assert.equal(entries.Type,'Application');assert.equal(entries.Terminal,'false');assert.ok(entries.Exec.includes(executable));
    assert.match(entries.Exec,/%[fFuU]/,'Desktop launch command forwards the selected file');
    for(const type of fileAssociations)assert.ok(entries.MimeType.split(';').includes(type.mimeType),`Missing supported MIME type: ${type.mimeType}`);
    for(const size of [16,24,32,48,64,128,256,512,1024]){
      const png=await fs.readFile(`/usr/share/icons/hicolor/${size}x${size}/apps/${entries.Icon}.png`);
      assert.equal(png.readUInt32BE(16),size);assert.equal(png.readUInt32BE(20),size);
    }
    // Query the same installed desktop entry and theme lookup used by GTK desktops.
    const icons=JSON.parse(execFileSync('/usr/bin/python3',['-c',`
import gi,json
gi.require_version('Gtk','3.0')
from gi.repository import Gio,Gtk
app=Gio.DesktopAppInfo.new('${manifest.desktopName}.desktop')
assert app is not None and app.should_show()
for mime in ${JSON.stringify(fileAssociations.map(type=>type.mimeType))}:
 assert any(candidate.get_id() == app.get_id() for candidate in Gio.AppInfo.get_all_for_type(mime)), mime
theme=Gtk.IconTheme.get_default()
icons={}
for size in [16,32,48,64,128,256,512]:
 info=theme.lookup_icon(app.get_icon().to_string(),size,0)
 assert info is not None, size
 icons[size]=info.get_filename()
print(json.dumps(icons))
`],{env,encoding:'utf8'}));
    assert.ok(Object.values(icons).every(filename=>filename.includes('/sheet-studio.png')));
    app=await _electron.launch({executablePath:executable,args:['--ozone-platform=x11'],env});
    const page=await app.firstWindow();await page.locator('.cell').first().waitFor();
    const windowId=await app.evaluate(({BrowserWindow})=>'0x'+BrowserWindow.getAllWindows()[0].getNativeWindowHandle().readUInt32LE(0).toString(16));
    const wmClass=execFileSync('xprop',['-id',windowId,'WM_CLASS'],{encoding:'utf8'}).trim();
    const actualClasses=[...wmClass.matchAll(/"([^"]+)"/g)].map(m=>m[1]);
    assert.ok(actualClasses.includes(entries.StartupWMClass),wmClass);
    await page.screenshot({path:path.join(output,'installed-x11.png')});
    await fs.writeFile(path.join(output,'x11.json'),JSON.stringify({version:manifest.version,desktopFile,entries,wmClass,icons},null,2)+'\n');
    console.log('PASS: installed desktop entry, themed icons at common sizes, and actual X11 WM_CLASS:',wmClass);
  }finally{
    if(app&&process.argv.includes('--wayland')){
      // This app has a disposable profile. Bound cleanup even if its hidden
      // renderer never completes a paint on the headless test compositor.
      const child=app.process();child.kill('SIGTERM');
      if(child.exitCode===null)await new Promise(resolve=>{const timer=setTimeout(()=>{child.kill('SIGKILL');resolve();},3000);child.once('exit',()=>{clearTimeout(timer);resolve();});});
    }else if(app)await app.close();
    await fs.rm(profile,{recursive:true,force:true});
  }
}
check().catch(error=>{console.error(error);process.exitCode=1;});
