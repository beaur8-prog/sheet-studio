'use strict';
// Exercise the installed package on a disposable Linux CI desktop.
const assert=require('node:assert/strict');
const fs=require('node:fs/promises');
const path=require('node:path');
const os=require('node:os');
const {spawn,execFileSync}=require('node:child_process');
const {_electron}=require('playwright');
const manifest=require('../package.json');

async function check(){
  assert.equal(process.platform,'linux');
  const executable=process.env.SHEET_STUDIO_EXECUTABLE;
  assert.ok(executable,'Set SHEET_STUDIO_EXECUTABLE to the installed application');
  const profile=await fs.mkdtemp(path.join(os.tmpdir(),'sheet-desktop-'));
  const output=path.resolve('test-results/linux-desktop');await fs.mkdir(output,{recursive:true});
  const env={...process.env,SHEET_STUDIO_DATA:profile};delete env.ELECTRON_RUN_AS_NODE;
  let app,child;
  try{
    if(process.argv.includes('--wayland')){
      assert.ok(env.WAYLAND_DISPLAY,'A disposable Wayland compositor is required');
      let log='',exited=false;
      child=spawn(executable,['--ozone-platform=wayland'],{env:{...env,WAYLAND_DEBUG:'1'},stdio:['ignore','pipe','pipe']});
      child.stdout.on('data',b=>{log+=b;});child.stderr.on('data',b=>{log+=b;});child.on('exit',()=>{exited=true;});
      const deadline=Date.now()+20000;
      while(!exited&&!/set_app_id\("[^"]+"\)/.test(log)&&Date.now()<deadline)await new Promise(r=>setTimeout(r,100));
      await fs.writeFile(path.join(output,'wayland-protocol.log'),log);
      assert.equal(exited,false,'The Wayland app must remain running');
      const ids=[...log.matchAll(/set_app_id\("([^"]+)"\)/g)].map(m=>m[1]);
      assert.ok(ids.length,'The real Wayland window must announce an app_id');
      assert.ok(ids.every(id=>id===manifest.desktopName),JSON.stringify(ids));
      await fs.writeFile(path.join(output,'wayland.json'),JSON.stringify({version:manifest.version,appIds:ids,desktopFile:manifest.desktopName+'.desktop'},null,2)+'\n');
      console.log('PASS: native Wayland app_id matches the installed desktop file:',ids.join(', '));
      return;
    }
    const desktopFile=`/usr/share/applications/${manifest.desktopName}.desktop`;
    execFileSync('desktop-file-validate',[desktopFile]);
    const desktop=await fs.readFile(desktopFile,'utf8');
    const entries=Object.fromEntries(desktop.split('\n').filter(line=>line.includes('=')).map(line=>[line.slice(0,line.indexOf('=')),line.slice(line.indexOf('=')+1)]));
    assert.equal(entries.Name,'Sheet Studio');assert.equal(entries.StartupWMClass,manifest.desktopName);assert.equal(entries.Icon,'sheet-studio');
    assert.equal(entries.Type,'Application');assert.equal(entries.Terminal,'false');assert.ok(entries.Exec.includes(executable));
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
    if(app)await app.close();
    if(child&&child.exitCode===null){child.kill('SIGTERM');await new Promise(resolve=>{const timer=setTimeout(()=>{child.kill('SIGKILL');resolve();},5000);child.once('exit',()=>{clearTimeout(timer);resolve();});});}
    await fs.rm(profile,{recursive:true,force:true});
  }
}
check().catch(error=>{console.error(error);process.exitCode=1;});
