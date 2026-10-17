import test from 'node:test';
import assert from 'node:assert/strict';
import path from 'node:path';
import platform from '../scripts/platform.cjs';
import menu from '../electron/menu.cjs';

test('build selects native Mac architecture and supports separate Intel and Apple Silicon output',()=>{
  assert.deepEqual(platform.buildTarget([], 'darwin','arm64'),{platform:'darwin',arch:'arm64',directory:false});
  assert.equal(platform.buildTarget(['--mac','--arch=x64'],'darwin','arm64').arch,'x64');
  assert.equal(platform.buildTarget(['--mac','--arm64','--dir'],'darwin','x64').directory,true);
  assert.equal(platform.buildTarget([], 'win32','x64').platform,'win32');
  assert.throws(()=>platform.buildTarget(['--mac'],'win32','x64'),/on macOS/);
  assert.throws(()=>platform.buildTarget(['--x64','--arm64'],'darwin','arm64'),/one architecture/);
  assert.throws(()=>platform.buildTarget(['--win','--arm64'],'win32','x64'),/Supported targets/);
  assert.throws(()=>platform.buildTarget(['--arch=typo'],'darwin','arm64'),/Unknown/);
  for(const arch of ['arm64','x64'])assert.ok(platform.packagedExecutable('/project','0.3.1','darwin',arch).endsWith(path.join('Sheet Studio.app','Contents','MacOS','Sheet Studio')));
  assert.notEqual(platform.outputDir('/project','0.3.1','darwin','arm64'),platform.outputDir('/project','0.3.1','darwin','x64'));
});

test('Mac menu routes workbook commands through the renderer and keeps system window roles',()=>{
  const calls=[],items=menu.applicationMenu('darwin',action=>calls.push(action));
  const commands=items.flatMap(item=>item.submenu||[]).filter(item=>item.click);
  for(const command of commands)command.click();
  assert.deepEqual(calls,['new','open','save','undo','redo','cut','copy','paste','select-all','find']);
  assert.ok(commands.every(item=>item.accelerator.startsWith('CmdOrCtrl+')));
  assert.ok(items.some(item=>item.role==='windowMenu'));
  assert.equal(menu.applicationMenu('win32',()=>{}),null);
});
