import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import path from 'node:path';
import os from 'node:os';
import {randomUUID} from 'node:crypto';
import {createFileLocations} from '../electron/file-locations.cjs';

test('feature-loss copies protect the original path, aliases and restart recovery without changing normal directory behavior',async()=>{
  const dir=await fs.mkdtemp(path.join(os.tmpdir(),'sheet-copy-'));try{
    const source=path.join(dir,'원본.xlsx'),hard=path.join(dir,'hard.xlsx'),link=path.join(dir,'linked.xlsx'),copy=path.join(dir,'원본-사본.xlsx'),id=randomUUID(),stateFile=path.join(dir,'locations.json');
    await fs.writeFile(source,'ORIGINAL');await fs.link(source,hard);
    const locations=createFileLocations({stateFile,documents:dir});await locations.confirm({token:locations.stage(source),documentId:id});
    await assert.rejects(locations.assertCopyTarget(id,source),/원본/);await assert.rejects(locations.assertCopyTarget(id,hard),/같은 파일/);
    if(process.platform!=='win32'){await fs.symlink(source,link);await assert.rejects(locations.assertCopyTarget(id,link),/원본/);}
    await locations.assertCopyTarget(id,copy);await locations.remember(id,copy);assert.equal(await locations.directory(id),dir);
    const restored=createFileLocations({stateFile,documents:dir});await assert.rejects(restored.assertCopyTarget(id,source),/원본/);await restored.assertCopyTarget(id,copy);
    await assert.rejects(restored.assertCopyTarget(randomUUID(),source,'원본.xlsx'),/다른 사본 이름/);
    assert.equal(await fs.readFile(source,'utf8'),'ORIGINAL');
  }finally{await fs.rm(dir,{recursive:true,force:true});}
});

test('when the original path is unknown or gone, feature-loss copies require a new file rather than an unverified alias',async()=>{
 const dir=await fs.mkdtemp(path.join(os.tmpdir(),'sheet-copy-unknown-'));try{
  const source=path.join(dir,'original.xlsx'),alias=path.join(dir,'different-name.xlsx'),fresh=path.join(dir,'new-copy.xlsx'),stateFile=path.join(dir,'locations.json'),locations=createFileLocations({stateFile,documents:dir}),id=randomUUID();await fs.writeFile(source,'KEEP');await fs.link(source,alias);
  await assert.rejects(locations.assertCopyTarget(id,alias,'original.xlsx'),/새 사본 이름/);await locations.assertCopyTarget(id,fresh,'original.xlsx');await locations.confirm({token:locations.stage(source),documentId:id});await fs.unlink(source);await assert.rejects(locations.assertCopyTarget(id,alias,'original.xlsx'),/새 사본 이름/);await locations.assertCopyTarget(id,fresh,'original.xlsx');assert.equal(await fs.readFile(alias,'utf8'),'KEEP');
 }finally{await fs.rm(dir,{recursive:true,force:true});}
});
