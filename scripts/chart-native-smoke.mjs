import fs from 'node:fs/promises';
import path from 'node:path';
import assert from 'node:assert/strict';
import {execFile} from 'node:child_process';
import {promisify} from 'node:util';
import {chartFixture,chartCases,addChart} from '../tests/fixtures/charts.mjs';
import {readSpreadsheet,writeSpreadsheet} from '../src/io.mjs';
import {chartData} from '../src/chart-model.mjs';
const exec=promisify(execFile),root=path.resolve(import.meta.dirname,'..'),dir=path.join(root,'.verification/charts'),ps=path.join(process.env.SystemRoot,'System32/WindowsPowerShell/v1.0/powershell.exe');
await fs.mkdir(dir,{recursive:true});
const entries=[];
for(const spec of chartCases){const book=chartFixture();addChart(book,spec);const name=spec.type+'-'+spec.grouping,file=path.join(dir,name+'.xlsx');await fs.writeFile(file,(await writeSpreadsheet(book)).data);entries.push({name,path:file,count:1});}
const source=path.join(dir,'excel-original.xlsx');
const created=JSON.parse((await exec(ps,['-NoProfile','-NonInteractive','-STA','-File',path.join(root,'scripts/create-chart-fixture.ps1'),'-Destination',source],{windowsHide:true,timeout:120000})).stdout.trim());
const bytes=await fs.readFile(source),loaded=await readSpreadsheet(bytes,'original.xlsx'),charts=loaded.sheets[0].charts;
assert.equal(charts.length,created.filter(c=>c.created).length,'All embedded Excel charts imported, including modern preservation placeholders');
loaded.sheets[0].cells.B2.raw='75';charts[0].title='Changed title';charts[0].width=700;charts[0].height=420;
assert.equal(chartData(charts[0],loaded.sheets[0],loaded)[0].values[0],75);
const out=path.join(dir,'original-roundtrip.xlsx');await fs.writeFile(out,(await writeSpreadsheet(loaded)).data);entries.push({name:'original-roundtrip',path:out,count:charts.length});
await fs.writeFile(path.join(dir,'manifest.json'),JSON.stringify(entries));
const report=JSON.parse((await exec(ps,['-NoProfile','-NonInteractive','-STA','-File',path.join(root,'scripts/check-charts.ps1'),'-Manifest',path.join(dir,'manifest.json')],{windowsHide:true,timeout:180000,maxBuffer:4*1024*1024})).stdout.trim());
const expected=[51,52,53,57,58,59,65,66,67,1,76,77,5,-4120,-4169,15,81,-4111];
for(const [i,item] of report.entries()){assert.equal(item.ok,true);const reloaded=await readSpreadsheet(await fs.readFile(item.saved),'resaved.xlsx');assert.equal(reloaded.sheets[0].charts.length,item.items.length);if(i<18){assert.equal(item.items[0].type,expected[i],item.name);assert.deepEqual(item.items[0].firstValues,[40,25,-10,35]);assert.equal(item.items[0].title,'분기 매출');}}
assert.equal(report.at(-1).items[0].title,'Changed title');assert.equal(report.at(-1).items[0].width,525);
assert.equal(report.at(-1).items[0].firstValues[0],75);
assert.deepEqual(report[17].items[0].seriesTypes,[51,65]);assert.deepEqual(report[17].items[0].axisGroups,[1,2]);
assert.deepEqual(await fs.readFile(source),bytes);
await fs.writeFile(path.join(dir,'report.json'),JSON.stringify({created,report,originalUnchanged:true},null,2));
console.log(JSON.stringify({ok:true,exportedVariants:chartCases.length,importedExcelCharts:charts.length,preservedOnly:charts.filter(c=>c.type==='unsupported').length}));
