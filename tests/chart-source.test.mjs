import test from 'node:test';
import assert from 'node:assert/strict';
import {createWorkbook} from '../src/engine.mjs';
import {chartSourceFormulas} from '../src/chart-model.mjs';

test('imported chart source descriptions use all actual references rather than the fallback rectangle',()=>{
 const b=createWorkbook(),s=b.sheets[0];s.name="Sales '26";
 const ref=range=>({sheetId:s.id,range,formula:'Sheet1!'+range});
 const chart={range:'A1:B2',source:{formulas:[ref('B1:B1'),ref('A2:A5'),ref('B2:B5'),ref('A2:A5'),{formula:'RevenueName'}]}};
 assert.deepEqual(chartSourceFormulas(chart,b),["'Sales ''26'!$B$1:$B$1","'Sales ''26'!$A$2:$A$5","'Sales ''26'!$B$2:$B$5",'RevenueName']);
 assert.equal(chart.range,'A1:B2');
});
test('source descriptions distinguish cache-only series and accept series references in older app documents',()=>{
 const b=createWorkbook(),s=b.sheets[0];
 assert.deepEqual(chartSourceFormulas({source:{formulas:[]},series:[{valRef:{cache:[1,2,3]}}]},b),[]);
 assert.deepEqual(chartSourceFormulas({source:{},series:[{valRef:{sheetId:s.id,range:'C2:C4'},catRef:{formula:'NamedCategories'}}]},b),['NamedCategories',"'Sheet1'!$C$2:$C$4"]);
});
