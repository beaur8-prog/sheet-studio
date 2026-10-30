import test from 'node:test';
import assert from 'node:assert/strict';
import {createWorkbook} from '../src/engine.mjs';
import {dataEdge,selectionStep} from '../src/cell-navigation.mjs';
test('data-edge navigation stops at the current contiguous block, next data or worksheet boundary instead of unrelated cells',()=>{
 const s=createWorkbook().sheets[0];s.cells={A1:{raw:'1'},A2:{raw:'2'},A3:{raw:'3'},A10:{raw:'10'},F50:{raw:'elsewhere'},A5:{raw:'',style:{bold:true}}};
 assert.deepEqual(dataEdge(s,{r:0,c:0},'r',1),{r:2,c:0});assert.deepEqual(dataEdge(s,{r:2,c:0},'r',1),{r:9,c:0});assert.deepEqual(dataEdge(s,{r:9,c:0},'r',1),{r:1048575,c:0});assert.deepEqual(dataEdge(s,{r:8,c:0},'r',-1),{r:2,c:0});assert.deepEqual(dataEdge(s,{r:2,c:0},'r',-1),{r:0,c:0});assert.deepEqual(dataEdge(s,{r:1,c:1},'c',1),{r:1,c:16383});
});
test('formula-empty cells and declared array followers remain occupied and hidden targets are skipped',()=>{
 const s=createWorkbook().sheets[0];s.cells={A1:{raw:'1'},A2:{raw:'=""'},A3:{raw:'3'},B1:{raw:'=A1:A3*2',xlsxFormula:{raw:'=A1:A3*2',formula:'A1:A3*2',shareType:'array',ref:'B1:B3'}}};assert.deepEqual(dataEdge(s,{r:0,c:0},'r',1),{r:2,c:0});assert.deepEqual(dataEdge(s,{r:0,c:1},'r',1),{r:2,c:1});assert.deepEqual(dataEdge(s,{r:0,c:0},'r',1,{hiddenRow:r=>r===2}),{r:1,c:0});
});
test('Tab/Enter traverse and wrap the selected rectangle in either direction with hidden rows and columns excluded',()=>{
 const b={r1:1,r2:3,c1:1,c2:3},opts={hiddenRow:r=>r===2,hiddenCol:c=>c===2};let p={r:1,c:1},seen=[];for(let i=0;i<4;i++){p=selectionStep(b,p,opts);seen.push(p);}assert.deepEqual(seen,[{r:1,c:3},{r:3,c:1},{r:3,c:3},{r:1,c:1}]);assert.deepEqual(selectionStep(b,p,{...opts,direction:-1}),{r:3,c:3});assert.deepEqual(selectionStep(b,p,{...opts,axis:'r'}),{r:3,c:1});assert.deepEqual(selectionStep(b,p,{...opts,axis:'r',direction:-1}),{r:3,c:3});assert.deepEqual(selectionStep(b,p,{hiddenRow:()=>true}),p);
});
test('merged followers are skipped without traversing a million covered cells',()=>{
 const b={r1:0,r2:1048575,c1:0,c2:1},merges=[{r1:0,r2:1048574,c1:0,c2:1}],p={r:0,c:0};assert.deepEqual(selectionStep(b,p,{merges}),{r:1048575,c:0});assert.deepEqual(selectionStep(b,p,{axis:'r',merges}),{r:1048575,c:0});const small={r1:0,r2:1,c1:0,c2:2};assert.deepEqual(selectionStep(small,p,{merges:[{r1:0,r2:0,c1:0,c2:1}]}),{r:0,c:2});
});
