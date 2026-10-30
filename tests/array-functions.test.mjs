import test from 'node:test';
import assert from 'node:assert/strict';
import {createWorkbook,evaluateFormulaArray,evaluateFormula,formulaCanSpill} from '../src/engine.mjs';

const book=createWorkbook();
const array=(formula,b=book)=>evaluateFormulaArray(b,b.sheets[0].id,formula);
const expected=(rows,cols,values)=>({rows,cols,values});
const error=formula=>array(formula).values[0];

test('pending spill detection is conservative without calculating any source values',()=>{
 for(const f of ['=SEQUENCE(100000)','=A1:A100000','={1;2}','=ISNUMBER(A1:A5)','=NamedRange','=XLOOKUP(2,A:A,B:C)'])assert.equal(formulaCanSpill(f),true,f);
 for(const f of ['=SUM(SEQUENCE(100000))','=A1+1','="SEQUENCE(3)"','=@A:A','=TYPE(A1:A5)','not a formula'])assert.equal(formulaCanSpill(f),false,f);
});

test('SEQUENCE generates row-major horizontal, vertical and two-dimensional arrays with optional arguments',()=>{
  assert.deepEqual(array('=SEQUENCE(4)'),expected(4,1,[1,2,3,4]));
  assert.deepEqual(array('=SEQUENCE(,4)'),expected(1,4,[1,2,3,4]));
  assert.deepEqual(array('=SEQUENCE(2,3,10,-2)'),expected(2,3,[10,8,6,4,2,0]));
  assert.deepEqual(array('=SEQUENCE(2,2,,0)'),expected(2,2,[1,1,1,1]));
  assert.deepEqual(array('=SEQUENCE(2.9,1,0.1,0.2)').values.map(n=>Math.round(n*1e8)/1e8),[.1,.3]);
  assert.equal(evaluateFormula(book,'sheet-1','=SUM(SEQUENCE(4))'),10);
  for(const formula of ['=SEQUENCE()','=SEQUENCE(0)','=SEQUENCE(-1)','=SEQUENCE(1,0)','=SEQUENCE("bad")','=SEQUENCE(1,1,1,1,1)'])assert.equal(error(formula),'#VALUE!',formula);
  assert.equal(error('=SEQUENCE(1001,1000)'),'#NUM!');
});
test('FILTER preserves rows and columns, evaluates only selected values and returns lazy if_empty expressions',()=>{
  assert.deepEqual(array('=FILTER({1,10;2,20;3,30},{TRUE;FALSE;TRUE})'),expected(2,2,[1,10,3,30]));
  assert.deepEqual(array('=FILTER({1,10;2,20;3,30},{TRUE,FALSE})'),expected(3,1,[1,2,3]));
  assert.deepEqual(array('=FILTER({1,#N/A;2,3},{FALSE,TRUE})'),expected(2,1,['#N/A',3]));
  assert.deepEqual(array('=FILTER({1;#DIV/0!},{TRUE;FALSE},1/0)'),expected(1,1,[1]));
  assert.deepEqual(array('=FILTER({1;2},{0;0},"없음")'),expected(1,1,['없음']));
  assert.deepEqual(array('=FILTER({1;2},{0;0},SEQUENCE(1,2,7))'),expected(1,2,[7,8]));
  assert.equal(error('=FILTER({1;2},{0;0})'),'#CALC!');
  assert.equal(error('=FILTER({1;2},{TRUE;#N/A})'),'#N/A');
  assert.equal(error('=FILTER({1;2},{"yes";TRUE})'),'#VALUE!');
  assert.equal(error('=FILTER({1,2;3,4},{TRUE,FALSE;FALSE,TRUE})'),'#VALUE!');
  assert.equal(error('=FILTER({1;2;3},{TRUE;FALSE})'),'#VALUE!');
});
test('SORT is stable, supports column sorting and composes with FILTER and LET',()=>{
  assert.deepEqual(array('=SORT({"b",2;"a",1;"c",1},2)'),expected(3,2,['a',1,'c',1,'b',2]));
  assert.deepEqual(array('=SORT({3,1,2;30,10,20},1,-1,TRUE)'),expected(2,3,[3,2,1,30,20,10]));
  assert.deepEqual(array('=LET(data,{3;1;2},SORT(FILTER(data,data>1)))'),expected(2,1,[2,3]));
  assert.deepEqual(array('=SORT({2;1},, -1)'),expected(2,1,[2,1]));
  assert.equal(error('=SORT({1;2},2)'),'#VALUE!');
  assert.equal(error('=SORT({1;2},1,0)'),'#VALUE!');
  assert.equal(error('=SORT({1;2},0)'),'#VALUE!');
});
test('UNIQUE compares complete rows/columns and preserves data types, errors and first occurrence order',()=>{
  assert.deepEqual(array('=UNIQUE({"A",1;"a",1;"B",1;"B",2})'),expected(3,2,['A',1,'B',1,'B',2]));
  assert.deepEqual(array('=UNIQUE({1,1,2;3,3,4},TRUE)'),expected(2,2,[1,2,3,4]));
  assert.deepEqual(array('=UNIQUE({1;1;2;3;3},,TRUE)'),expected(1,1,[2]));
  assert.deepEqual(array('=UNIQUE({1;"1";TRUE;#N/A;#N/A})'),expected(4,1,[1,'1',true,'#N/A']));
  assert.deepEqual(array('=SORT(UNIQUE({3;1;2;1;3}))'),expected(3,1,[1,2,3]));
  assert.equal(error('=UNIQUE({1;1},,TRUE)'),'#CALC!');
});

test('XLOOKUP returns complete matching rows or columns and a lazy array fallback',()=>{
 assert.deepEqual(array('=XLOOKUP(2,{1;2;3},{10,11;20,21;30,31})'),expected(1,2,[20,21]));
 assert.deepEqual(array('=XLOOKUP(2,{1,2,3},{10,20,30;11,21,31})'),expected(2,1,[20,21]));
 assert.deepEqual(array('=XLOOKUP(2,{1;2},{#N/A,11;20,#DIV/0!})'),expected(1,2,[20,'#DIV/0!']));
 assert.deepEqual(array('=XLOOKUP(4,{1;2},{10,11;20,21},SEQUENCE(2,2))'),expected(2,2,[1,2,3,4]));
 assert.deepEqual(array('=XLOOKUP(2,{1;2},{10,11;20,21},1/0)'),expected(1,2,[20,21]));
 const b=createWorkbook(),s=b.sheets[0];s.cells={A1:{raw:'1'},A1048576:{raw:'2'},B1048576:{raw:'last'},C1048576:{raw:'42'}};
 assert.deepEqual(array('=XLOOKUP(2,A:A,B:C)',b),expected(1,2,['last',42]));
 assert.equal(error('=XLOOKUP(2,{1;2},{10,11,12})'),'#VALUE!');
});

test('INDEX omitted axes return complete rows or columns while one-dimensional shorthand stays scalar',()=>{
 assert.deepEqual(array('=INDEX({1,2;3,4},2)'),expected(1,2,[3,4]));assert.deepEqual(array('=INDEX({1,2;3,4},,2)'),expected(2,1,[2,4]));assert.deepEqual(array('=INDEX({1,2;3,4},2,)'),expected(1,2,[3,4]));assert.deepEqual(array('=INDEX({1,2,3},2)'),expected(1,1,[2]));assert.deepEqual(array('=INDEX({1;2;3},2)'),expected(1,1,[2]));
 const b=createWorkbook(),s=b.sheets[0];s.cells={A1:{raw:'1'},B1:{raw:'2'},A2:{raw:'3'},B2:{raw:'4'},D1:{raw:'=INDEX(A1:B2,2)'},G1:{raw:'=INDEX(A1:B2,,2)'}};assert.deepEqual(array('=INDEX(A1:B2,2)',b),expected(1,2,[3,4]));
});

test('information predicates and NOT operate elementwise so FILTER can use typed and text-search masks',()=>{
 assert.deepEqual(array('=ISNUMBER({1,"2";TRUE,#N/A})'),expected(2,2,[true,false,false,false]));
 assert.deepEqual(array('=ISLOGICAL({TRUE;FALSE;1;#N/A})'),expected(4,1,[true,true,false,false]));
 assert.deepEqual(array('=ISERROR({1;#N/A;#DIV/0!})'),expected(3,1,[false,true,true]));
 assert.deepEqual(array('=ISERR({1;#N/A;#DIV/0!})'),expected(3,1,[false,false,true]));
 assert.deepEqual(array('=ISNA({1;#N/A;#DIV/0!})'),expected(3,1,[false,true,false]));
 assert.deepEqual(array('=ISTEXT({1;"1";"";TRUE;#N/A})'),expected(5,1,[false,true,true,false,false]));
 assert.deepEqual(array('=ISNONTEXT({1;"1";"";TRUE;#N/A})'),expected(5,1,[true,false,false,true,true]));
 assert.deepEqual(array('=ERROR.TYPE({#N/A,#DIV/0!;1,"x"})'),expected(2,2,[7,2,'#N/A','#N/A']));
 assert.deepEqual(array('=NOT({0,1;TRUE,FALSE})'),expected(2,2,[true,false,false,true]));
 assert.deepEqual(array('=N({TRUE,FALSE;"2",#N/A})'),expected(2,2,[1,0,0,'#N/A']));
 assert.deepEqual(array('=T({TRUE,2;"2",#N/A})'),expected(2,2,['','','2','#N/A']));
 assert.deepEqual(array('=FILTER({1;"1";TRUE;#N/A},ISNUMBER({1;"1";TRUE;#N/A}))'),expected(1,1,[1]));
 assert.deepEqual(array('=FILTER({"apple",1;"pear",2;"APPLE",3},ISNUMBER(SEARCH("apple",{"apple";"pear";"APPLE"})))'),expected(2,2,['apple',1,'APPLE',3]));
 assert.deepEqual(array('=FILTER({1;#N/A;3},NOT(ISERROR({1;#N/A;3})))'),expected(2,1,[1,3]));
 const b=createWorkbook(),s=b.sheets[0];s.cells={A1:{raw:"'"},A2:{raw:''},A3:{raw:'0'}};
 assert.deepEqual(array('=ISBLANK(A1:A3)',b),expected(3,1,[false,true,false]));
 // Reference and whole-array inspection are deliberately not scalar lifts.
 assert.equal(evaluateFormula(b,s.id,'=TYPE({1;2})'),64);assert.equal(evaluateFormula(b,s.id,'=ISREF(A1:A3)'),true);
});
