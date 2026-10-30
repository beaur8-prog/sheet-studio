import {validatePreservedData,validatePreservedXML} from './xlsx-preservation-schema.mjs';
import {MAX_DOCUMENT_CELLS, MAX_DOCUMENT_TEXT_CHARS, MAX_FORMULA_CELLS} from './limits.mjs';
import {unwrap} from './clone.mjs';
import {validateBorders} from './borders.mjs';
import {validateFreeze,validateConditionalRules} from './worksheet-schema.mjs';
import {validateAutoFilter} from './filter-schema.mjs';
import {validateChart} from './chart-schema.mjs';
import {validatePivot} from './pivot-schema.mjs';
import {createFunctionRuntime} from './function-runtime.mjs';
import {normalizeFunctionName,XLSX_FUTURE_FUNCTIONS} from './function-catalog.mjs';
/** A small, original spreadsheet evaluator. No JavaScript evaluation is used. */
const MAX_ROWS = 1048576;
const MAX_COLS = 16384;
const MAX_RANGE_CELLS = MAX_FORMULA_CELLS;
const MAX_STEPS = 8_000_000;
const MAX_DEPTH = 160;
const BLANK = Symbol('blank');
const DAY_MS = 86400000;
const externalEvaluators = new WeakMap();
export function setExternalEvaluator(workbook, evaluator) { externalEvaluators.set(unwrap(workbook), evaluator); }

class FormulaError extends Error {
  constructor(code) { super(code); this.code = code; }
}
// Internal scheduling signals are not formula errors: IFERROR/ISERROR must not
// swallow them while an unfinished dependency is being calculated.
class PendingCell {
  constructor(sheet,ref){this.sheet=sheet;this.ref=ref;this.key=sheet.id+'!'+ref;}
}
class CalculationLimit {}
const fail = code => { throw new FormulaError(code); };
const validPosition = (r, c) => Number.isInteger(r) && Number.isInteger(c) && r >= 0 && c >= 0 && r < MAX_ROWS && c < MAX_COLS;
const isRef = value => value && value.kind === 'reference';
const isRange = value => value && value.kind === 'range';
const scalar = value => {
  if (isRef(value)) return value.value;
  if (isRange(value)) return value.values.length === 1 ? value.values[0] : fail('#VALUE!');
  return value;
};
const checked = value => {
  value = scalar(value);
  if (value instanceof FormulaError) throw value;
  return value;
};
const numeric = value => {
  value = checked(value);
  if (value === BLANK) return 0;
  if (typeof value === 'boolean') return Number(value);
  if (typeof value === 'number') return finite(value);
  if (typeof value === 'string' && value.trim() !== '' && /^[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:e[+-]?\d+)?$/i.test(value.trim())) return finite(Number(value));
  return fail('#VALUE!');
};
const finite = value => Number.isFinite(value) ? value : fail('#NUM!');
const textual = value => {
  value = checked(value);
  if (value === BLANK) return '';
  if (typeof value === 'boolean') return value ? 'TRUE' : 'FALSE';
  return String(value);
};
const logical = value => {
  value = checked(value);
  if (value === BLANK) return false;
  if (typeof value === 'boolean') return value;
  if (typeof value === 'number') return value !== 0;
  if (/^true$/i.test(value)) return true;
  if (/^false$/i.test(value)) return false;
  return fail('#VALUE!');
};

export function colName(index) {
  if (!Number.isInteger(index) || index < 0 || index >= MAX_COLS) return '';
  let value = index + 1, name = '';
  while (value) { name = String.fromCharCode(65 + (value - 1) % 26) + name; value = Math.floor((value - 1) / 26); }
  return name;
}

export function colIndex(label) {
  if (typeof label !== 'string' || !/^[A-Za-z]{1,3}$/.test(label)) return -1;
  let result = 0;
  for (const char of label.toUpperCase()) result = result * 26 + char.charCodeAt(0) - 64;
  return result <= MAX_COLS ? result - 1 : -1;
}

export function parseAddress(value) {
  if (typeof value !== 'string') return null;
  const match = /^\$?([A-Za-z]{1,3})\$?([1-9]\d*)$/.exec(value);
  if (!match) return null;
  const r = Number(match[2]) - 1, c = colIndex(match[1]);
  return validPosition(r, c) ? { r, c } : null;
}

export function address(r, c) { return validPosition(r, c) ? `${colName(c)}${r + 1}` : '#REF!'; }

export function createWorkbook() {
  return {
    name: '통합 문서1', activeSheetId: 'sheet-1',
    sheets: [{ id: 'sheet-1', name: 'Sheet1', cells: {}, columnWidths: {}, rowHeights: {}, merges: [], charts: [] }],
  };
}

function tokenize(input) {
  if (input.length > 32768) fail('#VALUE!');
  const tokens = [];
  let at = 0;
  const push = (type, value, start, end) => {
    tokens.push({ type, value, start, end });
    if (tokens.length > 10000) fail('#VALUE!');
  };
  while (at < input.length) {
    const start = at, ch = input[at];
    if (/\s/.test(ch)) { at++; continue; }
    if (ch === '"' || ch === "'") {
      const quote = ch;
      let value = '', closed = false;
      at++;
      while (at < input.length) {
        if (input[at] === quote) {
          if (input[at + 1] === quote) { value += quote; at += 2; }
          else { at++; closed = true; break; }
        } else value += input[at++];
      }
      if (!closed) fail('#VALUE!');
      push(quote === '"' ? 'string' : 'sheet', value, start, at);
      continue;
    }
    const rest = input.slice(at);
    const error = /^(#(?:REF!|DIV\/0!|VALUE!|NAME\?|N\/A|NUM!|NULL!|CYCLE!|SPILL!))/i.exec(rest);
    if (error) { at += error[0].length; push('error', error[0].toUpperCase(), start, at); continue; }
    const sheetName=/^[\p{L}\p{N}_.$\\]+(?=!)/u.exec(rest);
    if(sheetName){at+=sheetName[0].length;push('word',sheetName[0],start,at);continue;}
    const number = /^(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?/.exec(rest);
    if (number) { at += number[0].length; push('number', Number(number[0]), start, at); continue; }
    const word = /^[\p{L}_$\\][\p{L}\p{N}_.$\\]*/u.exec(rest);
    if (word) { at += word[0].length; push('word', word[0], start, at); continue; }
    const comparison = /^(<>|<=|>=)/.exec(rest);
    if (comparison) { at += 2; push('op', comparison[0], start, at); continue; }
    if ('+-*/^&%=<>():,;!{}'.includes(ch)) { at++; push('op', ch, start, at); continue; }
    fail('#VALUE!');
  }
  tokens.push({ type: 'end', value: '', start: at, end: at });
  return tokens;
}

class Parser {
  constructor(formula) { this.tokens = tokenize(formula.replace(/^=/, '')); this.at = 0; this.depth = 0; }
  peek(value) { return this.tokens[this.at].type === 'op' && this.tokens[this.at].value === value; }
  accept(value) { if (this.peek(value)) { this.at++; return true; } return false; }
  expect(value) { if (!this.accept(value)) fail('#VALUE!'); }
  parse() {
    const result = this.expression();
    if (this.tokens[this.at].type !== 'end') fail('#VALUE!');
    return result;
  }
  expression() {
    if (++this.depth > MAX_DEPTH) fail('#VALUE!');
    let node = this.concat();
    while (['=', '<>', '<', '>', '<=', '>='].includes(this.tokens[this.at].value)) {
      const op = this.tokens[this.at++].value;
      node = { type: 'binary', op, left: node, right: this.concat() };
    }
    this.depth--;
    return node;
  }
  concat() {
    let node = this.sum();
    while (this.accept('&')) node = { type: 'binary', op: '&', left: node, right: this.sum() };
    return node;
  }
  sum() {
    let node = this.product();
    while (['+', '-'].includes(this.tokens[this.at].value)) {
      const op = this.tokens[this.at++].value;
      node = { type: 'binary', op, left: node, right: this.product() };
    }
    return node;
  }
  product() {
    let node = this.power();
    while (['*', '/'].includes(this.tokens[this.at].value)) {
      const op = this.tokens[this.at++].value;
      node = { type: 'binary', op, left: node, right: this.power() };
    }
    return node;
  }
  power() {
    let node = this.unary();
    while (this.accept('^')) node = { type: 'binary', op: '^', left: node, right: this.unary() };
    return node;
  }
  unary() {
    const ops = [];
    while (this.peek('+') || this.peek('-')) {
      ops.push(this.tokens[this.at++].value);
      if (ops.length > MAX_DEPTH) fail('#VALUE!');
    }
    let node = this.referenceExpression();
    while (this.accept('%')) node = { type: 'unary', op: '%', node };
    for (let i = ops.length - 1; i >= 0; i--) node = { type: 'unary', op: ops[i], node };
    return node;
  }
  referenceRange(){
    let node=this.primary();
    while(this.accept(':')){const right=this.primary();if(node.sheet&&right.type==='ref'&&!right.sheet)right.sheet=node.sheet;node={type:'referenceOp',op:':',left:node,right};}
    return node;
  }
  referenceExpression() {
    let node=this.referenceRange();
    while(true){
      const previous=this.tokens[this.at-1],next=this.tokens[this.at];
      if((['word','sheet','error'].includes(next.type)||this.peek('('))&&next.start>previous.end){node={type:'referenceOp',op:' ',left:node,right:this.referenceRange()};continue;}
      return node;
    }
  }
  primary() {
    if (this.accept('(')) { let value=this.expression();while(this.accept(','))value={type:'referenceOp',op:',',left:value,right:this.expression()};this.expect(')');return value; }
    if (this.accept('{')) {
      const values=[],widths=[];let width=0;
      const constant=node=>node.type==='literal'||node.type==='error'||node.type==='unary'&&constant(node.node);
      do {
        const node=this.expression();if(!constant(node))fail('#VALUE!');values.push(node);width++;
        if(this.accept(','))continue;
        widths.push(width);width=0;
        if(this.accept(';'))continue;
        this.expect('}');break;
      } while(values.length<MAX_RANGE_CELLS);
      if(widths.some(n=>n!==widths[0])||!widths.length)fail('#VALUE!');
      return {type:'array',values,rows:widths.length,cols:widths[0]};
    }
    const token = this.tokens[this.at++];
    if (token.type === 'number' || token.type === 'string') return { type: 'literal', value: token.value };
    if (token.type === 'error') return { type: 'error', value: token.value };
    if (token.type !== 'word' && token.type !== 'sheet') fail('#VALUE!');
    if (token.type === 'word' && this.accept('(')) {
      const args = [];
      if (!this.accept(')')) {
        do {
          args.push(this.peek(',') || this.peek(';') || this.peek(')') ? { type: 'literal', value: 0, omitted:true } : this.expression());
        } while (this.accept(',') || this.accept(';'));
        this.expect(')');
      }
      return { type: 'call', name: token.value.toUpperCase(), args };
    }
    let sheet = null, reference = token.value;
    if (this.accept('!')) {
      sheet = reference;
      const ref = this.tokens[this.at++];
      if (ref.type === 'error') return { type: 'error', value: ref.value };
      if (ref.type !== 'word') fail('#REF!');
      reference = ref.value;
    } else if (token.type === 'sheet') fail('#NAME?');
    if (!sheet && /^(true|false)$/i.test(reference)) return { type: 'literal', value: /^true$/i.test(reference) };
    const pos = parseAddress(reference);
    if (!pos) {
      if (/^\$?[A-Za-z]+\$?\d+$/.test(reference)) fail('#REF!');
      return { type: 'error', value: '#NAME?' };
    }
    const first = { type: 'ref', sheet, pos };
    return first;
  }
}

function context(workbook, sheetId) {
  const sheet = workbook?.sheets?.find(item => item.id === sheetId);
  if (!sheet) fail('#REF!');
  const dateOffset=workbook.date1904?1462:0;
  const calendar={dateSerial:(...args)=>dateSerial(...args)-dateOffset,serialParts:value=>serialParts(numeric(value)+dateOffset)};
  const runtime=dateOffset?createFunctionRuntime({...functionHelpers,...calendar,dateOffset,
    criterion:(value,options)=>criterion(value,{...options,dateOffset}),
    formatValue:(value,style)=>formatValue(value,style,workbook),
  }):extendedFunction;
  return { workbook, sheet, dateOffset, calendar, runtime, position:{r:0,c:0}, active: new Set(), cache: new Map(), steps: 0, depth: 0 };
}
function step(ctx) { if (++ctx.steps > MAX_STEPS) throw new CalculationLimit(); }
function resolveSheet(ctx, name) {
  if (!name) return ctx.sheet;
  return ctx.workbook.sheets.find(sheet => sheet.name.toLowerCase() === name.toLowerCase()) || fail('#REF!');
}
function literal(raw) {
  if (raw === undefined || raw === null || raw === '') return BLANK;
  if (typeof raw !== 'string') raw = String(raw);
  if (raw.startsWith("'")) return raw.slice(1);
  if (/^true$/i.test(raw)) return true;
  if (/^false$/i.test(raw)) return false;
  if (/^[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?$/.test(raw.trim())) return finite(Number(raw));
  if (/^[+-]?(?:\d+(?:\.\d*)?|\.\d+)%$/.test(raw.trim())) return finite(Number(raw.trim().slice(0, -1)) / 100);
  if (/^#(?:REF!|DIV\/0!|VALUE!|NAME\?|N\/A|NUM!|NULL!|CYCLE!|SPILL!)$/.test(raw)) return new FormulaError(raw);
  return raw;
}
function readCell(ctx, sheet, ref) {
  step(ctx);
  const pos = parseAddress(ref);
  if (!pos) return new FormulaError('#REF!');
  ref = address(pos.r, pos.c);
  const raw = sheet.cells?.[ref]?.raw ?? '';
  // Literal values are cheap; caching a million of them duplicates the model.
  if (typeof raw !== 'string' || !raw.startsWith('=')) return literal(raw);
  const key = `${sheet.id}!${ref}`;
  if (ctx.cache.has(key)) return ctx.cache.get(key);
  if (ctx.active.has(key)) return new FormulaError('#CYCLE!');
  if(ctx.scheduling&&ctx.frame.key!==key){
    if(ctx.waiting.has(key))return new FormulaError('#CYCLE!');
    throw new PendingCell(sheet,ref);
  }
  if (ctx.depth >= MAX_DEPTH) return new FormulaError('#NUM!');
  ctx.active.add(key);
  ctx.depth++;
  let value;
  const original = ctx.sheet, originalPosition=ctx.position;
  try {
    ctx.sheet = sheet;ctx.position=pos;
    const external = externalEvaluators.get(unwrap(ctx.workbook));
    const node=external?null:ctx.frame?.key===key?(ctx.frame.ast??=new Parser(raw).parse()):new Parser(raw).parse();
    value = external ? external(sheet.id, ref) : checked(evaluateNode(node, ctx, 0));
    if (value === BLANK && raw.startsWith('=')) value = 0;
  } catch (error) {
    if(error instanceof PendingCell||error instanceof CalculationLimit)throw error;
    value = error instanceof FormulaError ? error : new FormulaError('#VALUE!');
  } finally { ctx.sheet = original;ctx.position=originalPosition; ctx.depth--; ctx.active.delete(key); }
  ctx.cache.set(key, value);
  return value;
}
function runCalculation(ctx,compute){
  const frames=[{key:null,compute,memo:new Map(),ranges:new Map()}];
  ctx.scheduling=true;ctx.waiting=new Set();
  try{
    while(frames.length){
      const frame=frames.at(-1);ctx.frame=frame;
      try{
        const value=frame.compute();frames.pop();if(frame.key)ctx.waiting.delete(frame.key);
        if(!frames.length)return value;
      }catch(error){
        if(!(error instanceof PendingCell))throw error;
        if(frames.length>MAX_FORMULA_CELLS)throw new CalculationLimit();
        const {sheet,ref,key}=error;
        ctx.waiting.add(key);
        frames.push({key,compute:()=>readCell(ctx,sheet,ref),memo:new Map(),ranges:new Map()});
      }
    }
  }finally{ctx.frame=null;ctx.waiting=null;ctx.scheduling=false;}
}
function compare(left, right) {
  left = checked(left); right = checked(right);
  if (left === BLANK) left = typeof right === 'string' ? '' : typeof right === 'boolean' ? false : 0;
  if (right === BLANK) right = typeof left === 'string' ? '' : typeof left === 'boolean' ? false : 0;
  if (typeof left === 'string' && typeof right === 'string') {
    left = left.toLocaleLowerCase(); right = right.toLocaleLowerCase();
  }
  if (typeof left !== typeof right) {
    const rank = { number: 0, string: 1, boolean: 2 };
    return rank[typeof left] < rank[typeof right] ? -1 : 1;
  }
  return left === right ? 0 : left < right ? -1 : 1;
}
function evaluateNode(node, ctx, depth) {
  const memo=ctx.frame?.memo;
  if(memo?.has(node))return memo.get(node);
  const value=evaluateNodeUncached(node,ctx,depth);
  memo?.set(node,value);
  return value;
}
function evaluateNodeUncached(node, ctx, depth) {
  step(ctx);
  if (depth > MAX_DEPTH) fail('#NUM!');
  const evaluate = child => evaluateNode(child, ctx, depth + 1);
  if (node.type === 'literal') return node.value;
  if (node.type === 'error') fail(node.value);
  if (node.type === 'ref') {
    const sheet=resolveSheet(ctx,node.sheet);
    return { kind:'reference',sheet,top:node.pos.r,left:node.pos.c,rows:1,cols:1,value:readCell(ctx,sheet,address(node.pos.r,node.pos.c)) };
  }
  if(node.type==='array')return {kind:'range',rows:node.rows,cols:node.cols,values:node.values.map(child=>{try{return checked(evaluate(child));}catch(e){if(e instanceof FormulaError)return e;throw e;}})};
  if (node.type === 'range') {
    const sheet = resolveSheet(ctx, node.sheet);
    const top = Math.min(node.start.r, node.end.r), left = Math.min(node.start.c, node.end.c);
    const rows = Math.abs(node.start.r - node.end.r) + 1, cols = Math.abs(node.start.c - node.end.c) + 1;
    return getRange(ctx,sheet,top,left,rows,cols);
  }
  if(node.type==='referenceOp'){
    // Evaluate error endpoints first so malformed references preserve #REF!.
    const a=asRange(evaluate(node.left)),b=asRange(evaluate(node.right));
    if(node.op===',')return {kind:'range',rows:a.values.length+b.values.length,cols:1,values:[...a.values,...b.values]};
    if(!a.sheet||!b.sheet||a.sheet.id!==b.sheet.id)fail('#REF!');
    const intersection=node.op===' ',top=(intersection?Math.max:Math.min)(a.top,b.top),left=(intersection?Math.max:Math.min)(a.left,b.left);
    const bottom=(intersection?Math.min:Math.max)(a.top+a.rows-1,b.top+b.rows-1),right=(intersection?Math.min:Math.max)(a.left+a.cols-1,b.left+b.cols-1);
    if(bottom<top||right<left)fail('#NULL!');
    return getRange(ctx,a.sheet,top,left,bottom-top+1,right-left+1);
  }
  if (node.type === 'unary') {
    const source=evaluate(node.node),apply=v=>{if(node.op==='+')return checked(v);const value=numeric(v);return finite(node.op==='%'?value/100:node.op==='-'?-value:value);};
    if(isRange(source))return {...source,sheet:undefined,values:source.values.map(v=>{try{return apply(v);}catch(e){if(e instanceof FormulaError)return e;throw e;}})};
    return apply(source);
  }
  if (node.type === 'binary') {
    const left = evaluate(node.left), right = evaluate(node.right);
    const apply=(left,right)=>{
    if (node.op === '&') return textual(left) + textual(right);
    if (['=', '<>', '<', '>', '<=', '>='].includes(node.op)) {
      const order = compare(left, right);
      return ({ '=': order === 0, '<>': order !== 0, '<': order < 0, '>': order > 0, '<=': order <= 0, '>=': order >= 0 })[node.op];
    }
    const a = numeric(left), b = numeric(right);
    if (node.op === '/' && b === 0) fail('#DIV/0!');
    return finite(({ '+': () => a + b, '-': () => a - b, '*': () => a * b, '/': () => a / b, '^': () => a ** b })[node.op]());
    };
    if(isRange(left)||isRange(right)){
      const a=asRange(left),b=asRange(right),rows=Math.max(a.rows,b.rows),cols=Math.max(a.cols,b.cols);
      if(rows*cols>MAX_RANGE_CELLS)fail('#NUM!');
      if(a.rows!==b.rows&&a.rows!==1&&b.rows!==1||a.cols!==b.cols&&a.cols!==1&&b.cols!==1)fail('#VALUE!');
      const values=[];for(let r=0;r<rows;r++)for(let c=0;c<cols;c++){
        try{values.push(apply(a.values[(a.rows===1?0:r)*a.cols+(a.cols===1?0:c)],b.values[(b.rows===1?0:r)*b.cols+(b.cols===1?0:c)]));}
        catch(e){if(e instanceof FormulaError)values.push(e);else throw e;}
      }
      return {kind:'range',rows,cols,values};
    }
    return apply(left,right);
  }
  if (node.type === 'call') return callFunction(normalizeFunctionName(node.name), node.args, evaluate,ctx);
  return fail('#VALUE!');
}

const flatten = args => args.flatMap(arg => isRange(arg) ? arg.values.map(value => ({ value, reference: true })) : [{ value: scalar(arg), reference: isRef(arg) }]);
const asRange = arg => isRange(arg) ? arg : { ...(isRef(arg)?arg:{}), kind: 'range', values: [scalar(arg)], rows: 1, cols: 1 };
function numericItems(args) {
  const numbers = [];
  for (const item of flatten(args)) {
    const value = checked(item.value);
    if (value === BLANK) continue;
    if (typeof value === 'number') numbers.push(value);
    else if (!item.reference) numbers.push(numeric(value));
  }
  return numbers;
}
function wildcardRegex(pattern) {
  let source = '^';
  for (let i = 0; i < pattern.length; i++) {
    let ch = pattern[i];
    if (ch === '~' && i + 1 < pattern.length) ch = pattern[++i];
    else if (ch === '*') { source += '.*'; continue; }
    else if (ch === '?') { source += '.'; continue; }
    source += ch.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  }
  return new RegExp(source + '$', 'iu');
}
function criterion(test,{blankAsZero=false,booleanAsNumber=false,dateOffset=0}={}) {
  test = checked(test);
  if (test === BLANK) test = 0;
  let op = '=', target = test;
  if (typeof test === 'string') {
    const match = /^(<=|>=|<>|=|<|>)(.*)$/s.exec(test);
    if (match) { op = match[1]; target = match[2]; }
    if (/^[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:e[+-]?\d+)?$/i.test(target)) target = Number(target);
    else if (/^(true|false)$/i.test(target)) target = /^true$/i.test(target);
    else {
      const parsed=parseDateText(target);
      if(parsed!==null)target=parsed-dateOffset;
    }
  }
  if(booleanAsNumber&&typeof target==='boolean')target=Number(target);
  const wildcard = typeof target === 'string' && /[*?~]/.test(target) ? wildcardRegex(target) : null;
  return value => {
    value = scalar(value);
    if (value instanceof FormulaError) return typeof target === 'string' && op === '=' && value.code === target;
    if (wildcard && (op === '=' || op === '<>')) {
      const matches = typeof value === 'string' && wildcard.test(value);
      return op === '=' ? matches : !matches;
    }
    if(test==='=')return value===BLANK;
    if(test==='<>')return value!==BLANK;
    if(value===BLANK){if(blankAsZero&&typeof target==='number')value=0;else return target===''&&op==='=';}
    if(value==='')return target===''?(op==='='||op==='>='||op==='<='):op==='<>';
    if(booleanAsNumber&&typeof value==='boolean')value=Number(value);
    if(typeof target!==typeof value){
      if(typeof target==='number'&&typeof value==='string'&&['=','<>'].includes(op)&&value.trim()&&Number.isFinite(Number(value)))value=Number(value);
      else return op==='<>';
    }
    const order = compare(value, target);
    return ({ '=': order === 0, '<>': order !== 0, '<': order < 0, '>': order > 0, '<=': order <= 0, '>=': order >= 0 })[op];
  };
}

// Excel criteria accept date strings as serial values. Keep this parser
// explicit and deterministic instead of relying on Date.parse and OS locale.
function parseDateText(value){
  if(typeof value!=='string')return null;
  const text=value.trim();
  let match=/^(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})$/.exec(text),year,month,day;
  if(match){year=+match[1];month=+match[2];day=+match[3];}
  else {match=/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/.exec(text);if(!match)return null;month=+match[1];day=+match[2];year=+match[3];}
  if(year<1900||year>9999||month<1||month>12)return null;
  const days=year===1900&&month===2?29:new Date(Date.UTC(year,month,0)).getUTCDate();
  if(day<1||day>days)return null;
  return year===1900&&month===2&&day===29?60:dateSerial(year,month,day);
}

function dateSerial(year, month, day) {
  const date = new Date(0);
  date.setUTCHours(0,0,0,0);date.setUTCFullYear(year,month-1,1);
  const difference=Math.round((date.getTime()-Date.UTC(1899,11,31))/DAY_MS);
  const serial=difference+(difference>=60?1:0)+day-1;
  if (!Number.isFinite(serial) || serial < 0 || serial > 2958465) fail('#NUM!');
  return serial;
}
function serialParts(value) {
  const serial = Math.floor(numeric(value));
  if (serial < 0 || serial > 2958465) fail('#NUM!');
  if (serial === 60) return { year: 1900, month: 2, day: 29 };
  if (serial === 0) return { year: 1900, month: 1, day: 0 };
  const date = new Date(Date.UTC(1899, 11, 31) + (serial >= 60 ? serial - 1 : serial) * DAY_MS);
  return { year: date.getUTCFullYear(), month: date.getUTCMonth() + 1, day: date.getUTCDate() };
}
function rounded(value, digits, mode = 'nearest') {
  value = numeric(value); digits = Math.trunc(numeric(digits));
  if (digits > 308) return value;
  if (digits < -308) return 0;
  const abs = Math.abs(value), factor = 10 ** digits;
  const scaled = abs * factor;
  if (!Number.isFinite(scaled)) return value;
  let result;
  const epsilon = Math.min(1e-7, Math.max(1, scaled) * Number.EPSILON * 2);
  const nearest = Math.round(scaled);
  const corrected = nearest > 0 && Math.abs(scaled - nearest) < epsilon ? nearest : scaled;
  if (mode === 'up') result = Math.ceil(corrected);
  else if (mode === 'down') result = Math.floor(corrected);
  else result = Math.floor(scaled + 0.5 + epsilon);
  return finite(Math.sign(value) * result / factor);
}

function getRange(ctx,sheet,top,left,rows,cols) {
  if(!validPosition(top,left)||!validPosition(top+rows-1,left+cols-1)||rows<1||cols<1)fail('#REF!');
  if(rows*cols>MAX_RANGE_CELLS)fail('#NUM!');
  const key=[sheet.id,top,left,rows,cols].join(':'),ranges=ctx.frame?.ranges;
  let result=ranges?.get(key);
  if(!result){result={kind:'range',sheet,top,left,rows,cols,values:[]};ranges?.set(key,result);}
  // Resume a partially read range after a deferred formula completes. Restarting
  // at its first cell for every dependency makes large SUM ranges quadratic.
  for(let i=result.values.length;i<rows*cols;i++)result.values.push(readCell(ctx,sheet,address(top+Math.floor(i/cols),left+i%cols)));
  return result;
}
function referenceFunction(ctx,name,nodes,evaluate) {
  const origin={sheet:ctx.sheet,top:ctx.position.r,left:ctx.position.c,rows:1,cols:1};
  // Metadata functions must not read the referenced cell (ROW(A1) in A1 is not a cycle).
  const metadata=node=>{
    if(node.type==='ref')return {sheet:resolveSheet(ctx,node.sheet),top:node.pos.r,left:node.pos.c,rows:1,cols:1};
    if(node.type==='referenceOp'&&node.op!==','){const a=metadata(node.left),b=metadata(node.right);if(!a.sheet||a.sheet.id!==b.sheet?.id)fail('#REF!');const intersection=node.op===' ',top=(intersection?Math.max:Math.min)(a.top,b.top),left=(intersection?Math.max:Math.min)(a.left,b.left),bottom=(intersection?Math.min:Math.max)(a.top+a.rows-1,b.top+b.rows-1),right=(intersection?Math.min:Math.max)(a.left+a.cols-1,b.left+b.cols-1);if(bottom<top||right<left)fail('#NULL!');return {sheet:a.sheet,top,left,rows:bottom-top+1,cols:right-left+1};}
    if(node.type==='range')return {sheet:resolveSheet(ctx,node.sheet),top:Math.min(node.start.r,node.end.r),left:Math.min(node.start.c,node.end.c),rows:Math.abs(node.start.r-node.end.r)+1,cols:Math.abs(node.start.c-node.end.c)+1};
    const value=evaluate(node);if(isRef(value)||isRange(value))return value;return fail('#VALUE!');
  };
  if(name==='ROW'||name==='COLUMN'){const ref=nodes.length?metadata(nodes[0]):origin;if(!ref.sheet)fail('#VALUE!');return (name==='ROW'?ref.top:ref.left)+1;}
  if(name==='ROWS'||name==='COLUMNS'){const ref=metadata(nodes[0]);return name==='ROWS'?ref.rows:ref.cols;}
  if(name==='FORMULATEXT'||name==='ISFORMULA'){
    const ref=metadata(nodes[0]);if(!ref.sheet)return name==='ISFORMULA'?false:fail('#N/A');
    const raw=ref.sheet.cells?.[address(ref.top,ref.left)]?.raw??'',isFormula=raw.startsWith('=');
    return name==='ISFORMULA'?isFormula:isFormula?raw:fail('#N/A');
  }
  if(name==='SHEET'){
    let sheet=ctx.sheet;
    if(nodes.length){const v=evaluate(nodes[0]);sheet=v?.sheet||ctx.workbook.sheets.find(s=>s.name.toLowerCase()===textual(v).toLowerCase());if(!sheet)fail('#N/A');}
    return ctx.workbook.sheets.indexOf(sheet)+1;
  }
  if(name==='SHEETS'){if(!nodes.length)return ctx.workbook.sheets.length;const ref=metadata(nodes[0]);return ref.sheet?1:fail('#REF!');}
  if(name==='ADDRESS'){
    const args=nodes.map(evaluate),row=Math.trunc(numeric(args[0])),col=Math.trunc(numeric(args[1])),abs=args.length>2?Math.trunc(numeric(args[2])):1,a1=args.length>3?logical(args[3]):true;
    if(!validPosition(row-1,col-1)||![1,2,3,4].includes(abs))fail('#VALUE!');
    const rowAbs=abs===1||abs===2,colAbs=abs===1||abs===3;
    let result=a1?`${colAbs?'$':''}${colName(col-1)}${rowAbs?'$':''}${row}`:`R${rowAbs?row:'['+row+']'}C${colAbs?col:'['+col+']'}`;
    if(args.length>4){const sheet=textual(args[4]);result=(/^[\p{L}_][\p{L}\p{N}_]*$/u.test(sheet)&&!parseAddress(sheet)?sheet:`'${sheet.replace(/'/g,"''")}'`)+ '!'+result;}
    return result;
  }
  if(name==='INDIRECT'){
    let text=textual(evaluate(nodes[0]));const a1=nodes.length<2||logical(evaluate(nodes[1]));
    if(!a1){
      const convert=ref=>{const m=/^R(\[-?\d+\]|\d*)C(\[-?\d+\]|\d*)$/i.exec(ref);if(!m)fail('#REF!');const pos=(p,base)=>p.startsWith('[')?base+Number(p.slice(1,-1)):p?Number(p)-1:base;return address(pos(m[1],ctx.position.r),pos(m[2],ctx.position.c));};
      const split=text.lastIndexOf('!');text=text.slice(0,split+1)+text.slice(split+1).split(':').map(convert).join(':');
    }
    let node;try{node=new Parser(text).parse();}catch{fail('#REF!');}
    if(!['ref','range','referenceOp'].includes(node.type))fail('#REF!');return evaluate(node);
  }
  if(name==='OFFSET'){
    const ref=metadata(nodes[0]);if(!ref.sheet)fail('#VALUE!');
    const args=nodes.slice(1).map(node=>Math.trunc(numeric(evaluate(node))));
    return getRange(ctx,ref.sheet,ref.top+args[0],ref.left+args[1],args[2]??ref.rows,args[3]??ref.cols);
  }
  if(name==='SUBTOTAL'){
    const code=Math.trunc(numeric(evaluate(nodes[0]))),index=code>100?code-100:code;
    const target=['','AVERAGE','COUNT','COUNTA','MAX','MIN','PRODUCT','STDEV.S','STDEV.P','SUM','VAR.S','VAR.P'][index];
    if(!target||code!==index&&code!==index+100)fail('#VALUE!');
    const values=[];for(const node of nodes.slice(1)){
      const ref=metadata(node);if(!ref.sheet)fail('#VALUE!');if(ref.rows*ref.cols>MAX_RANGE_CELLS)fail('#NUM!');
      const key=JSON.stringify([code,ref.sheet.id,ref.top,ref.left,ref.rows,ref.cols]);
      const cache=ctx.frame?(ctx.frame.subtotals??=new Map()):new Map();
      let progress=cache.get(key);
      if(!progress){progress={next:0,values:[],hidden:new Set(ref.sheet.autoFilter?.hiddenRows||[])};cache.set(key,progress);}
      for(;progress.next<ref.rows*ref.cols;progress.next++){
        const r=ref.top+Math.floor(progress.next/ref.cols),c=ref.left+progress.next%ref.cols;
        const raw=ref.sheet.cells?.[address(r,c)]?.raw??'';
        if(/^\s*=\s*(?:_xlfn\.)?(?:SUBTOTAL|AGGREGATE)\s*\(/i.test(raw))continue;
        if(progress.hidden.has(r))continue;
        if(code>100&&ref.sheet.rowHeights?.[r]===0)continue;
        progress.values.push(readCell(ctx,ref.sheet,address(r,c)));
      }
      for(const value of progress.values)values.push(value);
    }
    return callFunction(target,[{type:'literal',value:{kind:'range',values,rows:values.length,cols:1}}],evaluate,ctx);
  }
  fail('#NAME?');
}
const functionHelpers={BLANK,FormulaError,fail,scalar,checked,numeric,textual,logical,finite,isRange,isRef,asRange,flatten,numericItems,criterion,compare,wildcardRegex,dateSerial,serialParts,rounded,colName,parseAddress,address,formatValue};
const extendedFunction=createFunctionRuntime(functionHelpers);

function callFunction(name, nodes, evaluate, ctx) {
  const arity = (min, max = min) => { if (nodes.length < min || nodes.length > max) fail('#VALUE!'); };
  if (name === 'IF') {
    arity(2, 3);
    return logical(evaluate(nodes[0])) ? evaluate(nodes[1]) : nodes.length === 3 ? evaluate(nodes[2]) : false;
  }
  if (name === 'IFERROR') {
    arity(2);
    try { const value=evaluate(nodes[0]);if(isRange(value))return {...value,values:value.values.map(v=>v instanceof FormulaError?checked(evaluate(nodes[1])):v)};return checked(value); } catch (error) {
      if (!(error instanceof FormulaError)) throw error;
      return evaluate(nodes[1]);
    }
  }
  const known = new Set(['SUM', 'AVERAGE', 'MIN', 'MAX', 'COUNT', 'COUNTA', 'AND', 'OR', 'NOT', 'ROUND', 'ROUNDUP', 'ROUNDDOWN', 'ABS', 'SQRT', 'POWER', 'MOD', 'INT', 'LEN', 'LEFT', 'RIGHT', 'MID', 'CONCAT', 'CONCATENATE', 'TRIM', 'UPPER', 'LOWER', 'SUMIF', 'COUNTIF', 'VLOOKUP', 'XLOOKUP', 'TODAY', 'DATE', 'YEAR', 'MONTH', 'DAY']);
  if (!known.has(name)) return ctx.runtime(name,nodes,evaluate,{
    referenceFunction:(name,nodes,evaluate)=>referenceFunction(ctx,name,nodes,evaluate),
    getRange:(sheet,top,left,rows,cols)=>getRange(ctx,sheet,top,left,rows,cols),
  });
  const args = nodes.map((node, index) => {if(name==='XLOOKUP'&&index===3)return BLANK;try{return evaluate(node);}catch(error){if(['COUNT','COUNTA'].includes(name)&&error instanceof FormulaError)return error;throw error;}});
  const num = index => numeric(args[index]);
  const text = index => textual(args[index]);
  const bool = index => logical(args[index]);
  if (['SUM', 'AVERAGE', 'MIN', 'MAX', 'COUNT', 'COUNTA'].includes(name)) {
    arity(1, 255);
    if (name === 'COUNTA') return flatten(args).filter(({ value }) => value !== BLANK).length;
    if (name === 'COUNT') return flatten(args).filter(({ value, reference }) => {
      if (typeof value === 'number') return true;
      if (value === BLANK || value instanceof FormulaError || reference) return false;
      try { numeric(value); return true; } catch { return false; }
    }).length;
    const numbers = numericItems(args);
    if (name === 'AVERAGE' && numbers.length === 0) fail('#DIV/0!');
    if (!numbers.length) return 0;
    if (name === 'MIN') return numbers.reduce((a, b) => Math.min(a, b));
    if (name === 'MAX') return numbers.reduce((a, b) => Math.max(a, b));
    const total = numbers.reduce((a, b) => a + b, 0);
    return finite(name === 'AVERAGE' ? total / numbers.length : total);
  }
  if (name === 'AND' || name === 'OR') {
    arity(1, 255);
    const values = [];
    for (const { value, reference } of flatten(args)) {
      checked(value);
      if (value === BLANK || (reference && typeof value === 'string')) continue;
      values.push(logical(value));
    }
    if (!values.length) fail('#VALUE!');
    return name === 'AND' ? values.every(Boolean) : values.some(Boolean);
  }
  if (name === 'NOT') { arity(1); return !bool(0); }
  if (['ROUND', 'ROUNDUP', 'ROUNDDOWN'].includes(name)) {
    arity(2); return rounded(args[0], args[1], name === 'ROUNDUP' ? 'up' : name === 'ROUNDDOWN' ? 'down' : 'nearest');
  }
  if (name === 'ABS') { arity(1); return Math.abs(num(0)); }
  if (name === 'SQRT') { arity(1); return finite(Math.sqrt(num(0))); }
  if (name === 'POWER') { arity(2); return finite(num(0) ** num(1)); }
  if (name === 'MOD') { arity(2); const a = num(0), b = num(1); if (!b) fail('#DIV/0!'); return finite(a - b * Math.floor(a / b)); }
  if (name === 'INT') { arity(1); return Math.floor(num(0)); }
  if (name === 'LEN') { arity(1); return text(0).length; }
  if (name === 'LEFT' || name === 'RIGHT') {
    arity(1, 2);
    const value = text(0), count = nodes.length === 1 ? 1 : Math.trunc(num(1));
    if (count < 0) fail('#VALUE!');
    return name === 'LEFT' ? value.slice(0, count) : count === 0 ? '' : value.slice(-count);
  }
  if (name === 'MID') {
    arity(3); const start = Math.trunc(num(1)), count = Math.trunc(num(2));
    if (start < 1 || count < 0) fail('#VALUE!');
    return text(0).slice(start - 1, start - 1 + count);
  }
  if (name === 'CONCAT' || name === 'CONCATENATE') {
    arity(1, 255); return flatten(args).map(({ value }) => textual(value)).join('');
  }
  if (name === 'TRIM') { arity(1); return text(0).replace(/^ +| +$/g, '').replace(/ +/g, ' '); }
  if (name === 'UPPER') { arity(1); return text(0).toUpperCase(); }
  if (name === 'LOWER') { arity(1); return text(0).toLowerCase(); }
  if (name === 'COUNTIF' || name === 'SUMIF') {
    arity(2, name === 'SUMIF' ? 3 : 2);
    const range = asRange(args[0]), matches = criterion(args[1],{dateOffset:ctx.dateOffset});
    if (name === 'COUNTIF') return range.values.filter(matches).length;
    let sumRange = args.length > 2 ? asRange(args[2]) : range;
    if(sumRange.sheet)sumRange=getRange(ctx,sumRange.sheet,sumRange.top,sumRange.left,range.rows,range.cols);
    if (sumRange.rows !== range.rows || sumRange.cols !== range.cols) fail('#VALUE!');
    let total = 0;
    range.values.forEach((value, i) => {
      if (!matches(value)) return;
      const item = checked(sumRange.values[i]);
      if (typeof item === 'number') total += item;
    });
    return finite(total);
  }
  if (name === 'VLOOKUP') {
    arity(3, 4);
    const lookup = checked(args[0]), range = asRange(args[1]), column = Math.trunc(num(2));
    if (column < 1) fail('#VALUE!');
    if (column > range.cols) fail('#REF!');
    const approximate = args.length < 4 ? true : bool(3);
    const pattern = typeof lookup === 'string' ? wildcardRegex(lookup) : null;
    const exactMatch = value => pattern ? typeof value === 'string' && pattern.test(value) : compare(value, lookup) === 0;
    let selected = -1;
    for (let r = 0; r < range.rows; r++) {
      const value = range.values[r * range.cols];
      if (value instanceof FormulaError||value===BLANK||typeof value!==typeof lookup) continue;
      if (!approximate && exactMatch(value)) { selected = r; break; }
      if (approximate && compare(value, lookup) <= 0) selected = r;
    }
    if (selected < 0) fail('#N/A');
    const result = checked(range.values[selected * range.cols + column - 1]);
    return result === BLANK ? 0 : result;
  }
  if (name === 'XLOOKUP') {
    arity(3, 6);
    const lookup = checked(args[0]), range = asRange(args[1]), result = asRange(args[2]);
    if ((range.rows > 1 && range.cols > 1) || result.values.length !== range.values.length) fail('#VALUE!');
    const mode = args.length > 4 ? Math.trunc(num(4)) : 0;
    const search = args.length > 5 ? Math.trunc(num(5)) : 1;
    if (![0, -1, 1, 2].includes(mode) || ![1, -1, 2, -2].includes(search)) fail('#VALUE!');
    let selected = -1, best;
    const indexes = Array.from({ length: range.values.length }, (_, index) => index);
    if (search < 0) indexes.reverse();
    for (const i of indexes) {
      const value = range.values[i];
      if (value instanceof FormulaError||value===BLANK||typeof value!==typeof lookup) continue;
      const order = compare(value, lookup);
      if (mode === 2 ? typeof lookup === 'string' && typeof value === 'string' && wildcardRegex(lookup).test(value) : order === 0) { selected = i; break; }
      if ((mode === -1 && order < 0) || (mode === 1 && order > 0)) {
        if (best === undefined || (mode === -1 ? compare(value, best) > 0 : compare(value, best) < 0)) { selected = i; best = value; }
      }
    }
    if (selected < 0) return args.length > 3&&!nodes[3].omitted ? checked(evaluate(nodes[3])) : fail('#N/A');
    const value = checked(result.values[selected]);
    return value === BLANK ? 0 : value;
  }
  if (name === 'TODAY') {
    arity(0); const now = new Date(); return ctx.calendar.dateSerial(now.getFullYear(), now.getMonth() + 1, now.getDate());
  }
  if (name === 'DATE') {
    arity(3); let year = Math.trunc(num(0));
    if (year >= 0 && year < 1900) year += 1900;
    if (year < 0 || year > 9999) fail('#NUM!');
    const month = Math.trunc(num(1)), day = Math.trunc(num(2));
    if (year === 1900 && month === 2 && day === 29) return 60-ctx.dateOffset;
    return ctx.calendar.dateSerial(year, month, day);
  }
  if (['YEAR', 'MONTH', 'DAY'].includes(name)) {
    arity(1); return ctx.calendar.serialParts(args[0])[name.toLowerCase()];
  }
  return fail('#NAME?');
}

function publicValue(value) {
  value = checked(value);
  return value === BLANK ? '' : typeof value === 'number' ? finite(value) : value;
}
function guard(callback) {
  try { return publicValue(callback()); } catch (error) { return error instanceof CalculationLimit ? '#NUM!' : error instanceof FormulaError ? error.code : '#VALUE!'; }
}

export function evaluateCell(workbook, sheetId, ref) {
  return createEvaluator(workbook,sheetId)(ref);
}

export function evaluateFormula(workbook, sheetId, formula) {
  return guard(() => {
    const ctx = context(workbook, sheetId);
    const node=new Parser(String(formula)).parse();
    const result = checked(runCalculation(ctx,()=>evaluateNode(node, ctx, 0)));
    return result === BLANK ? 0 : result;
  });
}

/** Returns a recalculation session for rendering many cells with a shared cache. */
export function createEvaluator(workbook, sheetId) {
  let ctx;
  try { ctx = context(workbook, sheetId); } catch { return () => '#REF!'; }
  return ref => guard(() => { ctx.steps = 0; return runCalculation(ctx,()=>readCell(ctx, ctx.sheet, ref)); });
}

export function formulaForXlsx(formula,originalFormula) {
  const offset=formula.startsWith('=')?1:0;
  let tokens;try{tokens=tokenize(formula.slice(offset));}catch{return formula;}
  const prefixes=new Map();if(originalFormula){try{const original=tokenize(originalFormula.replace(/^=/,''));original.forEach((token,i)=>{if(token.type==='word'&&original[i+1]?.value==='('&&/_(?:xlfn|xlws)\./i.test(token.value))prefixes.set(normalizeFunctionName(token.value),token.value);});}catch{}}
  const patches=[];
  tokens.forEach((token,i)=>{
    if(token.type!=='word'||tokens[i+1]?.value!=='(')return;
    const original=prefixes.get(normalizeFunctionName(token.value));
    if(original&&!/_(?:xlfn|xlws)\./i.test(token.value))patches.push({at:token.start+offset,end:token.end+offset,value:original});
    else if(XLSX_FUTURE_FUNCTIONS.has(token.value.toUpperCase()))patches.push({at:token.start+offset,end:token.start+offset,value:'_xlfn.'});
  });
  for(const {at,end,value} of patches.reverse())formula=formula.slice(0,at)+value+formula.slice(end);
  return formula;
}

export function numberFormatKind(format) {
  if (!format || /^general$/i.test(format)) return 'general';
  const cleaned = format.replace(/"[^"]*"|\\./g, '');
  if (/[ymdhs]/i.test(cleaned.replace(/\[[^\]]*\]/g, '')) || /\[(?:h{1,2}|m{1,2}|s{1,2})\]/i.test(cleaned)) return 'date';
  if (cleaned.includes('%')) return 'percent';
  if (cleaned.includes('*')) return 'accounting';
  if (/[$₩€£¥]|\[\$/.test(cleaned)) return 'currency';
  return 'number';
}

function formatDateValue(value, pattern, dateOffset=0) {
  const { year, month, day } = serialParts(value+dateOffset);
  const fraction = ((value % 1) + 1) % 1;
  const seconds = Math.round(fraction * 86400) % 86400;
  const hour = Math.floor(seconds / 3600), minute = Math.floor(seconds / 60) % 60, second = seconds % 60;
  const pad = number => String(number).padStart(2, '0');
  const date = new Date(Date.UTC(year,month-1,1));
  const weekday=['Sunday','Monday','Tuesday','Wednesday','Thursday','Friday','Saturday'][((Math.floor(value)+dateOffset+6)%7+7)%7];
  const hasAmPm = /AM\/PM|A\/P/i.test(pattern);
  const tokens = [...pattern.matchAll(/"[^"]*"|\\.|\[[^\]]*\]|AM\/PM|A\/P|yyyy|yy|mmmm|mmm|mm|m|dddd|ddd|dd|d|hh|h|ss|s|./gi)];
  let previousTimeUnit = '';
  return tokens.map((match, index) => {
    const token = match[0], lower = token.toLowerCase();
    if (token.startsWith('"')) return token.slice(1, -1);
    if (token.startsWith('\\')) return token.slice(1);
    if (lower === '[h]' || lower === '[hh]') { previousTimeUnit = 'h'; return String(Math.floor(value * 24)).padStart(lower === '[hh]' ? 2 : 1, '0'); }
    if (lower === '[m]' || lower === '[mm]') return String(Math.floor(value * 1440)).padStart(lower === '[mm]' ? 2 : 1, '0');
    if (lower === '[s]' || lower === '[ss]') return String(Math.round(value * 86400)).padStart(lower === '[ss]' ? 2 : 1, '0');
    if (token.startsWith('[')) return '';
    if (lower === 'am/pm') return hour >= 12 ? 'PM' : 'AM';
    if (lower === 'a/p') return hour >= 12 ? 'P' : 'A';
    if (lower === 'yyyy') return String(year);
    if (lower === 'yy') return String(year).slice(-2);
    if (lower === 'mmmm') return date.toLocaleString('en-US', { month: 'long', timeZone: 'UTC' });
    if (lower === 'mmm') return date.toLocaleString('en-US', { month: 'short', timeZone: 'UTC' });
    if (lower === 'mm' || lower === 'm') {
      const nextUnit = tokens.slice(index + 1).map(item => item[0].toLowerCase()).find(item => /^(?:h{1,2}|m{1,4}|s{1,2}|d{1,4}|y{2,4})$/.test(item));
      const isMinute = previousTimeUnit === 'h' || nextUnit === 's' || nextUnit === 'ss';
      previousTimeUnit = isMinute ? 'minute' : 'month';
      return lower === 'mm' ? pad(isMinute ? minute : month) : String(isMinute ? minute : month);
    }
    if (lower === 'dddd' || lower === 'ddd') return lower==='dddd'?weekday:weekday.slice(0,3);
    if (lower === 'dd' || lower === 'd') { previousTimeUnit = 'd'; return lower === 'dd' ? pad(day) : String(day); }
    if (lower === 'hh' || lower === 'h') { previousTimeUnit = 'h'; const displayHour = hasAmPm ? hour % 12 || 12 : hour; return lower === 'hh' ? pad(displayHour) : String(displayHour); }
    if (lower === 'ss' || lower === 's') { previousTimeUnit = 's'; return lower === 'ss' ? pad(second) : String(second); }
    return token;
  }).join('');
}

export function formatValue(value, style = {}, options = {}) {
  if (value === null || value === undefined || value === '') return '';
  if (typeof value === 'boolean') return value ? 'TRUE' : 'FALSE';
  if (typeof value !== 'number' || !Number.isFinite(value)) return String(value);
  let format = typeof style === 'string' ? style : style.format || 'general';
  if (typeof style === 'object' && typeof style.numberFormat === 'string' && (!style.format || numberFormatKind(style.numberFormat) === style.format)) format = style.numberFormat;
  const decimals = typeof style === 'object' && Number.isInteger(style.decimals) ? Math.max(0, Math.min(12, style.decimals)) : null;
  if (format === 'general' || format === 'General' || format === '@') return String(Number(value.toPrecision(12)));
  const standardFormats = ['general', 'number', 'currency', 'accounting', 'percent', 'text'];
  if (format === 'date' || (!standardFormats.includes(format) && numberFormatKind(format) === 'date')) {
    try {
      return formatDateValue(value, format === 'date' ? 'yyyy-mm-dd' : format,options.date1904?1462:0);
    } catch { return '########'; }
  }
  const customDecimals = (format.split(';')[0].match(/\.([0#]+)/) || [])[1];
  const precision = decimals ?? (customDecimals ? Math.min(12, customDecimals.length) : ['number', 'currency', 'accounting', 'percent'].includes(format) ? 2 : 0);
  const minPrecision = decimals ?? (customDecimals ? Math.min(12, customDecimals.replace(/#/g, '').length) : precision);
  if (format === 'percent' || format.includes('%')) return `${(value * 100).toLocaleString('ko-KR', { minimumFractionDigits: minPrecision, maximumFractionDigits: precision })}%`;
  const currency = format === 'currency' || format === 'accounting' || /[$₩€£¥]/.test(format);
  const symbol = format.match(/[$₩€£¥]/)?.[0] || '₩';
  const number = value.toLocaleString('ko-KR', { useGrouping: currency || format === 'number' || format.includes(','), minimumFractionDigits: minPrecision, maximumFractionDigits: precision });
  return currency ? `${symbol}${number}` : number;
}

/** Rewrite references after one structural row/column insertion or deletion.
 * Absolute locks affect copying, not structural edits. Sheet qualifiers and
 * original formula whitespace are retained. Deleted ranges contract in place.
 */
export function transformFormulaReferences(formula, { axis, index, delta, targetSheet, ownerSheet }) {
  if (typeof formula !== 'string' || !['r', 'c'].includes(axis) || !Number.isInteger(index) || index < 0 || ![-1, 1].includes(delta)) return formula;
  const offset = formula.startsWith('=') ? 1 : 0;
  let tokens;
  try { tokens = tokenize(formula.slice(offset)); } catch { return formula; }
  const patches = [];
  const target = String(targetSheet).toLowerCase();
  const readReference = (start, inheritedSheet) => {
    let at = start, token = tokens[at], explicitSheet = null;
    if (!token) return null;
    if ((token.type === 'word' || token.type === 'sheet') && tokens[at + 1]?.value === '!') {
      explicitSheet = token.value; at += 2; token = tokens[at];
    }
    if (token?.type !== 'word' || tokens[at + 1]?.value === '(' || tokens[at + 1]?.value === '!') return null;
    const position = parseAddress(token.value);
    if (!position) return null;
    const locks = /^(\$?)[A-Za-z]+(\$?)/.exec(token.value);
    return { position, token, start: tokens[start].start, end: token.end, next: at + 1, sheet: explicitSheet ?? inheritedSheet ?? ownerSheet, colLock: locks[1], rowLock: locks[2] };
  };
  const patchCoordinate = (ref, n) => {
    const r = axis === 'r' ? n : ref.position.r, c = axis === 'c' ? n : ref.position.c;
    const replacement = validPosition(r, c) ? `${ref.colLock}${colName(c)}${ref.rowLock}${r + 1}` : '#REF!';
    if (replacement !== ref.token.value) patches.push({ start: ref.token.start, end: ref.token.end, replacement });
  };
  const patchSingle = ref => {
    if (String(ref.sheet).toLowerCase() !== target) return;
    const n = ref.position[axis];
    if (n < index) return;
    patchCoordinate(ref, delta === -1 && n === index ? -1 : n + delta);
  };
  for (let at = 0; at < tokens.length;) {
    const first = readReference(at);
    if (!first) { at++; continue; }
    const last = tokens[first.next]?.value === ':' ? readReference(first.next + 1, first.sheet) : null;
    if (last && String(first.sheet).toLowerCase() === target && String(last.sheet).toLowerCase() === target) {
      const n1 = first.position[axis], n2 = last.position[axis];
      if (delta === 1) { patchSingle(first); patchSingle(last); }
      else {
        const lo = Math.min(n1, n2), hi = Math.max(n1, n2);
        if (lo === index && hi === index) patches.push({ start: first.start, end: last.end, replacement: '#REF!' });
        else {
          const nextLo = lo > index ? lo - 1 : lo, nextHi = hi >= index ? hi - 1 : hi;
          patchCoordinate(first, n1 <= n2 ? nextLo : nextHi);
          patchCoordinate(last, n1 <= n2 ? nextHi : nextLo);
        }
      }
    } else {
      patchSingle(first);
      if (last) patchSingle(last);
    }
    at = last ? last.next : first.next;
  }
  let result = formula;
  for (const patch of patches.sort((a, b) => b.start - a.start)) result = result.slice(0, patch.start + offset) + patch.replacement + result.slice(patch.end + offset);
  return result;
}

// Rewrite explicit sheet references, leaving strings (including INDIRECT text) untouched.
// A null newName permanently invalidates references to a deleted worksheet.
export function rewriteSheetReferences(formula, oldName, newName) {
  const offset = formula.startsWith('=') ? 1 : 0;
  let tokens;
  try { tokens = tokenize(formula.slice(offset)); } catch { return formula; }
  const patches = [], target = oldName.toLowerCase();
  for (let i = 0; i < tokens.length - 2; i++) {
    const token = tokens[i];
    if (!['word', 'sheet'].includes(token.type) || tokens[i + 1].value !== '!' || token.value.toLowerCase() !== target) continue;
    if (newName !== null) {
      patches.push({start: token.start, end: token.end, value: "'" + newName.replaceAll("'", "''") + "'"});
      continue;
    }
    let end = i + 2;
    // Consume an inherited range endpoint as part of the invalid reference.
    if (tokens[end + 1]?.value === ':') {
      if (tokens[end + 3]?.value === '!' && tokens[end + 2].value.toLowerCase() === target && parseAddress(tokens[end + 4]?.value)) end += 4;
      else if (parseAddress(tokens[end + 2]?.value) && tokens[end + 3]?.value !== '!') end += 2;
    }
    patches.push({start: token.start, end: tokens[end].end, value: '#REF!'});
    i = end;
  }
  for (const patch of patches.reverse()) formula = formula.slice(0, patch.start + offset) + patch.value + formula.slice(patch.end + offset);
  return formula;
}

export function rewriteWorkbookSheetReferences(workbook, oldName, newName) {
  for(const named of workbook.chartNames||[])named.formula=rewriteSheetReferences(named.formula,oldName,newName);
  for (const sheet of workbook.sheets) {
    for (const cell of Object.values(sheet.cells)) {
      if (cell.raw?.startsWith('=')) cell.raw = rewriteSheetReferences(cell.raw, oldName, newName);
    }
    for (const rule of sheet.conditionalRules || []) {
      if (rule.formulae) rule.formulae = rule.formulae.map(formula => rewriteSheetReferences(formula, oldName, newName));
    }
  }
}

export function shiftFormula(formula, dr, dc) {
  if (typeof formula !== 'string' || !formula.startsWith('=') || (!dr && !dc)) return formula;
  if (!Number.isInteger(dr) || !Number.isInteger(dc)) return formula;
  let tokens;
  try { tokens = tokenize(formula.slice(1)); } catch { return formula; }
  const patches = [];
  for (let index = 0; index < tokens.length; index++) {
    const token = tokens[index];
    if (token.type !== 'word' || tokens[index + 1]?.value === '!' || tokens[index + 1]?.value === '(') continue;
    const match = /^(\$?)([A-Za-z]{1,3})(\$?)([1-9]\d*)$/.exec(token.value);
    if (!match) continue;
    const pos = parseAddress(token.value);
    if (!pos) continue;
    const r = pos.r + (match[3] ? 0 : dr), c = pos.c + (match[1] ? 0 : dc);
    const replacement = validPosition(r, c) ? `${match[1]}${colName(c)}${match[3]}${r + 1}` : '#REF!';
    patches.push({ start: token.start + 1, end: token.end + 1, replacement });
  }
  let result = formula;
  for (const patch of patches.reverse()) result = result.slice(0, patch.start) + patch.replacement + result.slice(patch.end);
  return result;
}

export function formulaForEditing(formula){
  let tokens;try{tokens=tokenize(formula);}catch{return formula;}
  for(let i=tokens.length-2;i>=0;i--){const t=tokens[i];if(t.type==='word'&&tokens[i+1]?.value==='('){const name=t.value.replace(/_(?:xlfn|xlws)\./gi,'');if(name!==t.value)formula=formula.slice(0,t.start)+name+formula.slice(t.end);}}
  return formula;
}
export function transposeFormula(formula,source,destination){
  if(!formula?.startsWith('='))return formula;
  let tokens;try{tokens=tokenize(formula.slice(1));}catch{return formula;}
  const patches=[];
  for(let i=0;i<tokens.length;i++){
    const t=tokens[i],m=/^(\$?)([A-Za-z]{1,3})(\$?)([1-9]\d*)$/.exec(t.value);
    if(t.type!=='word'||!m||['!','('].includes(tokens[i+1]?.value))continue;
    const p=parseAddress(t.value);if(!p)continue;
    if(m[1]&&m[3])continue;
    const r=m[1]?p.c:destination.r+p.c-source.c,c=m[3]?p.r:destination.c+p.r-source.r;
    patches.push({start:t.start+1,end:t.end+1,value:validPosition(r,c)?`${m[3]}${colName(c)}${m[1]}${r+1}`:'#REF!'});
  }
  for(const p of patches.reverse())formula=formula.slice(0,p.start)+p.value+formula.slice(p.end);
  return formula;
}
export function moveFormulaReferences(formula,{sourceSheet,targetSheet,ownerSheet,range,dr,dc,originSheet}){
  if(typeof formula!=='string')return formula;
  const offset=formula.startsWith('=')?1:0;let tokens;try{tokens=tokenize(formula.slice(offset));}catch{return formula;}
  const patches=[],quote=name=>"'"+name.replaceAll("'","''")+"'!";
  const read=(i,inherited)=>{
    const first=tokens[i];let token=first,explicit=null;
    if(['word','sheet'].includes(token?.type)&&tokens[i+1]?.value==='!'){explicit=token.value;token=tokens[i+2];i+=2;}
    const pos=token?.type==='word'&&tokens[i+1]?.value!=='('?parseAddress(token.value):null;if(!pos)return null;
    const m=/^(\$?)[A-Za-z]+(\$?)/.exec(token.value);
    return {pos,token,start:first.start,end:token.end,next:i+1,sheet:explicit||inherited||originSheet||ownerSheet,explicit,colLock:m[1],rowLock:m[2]};
  };
  const point=(r,c,ref)=>`${ref.colLock}${colName(c)}${ref.rowLock}${r+1}`;
  const inside=p=>p.r>=range.r1&&p.r<=range.r2&&p.c>=range.c1&&p.c<=range.c2;
  for(let i=0;i<tokens.length;){
    const a=read(i);if(!a){i++;continue;}const z=tokens[a.next]?.value===':'?read(a.next+1,a.sheet):null;
    const prefix=name=>name.toLowerCase()!==ownerSheet.toLowerCase()||a.explicit||originSheet?quote(name):'';
    let value;
    if(a.sheet.toLowerCase()===sourceSheet.toLowerCase()&&(!z||z.sheet.toLowerCase()===sourceSheet.toLowerCase())){
      if(!z&&inside(a.pos))value=prefix(targetSheet)+point(a.pos.r+dr,a.pos.c+dc,a);
      if(z){
        const box={r1:Math.min(a.pos.r,z.pos.r),c1:Math.min(a.pos.c,z.pos.c),r2:Math.max(a.pos.r,z.pos.r),c2:Math.max(a.pos.c,z.pos.c)};
        const cut={r1:Math.max(box.r1,range.r1),c1:Math.max(box.c1,range.c1),r2:Math.min(box.r2,range.r2),c2:Math.min(box.c2,range.c2)};
        if(cut.r1<=cut.r2&&cut.c1<=cut.c2){
          if(inside(a.pos)&&inside(z.pos))value=prefix(targetSheet)+point(a.pos.r+dr,a.pos.c+dc,a)+':'+point(z.pos.r+dr,z.pos.c+dc,z);
          else{
            const subtract=(box,cut)=>{
              const r1=Math.max(box.r1,cut.r1),r2=Math.min(box.r2,cut.r2),c1=Math.max(box.c1,cut.c1),c2=Math.min(box.c2,cut.c2);
              if(r1>r2||c1>c2)return [box];
              return [{...box,r2:r1-1},{...box,r1:r2+1},{r1,r2,c1:box.c1,c2:c1-1},{r1,r2,c1:c2+1,c2:box.c2}].filter(b=>b.r1<=b.r2&&b.c1<=b.c2);
            };
            const moved={r1:cut.r1+dr,r2:cut.r2+dr,c1:cut.c1+dc,c2:cut.c2+dc};
            let remaining=subtract(box,cut);if(sourceSheet.toLowerCase()===targetSheet.toLowerCase())remaining=remaining.flatMap(b=>subtract(b,moved));
            const rect=(b,name)=>prefix(name)+point(b.r1,b.c1,a)+':'+point(b.r2,b.c2,z);
            value='('+[...remaining.map(b=>rect(b,sourceSheet)),rect(moved,targetSheet)].join(',')+')';
          }
        }
      }
    }
    if(value===undefined&&originSheet&&!a.explicit)value=quote(a.sheet)+formula.slice(a.token.start+offset,(z?.end??a.end)+offset);
    if(value!==undefined)patches.push({start:a.start,end:z?.end??a.end,value});
    i=z?.next??a.next;
  }
  for(const p of patches.reverse())formula=formula.slice(0,p.start+offset)+p.value+formula.slice(p.end+offset);
  return formula;
}

export function parseCSV(text) {
  if (typeof text !== 'string') throw new TypeError('CSV 텍스트가 필요합니다.');
  text = text.replace(/^\uFEFF/, '');
  if (text.length === 0) return [];
  const rows = [];
  let row = [], field = '', quoted = false, afterQuote = false, atFieldStart = true;
  const pushField = () => { row.push(field); field = ''; atFieldStart = true; afterQuote = false; };
  for (let i = 0; i < text.length; i++) {
    const ch = text[i];
    if (quoted) {
      if (ch === '"') {
        if (text[i + 1] === '"') { field += '"'; i++; }
        else { quoted = false; afterQuote = true; }
      } else field += ch;
    } else if (ch === '"' && atFieldStart) { quoted = true; atFieldStart = false; }
    else if (ch === ',') pushField();
    else if (ch === '\r' || ch === '\n') {
      pushField(); rows.push(row); row = [];
      if (ch === '\r' && text[i + 1] === '\n') i++;
    } else {
      if (afterQuote && !/\s/.test(ch)) throw new Error('CSV 따옴표 뒤에 잘못된 문자가 있습니다.');
      if (!afterQuote) field += ch;
      atFieldStart = false;
    }
  }
  if (quoted) throw new Error('CSV 따옴표가 닫히지 않았습니다.');
  if (field !== '' || row.length > 0 || !/[\r\n]$/.test(text)) { pushField(); rows.push(row); }
  return rows;
}

export function toCSV(rows) {
  if (!Array.isArray(rows)) throw new TypeError('CSV 행 배열이 필요합니다.');
  return rows.map(row => row.map(value => {
    const text = value === null || value === undefined ? '' : String(value);
    return /[",\r\n]/.test(text) ? `"${text.replace(/"/g, '""')}"` : text;
  }).join(',')).join('\r\n');
}

export function validateCell(ref, cell) {
  const object=value=>value!==null&&typeof value==='object'&&!Array.isArray(value);
  if (!parseAddress(ref) || ref.includes('$') || ref !== ref.toUpperCase() || !object(cell) || typeof cell.raw !== 'string' || cell.raw.length > 32768) throw new Error(`셀 데이터가 올바르지 않습니다: ${ref}`);
  if (cell.style !== undefined && !object(cell.style)) throw new Error(`셀 서식이 올바르지 않습니다: ${ref}`);
  if (cell.cachedFormula !== undefined && (typeof cell.cachedFormula !== 'string' || cell.cachedFormula.length > 32768)) throw new Error(`저장된 수식이 올바르지 않습니다: ${ref}`);
  if(cell.xlsxStyle!==undefined&&(!Number.isInteger(cell.xlsxStyle)||cell.xlsxStyle<0||cell.xlsxStyle>=65536))throw Error(`보존된 셀 서식이 올바르지 않습니다: ${ref}`);
  if(cell.xlsxValue!==undefined){const v=cell.xlsxValue;if(!object(v)||typeof v.raw!=='string'||v.raw.length>32768||!object(v.value)||!v.value.richText&&!v.value.hyperlink)throw Error(`보존된 셀 값이 올바르지 않습니다: ${ref}`);validatePreservedData(v,1_000_000);}
  if(cell.xlsxFormula!==undefined){const f=cell.xlsxFormula;if(!object(f)||typeof f.raw!=='string'||f.raw.length>32768||typeof f.formula!=='string'||f.formula.length>32768||f.shareType!==undefined&&f.shareType!=='array'||f.ref!==undefined&&(typeof f.ref!=='string'||!/^\$?[A-Z]+\$?\d+(?::\$?[A-Z]+\$?\d+)?$/.test(f.ref)))throw Error(`보존된 배열 수식이 올바르지 않습니다: ${ref}`);}
  if (cell.style) {
    if (cell.style.borders !== undefined) validateBorders(cell.style.borders);
    for (const field of ['bold', 'italic', 'underline', 'wrap', 'border']) if (cell.style[field] !== undefined && typeof cell.style[field] !== 'boolean') throw new Error(`셀 서식이 올바르지 않습니다: ${ref}`);
    for (const field of ['fontFamily', 'color', 'background', 'align', 'verticalAlign', 'format', 'numberFormat']) if (cell.style[field] !== undefined && (typeof cell.style[field] !== 'string' || cell.style[field].length > 512)) throw new Error(`셀 서식이 올바르지 않습니다: ${ref}`);
    if (cell.style.fontSize !== undefined && (typeof cell.style.fontSize !== 'number' || !Number.isFinite(cell.style.fontSize) || cell.style.fontSize <= 0 || cell.style.fontSize > 4096)) throw new Error(`글꼴 크기가 올바르지 않습니다: ${ref}`);
    if (cell.style.decimals !== undefined && (!Number.isInteger(cell.style.decimals) || cell.style.decimals < 0 || cell.style.decimals > 12)) throw new Error(`소수점 자릿수가 올바르지 않습니다: ${ref}`);
    for (const field of ['color', 'background']) if (cell.style[field] && !/^(?:#[\da-f]{3,8}|[a-z]+|(?:rgb|hsl)a?\([\d.%+,\s/-]+\))$/i.test(cell.style[field])) throw new Error(`셀 색상이 올바르지 않습니다: ${ref}`);
    if (cell.style.fontFamily && /[;{}<>]/.test(cell.style.fontFamily)) throw new Error(`글꼴 이름이 올바르지 않습니다: ${ref}`);
  }
}

export function validateWorkbookMetadata(book) {
  return validateWorkbook({...book,sheets:book.sheets.map(sheet=>({...sheet,cells:{}}))});
}

export function validateWorkbook(data) {
  const object = value => value && typeof value === 'object' && !Array.isArray(value);
  const validRange = value => {
    if (typeof value !== 'string') return false;
    const parts = value.split(':');
    if (parts.length !== 2) return false;
    const [first, last] = parts.map(parseAddress);
    return first && last && first.r <= last.r && first.c <= last.c;
  };
  if (!object(data) || !Array.isArray(data.sheets) || data.sheets.length < 1 || data.sheets.length > 100) throw new Error('유효한 통합 문서가 아닙니다.');
  if (typeof data.name !== 'string' || data.name.length > 255) throw new Error('통합 문서 이름이 올바르지 않습니다.');
  if(data.date1904!==undefined&&typeof data.date1904!=='boolean')throw new Error('통합 문서의 날짜 체계가 올바르지 않습니다.');
  const ids = new Set(), names = new Set();
  let cellCount = 0, textCount = 0;
  for (const sheet of data.sheets) {
    if (!object(sheet) || typeof sheet.id !== 'string' || !sheet.id || ids.has(sheet.id)) throw new Error('시트 ID가 올바르지 않습니다.');
    if (typeof sheet.name !== 'string' || !sheet.name.trim() || sheet.name.length > 31 || /[:\\/?*\[\]]/.test(sheet.name) || names.has(sheet.name.toLowerCase())) throw new Error('시트 이름이 올바르지 않습니다.');
    ids.add(sheet.id); names.add(sheet.name.toLowerCase());
    for(const key of ['defaultColumnWidth','defaultRowHeight'])if(sheet[key]!==undefined&&(!Number.isFinite(sheet[key])||sheet[key]<1||sheet[key]>10000))throw new Error('시트 기본 크기가 올바르지 않습니다.');
    if(sheet.hiddenSizes!==undefined){if(!object(sheet.hiddenSizes))throw Error('숨김 크기가 올바르지 않습니다.');for(const k of ['rows','cols']){if(!object(sheet.hiddenSizes[k]))throw Error('숨김 크기가 올바르지 않습니다.');for(const [i,v] of Object.entries(sheet.hiddenSizes[k]))if(!/^\d+$/.test(i)||+i>=(k==='rows'?1048576:16384)||!Number.isFinite(v)||v<0||v>10000)throw Error('숨김 크기가 올바르지 않습니다.');}}
    if(sheet.xlsxFeatures!==undefined){if(!object(sheet.xlsxFeatures))throw Error('보존된 시트 정보가 올바르지 않습니다.');for(const key of ['dataValidations','sheetProtection'])if(sheet.xlsxFeatures[key]!==undefined)validatePreservedXML(sheet.xlsxFeatures[key],key);if(sheet.xlsxFeatures.conditionalFormatting!==undefined){const rules=sheet.xlsxFeatures.conditionalFormatting;if(!Array.isArray(rules)||rules.length>10000)throw Error('보존된 조건부 서식이 올바르지 않습니다.');for(const rule of rules)validatePreservedXML(rule,'conditionalFormatting');}}
    if(sheet.freeze!==undefined)validateFreeze(sheet.freeze);
    if(sheet.conditionalRules!==undefined)validateConditionalRules(sheet.conditionalRules);
    if(sheet.unsupportedConditionalCount!==undefined&&(!Number.isInteger(sheet.unsupportedConditionalCount)||sheet.unsupportedConditionalCount<0||sheet.unsupportedConditionalCount>10000))throw Error('조건부 서식 개수가 올바르지 않습니다.');
    if (sheet.autoFilter !== undefined) validateAutoFilter(sheet.autoFilter);
    if (sheet.pivot !== undefined) { validatePivot(sheet.pivot); if (sheet.pivot.sourceSheetId === sheet.id) throw new Error('피벗 원본은 다른 시트여야 합니다.'); }
    if (!object(sheet.cells)) throw new Error('셀 데이터가 올바르지 않습니다.');
    for (const [ref, cell] of Object.entries(sheet.cells)) {
      if (++cellCount > MAX_DOCUMENT_CELLS) throw new Error('통합 문서의 셀 수가 제한을 초과했습니다.');
      validateCell(ref, cell);
      if((textCount+=cell.raw.length)>MAX_DOCUMENT_TEXT_CHARS)throw Error('문서의 셀 텍스트 총량이 제한을 초과했습니다.');
    }
    for (const field of ['columnWidths', 'rowHeights']) {
      if (sheet[field] === undefined) sheet[field] = {};
      if (!object(sheet[field])) throw new Error('행 또는 열 크기가 올바르지 않습니다.');
      for (const size of Object.values(sheet[field])) if (typeof size !== 'number' || !Number.isFinite(size) || size < 0 || size > 10000) throw new Error('행 또는 열 크기가 올바르지 않습니다.');
    }
    for (const field of ['merges', 'charts']) {
      if (sheet[field] === undefined) sheet[field] = [];
      if (!Array.isArray(sheet[field]) || sheet[field].length > 10000) throw new Error('시트 개체가 올바르지 않습니다.');
    }
    if (sheet.merges.some(range => !validRange(range))) throw new Error('병합 범위가 올바르지 않습니다.');
    for (const chart of sheet.charts) {
      validateChart(chart);
      if (!object(chart) || !validRange(chart.range)) throw new Error('차트 데이터 범위가 올바르지 않습니다.');
      for (const field of ['id', 'title', 'type']) if (chart[field] !== undefined && typeof chart[field] !== 'string') throw new Error('차트 속성이 올바르지 않습니다.');
      for (const field of ['width', 'height']) if (chart[field] !== undefined && (typeof chart[field] !== 'number' || !Number.isFinite(chart[field]) || chart[field] < 100 || chart[field] > 10000)) throw new Error('차트 크기가 올바르지 않습니다.');
      for (const field of ['x', 'y']) if (chart[field] !== undefined && (typeof chart[field] !== 'number' || !Number.isFinite(chart[field]) || chart[field] < 0)) throw new Error('차트 위치가 올바르지 않습니다.');
    }
  }
  if (typeof data.activeSheetId !== 'string' || !ids.has(data.activeSheetId)) throw new Error('활성 시트가 없습니다.');
  if(data.xlsxStyles!==undefined){if(!Array.isArray(data.xlsxStyles)||data.xlsxStyles.length>65536||data.xlsxStyles.some(s=>!object(s)||!object(s.style)||!object(s.baseline)))throw Error('보존된 Excel 서식 목록이 올바르지 않습니다.');validatePreservedData(data.xlsxStyles,32_000_000);}
  if(data.xlsxDxfs!==undefined){if(!Array.isArray(data.xlsxDxfs)||data.xlsxDxfs.length>65536)throw Error('보존된 Excel 조건부 서식이 올바르지 않습니다.');for(const rule of data.xlsxDxfs)validatePreservedXML(rule,'dxf');}
  if(data.chartTheme!==undefined&&(typeof data.chartTheme!=='string'||data.chartTheme.length>2000000||/<!DOCTYPE|<!ENTITY/i.test(data.chartTheme)))throw new Error('차트 테마가 올바르지 않습니다.');
  if(data.chartNames!==undefined&&(!Array.isArray(data.chartNames)||data.chartNames.length>10000||data.chartNames.some(n=>!object(n)||typeof n.name!=='string'||n.name.length>255||typeof n.formula!=='string'||n.formula.length>8192||n.sheetId!==undefined&&typeof n.sheetId!=='string'||n.hidden!==undefined&&typeof n.hidden!=='boolean')))throw new Error('차트 이름 참조가 올바르지 않습니다.');
  if(data.chartNameSheets!==undefined&&(!Array.isArray(data.chartNameSheets)||data.chartNameSheets.length>100||data.chartNameSheets.some(s=>!object(s)||typeof s.id!=='string'||typeof s.name!=='string'||s.name.length>31)))throw new Error('차트 원본 시트가 올바르지 않습니다.');
  if (data.nativeSource !== undefined) {
    const source=data.nativeSource;
    if(!object(source)||source.format!=='xlsm'||typeof source.filename!=='string'||source.filename.length>255||typeof source.base64!=='string'||source.base64.length>20*1024*1024||!/^[A-Za-z0-9+/]*={0,2}$/.test(source.base64)||!Array.isArray(source.baseline)||!source.baseline.length||!Array.isArray(source.operations)||source.operations.length>10000)throw new Error('VBA 원본 데이터가 올바르지 않습니다.');
    if(source.originToken!==undefined&&(typeof source.originToken!=='string'||source.originToken.length>150))throw new Error('VBA 출처 정보가 올바르지 않습니다.');
    validateWorkbook({name:data.name,activeSheetId:source.baseline[0].id,sheets:source.baseline});
    for(const op of source.operations)if(!object(op)||typeof op.sheetId!=='string'||!['r','c'].includes(op.axis)||!Number.isInteger(op.index)||op.index<0||op.index>=(op.axis==='r'?1048576:16384)||![-1,1].includes(op.delta))throw new Error('VBA 구조 변경 정보가 올바르지 않습니다.');
  }
  return data;
}

export function parseWorkbook(text) {
  if (typeof text !== 'string' || text.length > 100000000) throw new Error('통합 문서 크기가 제한을 초과했습니다.');
  const data = JSON.parse(text, (key, value) => {
    if (key === '__proto__' || key === 'prototype' || key === 'constructor') throw new Error('안전하지 않은 통합 문서 키입니다.');
    return value;
  });
  return validateWorkbook(data);
}

export function serializeWorkbook(workbook) {
  const text = JSON.stringify(workbook);
  parseWorkbook(text);
  return text;
}
