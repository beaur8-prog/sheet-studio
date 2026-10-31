import test from 'node:test';
import assert from 'node:assert/strict';
import {createDocumentLocation} from '../src/document-location.mjs';

test('a blank startup gets a new location identity while explicit recovery can restore the original folder identity',()=>{
 const previous=crypto.randomUUID(),values=new Map([['sheet-studio.document-location',previous]]),storage={getItem:key=>values.get(key),setItem:(key,value)=>values.set(key,value)};
 let book={};const fresh=createDocumentLocation(()=>book,storage,{restore:false});
 assert.notEqual(fresh.id(),previous);
 book={};fresh.carry(previous);assert.equal(fresh.id(),previous);
 const continued=createDocumentLocation(()=>book,storage);assert.equal(continued.id(),previous);
});
