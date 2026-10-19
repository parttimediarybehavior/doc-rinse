// Execute the built bundle against a minimal DOM. This is not a browser/layout test.
import assert from 'node:assert/strict';
import vm from 'node:vm';
import { readFileSync } from 'node:fs';
import { officeFixtureParts } from './office-fixtures.js';
import { writeZip } from '../src/zipwrite.js';
class Element {
 constructor(tag){this.tag=tag;this.children=[];this.textContent='';this.classList={add(){},remove(){}};this.dataset={};}
 appendChild(node){this.children.push(node);return node;}
 set innerHTML(value){assert.equal(value,'','Office UI must not interpolate file content as HTML');this.children=[];}
 addEventListener(){}
 click(){this.clicked=true;}
}
const elements=Object.fromEntries(['#intro','#out','#drop','#file'].map(k=>[k,new Element(k)]));
const document={querySelector:selector=>elements[selector],querySelectorAll:()=>[],createElement:tag=>new Element(tag)};
const html=readFileSync(new URL('../dist/doc-rinse.html',import.meta.url),'utf8');
const script=html.match(/<script[^>]*>([\s\S]*)<\/script>/)[1];
let downloads=0;
class TestURL extends URL {static createObjectURL(blob){assert.ok(blob.size);downloads++;return 'blob:test';}static revokeObjectURL(){}}
const context=vm.createContext({document,window:{addEventListener(){},scrollTo(){}},location:{search:''},URLSearchParams,URL:TestURL,TextEncoder,TextDecoder,Uint8Array,DataView,ArrayBuffer,Blob,console,setTimeout:()=>0,performance,atob,btoa});
vm.runInContext(script,context);
const all=n=>[n,...n.children.flatMap(all)];
for(const kind of ['xlsx','pptx']){
 const entries=Object.entries(officeFixtureParts(kind)).map(([name,data])=>({name,data:typeof data==='string'?new TextEncoder().encode(data):data}));
 const bytes=await writeZip(entries,{deflate:null});
 context.inputBytes=bytes;
 await vm.runInContext(`(async()=>{const r=await analyzePackage(inputBytes,{name:'unsafe <img src=x>.${kind}'});r._bytes=inputBytes;renderReport(r,r.file.name);})()`,context);
 const nodes=all(elements['#out']);
 assert.ok(nodes.some(n=>n.textContent.includes?.(`unsafe <img src=x>.${kind}`)));
 assert.ok(!nodes.some(n=>n.tag==='img'),'file text cannot create image elements');
 const inputs=nodes.filter(n=>n.tag==='input');assert.ok(inputs.length>0);assert.ok(inputs.every(n=>n.checked===false));
 const go=nodes.find(n=>n.tag==='button' && n.textContent==='clean selected fields');
 const before=downloads;await go.onclick();assert.equal(downloads,before,'no selection must not download');
 inputs[0].checked=true;await go.onclick();assert.equal(downloads,before+1);
 assert.ok(all(elements['#out']).some(n=>n.textContent.startsWith?.('Re-analysis:')));
 assert.ok(all(elements['#out']).some(n=>n.textContent.includes?.('will remain')));
 console.log(kind+': built-bundle dispatch, escaped text, opt-in cleanup and re-analysis PASS');
}
