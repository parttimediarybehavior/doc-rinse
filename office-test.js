import test from 'node:test';
import assert from 'node:assert/strict';
import { makeOfficeFixture, officeInflate } from './tests/office-fixtures.js';
let office = {};
try { office = await import('./src/office.js'); } catch(e) { if(e.code !== 'ERR_MODULE_NOT_FOUND') throw e; }
const analyze = async (bytes) => { assert.equal(typeof office.analyzeOffice,'function','read-only Office analyzer must exist'); return office.analyzeOffice(bytes,{inflate:officeInflate,name:'renamed.bin'}); };

test('XLSX inspection finds candidate data without modifying input',async()=>{
 const bytes=await makeOfficeFixture('xlsx'), before=bytes.slice(); const r=await analyze(bytes);
 assert.equal(r.format,'xlsx'); assert.deepEqual(bytes,before);
 assert.deepEqual(r.sheets.map(s=>s.state),['visible','hidden','veryHidden']);
 assert.equal(r.sheets[1].name,'<img src=x>'); assert.equal(r.sheets[2].part,'xl/worksheets/sheet3.xml');
 assert.equal(r.metadata.creator,'Alice & Bob'); assert.equal(r.metadata.lastModifiedBy,'Carol');
 assert.equal(r.definedNames[0].formula,'Secret!$A$1'); assert.equal(r.definedNames[0].scope,'0');
 assert.equal(r.formulaCaches[0].value,'2'); assert.equal(r.calcChains.length,1);
 assert.equal(r.comments.length,2); assert.match(r.comments.map(c=>c.text).join(' '),/Secret thread/);
 assert.equal(r.authors.some(a=>a.name==='Dana'),true);
 assert.equal(r.external[0].target,'file:///private/source.xlsx'); assert.equal(r.thumbnails.length,1);
 assert.equal(r.issues.length,0);
});
test('PPTX inspection honors presentation order and separates notes from masters',async()=>{
 const bytes=await makeOfficeFixture('pptx'),before=bytes.slice();const r=await analyze(bytes);
 assert.equal(r.format,'pptx');assert.deepEqual(bytes,before);
 assert.equal(r.slides[0].part,'ppt/slides/slide2.xml');assert.equal(r.slides[0].hidden,true);
 assert.equal(r.slides[1].hidden,false);assert.equal(r.notes.length,1);assert.equal(r.notes[0].text,'Secret speaker note');
 assert.equal(r.comments[0].text,'Secret slide comment');assert.deepEqual(r.orphanMedia,['ppt/media/orphan.png']);
 assert.equal(r.issues.length,0);
});
test('broken sheet target is reported without asserting absence of hidden data',async()=>{
 const r=await analyze(await makeOfficeFixture('xlsx',{'xl/worksheets/sheet2.xml':null}));
 assert.match(r.issues.join(' '),/sheet2.xml/);assert.equal(r.sheets[1].missing,true);
});
test('ambiguous package is rejected',async()=>{
 const b=await makeOfficeFixture('xlsx',{'ppt/presentation.xml':'<presentation/>'});
 await assert.rejects(()=>analyze(b),/ambiguous/i);
});

import { readZip, parseXML, walk } from './src/zipxml.js';
import { officeDeflate } from './tests/office-fixtures.js';
const cleanOffice=async(bytes,rep,fixes)=>{assert.equal(typeof office.scrubOffice,'function','selected Office cleanup must exist');return office.scrubOffice(bytes,rep,{inflate:officeInflate,deflate:officeDeflate,fixes});};
for(const kind of ['xlsx','pptx']) {
 test(`${kind}: cleanup is opt-in; individual removals preserve unrelated content`,async()=>{
  const bytes=await makeOfficeFixture(kind),r=await analyze(bytes);
  assert.equal(typeof office.planOffice,'function');assert.ok(office.planOffice(r).every(p=>!p.on));
  assert.deepEqual((await cleanOffice(bytes,r,[])).bytes,bytes);
  const z=await readZip(bytes,officeInflate);
  for(const action of ['identity','comments','thumbnail',...(kind==='pptx'?['notes']:[])]) {
   const result=await cleanOffice(bytes,r,[action]);const after=await analyze(result.bytes);const cleaned=await readZip(result.bytes,officeInflate);
   assert.equal(after.issues.length,0,JSON.stringify(after.issues));
   assert.equal(after.metadata.title,'Keep title');
   if(action==='identity') {assert.equal(after.metadata.creator,null);assert.equal(after.metadata.lastModifiedBy,null);assert.equal(after.metadata.Company,null);}
   else assert.equal(after.metadata.creator,r.metadata.creator);
   if(action==='comments') {assert.equal(after.comments.length,0);assert.equal(after.authors.length,0);}
   else assert.equal(after.comments.length,r.comments.length);
   if(action==='thumbnail')assert.equal(after.thumbnails.length,0);
   if(action==='notes')assert.equal(after.notes.length,0);
   for(const p of z.order.filter(n=>/^(xl\/worksheets\/[^/]+\.xml|xl\/workbook.xml|xl\/externalLinks\/[^/]+\.xml|xl\/calcChain.xml|ppt\/slides\/[^/]+\.xml|ppt\/presentation.xml|ppt\/media\/)/.test(n)))assert.deepEqual(await cleaned.bytes(p),await z.bytes(p),p);
   const ct=parseXML(await cleaned.text('[Content_Types].xml'));
   for(const n of walk(ct)) if(n.tag.split(':').pop()==='Override')assert.ok(cleaned.has(n.attrs.PartName.slice(1)),n.attrs.PartName);
   const twice=await cleanOffice(result.bytes,after,[action]);assert.deepEqual(twice.bytes,result.bytes,'idempotent '+action);
  }
 });
}
test('signed and unresolved packages cannot be modified under a clean verdict',async()=>{
 const signed=await makeOfficeFixture('xlsx',{'_xmlsignatures/sig1.xml':'<Signature/>'});
 await assert.rejects(()=>analyze(signed).then(r=>cleanOffice(signed,r,['identity'])),/signed/i);
 const broken=await makeOfficeFixture('xlsx',{'xl/worksheets/sheet2.xml':null});
 await assert.rejects(()=>analyze(broken).then(r=>cleanOffice(broken,r,['identity'])),/unresolved/i);
 const plain=await makeOfficeFixture('xlsx');
 await assert.rejects(()=>analyze(plain).then(r=>cleanOffice(plain,r,['hidden-sheets'])),/unsupported/i);
});
