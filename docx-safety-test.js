import test from 'node:test';
import assert from 'node:assert/strict';
import { writeZip } from './src/zipwrite.js';
import { readZip } from './src/zipxml.js';
import { analyze } from './src/engine.js';
import { scrub, plan } from './src/scrub.js';
import { officeInflate as inflate, officeDeflate as deflate } from './tests/office-fixtures.js';
const enc=new TextEncoder();
async function docx(parts){return writeZip(Object.entries({'[Content_Types].xml':'<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="png" ContentType="image/png"/></Types>',...parts}).map(([name,data])=>({name,data:typeof data==='string'?enc.encode(data):data})),{deflate});}
const document=body=>`<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:v="urn:schemas-microsoft-com:vml"><w:body>${body}</w:body></w:document>`;
const rel=target=>`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="im" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="${target}"/></Relationships>`;
test('language cleanup describes proofing tags without asserting origin or visual preservation',()=>{
 const action=plan({raw:{lang:3}}).find(p=>p.id==='locales');
 assert.match(action.detail,/proofing/i);
 assert.match(action.detail,/do not establish/i);
 assert.doesNotMatch(action.detail,/records the keyboard|assembled from three places|changes no pixel/);
});
for(const [owner,target,body] of [
 ['word/_rels/footer1.xml.rels','media/image1.png','<w:p><w:r><w:t>body</w:t></w:r></w:p>'],
 ['word/_rels/header1.xml.rels','/word/media/image1.png','<w:p/>'],
 ['word/_rels/document.xml.rels','media/image1.png','<w:p><w:r><w:object><v:imagedata r:id="im"/></w:object></w:r></w:p>'],
])test(`referenced image survives default cleanup: ${owner} ${target}`,async()=>{
 const input=await docx({'word/document.xml':document(body),[owner]:rel(target),'word/footer1.xml':'<w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>','word/header1.xml':'<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>','word/media/image1.png':new Uint8Array([1,2,3]),'word/media/orphan.png':new Uint8Array([4,5,6])});
 const r=await analyze(input,{inflate});assert.equal(r.media.find(m=>m.part==='word/media/image1.png').unused,false);
 assert.equal(r.media.find(m=>m.part==='word/media/orphan.png').unused,true);
 const out=await scrub(input,r,{inflate,deflate});const z=await readZip(out.bytes,inflate);
 assert.deepEqual(await z.bytes('word/media/image1.png'),new Uint8Array([1,2,3]));assert.equal(z.has('word/media/orphan.png'),false);
});

test('white and tiny text remain by default; explicit hidden text is removed',async()=>{
 const body='<w:p><w:r><w:rPr><w:vanish/></w:rPr><w:t>same words</w:t></w:r></w:p><w:p><w:r><w:rPr><w:color w:val="FFFFFF"/></w:rPr><w:t>same words</w:t></w:r></w:p><w:p><w:r><w:rPr><w:sz w:val="2"/></w:rPr><w:t>Tiny visible caption</w:t></w:r></w:p>';
 const bytes=await docx({'word/document.xml':document(body)}),r=await analyze(bytes,{inflate});
 assert.equal(plan(r).find(p=>p.id==='appearance')?.on,false,'appearance deletion is explicit opt-in');
 const out=await scrub(bytes,r,{inflate,deflate});const z=await readZip(out.bytes,inflate);const xml=await z.text('word/document.xml');
 assert.equal((xml.match(/same words/g)||[]).length,1);assert.match(xml,/Tiny visible caption/);assert.doesNotMatch(xml,/<w:vanish/);
 const after=await analyze(out.bytes,{inflate});assert.ok(after.hiddenText.every(h=>h.uncertain));
 const explicit=await scrub(bytes,r,{inflate,deflate,fixes:['appearance']});const ez=await readZip(explicit.bytes,inflate);const ex=await ez.text('word/document.xml');
 assert.match(ex,/<w:vanish/);assert.doesNotMatch(ex,/Tiny visible caption/);assert.equal((ex.match(/same words/g)||[]).length,1);
});
