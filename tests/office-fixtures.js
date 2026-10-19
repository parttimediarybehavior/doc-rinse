import { writeZip } from '../src/zipwrite.js';
import zlib from 'node:zlib';
export const officeInflate = async b => new Uint8Array(zlib.inflateRawSync(b));
export const officeDeflate = async b => new Uint8Array(zlib.deflateRawSync(b));
const enc = new TextEncoder();
const relNS = 'http://schemas.openxmlformats.org/package/2006/relationships';
const typesNS = 'http://schemas.openxmlformats.org/package/2006/content-types';
const docRel = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/';
export function officeFixtureParts(kind) {
  const main = kind === 'xlsx' ? 'xl/workbook.xml' : 'ppt/presentation.xml';
  const parts = {
    '_rels/.rels': `<Relationships xmlns="${relNS}"><Relationship Id="office" Type="${docRel}officeDocument" Target="${main}"/><Relationship Id="core" Type="${docRel}metadata/core-properties" Target="docProps/core.xml"/><Relationship Id="thumb" Type="${docRel}metadata/thumbnail" Target="docProps/thumbnail.png"/></Relationships>`,
    'docProps/core.xml': '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:d="http://purl.org/dc/elements/1.1/"><d:creator>Alice &amp; Bob</d:creator><cp:lastModifiedBy>Carol</cp:lastModifiedBy><d:title>Keep title</d:title></cp:coreProperties>',
    'docProps/app.xml': '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"><Company>Example</Company><Application>Fixture</Application></Properties>',
    'docProps/thumbnail.png': new Uint8Array([1,2,3]),
  };
  if (kind === 'xlsx') Object.assign(parts, {
    'xl/workbook.xml': '<s:workbook xmlns:s="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:q="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><s:sheets><s:sheet name="Visible" sheetId="1" q:id="r1"/><s:sheet name="&lt;img src=x&gt;" sheetId="2" state="hidden" q:id="r2"/><s:sheet name="Very hidden" sheetId="3" state="veryHidden" q:id="r3"/></s:sheets><s:definedNames><s:definedName name="PrivatePath" localSheetId="0">Secret!$A$1</s:definedName></s:definedNames><s:externalReferences><s:externalReference q:id="ext"/></s:externalReferences></s:workbook>',
    'xl/_rels/workbook.xml.rels': `<Relationships xmlns="${relNS}"><Relationship Id="r1" Type="${docRel}worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="r2" Type="${docRel}worksheet" Target="worksheets/sheet2.xml"/><Relationship Id="r3" Type="${docRel}worksheet" Target="worksheets/sheet3.xml"/><Relationship Id="ext" Type="${docRel}externalLink" Target="externalLinks/externalLink1.xml"/><Relationship Id="calc" Type="${docRel}calcChain" Target="calcChain.xml"/></Relationships>`,
    'xl/worksheets/sheet1.xml': '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1"><f>1+1</f><v>2</v></c><c r="B1" t="inlineStr"><is><t>Visible cell</t></is></c></row></sheetData></worksheet>',
    'xl/worksheets/sheet2.xml': '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/></worksheet>',
    'xl/worksheets/sheet3.xml': '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/></worksheet>',
    'xl/worksheets/_rels/sheet1.xml.rels': `<Relationships xmlns="${relNS}"><Relationship Id="c" Type="${docRel}comments" Target="../comments1.xml"/><Relationship Id="tc" Type="http://schemas.microsoft.com/office/2017/10/relationships/threadedComment" Target="../threadedComments/threadedComment1.xml"/></Relationships>`,
    'xl/comments1.xml': '<comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><authors><author>Alice</author></authors><commentList><comment ref="A1" authorId="0"><text><t>Secret legacy comment</t></text></comment></commentList></comments>',
    'xl/threadedComments/threadedComment1.xml': '<ThreadedComments xmlns="http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments"><threadedComment ref="A1" personId="p1"><text>Secret thread</text></threadedComment></ThreadedComments>',
    'xl/persons/person.xml': '<personList xmlns="http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments"><person id="p1" displayName="Dana" userId="dana@example.test"/></personList>',
    'xl/externalLinks/externalLink1.xml': '<externalLink xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>',
    'xl/externalLinks/_rels/externalLink1.xml.rels': `<Relationships xmlns="${relNS}"><Relationship Id="file" Type="${docRel}externalLinkPath" Target="file:///private/source.xlsx" TargetMode="External"/></Relationships>`,
    'xl/calcChain.xml': '<calcChain xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><c r="A1" i="1"/></calcChain>',
  });
  else Object.assign(parts, {
    'ppt/presentation.xml': '<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:z="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><p:sldIdLst><p:sldId id="256" z:id="s2"/><p:sldId id="257" z:id="s1"/></p:sldIdLst></p:presentation>',
    'ppt/_rels/presentation.xml.rels': `<Relationships xmlns="${relNS}"><Relationship Id="s1" Type="${docRel}slide" Target="slides/slide1.xml"/><Relationship Id="s2" Type="${docRel}slide" Target="slides/slide2.xml"/></Relationships>`,
    'ppt/slides/slide1.xml': '<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><p:cSld><a:t>Visible slide</a:t></p:cSld></p:sld>',
    'ppt/slides/slide2.xml': '<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" show="0"><p:cSld/></p:sld>',
    'ppt/slides/_rels/slide1.xml.rels': `<Relationships xmlns="${relNS}"><Relationship Id="n" Type="${docRel}notesSlide" Target="../notesSlides/notesSlide1.xml"/><Relationship Id="c" Type="${docRel}comments" Target="../comments/comment1.xml"/><Relationship Id="im" Type="${docRel}image" Target="../media/used.png"/></Relationships>`,
    'ppt/notesSlides/notesSlide1.xml': '<p:notes xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><p:cSld><a:t>Secret speaker note</a:t></p:cSld></p:notes>',
    'ppt/notesSlides/_rels/notesSlide1.xml.rels': `<Relationships xmlns="${relNS}"><Relationship Id="back" Type="${docRel}slide" Target="../slides/slide1.xml"/></Relationships>`,
    'ppt/notesMasters/notesMaster1.xml': '<p:notesMaster xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:t>Boilerplate</a:t></p:notesMaster>',
    'ppt/comments/comment1.xml': '<p:cmLst xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cm authorId="0"><p:text>Secret slide comment</p:text></p:cm></p:cmLst>',
    'ppt/commentAuthors.xml': '<p:cmAuthorLst xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cmAuthor id="0" name="Carol" initials="C"/></p:cmAuthorLst>',
    'ppt/media/used.png': new Uint8Array([4,5,6]),
    'ppt/media/orphan.png': new Uint8Array([7,8,9]),
  });
  parts['[Content_Types].xml'] = `<Types xmlns="${typesNS}"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="png" ContentType="image/png"/>${Object.keys(parts).filter(n=>n.endsWith('.xml')).map(n=>`<Override PartName="/${n}" ContentType="application/xml"/>`).join('')}</Types>`;
  return parts;
}
export async function makeOfficeFixture(kind, changes={}) {
  const parts = {...officeFixtureParts(kind), ...changes};
  return writeZip(Object.entries(parts).filter(([,v])=>v!==null).map(([name,data])=>({name,data:typeof data==='string'?enc.encode(data):data})), {deflate:officeDeflate});
}
