// XLSX/PPTX inspection. No formulas, macros, links or embedded code are executed.
import { readZip, parseXML, walk, textOf, DEFAULT_INFLATE } from './zipxml.js';
import { writeZip } from './zipwrite.js';
import { editElements } from './scrub.js';
import { analyze } from './engine.js';

const OF_LOCAL = s => s.split(':').pop();
const OF_NODES = (tree, tag) => [...walk(tree)].filter(n => OF_LOCAL(n.tag) === tag);
const OF_ATTR = (n, key) => Object.entries(n.attrs).find(([k]) => OF_LOCAL(k) === key)?.[1];
const OF_FIRST = (tree, tag) => OF_NODES(tree, tag)[0];
const OF_VALUE = (tree, tag) => { const n=OF_FIRST(tree,tag);return n ? textOf(n).trim() : null; };
const OF_RELS = part => part.replace(/([^/]+)$/, '_rels/$1.rels');

export function detectOffice(zip) {
  const kinds = [['xlsx','xl/workbook.xml'],['pptx','ppt/presentation.xml'],['docx','word/document.xml']].filter(([,p])=>zip.has(p));
  if(kinds.length>1) throw new Error('Ambiguous Office package: multiple document roots');
  return kinds.length && kinds[0][0]!=='docx' ? kinds[0][0] : null;
}
export async function analyzePackage(bytes, opts={}) {
  const zip=await readZip(bytes,opts.inflate || DEFAULT_INFLATE);
  return detectOffice(zip) ? analyzeOffice(bytes,opts) : analyze(bytes,opts);
}
function ofResolve(owner, target) {
  let path;
  try { path=decodeURIComponent(target.split('#')[0]); } catch { return null; }
  if(!path || /[\\\x00-\x1f]/.test(path) || /^[a-z]+:/i.test(path)) return null;
  const segments = path.startsWith('/') ? [] : owner.split('/').slice(0,-1);
  for(const s of path.split('/')) {
    if(!s || s==='.')continue;
    if(s==='..'){ if(!segments.length)return null; segments.pop(); }
    else segments.push(s);
  }
  return segments.join('/');
}
function ofOwner(rels) {
  if(rels==='_rels/.rels')return '';
  return rels.replace(/(^|\/)\_rels\//,'$1').replace(/\.rels$/,'');
}
async function ofRelationships(zip, issues) {
  const rows=[];
  for(const part of zip.order.filter(p=>p.endsWith('.rels'))) {
    const owner=ofOwner(part),tree=parseXML(await zip.text(part));
    for(const n of OF_NODES(tree,'Relationship')) {
      const external=n.attrs.TargetMode==='External';
      const resolved=external ? null : ofResolve(owner,n.attrs.Target || '');
      const row={part,owner,id:n.attrs.Id,type:n.attrs.Type || '',target:n.attrs.Target || '',external,resolved};
      if(!external && (!resolved || !zip.has(resolved)))issues.push(`${part}: unresolved target ${row.target}`);
      rows.push(row);
    }
  }
  return rows;
}

export async function analyzeOffice(bytes, opts={}) {
  const zip=await readZip(bytes,opts.inflate || DEFAULT_INFLATE);
  const format=detectOffice(zip);
  if(!format)throw new Error('Not a supported XLSX or PPTX package');
  const r={format,file:{name:opts.name || `document.${format}`,bytes:bytes.length},container:{partCount:zip.order.length},
    metadata:{},findings:[],sheets:[],slides:[],definedNames:[],formulaCaches:[],calcChains:[],comments:[],authors:[],
    notes:[],external:[],thumbnails:[],orphanMedia:[],issues:[],signed:zip.order.some(p=>p.startsWith('_xmlsignatures/'))};
  const tree = async p => parseXML((await zip.text(p)) || '');
  const core=await tree('docProps/core.xml'),app=await tree('docProps/app.xml');
  for(const key of ['creator','lastModifiedBy','title','subject','description','created','modified'])r.metadata[key]=OF_VALUE(core,key);
  for(const key of ['Company','Manager','Application'])r.metadata[key]=OF_VALUE(app,key);
  const rels=await ofRelationships(zip,r.issues);
  r.external=rels.filter(x=>x.external);
  const incoming=new Set(rels.filter(x=>!x.external).map(x=>x.resolved));
  const main=format==='xlsx' ? 'xl/workbook.xml' : 'ppt/presentation.xml';
  const root=await tree(main);
  const resolveId=n=>rels.find(x=>x.owner===main && x.id===Object.entries(n.attrs).find(([k])=>k.endsWith(':id'))?.[1])?.resolved || null;
  if(format==='xlsx') {
    for(const n of OF_NODES(root,'sheet')) {
      const part=resolveId(n),missing=!part || !zip.has(part);
      r.sheets.push({name:n.attrs.name || '',state:n.attrs.state || 'visible',part,missing});
      if(missing)r.issues.push(`Sheet ${n.attrs.name}: missing worksheet ${part || '(unresolved id)'}`);
    }
    r.definedNames=OF_NODES(root,'definedName').map(n=>({name:n.attrs.name || '',scope:n.attrs.localSheetId ?? 'workbook',formula:textOf(n)}));
    for(const part of zip.order.filter(p=>/^xl\/worksheets\/.*\.xml$/.test(p))) {
      for(const n of OF_NODES(await tree(part),'c')) {
        const formula=OF_FIRST(n,'f'),value=OF_FIRST(n,'v');
        if(formula && value)r.formulaCaches.push({part,cell:n.attrs.r || '',formula:textOf(formula),value:textOf(value)});
      }
    }
    r.calcChains=zip.order.filter(p=>/^xl\/calcChain.*\.xml$/.test(p));
  } else {
    for(const n of OF_NODES(root,'sldId')) {
      const part=resolveId(n),missing=!part || !zip.has(part);
      const slide=missing ? null : OF_FIRST(await tree(part),'sld');
      r.slides.push({part,hidden:slide ? ['0','false'].includes(slide.attrs.show) : null,missing});
      if(missing)r.issues.push(`Slide ${n.attrs.id}: missing slide ${part || '(unresolved id)'}`);
    }
    for(const part of zip.order.filter(p=>/^ppt\/notesSlides\/[^/]+\.xml$/.test(p))) {
      const t=await tree(part);
      r.notes.push({part,text:OF_NODES(t,'t').map(textOf).join('\n')});
    }
  }
  for(const part of zip.order.filter(p=>/\.xml$/.test(p))) {
    if(/(?:^xl\/(?:comments[^/]*|threadedComments\/[^/]+)|^ppt\/comments\/[^/]+)\.xml$/.test(part)) {
      const t=await tree(part);
      for(const n of [...OF_NODES(t,'comment'),...OF_NODES(t,'threadedComment'),...OF_NODES(t,'cm')]) {
        r.comments.push({part,ref:n.attrs.ref || '',authorId:n.attrs.authorId || n.attrs.personId || '',text:OF_VALUE(n,'text') || OF_NODES(n,'t').map(textOf).join('\n')});
      }
      for(const n of OF_NODES(t,'author'))r.authors.push({part,name:textOf(n)});
    }
    if(/^xl\/persons\/|^ppt\/(?:commentAuthors|authors)/.test(part)) {
      const t=await tree(part);
      for(const n of [...OF_NODES(t,'person'),...OF_NODES(t,'cmAuthor'),...OF_NODES(t,'author')])r.authors.push({part,name:n.attrs.displayName || n.attrs.name || '',userId:n.attrs.userId || ''});
    }
  }
  r.thumbnails=zip.order.filter(p=>/^docProps\/thumbnail\./i.test(p));
  r.orphanMedia=zip.order.filter(p=>/^(xl|ppt)\/media\/[^/]+$/.test(p) && !incoming.has(p));
  const add=(kind,part,detail)=>r.findings.push({kind,part,detail});
  for(const key of ['creator','lastModifiedBy','Company','Manager'])if(r.metadata[key])add('identity',key.startsWith('C')||key==='Manager'?'docProps/app.xml':'docProps/core.xml',`${key}: ${r.metadata[key]}`);
  for(const s of r.sheets.filter(s=>s.state!=='visible'))add('hidden-sheet',s.part,`${s.name}: ${s.state}; hidden data may still be used by formulas.`);
  r.slides.forEach((s,i)=>{if(s.hidden)add('hidden-slide',s.part,`Slide ${i+1} is marked hidden; its content remains in the package.`);});
  for(const n of r.notes)add('speaker-notes',n.part,n.text);
  for(const c of r.comments)add('comments',c.part,c.text);
  for(const a of r.authors)add('comment-author',a.part,[a.name,a.userId].filter(Boolean).join(' '));
  for(const x of r.external)add('external-link',x.part,x.target+' (not fetched)');
  for(const n of r.definedNames)add('defined-name',main,`${n.name} [${n.scope}]: ${n.formula}`);
  for(const v of r.formulaCaches)add('formula-cache',v.part,`${v.cell}: cached ${v.value}; may be stale; formula not executed.`);
  for(const p of r.calcChains)add('calculation-chain',p,'Stored calculation dependency information; not a history of user activity.');
  for(const p of r.thumbnails)add('thumbnail',p,'Saved preview image may reveal document content.');
  for(const p of r.orphanMedia)add('unreferenced-media',p,'No package relationship targets this media; not proof it is safe to delete.');
  if(r.signed)add('signature','_xmlsignatures/','Signed package: modification invalidates signatures; cleanup unavailable.');
  for(const issue of r.issues)add('unresolved-package',null,issue);
  r.limitations=['Inspection covers listed OOXML fields, not every possible embedded data store.','Hidden content, formulas, external references and media are not executed or fetched.','Re-analysis is not renderer validation or a guarantee of anonymity.'];
  return r;
}

export function planOffice(r) {
  return [
    {id:'identity',label:'Remove stored author and organization fields',count:['creator','lastModifiedBy','Company','Manager'].filter(k=>r.metadata[k]).length,risk:'metadata'},
    {id:'comments',label:'Remove comments and comment-author records',count:r.comments.length+r.authors.length,risk:'content'},
    {id:'notes',label:'Remove speaker-note parts',count:r.notes.length,risk:'content'},
    {id:'thumbnail',label:'Remove saved thumbnail',count:r.thumbnails.length,risk:'preview'},
  ].filter(p=>p.count).map(p=>({...p,on:false}));
}

// Use the established byte-preserving XML surgery; only selected nodes change.
function ofCut(xml,local,predicate=()=>true) {
  const tags=new Set(OF_NODES(parseXML(xml),local).map(n=>n.tag));
  let out=xml;
  for(const tag of tags)out=editElements(out,tag,'cut',(_inner,start)=>{
    const n=OF_FIRST(parseXML(start),local);return !!n && predicate(n);
  }).out;
  return out;
}
export async function scrubOffice(bytes, report, opts={}) {
  // Re-inspect the actual bytes, not stale or caller-supplied finding counts.
  const r=await analyzeOffice(bytes,{inflate:opts.inflate,name:report.file.name});
  const fixes=new Set(opts.fixes || []);
  const known=new Set(['identity','comments','notes','thumbnail']);
  for(const id of fixes)if(!known.has(id))throw new Error('Unsupported cleanup action: '+id);
  if(!fixes.size)return {bytes,applied:[],notes:['No cleanup selected.']};
  if(r.signed)throw new Error('Signed package: cleanup would invalidate digital signatures');
  if(r.issues.length)throw new Error('Unresolved package relationships: cleanup refused');
  const zip=await readZip(bytes,opts.inflate || DEFAULT_INFLATE);
  const removed=new Set(),applied=[];
  const changes=new Map();
  const remove = part => {if(zip.has(part))removed.add(part);};
  if(fixes.has('comments')) {
    for(const c of [...r.comments,...r.authors])remove(c.part);
    // Empty known comment/author parts can still carry private attributes.
    for(const p of zip.order)if(/^(xl\/(?:comments[^/]*\.xml|threadedComments\/[^/]+\.xml|persons\/[^/]+\.xml)|ppt\/(?:comments\/[^/]+\.xml|commentAuthors\.xml|authors\/[^/]+\.xml))$/.test(p))remove(p);
  }
  if(fixes.has('notes'))for(const n of r.notes)remove(n.part);
  if(fixes.has('thumbnail'))for(const p of r.thumbnails)remove(p);
  for(const p of [...removed])remove(OF_RELS(p));
  if(fixes.has('identity')) {
    for(const [p,fields] of [['docProps/core.xml',['creator','lastModifiedBy']],['docProps/app.xml',['Company','Manager']]]) {
      if(!zip.has(p))continue;
      let xml=await zip.text(p);
      for(const key of fields)xml=ofCut(xml,key);
      changes.set(p,xml);
    }
  }
  for(const p of zip.order.filter(p=>p.endsWith('.rels') && !removed.has(p))) {
    const owner=ofOwner(p),xml=await zip.text(p);
    changes.set(p,ofCut(xml,'Relationship',n=>n.attrs.TargetMode!=='External' && removed.has(ofResolve(owner,n.attrs.Target || ''))));
  }
  if(zip.has('[Content_Types].xml'))changes.set('[Content_Types].xml',ofCut(await zip.text('[Content_Types].xml'),'Override',n=>removed.has((n.attrs.PartName || '').replace(/^\//,''))));
  const files=[];
  const encode=new TextEncoder();
  let changed=removed.size>0;
  for(const p of zip.order) {
    if(removed.has(p))continue;
    const original=await zip.bytes(p);
    const replacement=changes.has(p) ? encode.encode(changes.get(p)) : original;
    if(changes.has(p) && changes.get(p)!==await zip.text(p))changed=true;
    files.push({name:p,data:replacement});
  }
  if(!changed)return {bytes,applied:[],notes:['Selected fields were already absent.']};
  const output=await writeZip(files,{deflate:opts.deflate});
  const after=await analyzeOffice(output,{inflate:opts.inflate,name:report.file.name});
  if(after.issues.length)throw new Error('Cleanup produced unresolved relationships; output withheld');
  for(const p of planOffice(r))if(fixes.has(p.id))applied.push({id:p.id,n:p.count});
  return {bytes:output,applied,notes:['Only selected categories were removed. Review remaining findings and open the result in your Office application.','Hidden sheets/slides, formulas, cached values, defined names, external links and unreferenced media are retained.']};
}
