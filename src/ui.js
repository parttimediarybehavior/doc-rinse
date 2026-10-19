// ui.js — Doc Rinse front end. Renders an excavation of one .docx.
import { compareDocs, textReport } from './engine.js';
import { analyzePackage, analyzeOffice, planOffice, scrubOffice } from './office.js';
import { plan, scrub, verifyScrub } from './scrub.js';

const $ = (sel, r = document) => r.querySelector(sel);
const el = (tag, cls, txt) => {
  const n = document.createElement(tag);
  if (cls) n.className = cls;
  if (txt !== undefined) n.textContent = txt;
  return n;
};
const esc = (s) => String(s == null ? '' : s).replace(/[&<>"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
const fmt = (n) => (n == null ? '—' : n.toLocaleString());
const shortDate = (s) => (s ? String(s).slice(0, 16).replace('T', ' ') : '—');

let CURRENT = null;
let COMPARE = null;

/* ---------------------------------------------------------- colouring --- */
// stable colour per session id, spread around the wheel by hash
function hueOf(rsid) {
  let h = 0;
  for (let i = 0; i < rsid.length; i++) h = (h * 31 + rsid.charCodeAt(i)) >>> 0;
  return (h * 137.508) % 360;
}
const sessColor = (rsid, a = 0.3) => `hsla(${hueOf(rsid).toFixed(0)}, 58%, 52%, ${a})`;
const sessSolid = (rsid) => `hsl(${hueOf(rsid).toFixed(0)}, 58%, 58%)`;

/* ------------------------------------------------------------- render --- */

export function renderReport(rep, name) {
  CURRENT = rep;
  COMPARE = null;
  $('#intro').classList.add('hidden');
  const out = $('#out');
  out.innerHTML = '';
  if (rep.format === 'xlsx' || rep.format === 'pptx') {
    out.appendChild(officeReportPanel(rep, name));
    window.scrollTo({top:0,behavior:'instant'});
    return;
  }
  out.appendChild(verdictCard(rep, name));
  const cols = el('div', 'cols');
  const side = el('div', 'side');
  const main = el('div', 'main');
  side.appendChild(card('what this file admits', findingsList(rep), true));
  side.appendChild(card('sessions', sessionTable(rep), true));
  side.appendChild(card('compare with another version', compareBox(), true));
  main.appendChild(card('clean this file', scrubPanel(rep, name), true));
  main.appendChild(stratigraphy(rep));
  if (rep.media.length) main.appendChild(card('images · crop excavation', imageGrid(rep), true));
  const recovered = recoveredPanel(rep);
  if (recovered) main.appendChild(card('recovered content', recovered, true));
  main.appendChild(card('timeline', timelinePanel(rep), true));
  main.appendChild(card('metadata', metaPanel(rep), true));
  main.appendChild(card('container', containerPanel(rep), true));
  cols.appendChild(side);
  cols.appendChild(main);
  out.appendChild(cols);
  out.appendChild(footnote());
  window.scrollTo({ top: 0, behavior: 'instant' });
}

function officeReportPanel(rep, name) {
  const box=el('div','cols');
  const detail=el('div','main');
  detail.appendChild(el('h2',null,`${name} · ${rep.format.toUpperCase()}`));
  detail.appendChild(el('p',null,`${rep.findings.length} findings. Inspection only until you select cleanup actions. No formulas, macros or external links are executed.`));
  const entries=el('div');
  for(const f of rep.findings) {
    const row=el('div','finding');
    row.appendChild(el('b',null,f.kind));
    row.appendChild(el('p',null,f.part || 'package'));
    row.appendChild(el('pre',null,f.detail));
    entries.appendChild(row);
  }
  if(!rep.findings.length)entries.appendChild(el('p',null,'No findings in the fields inspected. This does not establish absence of all private data.'));
  detail.appendChild(card('stored content and metadata',entries,true));
  const limits=el('div');for(const text of rep.limitations)limits.appendChild(el('p',null,text));
  detail.appendChild(card('inspection limits',limits,true));
  const actions=el('div','side'),controls=el('div'),checks=[];
  controls.appendChild(el('p',null,'Choose removals explicitly. Hidden sheets/slides, formulas, cached values, names, external links and unreferenced media will remain. Comments and speaker notes are content removals.'));
  for(const action of planOffice(rep)) {
    const label=el('label'),input=el('input');input.type='checkbox';input.checked=false;
    input.disabled=rep.signed || rep.issues.length>0;
    checks.push({input,id:action.id});label.appendChild(input);
    label.appendChild(el('span',null,`${action.label} (${action.count}; ${action.risk})`));controls.appendChild(label);
    controls.appendChild(el('br'));
  }
  const button=el('button',null,'clean selected fields'),status=el('div');
  button.disabled=rep.signed || rep.issues.length>0 || !checks.length;
  if(rep.signed || rep.issues.length)status.textContent='Cleanup unavailable: this package is signed or has unresolved relationships.';
  button.onclick=async()=>{
    const fixes=checks.filter(c=>c.input.checked).map(c=>c.id);
    if(!fixes.length){status.textContent='Select at least one removal.';return;}
    button.disabled=true;
    try {
      const res=await scrubOffice(rep._bytes,rep,{fixes});
      const cleanedName=name.replace(/\.[^.]*$/,'')+`.clean.${rep.format}`;
      const after=await analyzeOffice(res.bytes,{name:cleanedName});
      const url=URL.createObjectURL(new Blob([res.bytes],{type:'application/octet-stream'}));
      const a=el('a');a.href=url;a.download=cleanedName;a.click();
      setTimeout(()=>URL.revokeObjectURL(url),4000);
      status.innerHTML='';
      status.appendChild(el('p',null,`Re-analysis: ${rep.findings.length} → ${after.findings.length} findings. Review the downloaded file before sharing.`));
      for(const note of res.notes)status.appendChild(el('p',null,note));
      const remaining=el('ul');for(const f of after.findings)remaining.appendChild(el('li',null,`${f.kind}: ${f.detail}`));status.appendChild(remaining);
    } catch(e){status.textContent=`Cleanup failed: ${e.message}`;}
    button.disabled=false;
  };
  controls.appendChild(button);controls.appendChild(status);actions.appendChild(card('selected cleanup',controls,true));
  box.appendChild(actions);box.appendChild(detail);return box;
}

function card(title, body, tight) {
  const c = el('div', 'card');
  c.appendChild(el('h3', null, title));
  const b = el('div', 'body' + (tight ? ' tight' : ''));
  b.appendChild(body);
  c.appendChild(b);
  return c;
}

function verdictCard(rep, name) {
  const v = el('div', 'verdict');
  const top = el('div', 'top');
  top.appendChild(el('div', 'fname', `${name} · ${(rep.file.bytes / 1024).toFixed(0)} KB · ${rep.container.partCount} parts · ${rep.props.application || 'unknown producer'}${rep.props.appVersionLabel ? ` (${rep.props.appVersionLabel})` : ''}`));
  top.appendChild(el('h2', null, rep.verdict.headline));
  const ul = el('ul');
  for (const l of rep.verdict.lines) ul.appendChild(el('li', null, l));
  top.appendChild(ul);
  v.appendChild(top);

  const native = rep.sessions.filter((s) => s.native);
  const foreign = rep.sessions.filter((s) => !s.native);
  const stats = el('div', 'stats');
  const st = (label, value, hot) => {
    const d = el('div', 'stat' + (hot ? ' hot' : ''));
    d.appendChild(el('b', null, value));
    d.appendChild(el('span', null, label));
    stats.appendChild(d);
  };
  st('editing sessions', fmt(rep.settings.rsidCount || native.filter((s) => s.chars).length));
  st('words', fmt(rep.text.words));
  st('edit minutes', rep.props.totalTime == null ? '—' : fmt(rep.props.totalTime), rep.props.totalTime === 0 && rep.text.words > 150);
  st('saves', fmt(rep.timeline.revisions));
  st('names', fmt(rep.identities.length), rep.identities.length > 1);
  if (foreign.length) st('foreign ids', fmt(foreign.length), foreign.length >= 10);
  if (rep.deletedText.length) st('deleted passages', fmt(rep.deletedText.length), true);
  if (rep.hiddenText.length) st('invisible runs', fmt(rep.hiddenText.length), true);
  if (rep.comments.length) st('comments', fmt(rep.comments.length), true);
  v.appendChild(stats);
  const acts = el('div', 'acts');
  const dl = (filename, text, mime) => {
    const url = URL.createObjectURL(new Blob([text], { type: mime }));
    const a = document.createElement('a');
    a.href = url;
    a.download = filename;
    a.click();
    setTimeout(() => URL.revokeObjectURL(url), 4000);
  };
  const base = name.replace(/\.docx?$/i, '');
  const b1 = el('button', 'tgl', 'copy report');
  b1.onclick = async () => {
    const txt = textReport(rep);
    try {
      await navigator.clipboard.writeText(txt);
      b1.textContent = 'copied';
    } catch {
      dl(base + '.doc-rinse.txt', txt, 'text/plain');
      b1.textContent = 'downloaded';
    }
    setTimeout(() => (b1.textContent = 'copy report'), 2200);
  };
  const b2 = el('button', 'tgl', 'download .txt');
  b2.onclick = () => dl(base + '.doc-rinse.txt', textReport(rep), 'text/plain');
  const b3 = el('button', 'tgl', 'download .json');
  b3.onclick = () => {
    const clone = { ...rep };
    delete clone._model;
    clone.media = clone.media.map((m) => ({ ...m, bytes: undefined }));
    dl(base + '.doc-rinse.json', JSON.stringify(clone, null, 1), 'application/json');
  };
  for (const b of [b1, b2, b3]) acts.appendChild(b);
  const note = el('span', 'note', 'evidence export · nothing is sent anywhere');
  acts.appendChild(note);
  v.appendChild(acts);
  return v;
}

function findingsList(rep) {
  const box = el('div');
  if (!rep.findings.length) box.appendChild(el('div', 'finding', 'Nothing notable — unusually clean.'));
  for (const f of rep.findings) {
    const d = el('div', 'finding ' + f.level);
    const t = el('div', 't');
    t.appendChild(el('span', 'dot'));
    t.appendChild(el('span', null, f.title));
    d.appendChild(t);
    d.appendChild(el('div', 'd', f.detail));
    box.appendChild(d);
  }
  return box;
}

/* ------------------------------------------------------- stratigraphy --- */

function stratigraphy(rep) {
  const c = el('div', 'card');
  c.appendChild(el('h3', null, 'stratigraphy · the document by editing session'));
  const bar = el('div', 'toolbar');
  const doc = el('div', 'doc mode-session');
  const state = { 'mode-session': true, 'show-del': true, 'show-hidden': false, 'mark-import': false, plain: false };
  const sync = () => {
    doc.className =
      'doc' +
      (state['mode-session'] && !state.plain ? ' mode-session' : '') +
      (state['show-hidden'] ? ' show-hidden' : '') +
      (state['mark-import'] && !state.plain ? ' mark-import' : '') +
      (doc.classList.contains('focus') ? ' focus' : '');
    doc.querySelectorAll('.del').forEach((n) => (n.style.display = state['show-del'] ? '' : 'none'));
  };
  for (const [label, key] of [
    ['sessions', 'mode-session'],
    ['deleted text', 'show-del'],
    ['invisible text', 'show-hidden'],
    ['imported ids', 'mark-import'],
    ['plain', 'plain'],
  ]) {
    const b = el('button', 'tgl', label);
    b.setAttribute('aria-pressed', String(state[key]));
    b.onclick = () => {
      state[key] = !state[key];
      b.setAttribute('aria-pressed', String(state[key]));
      sync();
    };
    bar.appendChild(b);
  }
  bar.appendChild(el('div', 'sp'));
  const note = el('span', 'tgl');
  note.style.border = '0';
  note.style.background = 'none';
  note.textContent = 'click a session in the list to isolate its text';
  bar.appendChild(note);
  c.appendChild(bar);

  const model = rep._model;
  const foreignSet = new Set(rep.sessions.filter((s) => !s.native).map((s) => s.rsid));
  const ribbonRows = [];
  let lastRsid = null;
  for (const p of model.paragraphs) {
    if (!p.runs.length && !p.style) continue;
    const kind = /heading|^title|^subtitle/i.test(p.styleId || '') ? 'h' : p.listLevel !== null ? 'li' : p.tableDepth ? 'tbl' : '';
    const para = el('p', kind);
    para.dataset.i = p.index;
    let wrote = false;
    const weight = new Map();
    let prevRsid = lastRsid;
    for (const r of p.runs) {
      if (r.kind === 'commentMark') {
        if (r.which === 'start') para.appendChild(el('span', 'cm', '[' + r.commentId + ']'));
        continue;
      }
      if (!r.text) continue;
      const s = el('span', 'r');
      let cl = 'r';
      if (r.rev && (r.rev.type === 'del' || r.rev.type === 'moveFrom')) cl += ' del';
      else if (r.rev && (r.rev.type === 'ins' || r.rev.type === 'moveTo')) cl += ' ins';
      const whiteOnWhite = r.props.color && /^ffffff$/i.test(r.props.color) && r.props.highlight === undefined && !p.bgFill;
      const invisible = r.props.vanish || whiteOnWhite || (r.props.sz !== undefined && r.props.sz <= 1.5);
      if (invisible) cl += ' hid invis';
      if (r.rsid) {
        s.dataset.rsid = r.rsid;
        if (foreignSet.has(r.rsid)) cl += ' foreign';
        if (r.rsid !== prevRsid) cl += ' seam';
        prevRsid = r.rsid;
        s.style.setProperty('--c', sessColor(r.rsid, 0.17));
        s.style.setProperty('--cs', sessColor(r.rsid, 0.85));
        weight.set(r.rsid, (weight.get(r.rsid) || 0) + r.text.length);
      }
      if (r.rsidRPr && foreignSet.has(r.rsidRPr)) {
        cl += ' ffmt';
        s.dataset.ffmt = r.rsidRPr;
      }
      s.className = cl;
      s.textContent = r.text;
      if (r.props.b) s.style.fontWeight = '700';
      if (r.props.i) s.style.fontStyle = 'italic';
      s.onmouseenter = (ev) => showTip(ev, r, p, rep);
      s.onmouseleave = hideTip;
      para.appendChild(s);
      wrote = true;
    }
    if (!wrote && !kind) para.appendChild(el('span', 'r', '\u00a0'));
    lastRsid = prevRsid;
    doc.appendChild(para);
    const dom = [...weight.entries()].sort((a, b) => b[1] - a[1])[0];
    ribbonRows.push({ index: p.index, chars: Math.max(1, [...weight.values()].reduce((a, b) => a + b, 0)), rsid: dom ? dom[0] : null });
  }

  // sediment core: one band per paragraph, height ∝ text volume
  const ribbon = el('div', 'ribbon');
  ribbon.title = 'sediment core — one band per paragraph, coloured by the session that wrote it';
  for (const row of ribbonRows) {
    const i = el('i');
    i.style.flexGrow = String(Math.sqrt(row.chars));
    i.style.background = row.rsid ? sessSolid(row.rsid) : '#2a2320';
    if (row.rsid && foreignSet.has(row.rsid)) i.style.background = 'repeating-linear-gradient(45deg,#9a86c8,#9a86c8 2px,#2a2320 2px,#2a2320 4px)';
    i.onclick = () => {
      const target = doc.querySelector(`p[data-i="${row.index}"]`);
      if (!target) return;
      target.scrollIntoView({ block: 'center', behavior: 'smooth' });
      target.classList.remove('flash');
      void target.offsetWidth;
      target.classList.add('flash');
    };
    ribbon.appendChild(i);
  }
  const wrap = el('div', 'docwrap');
  wrap.appendChild(ribbon);
  wrap.appendChild(doc);
  c.appendChild(wrap);
  return c;
}

function showTip(ev, r, p, rep) {
  const tip = $('#tip');
  const sess = rep.sessions.find((s) => s.rsid === r.rsid);
  const rows = [];
  if (r.rsid) rows.push(`session <b>${r.rsid}</b>${sess && !sess.native ? ' (foreign — from another document)' : ''}`);
  if (sess && sess.native) rows.push(`${sess.chars.toLocaleString()} chars in this session · ${sess.share}% of the text${sess.isRoot ? ' · created the file' : ''}`);
  if (sess && sess.anchoredAt) rows.push(`dated ${shortDate(sess.anchoredAt)}`);
  if (r.rsidRPr && r.rsidRPr !== r.rsid) rows.push(`formatting id ${r.rsidRPr}`);
  if (r.props.lang || r.props.langEA) rows.push(`lang ${[r.props.lang, r.props.langEA].filter(Boolean).join(' / ')}`);
  if (r.props.font) rows.push(`font ${r.props.font}${r.props.sz ? ' ' + r.props.sz + 'pt' : ''}`);
  if (r.rev) rows.push(`tracked ${r.rev.type} by ${r.rev.author || '?'} ${shortDate(r.rev.date)}`);
  if (r.props.vanish) rows.push('marked hidden (w:vanish)');
  rows.push(`¶${p.index + 1}${p.style ? ' · ' + p.style : ''}${p.paraId ? ' · paraId ' + p.paraId : ''}`);
  tip.innerHTML = rows.join('<br>');
  tip.style.opacity = '1';
  const r2 = ev.target.getBoundingClientRect();
  tip.style.left = Math.min(window.innerWidth - 320, r2.left) + 'px';
  tip.style.top = (r2.bottom + 8 > window.innerHeight - 120 ? r2.top - tip.offsetHeight - 8 : r2.bottom + 8) + 'px';
}
const hideTip = () => ($('#tip').style.opacity = '0');

/* ---------------------------------------------------------- sessions --- */

function sessionTable(rep) {
  const wrap = el('div');
  const t = el('table', 'sess');
  const head = el('tr');
  for (const h of ['', 'id', 'text', 'share', 'where']) head.appendChild(el('th', null, h));
  t.appendChild(head);
  const shown = rep.sessions.filter((s) => s.chars > 0 || s.native).slice(0, 60);
  for (const s of shown) {
    const tr = el('tr');
    tr.dataset.rsid = s.rsid;
    const sw = el('td');
    const i = el('span', 'sw');
    i.style.background = sessSolid(s.rsid);
    if (!s.native) i.style.background = 'repeating-linear-gradient(45deg,#9a86c8,#9a86c8 2px,#2a2320 2px,#2a2320 4px)';
    sw.appendChild(i);
    tr.appendChild(sw);
    const id = el('td');
    id.appendChild(el('span', 'id', s.rsid));
    if (s.isRoot) id.appendChild(el('div', 'when', 'created the file'));
    else if (s.anchoredAt) id.appendChild(el('div', 'when', shortDate(s.anchoredAt)));
    else if (!s.native) id.appendChild(el('div', 'tagf', 'from another doc'));
    tr.appendChild(id);
    tr.appendChild(el('td', null, s.chars ? s.chars.toLocaleString() : '—'));
    const bt = el('td');
    const bar = el('div', 'bar');
    const fill = el('i');
    fill.style.width = Math.max(1, s.share) + '%';
    fill.style.background = sessSolid(s.rsid);
    bar.appendChild(fill);
    bt.appendChild(bar);
    tr.appendChild(bt);
    tr.appendChild(el('td', null, s.span ? `¶${s.span[0] + 1}${s.span[1] !== s.span[0] ? '–' + (s.span[1] + 1) : ''}` : '—'));
    tr.onclick = () => focusSession(s.rsid, tr);
    t.appendChild(tr);
  }
  wrap.appendChild(t);
  const ghosts = rep.sessions.filter((s) => s.native && !s.chars).length;
  const note = el('div', 'meta');
  note.style.cssText = 'padding:9px 11px;font:11px/1.5 var(--mono);color:var(--dimmer)';
  note.innerHTML = `Word mints a random id per save. Values are 24-bit random and listed sorted, so they carry no clock — only partition.<br>${ghosts} of this file's sessions left no surviving text.`;
  wrap.appendChild(note);
  return wrap;
}

function focusSession(rsid, tr) {
  const doc = $('.doc');
  if (!doc) return;
  const already = tr.classList.contains('on');
  document.querySelectorAll('.sess tr.on').forEach((n) => n.classList.remove('on'));
  doc.querySelectorAll('.r.on').forEach((n) => n.classList.remove('on'));
  if (already) {
    doc.classList.remove('focus');
    return;
  }
  tr.classList.add('on');
  doc.classList.add('focus');
  let first = null;
  doc.querySelectorAll(`.r[data-rsid="${rsid}"]`).forEach((n) => {
    n.classList.add('on');
    if (!first) first = n;
  });
  if (first) first.scrollIntoView({ block: 'center', behavior: 'smooth' });
}

/* ------------------------------------------------------------ images --- */

function imageGrid(rep) {
  const g = el('div', 'imgs');
  for (const m of rep.media) {
    const c = el('div', 'imgcard');
    const stage = el('div', 'stage');
    let url = null;
    if (m.bytes) {
      url = URL.createObjectURL(new Blob([m.bytes], { type: m.mime }));
      const img = el('img');
      img.src = url;
      img.alt = m.part;
      stage.appendChild(img);
      if (m.crop) {
        const win = el('div', 'win');
        const place = () => {
          const b = img.getBoundingClientRect(), s = stage.getBoundingClientRect();
          win.style.left = b.left - s.left + (b.width * m.crop.l) / 100 + 'px';
          win.style.top = b.top - s.top + (b.height * m.crop.t) / 100 + 'px';
          win.style.width = (b.width * (100 - m.crop.l - m.crop.r)) / 100 + 'px';
          win.style.height = (b.height * (100 - m.crop.t - m.crop.b)) / 100 + 'px';
        };
        img.onload = place;
        stage.appendChild(win);
        // two states: what Word shows you, and what the file actually contains
        const btn = el('button', 'reveal', 'excavate');
        let revealed = false;
        const apply = () => {
          if (revealed) {
            img.style.clipPath = 'none';
            win.style.display = 'block';
            win.style.boxShadow = 'none';
            btn.textContent = 'as Word shows it';
            place();
          } else {
            img.style.clipPath = `inset(${m.crop.t}% ${m.crop.r}% ${m.crop.b}% ${m.crop.l}%)`;
            win.style.display = 'none';
            btn.textContent = 'excavate';
          }
        };
        btn.onclick = (e) => {
          e.stopPropagation();
          revealed = revealed === false;
          apply();
        };
        apply();
        stage.appendChild(btn);
      }
    } else stage.appendChild(el('div', 'meta', 'no preview'));
    c.appendChild(stage);
    const meta = el('div', 'meta');
    const bits = [`<b>${esc(m.part.replace('word/media/', ''))}</b>`, `${m.format || '?'} ${m.width || '?'}×${m.height || '?'} · ${(m.size / 1024).toFixed(0)} KB`];
    if (m.crop) bits.push(`<span class="crop">cropped → ${m.visibleFraction}% shown; ${(100 - m.visibleFraction).toFixed(1)}% recoverable</span>`);
    if (m.uses[0] && m.uses[0].displayPx) bits.push(`displayed at ${m.uses[0].displayPx.w}×${m.uses[0].displayPx.h} px`);
    if (m.unused) bits.push('<span class="crop">not displayed anywhere in the document</span>');
    if (m.uses[0] && m.uses[0].descr) bits.push('alt: ' + esc(m.uses[0].descr).slice(0, 120));
    for (const [k, v] of Object.entries(m.meta || {})) bits.push(`${esc(k)}: <b>${esc(String(v)).slice(0, 90)}</b>`);
    meta.innerHTML = bits.join('<br>');
    c.appendChild(meta);
    g.appendChild(c);
  }
  return g;
}

/* --------------------------------------------------------- recovered --- */

function recoveredPanel(rep) {
  if (!rep.deletedText.length && !rep.hiddenText.length && !rep.comments.length && !rep.embedded.length) return null;
  const box = el('div');
  box.style.padding = '12px 15px';
  const section = (title, items, fmtItem) => {
    if (!items.length) return;
    const h = el('div');
    h.style.cssText = 'font:11px var(--mono);letter-spacing:.14em;text-transform:uppercase;color:var(--dimmer);margin:10px 0 4px';
    h.textContent = title;
    box.appendChild(h);
    for (const it of items.slice(0, 40)) box.appendChild(fmtItem(it));
  };
  section('deleted but still present', rep.deletedText, (d) => {
    const q = el('div', 'quote');
    q.appendChild(el('span', null, '“' + d.text.trim().slice(0, 400) + '”'));
    q.appendChild(el('span', 'who', `${d.author || 'unknown'} · ${shortDate(d.date)} · ¶${d.paraIndex + 1}`));
    return q;
  });
  section('invisible on screen', rep.hiddenText, (d) => {
    const q = el('div', 'quote');
    q.style.borderColor = 'var(--amber)';
    q.appendChild(el('span', null, '“' + d.text.trim().slice(0, 400) + '”'));
    q.appendChild(el('span', 'who', `${d.kind} · ¶${d.paraIndex + 1}`));
    return q;
  });
  section('comments', rep.comments, (c) => {
    const q = el('div', 'quote');
    q.style.borderColor = 'var(--violet)';
    q.appendChild(el('span', null, '“' + c.text.slice(0, 400) + '”'));
    q.appendChild(el('span', 'who', `${c.author || '?'} · ${shortDate(c.date)}${c.resolved ? ' · resolved' : ''}`));
    return q;
  });
  section('carried inside the file', rep.embedded, (e) => {
    const d = el('div', 'quote');
    d.style.borderColor = 'var(--green)';
    d.appendChild(el('span', null, `${e.kind} — ${e.part}${e.size ? ` (${(e.size / 1024).toFixed(1)} KB)` : ''}`));
    if (e.preview) d.appendChild(el('span', 'who', e.preview.slice(0, 240)));
    else if (e.urls && e.urls.length) d.appendChild(el('span', 'who', e.urls.join(' ')));
    return d;
  });
  return box;
}

/* ---------------------------------------------------------- timeline --- */

function timelinePanel(rep) {
  const box = el('div', 'tl');
  const rows = [
    ['created', rep.timeline.created],
    ['last saved', rep.timeline.modified],
    ['last printed', rep.timeline.lastPrinted],
  ];
  const kv = el('table', 'kv');
  for (const [k, v] of rows) {
    if (!v) continue;
    const tr = el('tr');
    tr.appendChild(el('td', null, k));
    tr.appendChild(el('td', null, String(v).replace('T', ' ').replace('.000Z', ' UTC')));
    kv.appendChild(tr);
  }
  const add = (k, v) => {
    if (v === null || v === undefined || v === '') return;
    const tr = el('tr');
    tr.appendChild(el('td', null, k));
    tr.appendChild(el('td', null, String(v)));
    kv.appendChild(tr);
  };
  add('saves', rep.timeline.revisions);
  add('editing minutes', rep.timeline.totalEditMinutes);
  add('wall clock', rep.timeline.wallClockMinutes != null ? humanMinutes(rep.timeline.wallClockMinutes) : null);
  if (rep.timeline.revisionDates.length) add('dated edits', `${rep.timeline.revisionDates.length} (${shortDate(rep.timeline.revisionDates[0])} → ${shortDate(rep.timeline.revisionDates.at(-1))})`);
  add('zip part timestamps', rep.timeline.zipMtimes.join(', ').slice(0, 120));
  box.appendChild(kv);
  if (rep.timeline.hourHistogram) {
    const h = el('div', 'hours');
    const max = Math.max(...rep.timeline.hourHistogram);
    rep.timeline.hourHistogram.forEach((n) => {
      const i = el('i');
      i.style.height = (n / max) * 100 + '%';
      i.style.opacity = n ? '.85' : '.12';
      i.title = `${n} dated edits`;
      h.appendChild(i);
    });
    box.appendChild(h);
    const lab = el('div', 'hourlab');
    lab.innerHTML = '<span>00</span><span>06</span><span>12</span><span>18</span><span>23 UTC</span>';
    box.appendChild(lab);
  }
  return box;
}
const humanMinutes = (m) => {
  if (m < 90) return `${m} min`;
  if (m < 60 * 48) return `${(m / 60).toFixed(1)} h`;
  return `${(m / 1440).toFixed(1)} days`;
};

/* ---------------------------------------------------------- metadata --- */

function metaPanel(rep) {
  const box = el('div');
  box.style.padding = '10px 15px';
  const kv = el('table', 'kv');
  const add = (k, v) => {
    if (v === null || v === undefined || v === '' || (Array.isArray(v) && !v.length)) return;
    const tr = el('tr');
    tr.appendChild(el('td', null, k));
    const td = el('td');
    td.textContent = Array.isArray(v) ? v.join(', ') : String(v);
    tr.appendChild(td);
    kv.appendChild(tr);
  };
  const p = rep.props;
  add('title', p.title);
  add('subject', p.subject);
  add('keywords', p.keywords);
  add('description', p.description);
  add('creator', p.creator);
  add('last saved by', p.lastModifiedBy);
  add('manager', p.manager);
  add('company', p.company);
  add('category', p.category);
  add('content status', p.contentStatus);
  add('template', p.template);
  add('attached template', rep.settings.attachedTemplate);
  add('producer', [p.application, p.appVersion, p.appVersionLabel].filter(Boolean).join(' · '));
  add('pages / words / chars', [p.pages, p.words, p.characters].filter((x) => x != null).join(' / '));
  add('spell/grammar state', rep.settings.proofState ? `${rep.settings.proofState.spelling || '?'} / ${rep.settings.proofState.grammar || '?'}` : null);
  add('theme language', rep.settings.themeFontLang);
  add('document protection', rep.settings.documentProtection ? `${rep.settings.documentProtection.edit || ''} ${rep.settings.documentProtection.hasHash ? '(password hash present)' : ''}` : null);
  add('track changes now', rep.settings.trackChangesOn ? 'ON' : null);
  add('rsidRoot', rep.settings.rsidRoot);
  add('language runs', rep.locales.map((l) => `${l.name} (${l.n})`));
  add('fonts in body', rep.fonts.map((f) => f.name));
  add('font table', rep.fontTable);
  for (const i of rep.identities) add('name: ' + i.name, i.roles.join('; '));
  for (const i of rep.identityIds || []) add('account: ' + (i.author || '?'), [i.provider, i.userId].filter(Boolean).join(' '));
  for (const c of rep.customProps) add('prop: ' + c.name, c.value);
  for (const e of rep.external) add('external ' + e.type, e.target);
  for (const f of rep.fields.slice(0, 20)) add('field', f.instr);
  box.appendChild(kv);
  return box;
}

function containerPanel(rep) {
  const box = el('div');
  box.style.padding = '10px 15px';
  const t = el('table', 'kv');
  for (const part of rep.container.parts) {
    const tr = el('tr');
    tr.appendChild(el('td', null, part.name));
    tr.appendChild(el('td', null, `${fmt(part.size)} B${part.mtime ? ' · ' + part.mtime.replace('T', ' ') : ''}${part.method === 0 ? ' · stored' : ''}`));
    t.appendChild(tr);
  }
  box.appendChild(t);
  return box;
}

function footnote() {
  const f = el('div', 'footnote');
  f.innerHTML = `
    <h4>what you are looking at</h4>
    <p>Word documents may contain <code>w:rsid</code> revision-save identifiers, language settings and editing-time fields. This view groups the stored identifiers and shows metadata as recorded. These fields can be incomplete, inherited, reset or changed by software; they do not prove writing sessions, paste events, authorship or time actually spent. Recoverable text and image content are separate, directly inspectable findings.</p>
    <h4>how to actually clean a document</h4>
    <p>File → Info → <em>Inspect Document</em> removes most of it. Accepting/rejecting all tracked changes, deleting comments, using <em>Compress Pictures → delete cropped areas</em>, and finally "Save As" a fresh copy (or print to PDF) removes the rest. Sending a PDF does <em>not</em> automatically remove EXIF from embedded images or cropped-away pixels.</p>
    <p style="color:var(--dimmer)">Doc Rinse runs entirely in this page. Your file is never uploaded, and there is no server to upload it to.</p>`;
  return f;
}

/* ----------------------------------------------------------- compare --- */

function compareBox() {
  const box = el('div', 'cmp');
  const mini = el('div', 'mini', 'Drop a second version here to see what changed between saves');
  const input = el('input');
  input.type = 'file';
  input.accept = '.docx';
  input.className = 'hidden';
  mini.onclick = () => input.click();
  mini.ondragover = (e) => {
    e.preventDefault();
    mini.style.borderColor = 'var(--amber)';
  };
  mini.ondragleave = () => (mini.style.borderColor = '');
  mini.ondrop = async (e) => {
    e.preventDefault();
    mini.style.borderColor = '';
    if (e.dataTransfer.files[0]) await loadCompare(e.dataTransfer.files[0], box);
  };
  input.onchange = async () => input.files[0] && (await loadCompare(input.files[0], box));
  box.appendChild(mini);
  box.appendChild(input);
  return box;
}

async function loadCompare(file, box) {
  const bytes = new Uint8Array(await file.arrayBuffer());
  let rep;
  try {
    rep = await analyzePackage(bytes, { name: file.name, keepBytes: false });
  } catch (e) {
    box.appendChild(el('div', 'err', 'could not read: ' + e.message));
    return;
  }
  COMPARE = rep;
  box.querySelectorAll('.cmpout').forEach((n) => n.remove());
  const out = el('div', 'cmpout');
  out.appendChild(renderLineage(CURRENT, rep, file.name));
  box.appendChild(out);
}

function renderLineage(a, b, bname) {
  const box = el('div');
  const L = compareDocs(a, b, bname);
  const h = el('div');
  h.style.cssText = 'font-size:13.5px;margin:12px 0 8px;color:var(--ink)';
  h.textContent = L.verdict;
  box.appendChild(h);

  const k = el('table', 'kv');
  const add = (x, y) => {
    const tr = el('tr');
    tr.appendChild(el('td', null, x));
    tr.appendChild(el('td', null, y));
    k.appendChild(tr);
  };
  add('shared sessions', `${L.shared.length}`);
  add('only in open file', `${L.onlyA.length}`);
  add(`only in ${bname.slice(0, 22)}`, `${L.onlyB.length}`);
  add('rsidRoot', `${a.settings.rsidRoot || '—'} vs ${b.settings.rsidRoot || '—'}`);
  add('saves', `${a.timeline.revisions || '?'} vs ${b.timeline.revisions || '?'}`);
  add('edit minutes', `${a.props.totalTime ?? '?'} vs ${b.props.totalTime ?? '?'}`);
  box.appendChild(k);

  if (L.paragraphs) {
    const P = L.paragraphs;
    const h2 = el('div');
    h2.style.cssText = 'font:11px var(--mono);letter-spacing:.14em;text-transform:uppercase;color:var(--dimmer);margin:14px 0 6px';
    h2.textContent = `paragraph identity · ${P.survived} survive, ${P.rewritten.length} rewritten, ${P.added.length} new, ${P.removed.length} gone`;
    box.appendChild(h2);
    const d = el('div', 'diff');
    for (const r of P.rewritten.slice(0, 25)) {
      const row = el('div', 'diffrow');
      row.appendChild(el('div', 'lab', 'paraId ' + r.paraId));
      row.appendChild(wordDiff(r.before, r.after));
      d.appendChild(row);
    }
    for (const r of P.added.slice(0, 10)) {
      const row = el('div', 'diffrow');
      row.appendChild(el('div', 'lab', 'new paragraph ' + r.paraId));
      row.appendChild(el('ins', null, r.text.slice(0, 300)));
      d.appendChild(row);
    }
    for (const r of P.removed.slice(0, 10)) {
      const row = el('div', 'diffrow');
      row.appendChild(el('div', 'lab', 'removed paragraph ' + r.paraId));
      row.appendChild(el('del', null, r.text.slice(0, 300)));
      d.appendChild(row);
    }
    box.appendChild(d);
  }
  return box;
}

// word-level LCS diff, good enough for prose
function wordDiff(x, y) {
  const A = x.split(/(\s+)/), B = y.split(/(\s+)/);
  const n = A.length, m = B.length;
  const cap = 600;
  const out = el('div');
  if (n > cap || m > cap) {
    out.appendChild(el('del', null, x.slice(0, 200)));
    out.appendChild(el('ins', null, y.slice(0, 200)));
    return out;
  }
  const L = Array.from({ length: n + 1 }, () => new Uint16Array(m + 1));
  for (let i = n - 1; i >= 0; i--) for (let j = m - 1; j >= 0; j--) L[i][j] = A[i] === B[j] ? L[i + 1][j + 1] + 1 : Math.max(L[i + 1][j], L[i][j + 1]);
  let i = 0, j = 0;
  const push = (tag, txt) => {
    if (!txt) return;
    const last = out.lastChild;
    if (last && last.nodeName.toLowerCase() === tag) last.textContent += txt;
    else out.appendChild(tag === 'span' ? el('span', null, txt) : el(tag, null, txt));
  };
  while (i < n && j < m) {
    if (A[i] === B[j]) { push('span', A[i]); i++; j++; }
    else if (L[i + 1][j] >= L[i][j + 1]) { push('del', A[i]); i++; }
    else { push('ins', B[j]); j++; }
  }
  while (i < n) push('del', A[i++]);
  while (j < m) push('ins', B[j++]);
  return out;
}

/* --------------------------------------------------------- bootstrap --- */

/* ---------------------------------------------------------------- scrub --- */

// Crop baking in the browser: decode the picture, draw only the window the
// document actually shows, hand back the re-encoded bytes. This is the one part
// of the scrubber that needs a host — the CLI has no image codec.
async function bakeImage({ bytes, mime, crop }) {
  const blob = new Blob([bytes], { type: mime || 'image/png' });
  const bmp = await createImageBitmap(blob);
  const x = Math.round((bmp.width * crop.l) / 100);
  const y = Math.round((bmp.height * crop.t) / 100);
  const w = Math.max(1, Math.round(bmp.width * (1 - (crop.l + crop.r) / 100)));
  const h = Math.max(1, Math.round(bmp.height * (1 - (crop.t + crop.b) / 100)));
  const canvas = typeof OffscreenCanvas !== 'undefined' ? new OffscreenCanvas(w, h) : Object.assign(document.createElement('canvas'), { width: w, height: h });
  const ctx = canvas.getContext('2d');
  ctx.drawImage(bmp, x, y, w, h, 0, 0, w, h);
  bmp.close && bmp.close();
  const type = mime === 'image/jpeg' ? 'image/jpeg' : 'image/png';
  const out = canvas.convertToBlob ? await canvas.convertToBlob({ type, quality: 0.92 }) : await new Promise((r) => canvas.toBlob(r, type, 0.92));
  return new Uint8Array(await out.arrayBuffer());
}

function scrubPanel(rep, name) {
  const box = el('div', 'scrub');
  const fixes = plan(rep);
  const lead = el('p', 'lead');
  lead.innerHTML = `Choose the supported removals below. Doc Rinse rewrites the container and re-analyzes the result in this tab. Some options remove content; review the downloaded file in your Office application. Re-analysis is not a guarantee of unchanged layout or removal of every possible trace. <b>The cleaned file never leaves your machine either.</b>`;
  box.appendChild(lead);

  const list = el('div', 'fixlist');
  const boxes = new Map();
  for (const f of fixes) {
    const row = el('label', 'fix' + (f.risk === 'content' ? ' content' : ''));
    const cb = el('input');
    cb.type = 'checkbox';
    cb.checked = f.on;
    boxes.set(f.id, cb);
    row.appendChild(cb);
    const txt = el('div');
    const t = el('div', 't');
    t.appendChild(el('b', null, f.label));
    t.appendChild(el('span', 'n', f.count ? `${fmt(f.count)} ${f.unit}` : ''));
    if (f.risk === 'content') t.appendChild(el('span', 'warn', 'changes content'));
    if (f.risk === 'reencode') t.appendChild(el('span', 'warn', 're-encodes the image'));
    txt.appendChild(t);
    txt.appendChild(el('div', 'd', f.detail));
    row.appendChild(txt);
    list.appendChild(row);
  }
  box.appendChild(list);

  const acts = el('div', 'acts');
  const go = el('button', 'tgl primary', 'clean this file');
  const res = el('div', 'scrubres');
  acts.appendChild(go);
  const all = el('button', 'tgl', 'select all');
  all.onclick = () => {
    const target = [...boxes.values()].some((c) => c.checked === false);
    for (const c of boxes.values()) c.checked = target;
    all.textContent = target ? 'select none' : 'select all';
  };
  acts.appendChild(all);
  box.appendChild(acts);
  box.appendChild(res);

  go.onclick = async () => {
    go.disabled = true;
    go.textContent = 'cleaning…';
    res.innerHTML = '';
    try {
      const chosen = [...boxes].filter(([, c]) => c.checked).map(([id]) => id);
      const t0 = performance.now();
      const out = await scrub(rep._bytes, rep, { fixes: chosen, bakeImage });
      const after = await analyzePackage(out.bytes, { name: name, keepBytes: false });
      const v = verifyScrub(rep, after);
      const ms = Math.round(performance.now() - t0);
      window.__clean = out; // same hook as window.__rep: inspect the result in the console

      const head = el('div', 'sum');
      const ok = v.text && v.text.ok;
      head.appendChild(el('b', ok ? 'ok' : 'bad', ok ? 'parsed-word check passed; rendering not verified' : 'parsed text changed — check the diff below'));
      head.appendChild(el('span', null, `${v.before.findings} findings → ${v.after.findings} · ${(rep.file.bytes / 1024).toFixed(0)} KB → ${(out.bytes.length / 1024).toFixed(0)} KB · ${ms} ms`));
      res.appendChild(head);

      const tbl = el('table', 'delta');
      const add = (k, a, b) => {
        if (!a && !b) return;
        const tr = el('tr');
        tr.appendChild(el('td', null, k));
        tr.appendChild(el('td', 'num', fmt(a)));
        tr.appendChild(el('td', 'arrow', '→'));
        const cell = el('td', 'num ' + (b === 0 && a > 0 ? 'good' : b < a ? 'good' : b > a ? 'bad' : ''), fmt(b));
        tr.appendChild(cell);
        tbl.appendChild(tr);
      };
      add('editing sessions', v.before.sessions, v.after.sessions);
      add('names', v.before.identities, v.after.identities);
      add('invisible runs', v.before.hidden, v.after.hidden);
      add('recoverable deletions', v.before.deleted, v.after.deleted);
      add('comments', v.before.comments, v.after.comments);
      add('images with metadata', v.before.imgMeta, v.after.imgMeta);
      add('cropped images', v.before.cropped, v.after.cropped);
      add('undisplayed images', v.before.unusedImages, v.after.unusedImages);
      add('words', v.before.words, v.after.words);
      res.appendChild(tbl);

      if (v.remaining.length) {
        const rem = el('div', 'remain');
        rem.appendChild(el('b', null, 'still standing:'));
        rem.appendChild(el('span', null, v.remaining.join(' · ')));
        res.appendChild(rem);
      }
      for (const n of out.notes) res.appendChild(el('div', 'note', n));

      const dlrow = el('div', 'acts');
      const b = el('button', 'tgl primary', `download cleaned .docx (${(out.bytes.length / 1024).toFixed(0)} KB)`);
      b.onclick = () => {
        const url = URL.createObjectURL(new Blob([out.bytes], { type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' }));
        const a = document.createElement('a');
        a.href = url;
        a.download = name.replace(/\.docx$/i, '') + '.clean.docx';
        a.click();
        setTimeout(() => URL.revokeObjectURL(url), 4000);
      };
      dlrow.appendChild(b);
      const insp = el('button', 'tgl', 'analyze the result');
      insp.onclick = () => loadFile(new File([out.bytes], name.replace(/\.docx$/i, '') + '.clean.docx'));
      dlrow.appendChild(insp);
      res.appendChild(dlrow);
    } catch (e) {
      res.appendChild(el('div', 'err', 'scrub failed: ' + e.message));
      console.error(e);
    }
    go.disabled = false;
    go.textContent = 'clean this file';
  };
  return box;
}

export async function loadFile(file) {
  const out = $('#out');
  const bytes = new Uint8Array(await file.arrayBuffer());
  try {
    const t0 = performance.now();
    const rep = await analyzePackage(bytes, { name: file.name, keepBytes: true });
    rep._ms = Math.round(performance.now() - t0);
    rep._bytes = bytes; // the scrub panel rewrites the original container
    renderReport(rep, file.name);
    window.__rep = rep;
  } catch (e) {
    out.innerHTML = '';
    out.appendChild(el('div', 'err', `${file.name}: ${e.message}`));
    console.error(e);
  }
}

// used by the embedded-sample buttons to demo the version comparison
export async function compareWith(file) {
  const box = document.querySelector('.cmp');
  if (box) await loadCompare(file, box);
}

export function wireUp() {
  const dz = $('#drop');
  const input = $('#file');
  dz.onclick = () => input.click();
  input.onchange = () => input.files[0] && loadFile(input.files[0]);
  for (const ev of ['dragenter', 'dragover']) {
    window.addEventListener(ev, (e) => {
      e.preventDefault();
      dz.classList.add('hot');
    });
  }
  window.addEventListener('dragleave', (e) => {
    if (e.relatedTarget === null) dz.classList.remove('hot');
  });
  window.addEventListener('drop', (e) => {
    e.preventDefault();
    dz.classList.remove('hot');
    const f = e.dataTransfer.files[0];
    if (f) loadFile(f);
  });

}
