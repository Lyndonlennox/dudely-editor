// Dudely Editor PC: "All tools" page, the screen that runs a tool, the page organiser and the form maker.
import { TOOLS, GROUPS, byId } from './pdftools.js';
import { icon } from './icons.js';
import { showView, saveFile } from './tools.js';
import { $, L, baseName, isPdf, isImg, fmtSize, esc, pdfBlob, libLoad, jsLoad, renderPage, zipOf, vis, rectToUser, tick, pdfjsLib } from './pdfcore.js';
const { PDFDocument, degrees, rgb } = L;
const GCOL = Object.fromEntries(GROUPS.map(([k, , c]) => [k, c]));
const busy = (t) => { const b = $('busy'); b.textContent = t || ''; b.style.display = t ? 'block' : 'none'; };
const ACCEPT = { pdf: 'application/pdf,.pdf', image: 'image/*', html: '.html,.htm,text/html', docx: '.docx', any: '' };
const accepts = (t, f) => t.accept === 'pdf' ? isPdf(f) : t.accept === 'image' ? isImg(f) : t.accept === 'html' ? /\.html?$/i.test(f.name) : t.accept === 'docx' ? /\.docx$/i.test(f.name) : true;
const KIND = { pdf: 'PDF', image: 'photos or pictures (JPG, PNG)', html: 'web page (.html)', docx: 'Word file', any: 'file' };

// ------------------------------------------------------------------ recent tools (only on this computer)
const recent = { get() { try { return JSON.parse(localStorage.getItem('dudely-recent') || '[]').filter((id) => byId[id]); } catch (e) { return []; } },
  add(id) { try { localStorage.setItem('dudely-recent', JSON.stringify([id, ...recent.get().filter((x) => x !== id)].slice(0, 6))); } catch (e) { /* private mode */ } } };

// ================================================================== ALL TOOLS PAGE
function toolBtn(t) {
  return `<button class="tl" data-tool-id="${t.id}" title="${esc(t.desc)}"><span class="tl-ic" style="--c:${GCOL[t.group]}">${icon(t.icon)}</span><span class="tl-nm">${esc(t.name)}</span>${t.isNew ? '<span class="new">New</span>' : ''}</button>`;
}
export function drawHome() {
  const q = ($('homeSearch').value || '').trim().toLowerCase();
  const match = (t) => !q || (t.name + ' ' + t.desc).toLowerCase().includes(q);
  const rec = recent.get().map((id) => byId[id]).filter(match);
  let html = '';
  if (rec.length && !q) html += `<section class="cat"><h3 style="--c:#5B6E70">${icon('clock', 'i sm')}Recent</h3>${rec.map(toolBtn).join('')}</section>`;
  for (const [g, label] of GROUPS) {
    const list = TOOLS.filter((t) => t.group === g && match(t)); if (!list.length) continue;
    html += `<section class="cat"><h3 style="--c:${GCOL[g]}">${esc(label)}</h3>${list.map(toolBtn).join('')}</section>`;
  }
  $('homeCats').innerHTML = html || `<p class="muted">No tool matches "${esc(q)}".</p>`;
  $('homeCount').textContent = `${TOOLS.length} tools`;
  document.querySelectorAll('[data-ntools]').forEach((e) => (e.textContent = TOOLS.length));
}
$('homeSearch').addEventListener('input', drawHome);
$('homeSearch').addEventListener('keydown', (e) => { if (e.key === 'Enter') { const b = $('homeCats').querySelector('.tl'); if (b) b.click(); } });
document.addEventListener('click', (e) => { const b = e.target.closest('[data-tool-id]'); if (b) openTool(b.dataset.toolId); });

export function openTool(id) {
  const t = byId[id]; if (!t) return; recent.add(id);
  if (t.go) {
    showView(t.go);
    const open = !$('fileChip').classList.contains('none');
    if (t.tool) document.querySelector(`.tool[data-tool="${t.tool}"]`)?.click();
    if (t.then === 'word') { if (open) $('btnWord').click(); else { $('hint').textContent = 'Open the PDF; then click "Save as Word" at the top.'; $('btnOpen').click(); } }
    else if (!open && t.go === 'edit') $('btnOpen').click();
    return;
  }
  if (t.pages) return openPages(t);
  if (t.forms) return openForms();
  openRunner(t);
}

// ================================================================== THE TOOL SCREEN
const RN = { t: null, files: [], defs: [], vals: {}, outs: [] };
function head(t) {
  return `<div class="t-head"><button class="ghost back" data-back>${icon('back')}All tools</button>
    <span class="t-ic" style="--c:${GCOL[t.group]}">${icon(t.icon)}</span><div><h2>${esc(t.name)}</h2><p>${esc(t.desc)}</p></div></div>`;
}
document.addEventListener('click', (e) => { if (e.target.closest('[data-back]')) { showView('home'); drawHome(); } });

function openRunner(t) {
  Object.assign(RN, { t, files: [], defs: t.opts.map((o) => ({ ...o })), vals: {}, outs: [] });
  for (const o of RN.defs) RN.vals[o.id] = o.value;
  showView('tool', t.name, t.desc);
  $('toolRoot').innerHTML = `${head(t)}
    <div class="t-body">
      <div class="panel t-files" id="rnFiles"></div>
      <div class="panel t-opts"><h4>Options</h4><form id="rnForm" class="opts" onsubmit="return false"></form>
        <div class="run-row"><button class="primary lg" id="rnGo" disabled>${icon('check')}<span id="rnGoLbl">${esc(t.name)}</span></button></div>
        <p class="status" id="rnStatus" role="status"></p><ul class="notes" id="rnNotes"></ul><div id="rnOut"></div></div>
    </div>
    <input type="file" id="rnInput" hidden ${t.multi ? 'multiple' : ''} accept="${ACCEPT[t.accept]}">`;
  $('rnInput').onchange = (e) => { addFiles([...e.target.files]); e.target.value = ''; };
  const zone = $('rnFiles');
  zone.addEventListener('dragover', (e) => { e.preventDefault(); zone.classList.add('over'); });
  zone.addEventListener('dragleave', () => zone.classList.remove('over'));
  zone.addEventListener('drop', (e) => { e.preventDefault(); e.stopPropagation(); zone.classList.remove('over'); addFiles([...e.dataTransfer.files]); });
  $('rnGo').onclick = run;
  drawFiles(); drawOpts();
}
async function addFiles(list) {
  const t = RN.t; const ok = list.filter((f) => accepts(t, f));
  if (list.length && !ok.length) { setStatus(`That is not a ${KIND[t.accept]}. ${t.name} needs a ${KIND[t.accept]}.`); return; }
  if (!t.multi) RN.files = [];
  for (const f of (t.multi ? ok : ok.slice(0, 1))) {
    const bytes = new Uint8Array(await f.arrayBuffer()); const it = { name: f.name, type: f.type, size: f.size, bytes, pages: null };
    RN.files.push(it);
    if (isPdf(f)) jsLoad(bytes).then((d) => { it.pages = d.numPages; d.destroy(); drawFiles(); }).catch((e) => { it.pages = e.name === 'PasswordException' ? 'locked' : '?'; drawFiles(); });
  }
  RN.outs = []; $('rnOut').innerHTML = ''; $('rnNotes').innerHTML = ''; setStatus('');
  if (t.fields && RN.files.length) {
    try { busy('Reading the file…'); RN.defs = (await t.fields(RN.files)).concat(t.opts.map((o) => ({ ...o }))); RN.vals = {}; for (const o of RN.defs) RN.vals[o.id] = o.value; }
    catch (e) { setStatus(e.message); RN.defs = []; } finally { busy(''); }
  }
  drawFiles(); drawOpts();
}
function drawFiles() {
  const t = RN.t; const z = $('rnFiles'); if (!z) return;
  if (!RN.files.length) {
    z.innerHTML = `<div class="dz">${icon('file', 'i big')}<b>Drop ${t.multi ? 'files' : 'a file'} here</b><span>${KIND[t.accept][0].toUpperCase() + KIND[t.accept].slice(1)}${t.multi ? ', one or more' : ''}</span>
      <button class="primary" id="rnPick">${icon('folder')}Choose ${t.multi ? 'files' : 'a file'}</button><small>${icon('lock', 'i sm')}Stays on this computer</small></div>`;
    $('rnPick').onclick = () => $('rnInput').click();
  } else {
    z.innerHTML = `<div class="fl-head"><h4>${RN.files.length} file${RN.files.length > 1 ? 's' : ''}</h4><button id="rnMore">${icon('plus')}${t.multi ? 'Add more' : 'Change file'}</button></div>
      <ol class="fl">${RN.files.map((f, i) => `<li><span class="fl-ic">${icon(isImg(f) ? 'image' : 'file')}</span><span class="fl-nm" title="${esc(f.name)}">${esc(f.name)}</span>
        <span class="fl-meta">${f.pages === 'locked' ? 'Password-protected' : f.pages ? `${f.pages} page${f.pages > 1 ? 's' : ''} · ` : ''}${f.pages === 'locked' ? '' : fmtSize(f.size)}</span>
        <span class="fl-act">${t.multi && RN.files.length > 1 ? `<button class="icon-btn" data-up="${i}" aria-label="Move up" ${i ? '' : 'disabled'}>${icon('up')}</button><button class="icon-btn" data-dn="${i}" aria-label="Move down" ${i < RN.files.length - 1 ? '' : 'disabled'}>${icon('down')}</button>` : ''}
        <button class="icon-btn" data-rm="${i}" aria-label="Remove">${icon('x')}</button></span></li>`).join('')}</ol>
      ${t.multi ? '<p class="muted small">Files are used in this order. Drop more files here to add them.</p>' : ''}`;
    $('rnMore').onclick = () => $('rnInput').click();
    z.querySelectorAll('[data-rm]').forEach((b) => (b.onclick = () => { RN.files.splice(+b.dataset.rm, 1); if (RN.t.fields && !RN.files.length) { RN.defs = RN.t.opts.map((o) => ({ ...o })); } drawFiles(); drawOpts(); }));
    z.querySelectorAll('[data-up]').forEach((b) => (b.onclick = () => { const i = +b.dataset.up; [RN.files[i - 1], RN.files[i]] = [RN.files[i], RN.files[i - 1]]; drawFiles(); }));
    z.querySelectorAll('[data-dn]').forEach((b) => (b.onclick = () => { const i = +b.dataset.dn; [RN.files[i + 1], RN.files[i]] = [RN.files[i], RN.files[i + 1]]; drawFiles(); }));
  }
  const need = RN.t.min || 1; $('rnGo').disabled = RN.files.length < need;
  if (RN.files.length && RN.files.length < need) setStatus(`Add at least ${need} files.`);
}
function drawOpts() {
  const f = $('rnForm'); const defs = RN.defs;
  if (!defs.length) { f.innerHTML = `<p class="muted">${RN.t.fields && !RN.files.length ? 'Choose a file first; its details appear here.' : 'No options needed. Add your file and click the button below.'}</p>`; return; }
  f.innerHTML = defs.map((o) => {
    if (o.show && !o.show(RN.vals)) return '';
    const v = RN.vals[o.id]; const id = 'o_' + o.id; const hint = o.hint ? `<small>${esc(o.hint)}</small>` : '';
    if (o.type === 'info') return `<p class="opt-info">${icon('info', 'i sm')}${esc(o.label)}</p>`;
    if (o.type === 'check') return `<label class="o-chk"><input type="checkbox" id="${id}" data-k="${o.id}" ${v ? 'checked' : ''}><span>${esc(o.label)}</span></label>`;
    let ctl;
    if (o.type === 'select') ctl = `<select id="${id}" data-k="${o.id}">${o.options.map(([a, b]) => `<option value="${esc(a)}" ${String(a) === String(v) ? 'selected' : ''}>${esc(b)}</option>`).join('')}</select>`;
    else if (o.type === 'textarea') ctl = `<textarea id="${id}" data-k="${o.id}" rows="${o.rows || 4}">${esc(v ?? '')}</textarea>`;
    else if (o.type === 'range') ctl = `<div class="rng"><input type="range" id="${id}" data-k="${o.id}" min="${o.min}" max="${o.max}" value="${v}"><output>${v}</output></div>`;
    else ctl = `<input id="${id}" data-k="${o.id}" type="${o.type === 'password' ? 'password' : o.type === 'number' ? 'number' : o.type === 'color' ? 'color' : 'text'}" value="${esc(v ?? '')}" ${o.min !== undefined ? `min="${o.min}"` : ''} ${o.max !== undefined ? `max="${o.max}"` : ''} ${o.step ? `step="${o.step}"` : ''} ${o.placeholder ? `placeholder="${esc(o.placeholder)}"` : ''} autocomplete="off">`;
    return `<label class="o-fld ${o.type === 'textarea' ? 'wide' : ''}" for="${id}"><span>${esc(o.label)}</span>${ctl}${hint}</label>`;
  }).join('');
  f.querySelectorAll('[data-k]').forEach((el) => {
    const upd = () => { const k = el.dataset.k; RN.vals[k] = el.type === 'checkbox' ? el.checked : el.value; if (el.type === 'range') el.nextElementSibling.textContent = el.value; };
    el.addEventListener('input', upd);
    el.addEventListener('change', () => { upd(); if (RN.defs.some((o) => o.show)) { const a = document.activeElement?.id; drawOpts(); if (a) $(a)?.focus(); } });
  });
}
const setStatus = (t) => { if ($('rnStatus')) $('rnStatus').textContent = t; };
async function run() {
  const t = RN.t; const notes = []; $('rnNotes').innerHTML = ''; $('rnOut').innerHTML = ''; $('rnGo').disabled = true;
  for (const o of RN.defs) if (o.required && (o.show ? o.show(RN.vals) : true) && !String(RN.vals[o.id] ?? '').trim()) { setStatus(`Fill in "${o.label}".`); $('rnGo').disabled = false; return; }
  const ctx = { progress: (s) => { setStatus(s); busy(s); }, note: (s) => notes.push(s), opts: RN.defs };
  const t0 = performance.now();
  try {
    ctx.progress('Working…'); await tick();
    RN.outs = await t.run(RN.files, { ...RN.vals }, ctx);
    const sec = ((performance.now() - t0) / 1000).toFixed(1);
    setStatus(RN.outs.length ? `Done in ${sec} s. ${RN.outs.length === 1 ? 'Your file is ready.' : `${RN.outs.length} files are ready.`}` : 'Done.');
    drawOuts();
  } catch (e) { console.error(e); setStatus(e.message || String(e)); }
  finally { busy(''); $('rnGo').disabled = false; $('rnNotes').innerHTML = notes.map((n) => `<li>${esc(n)}</li>`).join(''); }
}
function drawOuts() {
  const o = RN.outs; if (!o.length) return;
  const isP = (x) => /\.pdf$/i.test(x.name);
  $('rnOut').innerHTML = `<div class="results"><h4>${icon('check', 'i sm')}Ready</h4><ul>${o.slice(0, 200).map((x, i) => `<li><span class="fl-ic">${icon(isP(x) ? 'file' : /\.(jpe?g|png)$/i.test(x.name) ? 'image' : 'file')}</span>
      <span class="fl-nm" title="${esc(x.name)}">${esc(x.name)}</span><span class="fl-meta">${fmtSize(x.blob.size)}</span>
      <span class="fl-act">${isP(x) ? `<button data-edit="${i}" title="Open in the editor">${icon('edit')}Edit</button>` : ''}<button class="secondary" data-dl="${i}">${icon('download')}Download</button></span></li>`).join('')}</ul>
      ${o.length > 200 ? `<p class="muted small">…and ${o.length - 200} more in the zip.</p>` : ''}
      ${o.length > 1 ? `<div class="run-row"><button class="primary" id="rnZip">${icon('download')}Download all (zip)</button></div>` : ''}</div>`;
  $('rnOut').querySelectorAll('[data-dl]').forEach((b) => (b.onclick = () => { const x = o[+b.dataset.dl]; saveFile(x.blob, x.name, 'File'); }));
  $('rnOut').querySelectorAll('[data-edit]').forEach((b) => (b.onclick = async () => { const x = o[+b.dataset.edit]; showView('edit'); await window.dudelyOpen([new File([x.blob], x.name, { type: 'application/pdf' })]); }));
  if ($('rnZip')) $('rnZip').onclick = async () => { busy('Making the zip…'); try { await saveFile(await zipOf(o), `${baseName(RN.files[0]?.name || 'files')}_${RN.t.id}.zip`, 'Zip file'); } finally { busy(''); } };
}

// ================================================================== PAGE ORGANISER (organise, delete, extract, rotate)
const PG = { mode: 'organize', t: null, name: '', bytes: null, js: null, items: [], drag: -1, last: -1, obs: null };
const MODES = {
  organize: { go: 'Save PDF', tip: 'Drag pages to put them in order. Click to select, then turn, copy, add a blank page after, or delete.' },
  delete: { go: 'Delete selected pages', tip: 'Click the pages you want to remove (Shift+click selects a run of pages), then click Delete selected pages.' },
  extract: { go: 'Extract selected pages', tip: 'Click the pages you want to keep in a new PDF (Shift+click selects a run of pages).' },
  rotate: { go: 'Save PDF', tip: 'Click pages to select them, then turn them. With nothing selected, all pages turn.' },
};
function openPages(t) {
  Object.assign(PG, { mode: t.pages, t, name: '', bytes: null, items: [] }); PG.js?.destroy(); PG.js = null;
  showView('tool', t.name, t.desc);
  $('toolRoot').innerHTML = `${head(t)}
    <div class="panel pg" id="pgDrop">
      <div class="dz" id="pgEmpty">${icon('file', 'i big')}<b>Drop a PDF here</b><span>${esc(MODES[t.pages].tip)}</span><button class="primary" id="pgPick">${icon('folder')}Choose a PDF</button><small>${icon('lock', 'i sm')}Stays on this computer</small></div>
      <div id="pgWork" hidden>
        <div class="pg-bar">
          <span class="pg-file" id="pgFile"></span>
          <div class="seg"><button id="pgAll">Select all</button><button id="pgNone">None</button><button id="pgOdd">Odd</button><button id="pgEven">Even</button></div>
          <div class="seg"><button id="pgRotL" title="Turn left">${icon('rotl')}</button><button id="pgRotR" title="Turn right">${icon('rotate')}</button></div>
          <div class="seg" ${t.pages === 'rotate' ? 'hidden' : ''}><button id="pgDup" title="Copy selected pages">${icon('copy')}Copy</button><button id="pgBlank" title="Add a blank page after the selected page">${icon('blank')}Blank</button><button id="pgRev" title="Reverse the order">${icon('swap')}Reverse</button><button id="pgDel" title="Delete selected pages">${icon('trash')}Delete</button></div>
          <button id="pgChange">${icon('folder')}Other PDF</button>
        </div>
        <p class="muted small" id="pgTip">${esc(MODES[t.pages].tip)}</p>
        <div class="pg-grid" id="pgGrid"></div>
        <div class="pg-foot"><span id="pgCount" class="muted"></span>${t.pages === 'extract' ? '<label class="o-chk"><input type="checkbox" id="pgEach"><span>Each page as its own file</span></label>' : ''}
          <button class="primary lg" id="pgGo">${icon('check')}${MODES[t.pages].go}</button></div>
        <p class="status" id="pgStatus" role="status"></p>
      </div>
    </div><input type="file" id="pgInput" hidden accept="application/pdf,.pdf">`;
  $('pgPick').onclick = $('pgChange').onclick = () => $('pgInput').click();
  $('pgInput').onchange = (e) => { const f = e.target.files[0]; e.target.value = ''; if (f) loadPages(f); };
  const z = $('pgDrop');
  z.addEventListener('dragover', (e) => { if (e.dataTransfer.types.includes('Files')) { e.preventDefault(); z.classList.add('over'); } });
  z.addEventListener('dragleave', () => z.classList.remove('over'));
  z.addEventListener('drop', (e) => { z.classList.remove('over'); const f = [...e.dataTransfer.files].find(isPdf); if (f) { e.preventDefault(); e.stopPropagation(); loadPages(f); } });
  const sel = () => PG.items.filter((x) => x.sel);
  $('pgAll').onclick = () => { PG.items.forEach((x) => (x.sel = true)); drawPages(); };
  $('pgNone').onclick = () => { PG.items.forEach((x) => (x.sel = false)); drawPages(); };
  $('pgOdd').onclick = () => { PG.items.forEach((x, i) => (x.sel = i % 2 === 0)); drawPages(); };
  $('pgEven').onclick = () => { PG.items.forEach((x, i) => (x.sel = i % 2 === 1)); drawPages(); };
  const turn = (d) => { const s = sel().length ? sel() : PG.items; s.forEach((x) => (x.rot = (x.rot + d + 360) % 360)); drawPages(); };
  $('pgRotL').onclick = () => turn(-90); $('pgRotR').onclick = () => turn(90);
  $('pgDel').onclick = () => { if (!sel().length) return pgStatus('Select pages first.'); PG.items = PG.items.filter((x) => !x.sel); drawPages(); };
  $('pgDup').onclick = () => { if (!sel().length) return pgStatus('Select pages first.'); const n = []; PG.items.forEach((x) => { n.push(x); if (x.sel) n.push({ ...x, sel: false, key: Math.random() }); }); PG.items = n; drawPages(); };
  $('pgBlank').onclick = () => { const i = PG.items.findLastIndex((x) => x.sel); const at = i < 0 ? PG.items.length : i + 1; PG.items.splice(at, 0, { src: -1, rot: 0, sel: false, key: Math.random(), w: PG.items[Math.max(0, at - 1)]?.w || 595, h: PG.items[Math.max(0, at - 1)]?.h || 842 }); drawPages(); };
  $('pgRev').onclick = () => { PG.items.reverse(); drawPages(); };
  $('pgGo').onclick = savePages;
}
const pgStatus = (t) => { $('pgStatus').textContent = t; };
async function loadPages(f) {
  busy('Opening…');
  try {
    PG.bytes = new Uint8Array(await f.arrayBuffer()); PG.name = f.name; PG.js?.destroy();
    PG.js = await jsLoad(PG.bytes);
    PG.items = [];
    for (let i = 0; i < PG.js.numPages; i++) { const p = await PG.js.getPage(i + 1); const v = p.getViewport({ scale: 1 }); PG.items.push({ src: i, rot: 0, sel: false, key: Math.random(), w: v.width, h: v.height }); }
    $('pgEmpty').hidden = true; $('pgWork').hidden = false; $('pgFile').innerHTML = `${icon('file', 'i sm')}${esc(f.name)}`; pgStatus('');
    drawPages();
  } catch (e) { pgStatus(e.name === 'PasswordException' ? 'This PDF is password-protected. Use Unlock first.' : 'This PDF could not be opened.'); $('pgEmpty').hidden = false; }
  finally { busy(''); }
}
const thumbCache = new Map();
async function thumb(src) {
  const k = PG.name + '#' + src; if (thumbCache.has(k)) return thumbCache.get(k);
  const p = (async () => { const { canvas } = await renderPage(PG.js, src, 0.32); return canvas.toDataURL('image/jpeg', 0.75); })();
  thumbCache.set(k, p); return p;
}
function drawPages() {
  const g = $('pgGrid'); PG.obs?.disconnect();
  PG.obs = new IntersectionObserver((ents) => ents.forEach(async (en) => { if (!en.isIntersecting) return; PG.obs.unobserve(en.target); const src = +en.target.dataset.src; if (src >= 0) en.target.querySelector('img').src = await thumb(src); }), { root: $('viewTool'), rootMargin: '400px' });
  g.innerHTML = PG.items.map((x, i) => {
    const land = (x.rot % 180 ? x.h > x.w : x.w > x.h);
    return `<div class="pc ${x.sel ? 'sel' : ''}" draggable="true" data-i="${i}" data-src="${x.src}">
      <div class="pc-img ${land ? 'land' : ''}"><div class="pc-rot" style="transform:rotate(${x.rot}deg)${x.rot % 180 ? ` scale(${Math.min(x.w, x.h) / Math.max(x.w, x.h)})` : ''}">${x.src < 0 ? '<div class="blankpg">Blank</div>' : '<img alt="">'}</div>
        <span class="pc-chk">${icon('check', 'i sm')}</span>
        <span class="pc-tools"><button data-r="${i}" title="Turn right" aria-label="Turn right">${icon('rotate', 'i sm')}</button><button data-x="${i}" title="Remove" aria-label="Remove" ${PG.mode === 'rotate' ? 'hidden' : ''}>${icon('x', 'i sm')}</button></span></div>
      <span class="pc-n">${i + 1}${x.src >= 0 && x.src !== i ? `<small> (was ${x.src + 1})</small>` : ''}</span></div>`;
  }).join('');
  g.querySelectorAll('.pc').forEach((el) => {
    PG.obs.observe(el); const i = +el.dataset.i;
    el.onclick = (e) => {
      if (e.target.closest('[data-r]')) { PG.items[i].rot = (PG.items[i].rot + 90) % 360; return drawPages(); }
      if (e.target.closest('[data-x]')) { PG.items.splice(i, 1); return drawPages(); }
      if (e.shiftKey && PG.last >= 0) { const [a, b] = [Math.min(PG.last, i), Math.max(PG.last, i)]; for (let k = a; k <= b; k++) PG.items[k].sel = true; }
      else PG.items[i].sel = !PG.items[i].sel;
      PG.last = i; drawPages();
    };
    el.ondragstart = (e) => { PG.drag = i; el.classList.add('dragging'); e.dataTransfer.effectAllowed = 'move'; e.dataTransfer.setData('text/plain', String(i)); };
    el.ondragend = () => el.classList.remove('dragging');
    el.ondragover = (e) => { if (PG.drag < 0) return; e.preventDefault(); const r = el.getBoundingClientRect(); el.classList.toggle('before', e.clientX < r.left + r.width / 2); el.classList.toggle('after', e.clientX >= r.left + r.width / 2); };
    el.ondragleave = () => el.classList.remove('before', 'after');
    el.ondrop = (e) => {
      if (PG.drag < 0) return; e.preventDefault(); e.stopPropagation(); const r = el.getBoundingClientRect(); let to = i + (e.clientX >= r.left + r.width / 2 ? 1 : 0);
      const [m] = PG.items.splice(PG.drag, 1); if (PG.drag < to) to--; PG.items.splice(to, 0, m); PG.drag = -1; drawPages();
    };
  });
  const ns = PG.items.filter((x) => x.sel).length;
  $('pgCount').textContent = `${PG.items.length} page${PG.items.length === 1 ? '' : 's'}${ns ? ` · ${ns} selected` : ''}`;
  if (PG.mode === 'delete') $('pgGo').lastChild.textContent = ns ? `Delete ${ns} page${ns > 1 ? 's' : ''} and save` : 'Delete selected pages';
  if (PG.mode === 'extract') $('pgGo').lastChild.textContent = ns ? `Extract ${ns} page${ns > 1 ? 's' : ''}` : 'Extract selected pages';
}
async function savePages() {
  let items = PG.items; const each = PG.mode === 'extract' && $('pgEach')?.checked;
  if (PG.mode === 'delete') { if (!items.some((x) => x.sel)) return pgStatus('Click the pages you want to delete first.'); items = items.filter((x) => !x.sel); }
  if (PG.mode === 'extract') { items = items.filter((x) => x.sel); if (!items.length) return pgStatus('Click the pages you want to extract first.'); }
  if (!items.length) return pgStatus('There are no pages left to save.');
  busy('Saving…');
  try {
    const src = await libLoad(PG.bytes);
    const build = async (list) => {
      const d = await PDFDocument.create();
      for (const x of list) {
        if (x.src < 0) { d.addPage([x.w, x.h]); continue; }
        const [p] = await d.copyPages(src, [x.src]); p.setRotation(degrees((p.getRotation().angle + x.rot) % 360)); d.addPage(p);
      }
      return d.save({ useObjectStreams: true });
    };
    const suffix = { organize: 'organised', delete: 'pages_removed', extract: 'extract', rotate: 'rotated' }[PG.mode];
    if (each) {
      const files = []; for (const x of items) files.push({ name: `${baseName(PG.name)}_page${x.src + 1}.pdf`, blob: pdfBlob(await build([x])) });
      await saveFile(await zipOf(files), `${baseName(PG.name)}_pages.zip`, 'Zip file');
    } else await saveFile(pdfBlob(await build(items)), `${baseName(PG.name)}_${suffix}.pdf`, 'PDF document');
    pgStatus(`Saved ${items.length} page${items.length > 1 ? 's' : ''}. The original file is unchanged.`);
  } catch (e) { pgStatus(e.message); } finally { busy(''); }
}

// ================================================================== FORM MAKER
const FM = { name: '', bytes: null, js: null, page: 0, scale: 1, fields: [], type: 'text', drag: null, sel: -1, pw: 0, ph: 0 };
const FTYPES = [['text', 'Text box', 'textbox'], ['multi', 'Paragraph box', 'para'], ['check', 'Tick box', 'checkbox'], ['list', 'Drop-down list', 'list']];
function openForms() {
  const t = byId.forms; Object.assign(FM, { name: '', bytes: null, page: 0, fields: [], sel: -1 }); FM.js?.destroy(); FM.js = null;
  showView('tool', t.name, t.desc);
  $('toolRoot').innerHTML = `${head(t)}
    <div class="panel fm" id="fmDrop">
      <div class="dz" id="fmEmpty">${icon('file', 'i big')}<b>Drop a PDF here</b><span>Then draw boxes on the page where people should type or tick.</span><button class="primary" id="fmPick">${icon('folder')}Choose a PDF</button></div>
      <div id="fmWork" class="fm-work" hidden>
        <div class="fm-left">
          <div class="pg-bar"><div class="seg" id="fmTypes">${FTYPES.map(([k, l, ic]) => `<button data-ft="${k}" aria-pressed="${k === 'text'}">${icon(ic)}${l}</button>`).join('')}</div>
            <div class="seg"><button id="fmPrev" aria-label="Previous page">${icon('back')}</button><span id="fmPg" class="muted"></span><button id="fmNext" aria-label="Next page" style="transform:scaleX(-1)">${icon('back')}</button></div></div>
          <p class="muted small">Choose a box type, then drag on the page to draw it. Click a box to rename it or change it.</p>
          <div class="fm-stage" id="fmStage"><canvas id="fmCanvas"></canvas><div id="fmLayer"></div></div>
        </div>
        <div class="fm-side"><h4>Boxes on this form</h4><div id="fmList"></div><div id="fmEdit"></div>
          <button class="primary lg" id="fmGo">${icon('download')}Save form</button><p class="status" id="fmStatus" role="status"></p></div>
      </div>
    </div><input type="file" id="fmInput" hidden accept="application/pdf,.pdf">`;
  $('fmPick').onclick = () => $('fmInput').click();
  $('fmInput').onchange = (e) => { const f = e.target.files[0]; e.target.value = ''; if (f) loadForm(f); };
  const z = $('fmDrop'); z.addEventListener('dragover', (e) => e.preventDefault());
  z.addEventListener('drop', (e) => { const f = [...e.dataTransfer.files].find(isPdf); if (f) { e.preventDefault(); e.stopPropagation(); loadForm(f); } });
  $('fmTypes').querySelectorAll('[data-ft]').forEach((b) => (b.onclick = () => { FM.type = b.dataset.ft; $('fmTypes').querySelectorAll('[data-ft]').forEach((x) => x.setAttribute('aria-pressed', String(x === b))); }));
  $('fmPrev').onclick = () => { if (FM.page > 0) { FM.page--; drawFormPage(); } };
  $('fmNext').onclick = () => { if (FM.page < FM.js.numPages - 1) { FM.page++; drawFormPage(); } };
  $('fmGo').onclick = saveForm;
  const layer = $('fmLayer');
  layer.addEventListener('pointerdown', (e) => {
    if (e.target.closest('.fbox')) return; const r = layer.getBoundingClientRect();
    FM.drag = { x0: e.clientX - r.left, y0: e.clientY - r.top, el: document.createElement('div') }; FM.drag.el.className = 'fbox drawing'; layer.appendChild(FM.drag.el); layer.setPointerCapture(e.pointerId);
  });
  layer.addEventListener('pointermove', (e) => {
    if (!FM.drag) return; const r = layer.getBoundingClientRect(); const x = e.clientX - r.left, y = e.clientY - r.top; const d = FM.drag;
    Object.assign(d.el.style, { left: Math.min(x, d.x0) + 'px', top: Math.min(y, d.y0) + 'px', width: Math.abs(x - d.x0) + 'px', height: Math.abs(y - d.y0) + 'px' }); d.x1 = x; d.y1 = y;
  });
  layer.addEventListener('pointerup', () => {
    const d = FM.drag; FM.drag = null; if (!d) return; d.el.remove();
    let w = Math.abs((d.x1 ?? d.x0) - d.x0), h = Math.abs((d.y1 ?? d.y0) - d.y0); let x = Math.min(d.x0, d.x1 ?? d.x0), y = Math.min(d.y0, d.y1 ?? d.y0);
    const s = FM.scale; if (w < 6 && h < 6) { w = FM.type === 'check' ? 14 * s : 160 * s; h = FM.type === 'check' ? 14 * s : FM.type === 'multi' ? 60 * s : 20 * s; x -= 2; y -= h / 2; }
    if (FM.type === 'check') { const m = Math.max(w, h); w = h = Math.min(m, 30 * s); }
    const n = FM.fields.filter((f) => f.type === FM.type).length + 1; const base = { text: 'Text', multi: 'Paragraph', check: 'Tick', list: 'Choice' }[FM.type];
    FM.fields.push({ page: FM.page, type: FM.type, name: `${base} ${n}`, x: x / s, y: y / s, w: w / s, h: h / s, options: 'Yes, No', size: 11, required: false });
    FM.sel = FM.fields.length - 1; drawFormPage(false);
  });
}
async function loadForm(f) {
  busy('Opening…');
  try {
    FM.bytes = new Uint8Array(await f.arrayBuffer()); FM.name = f.name; FM.js?.destroy(); FM.js = await jsLoad(FM.bytes); FM.page = 0; FM.fields = []; FM.sel = -1;
    $('fmEmpty').hidden = true; $('fmWork').hidden = false; await drawFormPage();
  } catch (e) { $('fmStatus').textContent = 'This PDF could not be opened.'; } finally { busy(''); }
}
async function drawFormPage(render = true) {
  if (render) {
    const page = await FM.js.getPage(FM.page + 1); const v1 = page.getViewport({ scale: 1 });
    const avail = Math.max(320, $('fmStage').parentElement.clientWidth - 24); FM.scale = Math.min(1.6, avail / v1.width);
    const { canvas } = await renderPage(FM.js, FM.page, FM.scale * devicePixelRatio);
    const c = $('fmCanvas'); c.width = canvas.width; c.height = canvas.height; c.getContext('2d').drawImage(canvas, 0, 0);
    c.style.width = (v1.width * FM.scale) + 'px'; c.style.height = (v1.height * FM.scale) + 'px'; FM.pw = v1.width; FM.ph = v1.height;
    $('fmLayer').style.width = c.style.width; $('fmLayer').style.height = c.style.height;
    $('fmPg').textContent = `Page ${FM.page + 1} of ${FM.js.numPages}`;
  }
  const s = FM.scale;
  $('fmLayer').innerHTML = FM.fields.map((f, i) => f.page !== FM.page ? '' : `<div class="fbox t-${f.type} ${i === FM.sel ? 'on' : ''}" data-fi="${i}" style="left:${f.x * s}px;top:${f.y * s}px;width:${f.w * s}px;height:${f.h * s}px"><span>${esc(f.name)}</span></div>`).join('');
  $('fmLayer').querySelectorAll('.fbox').forEach((el) => (el.onpointerdown = (e) => {
    e.stopPropagation(); const i = +el.dataset.fi; FM.sel = i; const f = FM.fields[i]; const sx = e.clientX, sy = e.clientY, ox = f.x, oy = f.y;
    const mv = (ev) => { f.x = ox + (ev.clientX - sx) / s; f.y = oy + (ev.clientY - sy) / s; el.style.left = f.x * s + 'px'; el.style.top = f.y * s + 'px'; };
    const up = () => { window.removeEventListener('pointermove', mv); window.removeEventListener('pointerup', up); drawFormPage(false); };
    window.addEventListener('pointermove', mv); window.addEventListener('pointerup', up);
  }));
  $('fmList').innerHTML = FM.fields.length ? `<ol class="fm-list">${FM.fields.map((f, i) => `<li class="${i === FM.sel ? 'on' : ''}" data-pick="${i}">${icon(FTYPES.find((t) => t[0] === f.type)[2], 'i sm')}<span>${esc(f.name)}</span><small>p.${f.page + 1}</small></li>`).join('')}</ol>` : '<p class="muted small">No boxes yet.</p>';
  $('fmList').querySelectorAll('[data-pick]').forEach((li) => (li.onclick = () => { FM.sel = +li.dataset.pick; if (FM.fields[FM.sel].page !== FM.page) { FM.page = FM.fields[FM.sel].page; drawFormPage(); } else drawFormPage(false); }));
  const f = FM.fields[FM.sel];
  $('fmEdit').innerHTML = !f ? '' : `<div class="opts fm-edit"><label class="o-fld"><span>Name</span><input id="feName" value="${esc(f.name)}"></label>
    ${f.type === 'list' ? `<label class="o-fld"><span>Choices (separate with commas)</span><input id="feOpts" value="${esc(f.options)}"></label>` : ''}
    ${f.type !== 'check' ? `<label class="o-fld"><span>Text size (pt)</span><input id="feSize" type="number" min="0" max="48" value="${f.size}"></label>` : ''}
    <label class="o-chk"><input type="checkbox" id="feReq" ${f.required ? 'checked' : ''}><span>Must be filled in</span></label>
    <button id="feDel">${icon('trash')}Remove this box</button></div>`;
  if (f) {
    $('feName').oninput = (e) => { f.name = e.target.value; };
    $('feName').onchange = () => drawFormPage(false);
    if ($('feOpts')) $('feOpts').oninput = (e) => { f.options = e.target.value; };
    if ($('feSize')) $('feSize').oninput = (e) => { f.size = +e.target.value || 0; };
    $('feReq').onchange = (e) => { f.required = e.target.checked; };
    $('feDel').onclick = () => { FM.fields.splice(FM.sel, 1); FM.sel = -1; drawFormPage(false); };
  }
}
async function saveForm() {
  const st = $('fmStatus'); if (!FM.fields.length) { st.textContent = 'Draw at least one box first.'; return; }
  const names = new Set(); for (const f of FM.fields) { const n = f.name.trim(); if (!n) { st.textContent = 'Every box needs a name.'; return; } if (names.has(n)) { st.textContent = `Two boxes are called "${n}". Give each box its own name.`; return; } names.add(n); }
  busy('Saving…');
  try {
    const d = await libLoad(FM.bytes); const form = d.getForm();
    for (const f of FM.fields) {
      const p = d.getPage(f.page); const { H } = vis(p); const u = rectToUser(p, f.x, H - (f.y + f.h), f.x + f.w, H - f.y); const box = { x: u.x, y: u.y, width: u.w, height: u.h, borderWidth: 1, borderColor: rgb(0.55, 0.62, 0.63), backgroundColor: rgb(0.96, 0.98, 1), rotate: degrees(vis(p).r) };
      let fld;
      if (f.type === 'check') { fld = form.createCheckBox(f.name.trim()); fld.addToPage(p, box); }
      else if (f.type === 'list') { fld = form.createDropdown(f.name.trim()); fld.addOptions(f.options.split(',').map((s) => s.trim()).filter(Boolean)); fld.addToPage(p, box); if (f.size) fld.setFontSize(f.size); }
      else { fld = form.createTextField(f.name.trim()); if (f.type === 'multi') fld.enableMultiline(); fld.addToPage(p, box); if (f.size) fld.setFontSize(f.size); }
      if (f.required) fld.enableRequired();
    }
    await saveFile(pdfBlob(await d.save()), `${baseName(FM.name)}_form.pdf`, 'PDF document');
    st.textContent = `Saved with ${FM.fields.length} box${FM.fields.length > 1 ? 'es' : ''}. Open it in any PDF reader to fill it in.`;
  } catch (e) { st.textContent = /already exists/i.test(e.message) ? 'A box name is already used in this PDF. Choose another name.' : e.message; }
  finally { busy(''); }
}

window.dudelyTool = openTool;
if (/[?&]debug/.test(location.search)) { window.__RN = RN; window.__PG = PG; window.__FM = FM; }
window.dudelyHome = drawHome;
window.dudelyShowEdit = () => showView('edit');
drawHome();
if ($('fileChip').classList.contains('none')) showView('home');
