// Dudely Editor PC: Photos (improve and download), Merge and Split.
import * as pdfjsLib from './pdf.min.mjs';
pdfjsLib.GlobalWorkerOptions.workerSrc = './pdf.worker.min.mjs';
const { PDFDocument } = window.PDFLib;
const $ = (id) => document.getElementById(id);
const baseName = (n) => n.replace(/\.[^.]+$/, '');
const isPdf = (f) => f.type === 'application/pdf' || /\.pdf$/i.test(f.name);
const isImg = (f) => f.type.startsWith('image/');

// ------------------------------------------------------------------ tabs
const VIEWS = { home: 'viewHome', tool: 'viewTool', edit: 'viewEdit', photos: 'viewPhotos', merge: 'viewMerge', split: 'viewSplit' };
export function showView(v, title, sub) {
  for (const [k, id] of Object.entries(VIEWS)) $(id).hidden = k !== v;
  document.querySelectorAll('.tab').forEach((b) => b.setAttribute('aria-selected', String(b.dataset.view === v || (v === 'tool' && b.dataset.view === 'home'))));
  $('editActions').hidden = v !== 'edit';
  document.querySelectorAll('.acts[data-for]').forEach((el) => (el.hidden = el.dataset.for !== v));
  const [t0, s0] = VIEW_TEXT[v] || ['', ''];
  $('viewTitle').textContent = title || t0; $('viewSub').textContent = sub || s0;
  document.body.dataset.view = v;
}
const VIEW_TEXT = {
  home: ['All tools', 'Every PDF tool in one place. Your files never leave this computer.'],
  tool: ['', ''],
  edit: ['Edit PDF', 'Change words on the page and keep the original layout. Save as PDF or Word.'],
  photos: ['Photos', 'Make photos clearer and brighter, crop and turn them, then download them or put them in a PDF.'],
  merge: ['Merge', 'Join PDFs and photos into one PDF, in the order you choose.'],
  split: ['Split', 'Cut one PDF into several files: every few pages, or by page ranges you type.'],
};
document.querySelectorAll('.tab').forEach((b) => (b.onclick = () => { showView(b.dataset.view); if (b.dataset.view === 'home') window.dudelyHome?.(); }));

// ------------------------------------------------------------------ saving files
export async function saveFile(blob, name, desc = 'File') {
  const ext = (name.match(/\.[^.]+$/) || ['.bin'])[0];
  if (window.showSaveFilePicker) {
    try {
      const h = await window.showSaveFilePicker({ suggestedName: name, types: [{ description: desc, accept: { [blob.type || 'application/octet-stream']: [ext] } }] });
      const w = await h.createWritable(); await w.write(blob); await w.close();
      return h.name;
    } catch (e) { if (e.name === 'AbortError') return null; }
  }
  const a = document.createElement('a'); a.href = URL.createObjectURL(blob); a.download = name; a.click();
  setTimeout(() => URL.revokeObjectURL(a.href), 5000);
  return name;
}
function status(id, t) { $(id).textContent = t; }
async function openInEditor(bytes, name) {
  showView('edit');
  await window.dudelyOpen([new File([bytes], name, { type: 'application/pdf' })]);
}
async function imageToJpegBytes(src, maxSide = 2400) {
  const bmp = src instanceof ImageBitmap ? src : await createImageBitmap(src);
  const s = Math.min(1, maxSide / Math.max(bmp.width, bmp.height));
  const c = document.createElement('canvas'); c.width = Math.round(bmp.width * s); c.height = Math.round(bmp.height * s);
  c.getContext('2d').drawImage(bmp, 0, 0, c.width, c.height);
  const blob = await new Promise((r) => c.toBlob(r, 'image/jpeg', 0.9));
  return { bytes: new Uint8Array(await blob.arrayBuffer()), w: c.width, h: c.height };
}
async function addImagePage(doc, src) {
  const j = await imageToJpegBytes(src);
  const img = await doc.embedJpg(j.bytes);
  const pw = 595, ph = pw * j.h / j.w;
  doc.addPage([pw, ph]).drawImage(img, { x: 0, y: 0, width: pw, height: ph });
}

// ================================================================== PHOTOS
const PH = { list: [], cur: -1, cropMode: false, drag: null };
const newAdj = () => ({ b: 0, c: 0, s: 0, w: 0, sh: 0 });
$('phAdd').onclick = () => $('phInput').click();
$('phInput').onchange = async (e) => { const f = [...e.target.files].filter(isImg); e.target.value = ''; await addPhotos(f); };
async function addPhotos(files) {
  for (const f of files) {
    try {
      const bmp = await createImageBitmap(f);
      PH.list.push({ name: f.name, bmp, adj: newAdj(), rot: 0, flip: false, crop: null, auto: null, doc: false });
    } catch (e) { status('phStatus', `Could not open ${f.name}.`); }
  }
  if (PH.cur < 0 && PH.list.length) PH.cur = 0;
  drawList(); loadControls(); preview();
}
function drawList() {
  const ul = $('phList'); ul.innerHTML = '';
  PH.list.forEach((p, i) => {
    const b = document.createElement('button'); b.className = 'thumb' + (i === PH.cur ? ' on' : '');
    const c = document.createElement('canvas'); const s = 72 / Math.max(p.bmp.width, p.bmp.height);
    c.width = Math.round(p.bmp.width * s); c.height = Math.round(p.bmp.height * s); c.getContext('2d').drawImage(p.bmp, 0, 0, c.width, c.height);
    const t = document.createElement('span'); t.textContent = p.name;
    b.append(c, t); b.onclick = () => { PH.cur = i; PH.cropMode = false; drawList(); loadControls(); preview(); };
    ul.append(b);
  });
  const has = PH.list.length > 0;
  $('phEmpty').hidden = has; $('phWork').hidden = !has;
}
const cur = () => PH.list[PH.cur];
const SLIDERS = [['phB', 'b'], ['phC', 'c'], ['phS', 's'], ['phW', 'w'], ['phSh', 'sh']];
function loadControls() {
  const p = cur(); if (!p) return;
  for (const [id, k] of SLIDERS) { $(id).value = p.adj[k]; $(id + 'v').textContent = p.adj[k]; }
  $('phDoc').checked = p.doc;
  $('phCropBtn').textContent = PH.cropMode ? 'Apply crop' : 'Crop';
  $('phCropCancel').hidden = !PH.cropMode;
}
for (const [id, k] of SLIDERS) $(id).oninput = () => { const p = cur(); if (!p) return; p.adj[k] = +$(id).value; $(id + 'v').textContent = p.adj[k]; preview(); };
$('phDoc').onchange = () => { const p = cur(); if (p) { p.doc = $('phDoc').checked; if (p.doc && !p.auto) autoLevels(p); preview(); } };
$('phAuto').onclick = () => { const p = cur(); if (!p) return; autoLevels(p); p.adj.s = Math.max(p.adj.s, 15); p.adj.sh = Math.max(p.adj.sh, 35); loadControls(); preview(); };
$('phReset').onclick = () => { const p = cur(); if (!p) return; Object.assign(p, { adj: newAdj(), rot: 0, flip: false, crop: null, auto: null, doc: false }); PH.cropMode = false; loadControls(); preview(); };
$('phRotL').onclick = () => turn(-1); $('phRotR').onclick = () => turn(1);
$('phFlip').onclick = () => { const p = cur(); if (!p) return; p.flip = !p.flip; if (p.crop) p.crop = { x: 1 - p.crop.x - p.crop.w, y: p.crop.y, w: p.crop.w, h: p.crop.h }; preview(); };
function turn(dir) {
  const p = cur(); if (!p) return;
  p.rot = (p.rot + (dir > 0 ? 90 : 270)) % 360;
  if (p.crop) { const r = p.crop; p.crop = dir > 0 ? { x: 1 - r.y - r.h, y: r.x, w: r.h, h: r.w } : { x: r.y, y: 1 - r.x - r.w, w: r.h, h: r.w }; }
  preview();
}
/** Colour balance: stretch each colour so the darkest 1% is black and the brightest 1% is white. */
function autoLevels(p) {
  const s = Math.min(1, 600 / Math.max(p.bmp.width, p.bmp.height));
  const c = document.createElement('canvas'); c.width = Math.max(1, Math.round(p.bmp.width * s)); c.height = Math.max(1, Math.round(p.bmp.height * s));
  const x = c.getContext('2d'); x.drawImage(p.bmp, 0, 0, c.width, c.height);
  const d = x.getImageData(0, 0, c.width, c.height).data;
  const lo = [], hi = [];
  for (let k = 0; k < 3; k++) {
    const h = new Uint32Array(256); for (let i = k; i < d.length; i += 4) h[d[i]]++;
    const n = d.length / 4; let acc = 0, a = 0, b = 255;
    for (let v = 0; v < 256; v++) { acc += h[v]; if (acc > n * 0.01) { a = v; break; } }
    acc = 0; for (let v = 255; v >= 0; v--) { acc += h[v]; if (acc > n * 0.01) { b = v; break; } }
    lo.push(a); hi.push(Math.max(a + 10, b));
  }
  p.auto = { lo, hi };
}
/** Draws a photo with all its changes, at most maxSide pixels on the long side. */
function render(p, maxSide) {
  const rw = p.rot % 180 ? p.bmp.height : p.bmp.width, rh = p.rot % 180 ? p.bmp.width : p.bmp.height;
  const cr = p.crop || { x: 0, y: 0, w: 1, h: 1 };
  const outW0 = rw * cr.w, outH0 = rh * cr.h;
  const s = Math.min(1, maxSide / Math.max(outW0, outH0));
  // turned and mirrored full picture at this scale, then the crop
  const full = document.createElement('canvas'); full.width = Math.max(1, Math.round(rw * s)); full.height = Math.max(1, Math.round(rh * s));
  const f = full.getContext('2d');
  f.translate(full.width / 2, full.height / 2); f.rotate(p.rot * Math.PI / 180); if (p.flip) f.scale(-1, 1);
  const dw = p.bmp.width * s, dh = p.bmp.height * s;
  f.drawImage(p.bmp, -dw / 2, -dh / 2, dw, dh);
  const out = document.createElement('canvas'); out.width = Math.max(1, Math.round(outW0 * s)); out.height = Math.max(1, Math.round(outH0 * s));
  const o = out.getContext('2d', { willReadFrequently: true });
  o.drawImage(full, cr.x * full.width, cr.y * full.height, cr.w * full.width, cr.h * full.height, 0, 0, out.width, out.height);
  const a = p.adj;
  if (!(a.b || a.c || a.s || a.w || a.sh || p.auto || p.doc)) return out;
  const img = o.getImageData(0, 0, out.width, out.height), d = img.data;
  const cf = (259 * (a.c * 2.55 + 255)) / (255 * (259 - a.c * 2.55));
  const lv = p.auto, sat = 1 + a.s / 100, br = a.b * 1.28, wm = a.w * 0.6;
  for (let i = 0; i < d.length; i += 4) {
    let r = d[i], g = d[i + 1], b = d[i + 2];
    if (lv) { r = (r - lv.lo[0]) * 255 / (lv.hi[0] - lv.lo[0]); g = (g - lv.lo[1]) * 255 / (lv.hi[1] - lv.lo[1]); b = (b - lv.lo[2]) * 255 / (lv.hi[2] - lv.lo[2]); }
    if (p.doc) {                                   // clean document: grey, white paper, dark text
      let y = 0.299 * r + 0.587 * g + 0.114 * b;
      y = (y - 40) * 255 / 170; if (y > 215) y = 255;
      r = g = b = y;
    } else {
      const y = 0.299 * r + 0.587 * g + 0.114 * b;
      r = y + (r - y) * sat; g = y + (g - y) * sat; b = y + (b - y) * sat;
      r += wm; b -= wm;
    }
    r = cf * (r + br - 128) + 128; g = cf * (g + br - 128) + 128; b = cf * (b + br - 128) + 128;
    d[i] = r; d[i + 1] = g; d[i + 2] = b;   // Uint8ClampedArray keeps 0..255
  }
  if (a.sh > 0) sharpen(img, a.sh / 50);
  o.putImageData(img, 0, 0);
  return out;
}
/** Unsharp mask with a 3x3 blur: makes edges and text crisper. */
function sharpen(img, amount) {
  const { width: w, height: h, data: d } = img, src = new Uint8ClampedArray(d);
  for (let y = 1; y < h - 1; y++) for (let x = 1; x < w - 1; x++) {
    const i = (y * w + x) * 4;
    for (let k = 0; k < 3; k++) {
      const j = i + k;
      const blur = (src[j - 4 * w - 4] + src[j - 4 * w] + src[j - 4 * w + 4] + src[j - 4] + src[j] + src[j + 4] + src[j + 4 * w - 4] + src[j + 4 * w] + src[j + 4 * w + 4]) / 9;
      d[j] = src[j] + amount * (src[j] - blur);
    }
  }
}
let raf = 0;
function preview() {
  cancelAnimationFrame(raf);
  raf = requestAnimationFrame(() => {
    const p = cur(), c = $('phCanvas'); if (!p) return;
    const box = $('phStage').getBoundingClientRect();
    const maxSide = Math.min(1400, Math.max(box.width, box.height) * (window.devicePixelRatio || 1));
    const showP = PH.cropMode ? { ...p, crop: null } : p;      // while cropping, show the whole picture
    const out = render(showP, maxSide);
    c.width = out.width; c.height = out.height; c.getContext('2d').drawImage(out, 0, 0);
    const fit = Math.min((box.width - 20) / out.width, (box.height - 20) / out.height, 1 / (window.devicePixelRatio || 1) * 2);
    c.style.width = Math.round(out.width * fit) + 'px'; c.style.height = Math.round(out.height * fit) + 'px';
    drawCropBox();
    const rw = p.rot % 180 ? p.bmp.height : p.bmp.width, rh = p.rot % 180 ? p.bmp.width : p.bmp.height;
    const cr = p.crop || { w: 1, h: 1 };
    $('phInfo').textContent = `${p.name} · ${Math.round(rw * cr.w)} × ${Math.round(rh * cr.h)} px`;
  });
}
// crop: drag a box on the picture
function drawCropBox() {
  const box = $('phCropBox'), c = $('phCanvas'), p = cur();
  if (!PH.cropMode || !p) { box.hidden = true; return; }
  const r = PH.drag || p.crop || { x: 0.05, y: 0.05, w: 0.9, h: 0.9 };
  PH.pending = r;
  const cr = c.getBoundingClientRect(), st = $('phStage').getBoundingClientRect();
  box.hidden = false;
  Object.assign(box.style, { left: (cr.left - st.left + r.x * cr.width) + 'px', top: (cr.top - st.top + r.y * cr.height) + 'px', width: (r.w * cr.width) + 'px', height: (r.h * cr.height) + 'px' });
}
$('phCropBtn').onclick = () => {
  const p = cur(); if (!p) return;
  if (!PH.cropMode) { PH.cropMode = true; PH.drag = null; status('phStatus', 'Drag a box over the part to keep, then click Apply crop.'); }
  else { if (PH.pending && PH.pending.w > 0.02 && PH.pending.h > 0.02) p.crop = PH.pending; PH.cropMode = false; PH.drag = null; status('phStatus', 'Cropped.'); }
  loadControls(); preview();
};
$('phCropCancel').onclick = () => { PH.cropMode = false; PH.drag = null; loadControls(); preview(); };
$('phCanvas').addEventListener('pointerdown', (e) => {
  if (!PH.cropMode) return;
  const c = $('phCanvas'), r = c.getBoundingClientRect(); c.setPointerCapture(e.pointerId);
  const x0 = (e.clientX - r.left) / r.width, y0 = (e.clientY - r.top) / r.height;
  const move = (ev) => {
    const x1 = Math.min(1, Math.max(0, (ev.clientX - r.left) / r.width)), y1 = Math.min(1, Math.max(0, (ev.clientY - r.top) / r.height));
    PH.drag = { x: Math.min(x0, x1), y: Math.min(y0, y1), w: Math.abs(x1 - x0), h: Math.abs(y1 - y0) }; drawCropBox();
  };
  c.addEventListener('pointermove', move);
  c.addEventListener('pointerup', () => c.removeEventListener('pointermove', move), { once: true });
});
window.addEventListener('resize', () => { if (!$('viewPhotos').hidden) preview(); });
async function photoBlob(p) {
  const out = render(p, 5000);
  const type = $('phFormat').value;
  return await new Promise((r) => out.toBlob(r, type, 0.92));
}
const extFor = () => ($('phFormat').value === 'image/png' ? '.png' : '.jpg');
$('phDownload').onclick = async () => {
  const p = cur(); if (!p) return;
  status('phStatus', 'Preparing…');
  const saved = await saveFile(await photoBlob(p), baseName(p.name) + '_edited' + extFor(), 'Picture');
  status('phStatus', saved ? `Saved ${saved}.` : 'Not saved.');
};
$('phDownloadAll').onclick = async () => {
  if (!PH.list.length) return;
  status('phStatus', 'Preparing all pictures…');
  const zip = new JSZip();
  for (const p of PH.list) zip.file(baseName(p.name) + '_edited' + extFor(), await photoBlob(p));
  const saved = await saveFile(await zip.generateAsync({ type: 'blob', mimeType: 'application/zip' }), 'Dudely_photos.zip', 'Zip file');
  status('phStatus', saved ? `Saved ${saved} (${PH.list.length} pictures).` : 'Not saved.');
};
$('phToPdf').onclick = async () => {
  if (!PH.list.length) return;
  status('phStatus', 'Making a PDF…');
  const doc = await PDFDocument.create();
  for (const p of PH.list) await addImagePage(doc, render(p, 2400));
  await openInEditor(await doc.save(), 'Photos.pdf');
};
$('phRemove').onclick = () => {
  if (PH.cur < 0) return;
  PH.list.splice(PH.cur, 1); PH.cur = Math.min(PH.cur, PH.list.length - 1); PH.cropMode = false;
  drawList(); loadControls(); preview();
};
for (const id of ['phStage', 'phEmpty']) {
  $(id).addEventListener('dragover', (e) => e.preventDefault());
  $(id).addEventListener('drop', (e) => { e.preventDefault(); e.stopPropagation(); addPhotos([...e.dataTransfer.files].filter(isImg)); });
}

// ================================================================== MERGE
const MG = [];
$('mgAdd').onclick = () => $('mgInput').click();
$('mgInput').onchange = async (e) => { const f = [...e.target.files]; e.target.value = ''; await addMerge(f); };
async function addMerge(files) {
  for (const f of files) {
    if (isPdf(f)) {
      try {
        const bytes = new Uint8Array(await f.arrayBuffer());
        const d = await PDFDocument.load(bytes, { ignoreEncryption: true });
        MG.push({ name: f.name, kind: 'pdf', bytes, pages: d.getPageCount() });
      } catch (e) { status('mgStatus', `Could not open ${f.name} (it may be protected with a password).`); }
    } else if (isImg(f)) MG.push({ name: f.name, kind: 'img', file: f, pages: 1 });
  }
  drawMerge();
}
function drawMerge() {
  const tb = $('mgRows'); tb.innerHTML = '';
  MG.forEach((it, i) => {
    const tr = document.createElement('tr');
    tr.innerHTML = `<td>${i + 1}</td><td class="nm"></td><td>${it.kind === 'pdf' ? 'PDF' : 'Photo'}</td><td>${it.pages}</td><td class="act"></td>`;
    tr.querySelector('.nm').textContent = it.name;
    const act = tr.querySelector('.act');
    const btn = (t, lbl, fn, dis) => { const b = document.createElement('button'); b.textContent = t; b.setAttribute('aria-label', lbl); b.disabled = dis; b.onclick = fn; act.append(b); };
    btn('↑', 'Move up', () => { [MG[i - 1], MG[i]] = [MG[i], MG[i - 1]]; drawMerge(); }, i === 0);
    btn('↓', 'Move down', () => { [MG[i + 1], MG[i]] = [MG[i], MG[i + 1]]; drawMerge(); }, i === MG.length - 1);
    btn('✕', 'Remove', () => { MG.splice(i, 1); drawMerge(); }, false);
    tb.append(tr);
  });
  const total = MG.reduce((s, x) => s + x.pages, 0);
  $('mgTotal').textContent = MG.length ? `${MG.length} file(s), ${total} page(s) in total` : 'No files yet. Add PDFs or photos; they are joined in the order shown.';
  $('mgGo').disabled = $('mgOpen').disabled = MG.length < 1;
}
async function buildMerged() {
  const out = await PDFDocument.create();
  for (const it of MG) {
    if (it.kind === 'pdf') {
      const src = await PDFDocument.load(it.bytes, { ignoreEncryption: true });
      const pages = await out.copyPages(src, src.getPageIndices());
      pages.forEach((p) => out.addPage(p));
    } else await addImagePage(out, it.file);
  }
  return await out.save();
}
$('mgGo').onclick = async () => {
  status('mgStatus', 'Joining…');
  try {
    const name = ($('mgName').value.trim() || 'merged').replace(/\.pdf$/i, '') + '.pdf';
    const saved = await saveFile(new Blob([await buildMerged()], { type: 'application/pdf' }), name, 'PDF');
    status('mgStatus', saved ? `Saved ${saved}.` : 'Not saved.');
  } catch (e) { status('mgStatus', 'Could not join the files: ' + (e.message || e)); }
};
$('mgOpen').onclick = async () => { status('mgStatus', 'Joining…'); await openInEditor(await buildMerged(), ($('mgName').value.trim() || 'merged') + '.pdf'); status('mgStatus', ''); };
$('mgDrop').addEventListener('dragover', (e) => e.preventDefault());
$('mgDrop').addEventListener('drop', (e) => { e.preventDefault(); e.stopPropagation(); addMerge([...e.dataTransfer.files]); });

// ================================================================== SPLIT
const SP = { bytes: null, name: '', n: 0 };
$('spPick').onclick = () => $('spInput').click();
$('spInput').onchange = async (e) => { const f = e.target.files[0]; e.target.value = ''; if (f) await loadSplit(f); };
$('spDrop').addEventListener('dragover', (e) => e.preventDefault());
$('spDrop').addEventListener('drop', (e) => { e.preventDefault(); e.stopPropagation(); const f = [...e.dataTransfer.files].find(isPdf); if (f) loadSplit(f); });
async function loadSplit(f) {
  status('spStatus', 'Opening…');
  try {
    SP.bytes = new Uint8Array(await f.arrayBuffer()); SP.name = baseName(f.name);
    const pdf = await pdfjsLib.getDocument({ data: SP.bytes.slice() }).promise;
    SP.n = pdf.numPages;
    $('spFile').textContent = `${f.name} · ${SP.n} page(s)`;
    $('spWork').hidden = false; status('spStatus', '');
    planSplit();
    const grid = $('spThumbs'); grid.innerHTML = '';
    for (let i = 1; i <= Math.min(SP.n, 300); i++) {
      const fig = document.createElement('figure'); const c = document.createElement('canvas'); const cap = document.createElement('figcaption');
      cap.textContent = String(i); fig.append(c, cap); grid.append(fig);
      const pg = await pdf.getPage(i); const vp0 = pg.getViewport({ scale: 1 }); const vp = pg.getViewport({ scale: 110 / vp0.width });
      c.width = Math.round(vp.width); c.height = Math.round(vp.height);
      await pg.render({ canvasContext: c.getContext('2d'), viewport: vp }).promise;
    }
    if (SP.n > 300) { const m = document.createElement('p'); m.textContent = `Showing the first 300 of ${SP.n} pages.`; grid.append(m); }
  } catch (e) { status('spStatus', 'Could not open this PDF: ' + (e.message || e)); }
}
/** "1-3, 5, 8-" → [[1,3],[5,5],[8,N]] or an error message. */
function parseRanges(text, n) {
  const parts = text.split(/[,;]+/).map((s) => s.trim()).filter(Boolean);
  if (!parts.length) return 'Type page ranges, for example 1-3, 5, 8-10.';
  const out = [];
  for (const p of parts) {
    const m = p.match(/^(\d+)\s*(?:[-–]\s*(\d*))?$/);
    if (!m) return `"${p}" is not a page or range. Use numbers like 4 or 2-6.`;
    const a = +m[1], b = m[2] === undefined ? a : (m[2] === '' ? n : +m[2]);
    if (a < 1 || b > n || a > b) return `"${p}" is outside pages 1 to ${n}.`;
    out.push([a, b]);
  }
  return out;
}
function currentPlan() {
  if (!SP.n) return [];
  if ($('spModeEvery').checked) {
    const k = Math.max(1, Math.floor(+$('spEvery').value || 1)), out = [];
    for (let a = 1; a <= SP.n; a += k) out.push([a, Math.min(SP.n, a + k - 1)]);
    return out;
  }
  return parseRanges($('spRanges').value, SP.n);
}
function planSplit() {
  $('spEveryBox').classList.toggle('dim', !$('spModeEvery').checked);
  $('spRangeBox').classList.toggle('dim', $('spModeEvery').checked);
  const plan = currentPlan(), ul = $('spPlan'); ul.innerHTML = '';
  if (typeof plan === 'string') { ul.innerHTML = ''; status('spPlanMsg', plan); $('spGo').disabled = true; return; }
  status('spPlanMsg', plan.length ? `${plan.length} file(s) will be made:` : '');
  plan.slice(0, 60).forEach(([a, b], i) => { const li = document.createElement('li'); li.textContent = `File ${i + 1}: page${a === b ? ` ${a}` : `s ${a}–${b}`}`; ul.append(li); });
  if (plan.length > 60) { const li = document.createElement('li'); li.textContent = `… and ${plan.length - 60} more`; ul.append(li); }
  $('spGo').disabled = !plan.length;
}
for (const id of ['spModeEvery', 'spModeRanges', 'spEvery', 'spRanges']) { $(id).oninput = planSplit; $(id).onchange = planSplit; }
$('spGo').onclick = async () => {
  const plan = currentPlan(); if (typeof plan === 'string' || !plan.length) return;
  status('spStatus', 'Splitting…');
  try {
    const src = await PDFDocument.load(SP.bytes, { ignoreEncryption: true });
    const files = [];
    for (const [a, b] of plan) {
      const d = await PDFDocument.create();
      const idx = []; for (let i = a; i <= b; i++) idx.push(i - 1);
      (await d.copyPages(src, idx)).forEach((p) => d.addPage(p));
      files.push([`${SP.name}_p${a === b ? a : `${a}-${b}`}.pdf`, await d.save()]);
    }
    let saved;
    if (files.length === 1) saved = await saveFile(new Blob([files[0][1]], { type: 'application/pdf' }), files[0][0], 'PDF');
    else {
      const zip = new JSZip(); files.forEach(([n, b]) => zip.file(n, b));
      saved = await saveFile(await zip.generateAsync({ type: 'blob', mimeType: 'application/zip' }), `${SP.name}_split.zip`, 'Zip file');
    }
    status('spStatus', saved ? `Saved ${saved}${files.length > 1 ? ` with ${files.length} PDF files inside` : ''}.` : 'Not saved.');
  } catch (e) { status('spStatus', 'Could not split: ' + (e.message || e)); }
};
