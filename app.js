// Dudely Editor PC. Edits are laid on top of the original pages; nothing else in the PDF changes.
import * as pdfjsLib from './pdf.min.mjs';
pdfjsLib.GlobalWorkerOptions.workerSrc = './pdf.worker.min.mjs';
const { PDFDocument, StandardFonts, rgb, degrees, LineCapStyle } = window.PDFLib;

const TEXT_SIZES = [0.018, 0.025, 0.035];     // fraction of the page width
const PEN = [0.0025, 0.004, 0.007];
const STAMP = [0.025, 0.035, 0.05];
const PAD = 0.18;                              // cover a little beyond a changed line (fraction of its height)
const STRETCH = 1.10;                          // longer text may run 10% past the old line, then gets smaller

const $ = (id) => document.getElementById(id);
const S = { bytes: null, name: '', pdf: null, pages: [], tool: 'edit', color: '#000000', size: 1, zoom: 1, history: [], dirty: false, ocr: null, editing: null };

// ------------------------------------------------------------------ helpers
const hint = (t) => { $('hint').textContent = t; };
const busy = (t) => { const b = $('busy'); b.textContent = t || ''; b.style.display = t ? 'block' : 'none'; };
const hexToRgb = (h) => [parseInt(h.slice(1, 3), 16), parseInt(h.slice(3, 5), 16), parseInt(h.slice(5, 7), 16)];
const css = (c) => Array.isArray(c) ? `rgb(${c[0]},${c[1]},${c[2]})` : c;
const toPdfColor = (c) => { const a = Array.isArray(c) ? c : hexToRgb(c); return rgb(a[0] / 255, a[1] / 255, a[2] / 255); };
function tailsOf(s) { if (/[gjpqyQ;()\[\]{}|\/_@$]/.test(s)) return 2; if (s.includes(',')) return 1; return 0; }
const isNumber = (s) => /^[\d.,\s%$:\/-]+$/.test(s.trim());
// Fonts: sans (Arial / Helvetica), serif (Times New Roman), mono (Courier). Old edits only had "serif".
const FAMILY_CSS = { sans: 'Helvetica, Arial, sans-serif', serif: '"Times New Roman", Times, serif', mono: '"Courier New", Courier, monospace' };
const FAMILY_WORD = { sans: 'Arial', serif: 'Times New Roman', mono: 'Courier New' };
const famOf = (m) => m.family || (m.serif ? 'serif' : 'sans');
function fontCss(m, px) { return `${m.italic ? 'italic ' : ''}${m.bold ? 'bold ' : ''}${px}px ${FAMILY_CSS[famOf(m)]}`; }
const toHex = (c) => '#' + (Array.isArray(c) ? c : hexToRgb(c)).map((v) => Math.max(0, Math.min(255, Math.round(v))).toString(16).padStart(2, '0')).join('');
/** All standard PDF fonts, picked by family, bold and italic. */
async function embedAllFonts(doc) {
  const SF = StandardFonts, F = {};
  const list = { sans: [SF.Helvetica, SF.HelveticaBold, SF.HelveticaOblique, SF.HelveticaBoldOblique],
    serif: [SF.TimesRoman, SF.TimesRomanBold, SF.TimesRomanItalic, SF.TimesRomanBoldItalic],
    mono: [SF.Courier, SF.CourierBold, SF.CourierOblique, SF.CourierBoldOblique] };
  for (const [fam, arr] of Object.entries(list)) F[fam] = await Promise.all(arr.map((n) => doc.embedFont(n)));
  F.pick = (m) => F[famOf(m)][(m.bold ? 1 : 0) + (m.italic ? 2 : 0)];
  F.sans0 = F.sans[0];
  return F;
}

const TOOL_HINTS = {
  edit: 'Edit text: click a line in a blue box and type. Enter applies, Esc cancels. The layout stays as it is.',
  add: 'Add text: click where the text should start, type, press Enter. Click your text to change it; drag it to move it.',
  pen: 'Pen / sign: write or sign with the mouse, pen or finger.',
  white: 'White-out: drag a box over what you want to cover.',
  highlight: 'Highlight: drag a box over the words.',
  tick: 'Tick: click where the tick goes.',
  cross: 'Cross: click where the cross goes.',
};

// ------------------------------------------------------------------ opening
$('btnOpen').onclick = $('btnOpen2').onclick = () => { if (confirmLeave()) $('fileInput').click(); };
$('fileInput').onchange = (e) => { const f = [...e.target.files]; e.target.value = ''; if (f.length) openFiles(f); };
window.addEventListener('dragover', (e) => { e.preventDefault(); if (!$('viewEdit').hidden) document.body.classList.add('drop'); });
window.addEventListener('dragleave', () => document.body.classList.remove('drop'));
window.addEventListener('drop', (e) => {
  e.preventDefault(); document.body.classList.remove('drop');
  if ($('viewEdit').hidden) return;
  const f = [...e.dataTransfer.files].filter((x) => x.type === 'application/pdf' || x.type.startsWith('image/') || /\.pdf$/i.test(x.name));
  if (f.length && confirmLeave()) openFiles(f);
});
function confirmLeave() { return !S.dirty || confirm('Your edits are not saved yet. Continue anyway?'); }
window.addEventListener('beforeunload', (e) => { if (S.dirty) { e.preventDefault(); e.returnValue = ''; } });

async function openFiles(files) {
  if ($('viewEdit').hidden) window.dudelyShowEdit?.();
  busy('Opening…');
  try {
    const first = files[0];
    let bytes;
    if (/\.docx$/i.test(first.name) || first.type === 'application/vnd.openxmlformats-officedocument.wordprocessingml.document') {
      const buf = await first.arrayBuffer();
      const made = await wordToPdf(buf);
      if (!made) { busy(''); await printOtherWord(buf, first.name); return; }
      bytes = made;
    } else if (first.type === 'application/pdf' || /\.pdf$/i.test(first.name)) bytes = new Uint8Array(await first.arrayBuffer());
    else bytes = await imagesToPdf(files.filter((f) => f.type.startsWith('image/')));
    S.bytes = bytes;
    S.name = first.name.replace(/\.[^.]+$/, '');
    S.pdf = await pdfjsLib.getDocument({ data: bytes.slice() }).promise;
    S.history = []; S.dirty = false; S.pages = [];
    $('fileName').textContent = first.name + (files.length > 1 ? ` (+${files.length - 1})` : '');
    $('empty').style.display = 'none';
    $('btnSave').disabled = false;
    const holder = $('pages'); holder.innerHTML = '';
    for (let i = 0; i < S.pdf.numPages; i++) {
      const page = await S.pdf.getPage(i + 1);
      const el = document.createElement('div'); el.className = 'page';
      const num = document.createElement('div'); num.className = 'num'; num.textContent = `Page ${i + 1} of ${S.pdf.numPages}`;
      const canvas = document.createElement('canvas'); const overlay = document.createElement('canvas');
      el.append(num, canvas, overlay); holder.append(el);
      const P = { i, page, base: page.getViewport({ scale: 1 }), el, canvas, overlay, marks: [], lines: null, hover: null, task: null };
      S.pages.push(P);
      wirePointer(P);
    }
    await layoutAll();
    setTool(S.tool);
    S.pages.forEach((P) => { P.linesReady = findLines(P); });
    $('btnWord').disabled = false;
  } catch (err) {
    console.error(err);
    hint('Could not open this file: ' + (err && err.message ? err.message : err) + (String(err).includes('assword') ? ' (it is protected with a password)' : ''));
  } finally { busy(''); }
}

async function imagesToPdf(files) {
  const doc = await PDFDocument.create();
  for (const f of files) {
    const bmp = await createImageBitmap(f);
    const s = Math.min(1, 2400 / Math.max(bmp.width, bmp.height));
    const c = document.createElement('canvas'); c.width = Math.round(bmp.width * s); c.height = Math.round(bmp.height * s);
    c.getContext('2d').drawImage(bmp, 0, 0, c.width, c.height);
    const blob = await new Promise((r) => c.toBlob(r, 'image/jpeg', 0.9));
    const img = await doc.embedJpg(new Uint8Array(await blob.arrayBuffer()));
    const pw = 595, ph = pw * c.height / c.width;
    doc.addPage([pw, ph]).drawImage(img, { x: 0, y: 0, width: pw, height: ph });
  }
  return await doc.save();
}

// ------------------------------------------------------------------ showing pages
function fitWidth() { return Math.min($('viewer').clientWidth - 48, 1000); }
async function layoutAll() {
  const fw = fitWidth();
  for (const P of S.pages) {
    P.w = Math.round(fw * S.zoom); P.h = Math.round(P.w * P.base.height / P.base.width);
    P.el.style.width = P.w + 'px'; P.el.style.height = P.h + 'px';
  }
  for (const P of S.pages) await renderPage(P);
  $('zoomVal').textContent = Math.round(S.zoom * 100) + '%';
}
async function renderPage(P) {
  const dpr = window.devicePixelRatio || 1;
  const vp = P.page.getViewport({ scale: (P.w / P.base.width) * dpr });
  for (const c of [P.canvas, P.overlay]) {
    c.width = Math.round(vp.width); c.height = Math.round(vp.height);
    c.style.width = P.w + 'px'; c.style.height = P.h + 'px';
  }
  if (P.task) { try { P.task.cancel(); } catch (e) { /* ignore */ } }
  P.task = P.page.render({ canvasContext: P.canvas.getContext('2d'), viewport: vp });
  try { await P.task.promise; } catch (e) { /* cancelled by a newer render */ }
  draw(P);
}
$('btnZoomIn').onclick = () => { S.zoom = Math.min(3, +(S.zoom + 0.25).toFixed(2)); closeEditor(); layoutAll(); };
$('btnZoomOut').onclick = () => { S.zoom = Math.max(0.5, +(S.zoom - 0.25).toFixed(2)); closeEditor(); layoutAll(); };
let rT; window.addEventListener('resize', () => { clearTimeout(rT); rT = setTimeout(() => { if (S.pages.length) layoutAll(); }, 250); });

// ------------------------------------------------------------------ styled pieces of text ("runs")
// A run is { t: text, family, bold, italic, color: '#rrggbb', k: size as a multiple of the line's own size }.
const STYLE_KEYS = ['family', 'bold', 'italic', 'color', 'k'];
const styleOf = (r) => ({ family: r.family || 'sans', bold: !!r.bold, italic: !!r.italic, color: (r.color || '#000000').toLowerCase(), k: r.k || 1 });
const sameStyle = (a, b) => STYLE_KEYS.every((k) => a[k] === b[k]);
function mergeRuns(runs) {
  const out = [];
  for (const r of runs) {
    if (!r.t) continue;
    const st = styleOf(r), last = out[out.length - 1];
    if (last && sameStyle(last, st)) last.t += r.t; else out.push({ ...st, t: r.t });
  }
  return out;
}
const runsKey = (runs) => JSON.stringify(mergeRuns(runs).map((r) => [r.t, ...STYLE_KEYS.map((k) => r[k])]));
/** The runs of an edit; older edits have one style for the whole line. */
function runsOf(m) {
  if (m.runs && m.runs.length) return m.runs.map((r) => ({ ...r }));
  return m.text ? [{ t: m.text, family: famOf(m), bold: !!m.bold, italic: !!m.italic, color: toHex(m.type === 'replace' ? m.fg : m.color), k: 1 }] : [];
}
/** The line as it is on the page: mixed bold / italic / fonts inside one line are kept. */
function lineRuns(L, look) {
  if (!L.segs) return L.text ? [{ ...styleOf(look), t: L.text }] : [];
  return mergeRuns(L.segs.map((g) => ({ t: g.s, family: g.family, bold: g.bold, italic: g.italic, color: look.color, k: g.k })));
}
function cleanSegs(L) {
  const segs = L.segs.map((g) => ({ ...g, k: Math.abs(g.h / L.h - 1) < 0.08 ? 1 : Math.round(g.h / L.h * 100) / 100 }));
  if (!segs.length) return null;
  segs[0].s = segs[0].s.replace(/^\s+/, ''); segs[segs.length - 1].s = segs[segs.length - 1].s.replace(/\s+$/, '');
  const keep = segs.filter((g) => g.s);
  if (keep.map((g) => g.s).join('') !== L.text.trim()) return null;
  const mixed = keep.some((g) => g.family !== keep[0].family || g.bold !== keep[0].bold || g.italic !== keep[0].italic || g.k !== keep[0].k);
  return mixed ? keep : null;
}

// ------------------------------------------------------------------ finding the page's own text
async function findLines(P) {
  try {
    const tc = await P.page.getTextContent();
    const W = P.base.width, H = P.base.height, items = [];
    let chars = 0;
    for (const it of tc.items) {
      if (!it.str || !it.str.trim()) continue;
      const tx = pdfjsLib.Util.transform(P.base.transform, it.transform);
      const h = Math.hypot(tx[2], tx[3]);
      if (h < 2 || Math.abs(tx[1]) > 0.05 * h) continue;          // skip turned text
      const st = tc.styles[it.fontName] || {};
      let fo = null; try { if (P.page.commonObjs.has(it.fontName)) fo = P.page.commonObjs.get(it.fontName); } catch (e) { /* not loaded */ }
      const fname = (fo && fo.name) || '';
      const bold = !!(fo && (fo.bold || fo.black)) || /bold|black|heavy|semibold|demi/i.test(fname);
      const serif = /times|georgia|garamond|cambria|serif/i.test(fname) && !/sans/i.test(fname) || (/serif/i.test(st.fontFamily || '') && !/sans/i.test(st.fontFamily || ''));
      const mono = /courier|mono|consol/i.test(fname) || /monospace/i.test(st.fontFamily || '');
      const italic = !!(fo && fo.italic) || /italic|oblique/i.test(fname);
      const fontName = fname.replace(/^[A-Z]{6}\+/, '').replace(/[-,].*$/, '');
      items.push({ x: tx[4], y: tx[5], w: it.width * P.base.scale, h, s: it.str, serif, bold, italic, family: mono ? 'mono' : serif ? 'serif' : 'sans', fontName });
      chars += it.str.trim().length;
    }
    if (chars >= 12) {
      items.sort((a, b) => (Math.abs(a.y - b.y) < 0.3 * Math.min(a.h, b.h) ? a.x - b.x : a.y - b.y));
      const lines = [];
      for (const it of items) {
        const L = lines[lines.length - 1];
        if (L && Math.abs(L.y - it.y) < 0.3 * Math.min(L.h, it.h) && it.x - L.x1 < 0.9 * L.h && it.x >= L.x0) {
          const gap = it.x - L.x1;
          const piece = (gap > 0.15 * L.h && !L.text.endsWith(' ') && !it.s.startsWith(' ') ? ' ' : '') + it.s;
          L.text += piece; L.segs.push({ s: piece, family: it.family, bold: it.bold, italic: it.italic, h: it.h });
          L.x1 = Math.max(L.x1, it.x + it.w);
        } else lines.push({ x0: it.x, x1: it.x + it.w, y: it.y, h: it.h, text: it.s, serif: it.serif, bold: it.bold, italic: it.italic, family: it.family, fontName: it.fontName,
          segs: [{ s: it.s, family: it.family, bold: it.bold, italic: it.italic, h: it.h }] });
      }
      P.lines = lines.filter((L) => L.text.trim()).map((L) => ({
        l: L.x0 / W, r: L.x1 / W, t: (L.y - 0.80 * L.h) / H, b: (L.y + 0.22 * L.h) / H,
        text: L.text.trim(), size: L.h / H, base: L.y / H, serif: L.serif, bold: L.bold, italic: L.italic, family: L.family, fontName: L.fontName, src: 'pdf',
        segs: cleanSegs(L),
      }));
      // Some PDFs store characters (often brackets or symbols) as private codes that cannot be read back.
      // For those lines only, read the words from the page picture instead.
      for (const L of P.lines) if (/[\uE000-\uF8FF\uFFFD]/.test(L.text)) { L.text = await readLineFromPicture(P, L) || L.text.replace(/[\uE000-\uF8FF\uFFFD]/g, ''); L.segs = null; }
    } else {
      P.lines = [];
      await ocrPage(P);
    }
  } catch (e) { console.warn(e); P.lines = P.lines || []; }
  draw(P);
  if (S.tool === 'edit') setTool('edit');
}

async function ocrPage(P) {
  busy(`Reading the text on page ${P.i + 1} (scan)…`);
  try {
    await ocrWorker();
    const scale = Math.min(2.5, 2400 / P.base.width);
    const vp = P.page.getViewport({ scale });
    const c = document.createElement('canvas'); c.width = Math.round(vp.width); c.height = Math.round(vp.height);
    await P.page.render({ canvasContext: c.getContext('2d'), viewport: vp }).promise;
    const { data } = await S.ocr.recognize(c);
    const out = [];
    const push = (text, bb) => {
      text = text.trim(); if (!text) return;
      const t = bb.y0 / c.height, b = bb.y1 / c.height, hN = b - t;
      const tails = tailsOf(text);
      const sizeOfBox = tails === 0 ? 1.33 : tails === 1 ? 1.18 : 1.0;
      const up = tails === 0 ? 0 : tails === 1 ? 0.12 : 0.20;
      out.push({ l: bb.x0 / c.width, r: bb.x1 / c.width, t, b, text, size: hN * sizeOfBox, base: b - up * hN, serif: false, src: 'scan' });
    };
    for (const ln of (data.lines || [])) {
      if (!(ln.text || '').trim() || (ln.confidence !== undefined && ln.confidence < 30)) continue;
      const words = (ln.words || []).filter((w) => (w.text || '').trim());
      if (!words.length) { push(ln.text, ln.bbox); continue; }
      // a wide gap between words means a new cell (tables): each cell can be changed on its own
      const lh = ln.bbox.y1 - ln.bbox.y0;
      let seg = null;
      for (const w of words) {
        if (seg && w.bbox.x0 - seg.bb.x1 < 1.1 * lh) {
          seg.text += ' ' + w.text;
          seg.bb = { x0: seg.bb.x0, y0: Math.min(seg.bb.y0, w.bbox.y0), x1: w.bbox.x1, y1: Math.max(seg.bb.y1, w.bbox.y1) };
        } else { if (seg) push(seg.text, seg.bb); seg = { text: w.text, bb: { ...w.bbox } }; }
      }
      if (seg) push(seg.text, seg.bb);
    }
    P.lines = out;
  } catch (e) { console.warn(e); hint('Could not read the text on this scan. Use Add text and White-out instead.'); P.lines = []; }
  finally { busy(''); }
}

async function ocrWorker() {
  if (!S.ocr) S.ocr = await Tesseract.createWorker('eng', 1, {
    workerPath: new URL('./worker.min.js', location.href).href,
    corePath: new URL('./', location.href).href,
    langPath: new URL('.', location.href).href.replace(/\/$/, ''),
    gzip: true,
  });
  return S.ocr;
}
/** Reads one line of text from the page picture (used when the PDF's own text cannot be decoded). */
async function readLineFromPicture(P, L) {
  try {
    const sc = 4;
    const vp = P.page.getViewport({ scale: sc });
    const W = P.base.width, H = P.base.height, pad = (L.b - L.t) * 0.3;
    const x = Math.max(0, (L.l - 0.01) * W * sc), y = Math.max(0, (L.t - pad) * H * sc);
    const w = Math.min(vp.width - x, (L.r - L.l + 0.02) * W * sc), h = Math.min(vp.height - y, (L.b - L.t + 2 * pad) * H * sc);
    const c = document.createElement('canvas'); c.width = Math.ceil(w); c.height = Math.ceil(h);
    const ctx = c.getContext('2d'); ctx.translate(-x, -y);
    await P.page.render({ canvasContext: ctx, viewport: vp }).promise;
    const { data } = await (await ocrWorker()).recognize(c, { tessedit_pageseg_mode: '7' });
    return (data.text || '').replace(/\s+/g, ' ').trim();
  } catch (e) { return ''; }
}

/** Background and text colour around a line, read from the page picture; ink = share of the line covered (bold covers more). */
function sampleColors(P, L) { return sampleCanvas(P.canvas, L); }
function sampleCanvas(c, L) {
  const ctx = c.getContext('2d', { willReadFrequently: true });
  const l = Math.max(0, Math.floor(L.l * c.width)), r = Math.min(c.width, Math.ceil(L.r * c.width));
  const t = Math.max(0, Math.floor(L.t * c.height)), b = Math.min(c.height, Math.ceil(L.b * c.height));
  const pad = Math.max(2, Math.round((b - t) * 0.25));
  const X0 = Math.max(0, l - pad), Y0 = Math.max(0, t - pad), X1 = Math.min(c.width, r + pad), Y1 = Math.min(c.height, b + pad);
  if (X1 - X0 < 2 || Y1 - Y0 < 2) return { fg: [0, 0, 0], bg: [255, 255, 255], ink: 0 };
  const d = ctx.getImageData(X0, Y0, X1 - X0, Y1 - Y0).data, w = X1 - X0;
  const px = (x, y) => { const k = ((y - Y0) * w + (x - X0)) * 4; return [d[k], d[k + 1], d[k + 2]]; };
  const ring = [];
  for (let x = X0; x < X1; x += 2) { ring.push(px(x, Y0)); ring.push(px(x, Y1 - 1)); }
  for (let y = Y0; y < Y1; y += 2) { ring.push(px(X0, y)); ring.push(px(X1 - 1, y)); }
  const med = (k) => ring.map((p) => p[k]).sort((a, b2) => a - b2)[ring.length >> 1];
  const bg = [med(0), med(1), med(2)];
  const dist = (p) => Math.abs(p[0] - bg[0]) + Math.abs(p[1] - bg[1]) + Math.abs(p[2] - bg[2]);
  const inside = [];
  for (let y = t; y < b; y++) for (let x = l; x < r; x++) inside.push(px(x, y));
  if (!inside.length) return { fg: [0, 0, 0], bg, ink: 0 };
  const far = inside.slice().sort((a, b2) => dist(b2) - dist(a)).slice(0, Math.max(1, inside.length / 12 | 0));
  const fg = [0, 1, 2].map((k) => Math.round(far.reduce((s, p) => s + p[k], 0) / far.length));
  const md = Math.max(1, dist(fg));
  const ink = inside.filter((p) => dist(p) > md / 2).length / inside.length;
  return { fg, bg, ink };
}

/** On scans, measures the real height of the letters in a line, so new words get the same size and baseline. */
function refineScanLine(P, L, col) {
  if (L.refined) return; L.refined = true;
  const c = P.canvas, ctx = c.getContext('2d', { willReadFrequently: true });
  const hpx = (L.b - L.t) * c.height;
  const l = Math.max(0, Math.floor(L.l * c.width)), r = Math.min(c.width, Math.ceil(L.r * c.width));
  const t = Math.max(0, Math.floor(L.t * c.height - 0.3 * hpx)), b = Math.min(c.height, Math.ceil(L.b * c.height + 0.3 * hpx));
  if (r - l < 2 || b - t < 2) return;
  const d = ctx.getImageData(l, t, r - l, b - t).data, w = r - l;
  const bg = col.bg, md = Math.max(30, Math.abs(col.fg[0] - bg[0]) + Math.abs(col.fg[1] - bg[1]) + Math.abs(col.fg[2] - bg[2]));
  const rows = [];
  for (let y = 0; y < b - t; y++) {
    let n = 0;
    for (let x = 0; x < w; x++) { const k = (y * w + x) * 4; if (Math.abs(d[k] - bg[0]) + Math.abs(d[k + 1] - bg[1]) + Math.abs(d[k + 2] - bg[2]) > md / 2) n++; }
    rows.push(n > 0);
  }
  // the run of inked rows around the middle of the line (stops at the gap before the next line)
  const mid = Math.round((L.t + L.b) / 2 * c.height) - t;
  let s0 = -1;
  for (let k = 0; k <= Math.round(0.3 * hpx) + 2 && s0 < 0; k++) { if (rows[mid + k]) s0 = mid + k; else if (rows[mid - k]) s0 = mid - k; }
  if (s0 < 0) return;
  let top = s0, bot = s0 + 1;
  while (top > 0 && (rows[top - 1] || (top > 1 && rows[top - 2] && hpx > 14))) top--;
  while (bot < rows.length && (rows[bot] || (bot + 1 < rows.length && rows[bot + 1] && hpx > 14))) bot++;
  const inkT = (t + top) / c.height, inkB = (t + bot) / c.height, inkH = inkB - inkT;
  const tails = tailsOf(L.text);
  if (tails === 2) { L.size = inkH / 0.93; L.base = inkB - 0.21 * L.size; }
  else if (tails === 1) { L.size = inkH / 0.82; L.base = inkB - 0.10 * L.size; }
  else { L.size = inkH / 0.72; L.base = inkB; }
  L.t = Math.min(L.t, inkT); L.b = Math.max(L.b, inkB);
}

// ------------------------------------------------------------------ drawing the edits
function draw(P) {
  const c = P.overlay, ctx = c.getContext('2d');
  const W = c.width, H = c.height;
  ctx.clearRect(0, 0, W, H);
  for (const m of P.marks) drawMark(ctx, m, W, H);
  if (P.current) drawMark(ctx, P.current, W, H);
  if (S.tool === 'edit' && P.lines) {
    for (const L of P.lines) {
      if (P.marks.some((m) => m.type === 'replace' && m.line === L)) continue;
      ctx.strokeStyle = L === P.hover ? 'rgba(37,99,235,0.95)' : 'rgba(37,99,235,0.35)';
      ctx.lineWidth = L === P.hover ? 2.5 : 1.5;
      ctx.strokeRect(L.l * W - 2, L.t * H - 2, (L.r - L.l) * W + 4, (L.b - L.t) * H + 4);
    }
  }
}
function drawMark(ctx, m, W, H) {
  ctx.save();
  if (m.type === 'replace') {
    const bh = (m.b - m.t) * H, pad = PAD * bh;
    ctx.fillStyle = css(m.bg);
    ctx.fillRect(m.l * W - pad, m.t * H - pad, (m.r - m.l) * W + 2 * pad, bh + 2 * pad);
    const runs = runsOf(m);
    let px = m.size * H;
    const width = () => runs.reduce((a, r) => { ctx.font = fontCss(r, px * r.k); return a + ctx.measureText(r.t).width; }, 0);
    let tw = width(); const room = (m.r - m.l) * W * STRETCH;
    if (!m.fixedSize && tw > room && tw > 0) { px *= room / tw; tw = room; }
    ctx.textBaseline = 'alphabetic';
    let x = m.alignRight ? m.r * W - tw : m.l * W;
    for (const r of runs) { ctx.font = fontCss(r, px * r.k); ctx.fillStyle = r.color; ctx.fillText(r.t, x, m.base * H); x += ctx.measureText(r.t).width; }
  } else if (m.type === 'box') {
    const x = Math.min(m.l, m.r) * W, y = Math.min(m.t, m.b) * H, w = Math.abs(m.r - m.l) * W, h = Math.abs(m.b - m.t) * H;
    ctx.fillStyle = m.highlight ? 'rgba(255,225,0,0.38)' : '#fff';
    ctx.fillRect(x, y, w, h);
    if (!m.highlight) { ctx.setLineDash([6, 4]); ctx.strokeStyle = 'rgba(0,0,0,.35)'; ctx.strokeRect(x, y, w, h); }
  } else if (m.type === 'ink') {
    ctx.strokeStyle = m.color; ctx.lineWidth = Math.max(1.5, m.width * W); ctx.lineCap = 'round'; ctx.lineJoin = 'round';
    for (const s of m.strokes) {
      ctx.beginPath(); ctx.moveTo(s[0][0] * W, s[0][1] * H);
      if (s.length === 1) ctx.lineTo(s[0][0] * W + 0.5, s[0][1] * H);
      for (let k = 1; k < s.length; k++) ctx.lineTo(s[k][0] * W, s[k][1] * H);
      ctx.stroke();
    }
  } else if (m.type === 'text') {
    const px = m.size * W;
    if (m.runs) {
      const top = Math.max(...m.runs.map((r) => r.k)); let x = m.x * W;
      for (const r of m.runs) { ctx.font = fontCss(r, px * r.k); ctx.fillStyle = r.color; ctx.fillText(r.t, x, m.y * H + 0.8 * px * top); x += ctx.measureText(r.t).width; }
    } else {
      ctx.font = fontCss(m, px); ctx.fillStyle = css(m.color);
      m.text.split('\n').forEach((line, i) => ctx.fillText(line, m.x * W, m.y * H + (0.8 + i * 1.2) * px));
    }
    if (m === S.dragging) { const bb = textBounds(m, W, H); ctx.setLineDash([5, 4]); ctx.strokeStyle = '#2563EB'; ctx.strokeRect(bb[0], bb[1], bb[2] - bb[0], bb[3] - bb[1]); }
  }
  ctx.restore();
}
function textBounds(m, W, H) {
  const ctx = document.createElement('canvas').getContext('2d'); const px = m.size * W;
  if (m.runs) {
    const top = Math.max(...m.runs.map((r) => r.k));
    const tw = m.runs.reduce((a, r) => { ctx.font = fontCss(r, px * r.k); return a + ctx.measureText(r.t).width; }, 0);
    return [m.x * W - 6, m.y * H - 6, m.x * W + tw + 6, m.y * H + 1.2 * px * top + 6];
  }
  ctx.font = fontCss(m, px);
  const lines = m.text.split('\n'); const tw = Math.max(...lines.map((l) => ctx.measureText(l).width));
  return [m.x * W - 6, m.y * H - 6, m.x * W + tw + 6, m.y * H + lines.length * 1.2 * px + 6];
}

// ------------------------------------------------------------------ mouse / pen
function wirePointer(P) {
  const o = P.overlay;
  const norm = (e) => { const r = o.getBoundingClientRect(); return [(e.clientX - r.left) / r.width, (e.clientY - r.top) / r.height]; };
  let down = null, moved = false, start = null;
  o.addEventListener('pointermove', (e) => {
    const [x, y] = norm(e);
    if (!down) {
      if (S.tool === 'edit' && P.lines) {
        const h = lineAt(P, x, y); if (h !== P.hover) { P.hover = h; o.style.cursor = h || replaceAt(P, x, y) ? 'text' : 'default'; draw(P); }
      } else o.style.cursor = S.tool === 'add' ? 'text' : 'crosshair';
      return;
    }
    if (Math.abs(e.clientX - down[0]) + Math.abs(e.clientY - down[1]) > 4) moved = true;
    const cur = P.current;
    if (S.tool === 'pen' && cur) cur.strokes[0].push([x, y]);
    else if ((S.tool === 'white' || S.tool === 'highlight') && cur) { cur.r = x; cur.b = y; }
    else if (S.tool === 'add' && S.dragging && moved) { S.dragging.x = start[0] + (x - start[2]); S.dragging.y = start[1] + (y - start[3]); }
    draw(P);
  });
  o.addEventListener('pointerdown', (e) => {
    if (!S.pages.length) return;
    if (S.editing) { applyEditor(); }
    o.setPointerCapture(e.pointerId);
    const [x, y] = norm(e); down = [e.clientX, e.clientY]; moved = false;
    if (S.tool === 'pen') P.current = { type: 'ink', strokes: [[[x, y]]], color: S.color, width: PEN[S.size] };
    else if (S.tool === 'white' || S.tool === 'highlight') P.current = { type: 'box', l: x, t: y, r: x, b: y, highlight: S.tool === 'highlight' };
    else if (S.tool === 'add') { S.dragging = textAt(P, x, y); if (S.dragging) start = [S.dragging.x, S.dragging.y, x, y]; }
  });
  o.addEventListener('pointerup', (e) => {
    if (!down) return;
    const [x, y] = norm(e); down = null;
    const W = P.overlay.width, H = P.overlay.height;
    if (S.tool === 'pen' && P.current) addMark(P, P.current);
    else if ((S.tool === 'white' || S.tool === 'highlight') && P.current) {
      const m = P.current; if (Math.abs(m.r - m.l) * W > 10 && Math.abs(m.b - m.t) * H > 6) addMark(P, m);
    } else if (S.tool === 'tick' || S.tool === 'cross') {
      const s = STAMP[S.size], ar = W / H;
      const strokes = S.tool === 'tick'
        ? [[[x - s * 0.5, y], [x - s * 0.15, y + s * 0.4 * ar], [x + s * 0.6, y - s * 0.55 * ar]]]
        : [[[x - s * 0.45, y - s * 0.45 * ar], [x + s * 0.45, y + s * 0.45 * ar]], [[x + s * 0.45, y - s * 0.45 * ar], [x - s * 0.45, y + s * 0.45 * ar]]];
      addMark(P, { type: 'ink', strokes, color: S.color, width: PEN[S.size] * 1.6 });
    } else if (S.tool === 'add') {
      const d = S.dragging; S.dragging = null;
      if (d && moved) S.dirty = true;
      else if (d) openEditor(P, { mode: 'text', mark: d });
      else if (!moved) openEditor(P, { mode: 'text', x, y });
    } else if (S.tool === 'edit' && !moved) {
      const r = replaceAt(P, x, y);
      if (r) openEditor(P, { mode: 'replace', mark: r });
      else { const L = lineAt(P, x, y); if (L) openEditor(P, { mode: 'replace', line: L }); }
    }
    P.current = null; draw(P);
  });
}
const lineAt = (P, x, y) => (P.lines || []).filter((L) => x >= L.l - 0.005 && x <= L.r + 0.005 && y >= L.t - 0.004 && y <= L.b + 0.004)
  .sort((a, b) => (a.b - a.t) - (b.b - b.t))[0] || null;
const replaceAt = (P, x, y) => [...P.marks].reverse().find((m) => m.type === 'replace' && x >= m.l - 0.005 && x <= m.r + 0.005 && y >= m.t - 0.004 && y <= m.b + 0.004) || null;
const textAt = (P, x, y) => { const W = P.overlay.width, H = P.overlay.height; return [...P.marks].reverse().find((m) => { if (m.type !== 'text') return false; const b = textBounds(m, W, H); return x * W >= b[0] && x * W <= b[2] && y * H >= b[1] && y * H <= b[3]; }) || null; };
function addMark(P, m) { P.marks.push(m); S.history.push([P, m]); S.dirty = true; draw(P); }

// ------------------------------------------------------------------ the inline editor (like typing in Word)
// The box holds styled pieces. Select some words and change the font, size, colour, bold or italic: only those words change.
// With nothing selected, the change applies to the whole line. Unless changed, everything keeps the original look.
const ed = $('editor'), edInput = $('edInput');
function openEditor(P, o) {
  closeEditor();
  S.editing = { P, ...o, sel: null };
  const E = S.editing;
  P.el.appendChild(ed);
  const W = P.w, H = P.h, Wpt = P.base.width, Hpt = P.base.height;
  let runs, l, t, w, look;
  if (o.mode === 'replace') {
    const src = o.mark || o.line;
    if (o.mark) {
      E.basePt = o.mark.size * Hpt; runs = runsOf(o.mark);
      look = runs[0] ? styleOf(runs[0]) : styleOf({ family: famOf(o.mark), bold: o.mark.bold, italic: o.mark.italic, color: toHex(o.mark.fg) });
      E.orig = o.mark.orig || runs;
    } else {
      const L = o.line, col = sampleColors(P, L); E.col = col;
      if (L.src === 'scan') refineScanLine(P, L, col);
      look = styleOf({ family: L.family || (L.serif ? 'serif' : 'sans'), bold: L.bold !== undefined ? !!L.bold : col.ink > 0.31, italic: !!L.italic, color: toHex(col.fg) });
      E.basePt = L.size * Hpt; runs = lineRuns(L, look); E.orig = runs;
    }
    l = src.l * W; t = src.t * H; w = Math.max((src.r - src.l) * W * 1.25 + 30, 140);
    $('edOrig').hidden = !o.mark; $('edDel').hidden = true;
    E.label = o.line && o.line.fontName ? `Original font: ${o.line.fontName}` : (o.line && o.line.src === 'scan' ? 'Original look read from the scan' : 'Original look');
  } else {
    const m = o.mark;
    E.basePt = (m ? m.size : TEXT_SIZES[S.size]) * Wpt;
    runs = m ? runsOf(m) : [];
    look = runs[0] ? styleOf(runs[0]) : styleOf({ family: m ? famOf(m) : 'sans', bold: m && m.bold, italic: m && m.italic, color: toHex(m ? m.color : S.color) });
    E.orig = m ? (m.orig || runs) : [];
    l = (m ? m.x : o.x) * W; t = (m ? m.y : o.y) * H; w = 220;
    $('edOrig').hidden = true; $('edDel').hidden = !m;
    E.label = 'New text';
  }
  E.runs0 = mergeRuns(runs); E.look = look; E.defStyle = { ...look };
  E.scale = W / Wpt;                       // screen pixels per point
  ed.style.display = 'block';
  ed.style.left = Math.max(0, l - 4) + 'px'; ed.style.top = Math.max(0, t - 4) + 'px';
  edInput.style.minWidth = Math.max(60, Math.min(w, W - l + 4)) + 'px';
  renderRuns(runs);
  showStyle(E.defStyle); showScope();
  hint('Select some words to change only them (font, size, colour, B, I). With nothing selected, changes apply to the whole line. Enter applies, Esc cancels.');
  setTimeout(() => { edInput.focus(); setSel(0, textLen()); }, 0);
}
function closeEditor() {
  const was = !!S.editing;
  ed.style.display = 'none'; S.editing = null;
  if (was && S.pages.length && TOOL_HINTS[S.tool]) hint(TOOL_HINTS[S.tool]);
}
/** Puts styled pieces into the box, each showing its own font, size and colour. */
function renderRuns(runs) {
  const E = S.editing; if (!E) return;
  const px = E.basePt * E.scale;
  edInput.textContent = '';
  let top = 1;
  for (const r of mergeRuns(runs)) {
    const sp = document.createElement('span');
    sp.dataset.st = JSON.stringify(styleOf(r));
    sp.textContent = r.t;
    sp.style.font = fontCss(r, Math.max(8, px * r.k)); sp.style.color = r.color;
    edInput.appendChild(sp); top = Math.max(top, r.k);
  }
  edInput.style.font = fontCss(E.defStyle, Math.max(8, px * E.defStyle.k));
  edInput.style.color = E.defStyle.color;
  edInput.style.minHeight = Math.max(18, px * top * 1.3) + 'px';
}
function textNodes() { const w = document.createTreeWalker(edInput, NodeFilter.SHOW_TEXT); const a = []; while (w.nextNode()) a.push(w.currentNode); return a; }
const textLen = () => textNodes().reduce((a, n) => a + n.data.length, 0);
function offsetOf(node, off) { const r = document.createRange(); r.setStart(edInput, 0); r.setEnd(node, off); return r.toString().length; }
function setSel(a, b) {
  const nodes = textNodes(); if (!nodes.length) return;
  const at = (o) => { for (const n of nodes) { if (o <= n.data.length) return [n, o]; o -= n.data.length; } const n = nodes[nodes.length - 1]; return [n, n.data.length]; };
  const [n1, o1] = at(a), [n2, o2] = at(b);
  const r = document.createRange(); r.setStart(n1, o1); r.setEnd(n2, o2);
  const sel = getSelection(); sel.removeAllRanges(); sel.addRange(r);
}
/** Reads the box back as styled pieces. Text typed where no style is known takes the style at the cursor. */
function readRuns() {
  const E = S.editing, out = []; let last = E.defStyle;
  for (const n of textNodes()) {
    if (!n.data) continue;
    const sp = n.parentElement && n.parentElement.closest('[data-st]');
    const st = sp && edInput.contains(sp) ? JSON.parse(sp.dataset.st) : last;
    last = st; out.push({ ...st, t: n.data.replace(/ /g, ' ').replace(/[\r\n]+/g, ' ') });
  }
  return mergeRuns(out);
}
function styleAt(runs, pos) {
  let p = 0;
  for (const r of runs) { if (pos < p + r.t.length) return styleOf(r); p += r.t.length; }
  return runs.length ? styleOf(runs[runs.length - 1]) : null;
}
function showStyle(st) {
  const E = S.editing;
  $('edFont').value = st.family; $('edBold').checked = st.bold; $('edItalic').checked = st.italic;
  $('edColor').value = st.color; $('edSize').value = Math.round(E.basePt * st.k * 2) / 2;
}
function showScope() { const E = S.editing; if (E) $('edOrigFont').textContent = E.sel ? 'Changing the selected words' : E.label; }
document.addEventListener('selectionchange', () => {
  const E = S.editing; if (!E) return;
  const sel = getSelection(); if (!sel.rangeCount) return;
  const r = sel.getRangeAt(0);
  if (!edInput.contains(r.startContainer) || !edInput.contains(r.endContainer)) return;   // the person is using a control
  const a = offsetOf(r.startContainer, r.startOffset), b = offsetOf(r.endContainer, r.endOffset);
  E.sel = b > a ? [a, b] : null;
  const st = styleAt(readRuns(), b > a ? a : Math.max(0, a - 1));
  if (st) { E.defStyle = st; showStyle(st); }
  showScope();
});
/** Changes one property (font, size, colour, bold, italic) of the selected words, or of the whole line. */
function applyStyle(prop) {
  const E = S.editing; if (!E) return;
  let v;
  if (prop === 'family') v = $('edFont').value;
  else if (prop === 'bold') v = $('edBold').checked;
  else if (prop === 'italic') v = $('edItalic').checked;
  else if (prop === 'color') v = $('edColor').value.toLowerCase();
  else { const pt = Math.min(200, Math.max(3, parseFloat($('edSize').value) || E.basePt)); v = Math.round(pt / E.basePt * 1000) / 1000; prop = 'k'; }
  const runs = readRuns(), total = runs.reduce((a, r) => a + r.t.length, 0);
  const [a, b] = E.sel || [0, total];
  const out = []; let p = 0;
  for (const r of runs) {
    const s0 = p, s1 = p + r.t.length; p = s1;
    const cut = [s0, Math.max(s0, Math.min(s1, a)), Math.max(s0, Math.min(s1, b)), s1];
    for (let i = 0; i < 3; i++) {
      if (cut[i + 1] <= cut[i]) continue;
      const piece = { ...r, t: r.t.slice(cut[i] - s0, cut[i + 1] - s0) };
      if (i === 1) piece[prop] = v;
      out.push(piece);
    }
  }
  if (!E.sel) E.defStyle = { ...E.defStyle, [prop]: v };
  else E.defStyle = { ...styleAt(out, a) };
  renderRuns(out);
  edInput.focus();
  if (E.sel) setSel(a, b); else setSel(total, total);
}
function applyEditor() {
  const E = S.editing; if (!E) return;
  if (document.activeElement === $('edSize')) applyStyle('size');
  const P = E.P, Wpt = P.base.width, Hpt = P.base.height;
  let runs = readRuns();
  while (runs.length && !runs[runs.length - 1].t.trim()) runs.pop();
  if (runs.length) runs[runs.length - 1].t = runs[runs.length - 1].t.replace(/\s+$/, '');
  runs = mergeRuns(runs);
  const txt = runs.map((r) => r.t).join(''), first = runs[0] || E.defStyle;
  const style = { family: first.family, serif: first.family === 'serif', bold: first.bold, italic: first.italic };
  const fixed = runs.some((r) => r.k !== 1);
  if (E.mode === 'replace') {
    if (E.mark) {
      Object.assign(E.mark, style, { text: txt, runs, fg: hexToRgb(first.color), fixedSize: fixed || !!E.mark.fixedSize });
      S.dirty = true;
    } else if (runsKey(runs) !== runsKey(E.runs0)) {
      const L = E.line;
      addMark(P, { type: 'replace', line: L, l: L.l, t: L.t, r: L.r, b: L.b, text: txt, runs, orig: E.runs0, fg: hexToRgb(first.color), bg: E.col.bg,
        ...style, alignRight: isNumber(L.text), size: L.size, fixedSize: fixed, base: L.base });
    }
  } else {
    const fields = { text: txt, runs, color: first.color, size: E.basePt / Wpt, ...style };
    if (E.mark) { if (txt) { Object.assign(E.mark, fields); S.dirty = true; } else removeMark(P, E.mark); }
    else if (txt) addMark(P, { type: 'text', x: E.x, y: E.y, orig: [], ...fields });
  }
  closeEditor(); draw(P);
}
function cancelEditor() { const P = S.editing && S.editing.P; closeEditor(); if (P) draw(P); }
function removeMark(P, m) { P.marks = P.marks.filter((x) => x !== m); S.history = S.history.filter((h) => h[1] !== m); S.dirty = true; draw(P); }
edInput.addEventListener('keydown', (e) => {
  if (e.key === 'Enter') { e.preventDefault(); applyEditor(); }
  else if (e.key === 'Escape') { e.preventDefault(); cancelEditor(); }
  else if ((e.ctrlKey || e.metaKey) && /^[biu]$/i.test(e.key)) {          // Ctrl+B / Ctrl+I like in Word
    e.preventDefault();
    if (e.key.toLowerCase() === 'b') { $('edBold').checked = !$('edBold').checked; applyStyle('bold'); }
    if (e.key.toLowerCase() === 'i') { $('edItalic').checked = !$('edItalic').checked; applyStyle('italic'); }
  }
});
edInput.addEventListener('paste', (e) => {
  e.preventDefault();
  document.execCommand('insertText', false, (e.clipboardData.getData('text/plain') || '').replace(/[\r\n\t]+/g, ' '));
});
edInput.addEventListener('drop', (e) => e.preventDefault());
// the browser sometimes wraps typed letters in its own tags: turn them back into our styled pieces
edInput.addEventListener('input', () => {
  const stray = [...edInput.childNodes].some((n) => !(n.nodeType === 1 && n.tagName === 'SPAN' && n.dataset.st && n.children.length === 0));
  if (!stray) return;
  const sel = getSelection(); let caret = textLen();
  if (sel.rangeCount && edInput.contains(sel.getRangeAt(0).endContainer)) caret = offsetOf(sel.getRangeAt(0).endContainer, sel.getRangeAt(0).endOffset);
  renderRuns(readRuns()); setSel(caret, caret);
});
$('edOk').onclick = applyEditor;
$('edCancel').onclick = cancelEditor;
$('edOrig').onclick = () => { const E = S.editing; if (E && E.mark) { removeMark(E.P, E.mark); } closeEditor(); };
$('edDel').onclick = () => { const E = S.editing; if (E && E.mark) { removeMark(E.P, E.mark); } closeEditor(); };
ed.addEventListener('pointerdown', (e) => e.stopPropagation());
// keep the selected words selected while clicking B, I and the buttons
ed.querySelectorAll('.tg, button').forEach((el) => el.addEventListener('mousedown', (e) => e.preventDefault()));
$('edFont').addEventListener('change', () => applyStyle('family'));
for (const [id, prop] of [['edBold', 'bold'], ['edItalic', 'italic']]) {
  $(id).closest('label').addEventListener('click', (e) => { e.preventDefault(); $(id).checked = !$(id).checked; applyStyle(prop); });
}
$('edColor').addEventListener('change', () => applyStyle('color'));
$('edSize').addEventListener('change', () => applyStyle('size'));
for (const id of ['edFont', 'edBold', 'edItalic', 'edColor', 'edSize']) $(id).addEventListener('keydown', (e) => {
  if (e.key === 'Enter') { e.preventDefault(); if (id === 'edSize') applyStyle('size'); applyEditor(); }
  else if (e.key === 'Escape') { e.preventDefault(); cancelEditor(); }
});
/** Back to the original font, size and colour (the words you typed stay). */
$('edKeep').onclick = () => {
  const E = S.editing; if (!E) return;
  const now = readRuns(), txt = now.map((r) => r.t).join(''), orig = mergeRuns(E.orig || []);
  const restored = orig.map((r) => r.t).join('') === txt ? orig : (txt ? [{ ...styleOf(E.look), t: txt }] : []);
  E.defStyle = { ...styleOf(E.look) }; E.sel = null;
  renderRuns(restored); showStyle(E.defStyle); showScope();
  edInput.focus(); setSel(textLen(), textLen());
};

// ------------------------------------------------------------------ toolbar
function setTool(t) {
  S.tool = t; closeEditor();
  document.querySelectorAll('.tool').forEach((b) => b.setAttribute('aria-pressed', String(b.dataset.tool === t)));
  if (!S.pages.length) { hint('Open a PDF or a photo to start. You can also drag a file onto this window.'); return; }
  let msg = TOOL_HINTS[t];
  if (t === 'edit') {
    const waiting = S.pages.some((P) => P.lines === null);
    const none = S.pages.every((P) => P.lines && P.lines.length === 0);
    if (waiting) msg = 'Edit text: finding the text on the pages…';
    else if (none) msg = 'Edit text: no text found. Use Add text and White-out instead.';
  }
  hint(msg);
  S.pages.forEach(draw);
}
document.querySelectorAll('.tool').forEach((b) => (b.onclick = () => setTool(b.dataset.tool)));
document.querySelectorAll('.swatch').forEach((b) => (b.onclick = () => {
  S.color = b.dataset.color; document.querySelectorAll('.swatch').forEach((x) => x.setAttribute('aria-pressed', String(x === b)));
}));
document.querySelectorAll('.size').forEach((b) => (b.onclick = () => {
  S.size = +b.dataset.size; document.querySelectorAll('.size').forEach((x) => x.setAttribute('aria-pressed', String(x === b)));
}));
function undo() {
  const last = S.history.pop(); if (!last) { hint('Nothing to undo.'); return; }
  const [P, m] = last; P.marks = P.marks.filter((x) => x !== m); draw(P);
  S.dirty = S.pages.some((p) => p.marks.length);
}
$('btnUndo').onclick = undo;
window.addEventListener('keydown', (e) => {
  if (S.editing) return;
  if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'z') { e.preventDefault(); undo(); }
  if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 's') { e.preventDefault(); save(); }
});

// ------------------------------------------------------------------ saving
$('btnSave').onclick = () => save();
function printable(s, font) {
  let out = '';
  for (let ch of s.normalize('NFKC').replace(/[‘’]/g, "'").replace(/[“”]/g, '"').replace(/[–—]/g, '-').replace(/\t/g, ' ')) {
    try { font.encodeText(ch); out += ch; } catch (e) { out += '?'; }
  }
  return out;
}
async function save() {
  if (!S.bytes) return;
  if (S.editing) applyEditor();
  if (!S.pages.some((P) => P.marks.length)) { hint('No edits yet.'); return; }
  busy('Saving…');
  try {
    const doc = await PDFDocument.load(S.bytes, { ignoreEncryption: true });
    const F = await embedAllFonts(doc);
    for (const P of S.pages) if (P.marks.length) writeMarks(doc.getPage(P.i), P.marks, F);
    const out = await doc.save();
    const name = `${S.name}_edited.pdf`;
    const blob = new Blob([out], { type: 'application/pdf' });
    if (window.showSaveFilePicker) {
      try {
        const h = await window.showSaveFilePicker({ suggestedName: name, types: [{ description: 'PDF', accept: { 'application/pdf': ['.pdf'] } }] });
        const w = await h.createWritable(); await w.write(blob); await w.close();
        S.dirty = false; hint(`Saved as ${h.name}. The original file is unchanged.`); return;
      } catch (e) { if (e.name === 'AbortError') { hint('Not saved.'); return; } }
    }
    const a = document.createElement('a'); a.href = URL.createObjectURL(blob); a.download = name; a.click();
    setTimeout(() => URL.revokeObjectURL(a.href), 4000);
    S.dirty = false; hint(`Saved as ${name} in your Downloads folder. The original file is unchanged.`);
  } catch (e) { console.error(e); hint('Could not save: ' + (e.message || e)); }
  finally { busy(''); }
}

/** Writes the edits on one page in the page's own coordinates (handles turned pages). */
function writeMarks(page, marks, F) {
  const rot = ((page.getRotation().angle % 360) + 360) % 360;
  const cb = page.getCropBox();
  const dispW = rot % 180 === 0 ? cb.width : cb.height, dispH = rot % 180 === 0 ? cb.height : cb.width;
  const pt = (nx, ny) => {
    const u = nx * dispW, v = ny * dispH, x0 = cb.x, y0 = cb.y;
    if (rot === 90) return [x0 + v, y0 + u];
    if (rot === 180) return [x0 + cb.width - u, y0 + v];
    if (rot === 270) return [x0 + cb.width - v, y0 + cb.height - u];
    return [x0 + u, y0 + cb.height - v];
  };
  const rect = (l, t, r, b, opts) => {
    const a = pt(l, t), c = pt(r, b);
    page.drawRectangle({ x: Math.min(a[0], c[0]), y: Math.min(a[1], c[1]), width: Math.abs(a[0] - c[0]), height: Math.abs(a[1] - c[1]), borderWidth: 0, ...opts });
  };
  for (const m of marks) {
    if (m.type === 'replace') {
      const hN = m.b - m.t, padY = PAD * hN, padX = PAD * hN * dispH / dispW;
      rect(m.l - padX, m.t - padY, m.r + padX, m.b + padY, { color: toPdfColor(m.bg) });
      const runs = runsOf(m).map((r) => { const font = F.pick(r); return { ...r, font, safe: printable(r.t, font) }; }).filter((r) => r.safe);
      if (runs.length) {
        let size = m.size * dispH;
        let tw = runs.reduce((a, r) => a + r.font.widthOfTextAtSize(r.safe, size * r.k), 0);
        const room = (m.r - m.l) * dispW * STRETCH;
        if (!m.fixedSize && tw > room && tw > 0) { size *= room / tw; tw = room; }
        let xN = m.alignRight ? m.r - tw / dispW : m.l;
        for (const r of runs) {
          const p = pt(xN, m.base);
          page.drawText(r.safe, { x: p[0], y: p[1], size: size * r.k, font: r.font, color: toPdfColor(r.color), rotate: degrees(rot) });
          xN += r.font.widthOfTextAtSize(r.safe, size * r.k) / dispW;
        }
      }
    } else if (m.type === 'box') {
      rect(m.l, m.t, m.r, m.b, m.highlight ? { color: rgb(1, 0.88, 0), opacity: 0.35 } : { color: rgb(1, 1, 1) });
    } else if (m.type === 'ink') {
      const thickness = Math.max(0.4, m.width * dispW);
      for (const s of m.strokes) {
        const pts = s.length === 1 ? [s[0], [s[0][0] + 0.0002, s[0][1]]] : s;
        for (let k = 1; k < pts.length; k++) {
          const a = pt(pts[k - 1][0], pts[k - 1][1]), b = pt(pts[k][0], pts[k][1]);
          page.drawLine({ start: { x: a[0], y: a[1] }, end: { x: b[0], y: b[1] }, thickness, color: toPdfColor(m.color), lineCap: LineCapStyle.Round });
        }
      }
    } else if (m.type === 'text') {
      const size = m.size * dispW;
      if (m.runs) {
        const top = Math.max(...m.runs.map((r) => r.k)); let xN = m.x;
        for (const r of m.runs) {
          const font = F.pick(r), safe = printable(r.t, font); if (!safe) continue;
          const p = pt(xN, m.y + 0.8 * top * m.size * dispW / dispH);
          page.drawText(safe, { x: p[0], y: p[1], size: size * r.k, font, color: toPdfColor(r.color), rotate: degrees(rot) });
          xN += font.widthOfTextAtSize(safe, size * r.k) / dispW;
        }
        continue;
      }
      m.text.split('\n').forEach((line, i) => {
        const font = F.pick(m);
        const safe = printable(line, font); if (!safe) return;
        const p = pt(m.x, m.y + (0.8 + i * 1.2) * m.size * dispW / dispH);
        page.drawText(safe, { x: p[0], y: p[1], size, font, color: toPdfColor(m.color), rotate: degrees(rot) });
      });
    }
  }
}

// ------------------------------------------------------------------ install on this PC
let installEvent = null;
window.addEventListener('beforeinstallprompt', (e) => { e.preventDefault(); installEvent = e; $('btnInstall').hidden = false; });
$('btnInstall').onclick = async () => { if (!installEvent) return; installEvent.prompt(); await installEvent.userChoice; installEvent = null; $('btnInstall').hidden = true; };
if ('serviceWorker' in navigator && (location.protocol === 'https:' || location.hostname === 'localhost' || location.hostname === '127.0.0.1')) {
  navigator.serviceWorker.register('sw.js').catch(() => {});
}
// Open files passed by the system ("Open with" when installed)
if ('launchQueue' in window) {
  window.launchQueue.setConsumer(async (params) => {
    if (!params.files || !params.files.length) return;
    const files = await Promise.all(params.files.map((h) => h.getFile()));
    openFiles(files);
  });
}
setTool('edit');

// ------------------------------------------------------------------ Word: save as Word, and Word back to PDF
// A page becomes: the page picture behind the text (images, lines, colours, logos stay exactly as they are,
// with the old words covered) + every line of text as an editable Word text frame in the same place, size and colour.
const TW = 20;                                  // twips per point
const xmlEsc = (t) => [...String(t)].filter((ch) => { const c = ch.codePointAt(0); return c === 9 || c === 10 || (c >= 32 && c !== 0xFFFE && c !== 0xFFFF); })
  .join('').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
const hex6 = (c) => (Array.isArray(c) ? c : hexToRgb(c)).map((v) => v.toString(16).padStart(2, '0')).join('').toUpperCase();
const LINE_H = 1.2;                             // frame line height, in font sizes
const BASE_AT = 0.8;                            // baseline position inside a line, as a share of the line height

$('btnWord').onclick = () => saveWord();

async function saveWord() {
  if (!S.pdf) return;
  if (S.editing) applyEditor();
  busy('Making the Word file…');
  try {
    for (const P of S.pages) if (P.linesReady) await P.linesReady;
    const zip = new JSZip();
    const put = (n, d) => zip.file(n, d, { createFolders: false });
    const media = [], body = [], layout = { app: 'Dudely Editor PC', version: 1, pages: [] };
    const mctx = document.createElement('canvas').getContext('2d');
    let pid = 1;
    for (const P of S.pages) {
      const Wpt = P.base.width, Hpt = P.base.height;
      // 1) the page picture at 150 dpi, with the user's drawings, and the words covered in their own background colour
      const sc = 150 / 72;
      const vp = P.page.getViewport({ scale: sc });
      const c = document.createElement('canvas'); c.width = Math.round(vp.width); c.height = Math.round(vp.height);
      const ctx = c.getContext('2d');
      await P.page.render({ canvasContext: ctx, viewport: vp }).promise;
      const frames = [];
      const lines = P.lines || [];
      const covers = [];
      for (const L of lines) if (L.src === 'scan') refineScanLine(P, L, sampleColors(P, L));   // real letter height on scans
      for (const L of lines) {
        const rep = P.marks.find((m) => m.type === 'replace' && m.line === L);
        const col = sampleCanvas(c, L);
        covers.push([L, col.bg]);
        const text = rep ? rep.text : L.text;
        if (!text) continue;
        const basePt = (rep ? rep.size : L.size) * Hpt;
        const runs = rep ? runsOf(rep) : lineRuns(L, styleOf({ family: L.family || (L.serif ? 'serif' : 'sans'),
          bold: L.bold !== undefined ? L.bold : col.ink > 0.31, italic: !!L.italic, color: toHex(col.fg) }));
        frames.push({ runs: runs.map((r) => ({ ...r, size: basePt * r.k })), l: L.l, r: L.r, base: (rep ? rep.base : L.base) * Hpt, right: isNumber(L.text) });
      }
      for (const [L, bg] of covers) {
        const bh = (L.b - L.t) * c.height, pad = PAD * bh;
        ctx.fillStyle = css(bg); ctx.fillRect(L.l * c.width - pad, L.t * c.height - pad, (L.r - L.l) * c.width + 2 * pad, bh + 2 * pad);
      }
      for (const m of P.marks) {
        if (m.type === 'box' || m.type === 'ink') drawMark(ctx, m, c.width, c.height);
        if (m.type === 'text' && m.runs) {
          const top = Math.max(...m.runs.map((r) => r.k));
          frames.push({ runs: m.runs.map((r) => ({ ...r, size: m.size * Wpt * r.k })), l: m.x, r: m.x + 0.5, base: m.y * Hpt + 0.8 * top * m.size * Wpt, right: false });
        } else if (m.type === 'text') m.text.split('\n').forEach((t, i) => frames.push({ runs: [{ t, family: famOf(m), bold: !!m.bold, italic: !!m.italic, color: toHex(m.color), size: m.size * Wpt }],
          l: m.x, r: m.x + 0.5, base: m.y * Hpt + (0.8 + i * 1.2) * m.size * Wpt, right: false }));
      }
      const blob = await new Promise((r) => c.toBlob(r, 'image/jpeg', 0.88));
      const img = `bg${P.i + 1}.jpg`;
      put('word/media/' + img, await blob.arrayBuffer());
      media.push(img);
      // 2) Word XML for this page
      const cx = Math.round(Wpt * 12700), cy = Math.round(Hpt * 12700);
      const rid = 'rIdBg' + (P.i + 1);
      let x = `<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="20" w:lineRule="exact"/></w:pPr><w:r><w:drawing>` +
        `<wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="1" behindDoc="1" locked="1" layoutInCell="1" allowOverlap="1">` +
        `<wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="page"><wp:posOffset>0</wp:posOffset></wp:positionH>` +
        `<wp:positionV relativeFrom="page"><wp:posOffset>0</wp:posOffset></wp:positionV><wp:extent cx="${cx}" cy="${cy}"/>` +
        `<wp:effectExtent l="0" t="0" r="0" b="0"/><wp:wrapNone/><wp:docPr id="${pid}" name="Page ${P.i + 1}"/><wp:cNvGraphicFramePr/>` +
        `<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic>` +
        `<pic:nvPicPr><pic:cNvPr id="${pid}" name="${img}"/><pic:cNvPicPr/></pic:nvPicPr>` +
        `<pic:blipFill><a:blip r:embed="${rid}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>` +
        `<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>` +
        `</pic:pic></a:graphicData></a:graphic></wp:anchor></w:drawing></w:r></w:p>`;
      pid++;
      const fl = [];
      for (const f of frames) {
        const fsize = Math.max(...f.runs.map((r) => r.size));
        const tw = f.runs.reduce((a, r) => { mctx.font = fontCss(r, r.size); return a + mctx.measureText(r.t).width; }, 0);
        const lineH = fsize * LINE_H;
        const w = Math.max((f.r - f.l) * Wpt, tw) * 1.12 + 6;
        const left = f.right ? Math.max(0, f.r * Wpt - w) : f.l * Wpt;
        const top = f.base - BASE_AT * lineH;
        x += `<w:p><w:pPr><w:framePr w:w="${Math.round(w * TW)}" w:h="${Math.round(lineH * TW)}" w:hRule="atLeast" w:wrap="around" ` +
          `w:vAnchor="page" w:hAnchor="page" w:x="${Math.round(left * TW)}" w:y="${Math.max(0, Math.round(top * TW))}"/>` +
          `<w:spacing w:before="0" w:after="0" w:line="${Math.round(lineH * TW)}" w:lineRule="exact"/><w:jc w:val="${f.right ? 'right' : 'left'}"/></w:pPr>` +
          f.runs.map((r) => {
            const fam = FAMILY_WORD[r.family] || 'Arial', sz = Math.max(2, Math.round(r.size * 2));
            return `<w:r><w:rPr><w:rFonts w:ascii="${fam}" w:hAnsi="${fam}" w:cs="${fam}"/>` +
              `${r.bold ? '<w:b/>' : ''}${r.italic ? '<w:i/>' : ''}<w:color w:val="${hex6(r.color)}"/><w:sz w:val="${sz}"/><w:szCs w:val="${sz}"/></w:rPr>` +
              `<w:t xml:space="preserve">${xmlEsc(r.t)}</w:t></w:r>`;
          }).join('') + '</w:p>';
        fl.push({ w, h: lineH, x: left, top });
      }
      const sect = `<w:sectPr><w:pgSz w:w="${Math.round(Wpt * TW)}" w:h="${Math.round(Hpt * TW)}"/>` +
        `<w:pgMar w:top="0" w:right="0" w:bottom="0" w:left="0" w:header="0" w:footer="0" w:gutter="0"/></w:sectPr>`;
      const last = P.i === S.pages.length - 1;
      body.push(x + (last ? '' : `<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="20" w:lineRule="exact"/>${sect}</w:pPr></w:p>`));
      if (last) body.push(sect);
      layout.pages.push({ w: Wpt, h: Hpt, bg: img, frames: fl.length });
    }
    const ns = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ' +
      'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ' +
      'xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"';
    const head = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
    put('word/document.xml', `${head}<w:document ${ns}><w:body>${body.join('')}</w:body></w:document>`);
    put('word/_rels/document.xml.rels', `${head}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
      '<Relationship Id="rIdStyles" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>' +
      media.map((m, i) => `<Relationship Id="rIdBg${i + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/${m}"/>`).join('') +
      '</Relationships>');
    put('word/styles.xml', `${head}<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:docDefaults><w:rPrDefault><w:rPr>` +
      '<w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial" w:eastAsia="Arial"/><w:sz w:val="20"/><w:lang w:val="en-GB"/></w:rPr></w:rPrDefault>' +
      '<w:pPrDefault><w:pPr><w:spacing w:before="0" w:after="0"/></w:pPr></w:pPrDefault></w:docDefaults>' +
      '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style></w:styles>');
    put('word/dudely.json', JSON.stringify(layout));
    put('[Content_Types].xml', `${head}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>' +
      '<Default Extension="jpg" ContentType="image/jpeg"/><Default Extension="json" ContentType="application/json"/>' +
      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
      '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>' +
      '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/></Types>');
    put('_rels/.rels', `${head}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>' +
      '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/></Relationships>');
    const now = new Date().toISOString().replace(/\.\d+Z$/, 'Z');
    put('docProps/core.xml', `${head}<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" ` +
      'xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">' +
      `<dc:title>${xmlEsc(S.name)}</dc:title><dc:creator>Dudely Editor PC</dc:creator>` +
      `<dcterms:created xsi:type="dcterms:W3CDTF">${now}</dcterms:created><dcterms:modified xsi:type="dcterms:W3CDTF">${now}</dcterms:modified></cp:coreProperties>`);
    const out = await zip.generateAsync({ type: 'blob', mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', compression: 'DEFLATE' });
    await saveBlob(out, `${S.name}.docx`, 'Word document', { 'application/vnd.openxmlformats-officedocument.wordprocessingml.document': ['.docx'] });
  } catch (e) { console.error(e); hint('Could not make the Word file: ' + (e.message || e)); }
  finally { busy(''); }
}

async function saveBlob(blob, name, desc, accept) {
  if (window.showSaveFilePicker) {
    try {
      const h = await window.showSaveFilePicker({ suggestedName: name, types: [{ description: desc, accept }] });
      const w = await h.createWritable(); await w.write(blob); await w.close();
      hint(`Saved as ${h.name}.`); return;
    } catch (e) { if (e.name === 'AbortError') { hint('Not saved.'); return; } }
  }
  const a = document.createElement('a'); a.href = URL.createObjectURL(blob); a.download = name; a.click();
  setTimeout(() => URL.revokeObjectURL(a.href), 4000);
  hint(`Saved as ${name} in your Downloads folder.`);
}

/** A Word file made by this editor (edited in Word or not) becomes a PDF with the same layout. Other Word files: returns null. */
async function wordToPdf(buf) {
  const zip = await JSZip.loadAsync(buf);
  if (!zip.file('word/dudely.json') && !zip.file('word/tracedesk.json')) return null;   // made by this editor (old files too)
  const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
  const xml = new DOMParser().parseFromString(await zip.file('word/document.xml').async('string'), 'application/xml');
  const rels = new DOMParser().parseFromString(await zip.file('word/_rels/document.xml.rels').async('string'), 'application/xml');
  const target = {}; for (const r of rels.getElementsByTagName('Relationship')) target[r.getAttribute('Id')] = r.getAttribute('Target');
  const attr = (el, n) => el ? el.getAttributeNS(W, n) || el.getAttribute('w:' + n) : null;
  const pages = []; let cur = { frames: [], bg: null };
  const body = xml.getElementsByTagNameNS(W, 'body')[0];
  const closePage = (sect) => {
    const sz = sect && sect.getElementsByTagNameNS(W, 'pgSz')[0];
    cur.w = sz ? +attr(sz, 'w') / TW : 595; cur.h = sz ? +attr(sz, 'h') / TW : 842;
    pages.push(cur); cur = { frames: [], bg: null };
  };
  for (const node of [...body.children]) {
    if (node.localName === 'sectPr') { closePage(node); continue; }
    if (node.localName !== 'p') continue;
    const blip = node.getElementsByTagNameNS('http://schemas.openxmlformats.org/drawingml/2006/main', 'blip')[0];
    if (blip && !cur.bg) cur.bg = target[blip.getAttributeNS('http://schemas.openxmlformats.org/officeDocument/2006/relationships', 'embed')];
    const pPr = node.getElementsByTagNameNS(W, 'pPr')[0];
    const fp = pPr && pPr.getElementsByTagNameNS(W, 'framePr')[0];
    if (fp) {
      const runs = [...node.getElementsByTagNameNS(W, 'r')];
      let text = '';
      for (const r of runs) for (const ch of r.children) {
        if (ch.localName === 't') text += ch.textContent; else if (ch.localName === 'tab') text += ' '; else if (ch.localName === 'br') text += '\n';
      }
      const styleOfRun = (r) => {
        const pr = r.getElementsByTagNameNS(W, 'rPr')[0], g = (n) => pr && pr.getElementsByTagNameNS(W, n)[0];
        const on = (el) => !!el && attr(el, 'val') !== '0' && attr(el, 'val') !== 'false';
        const f = g('rFonts') ? attr(g('rFonts'), 'ascii') || '' : '';
        const c = g('color') && attr(g('color'), 'val');
        let t = ''; for (const ch of r.children) { if (ch.localName === 't') t += ch.textContent; else if (ch.localName === 'tab') t += ' '; else if (ch.localName === 'br') t += '\n'; }
        return { t, bold: on(g('b')), italic: on(g('i')), size: g('sz') ? +attr(g('sz'), 'val') / 2 : 10,
          family: /courier|mono|consol/i.test(f) ? 'mono' : /times|georgia|garamond|cambria|serif/i.test(f) && !/sans/i.test(f) ? 'serif' : 'sans',
          color: /^[0-9A-Fa-f]{6}$/.test(c || '') ? c : '000000' };
      };
      const wruns = runs.map(styleOfRun).filter((r) => r.t);
      const rPr = runs[0] && runs[0].getElementsByTagNameNS(W, 'rPr')[0];
      const get = (n) => rPr && rPr.getElementsByTagNameNS(W, n)[0];
      const sp = pPr.getElementsByTagNameNS(W, 'spacing')[0], jc = pPr.getElementsByTagNameNS(W, 'jc')[0];
      const b = get('b'), it = get('i');
      const rf = get('rFonts') ? attr(get('rFonts'), 'ascii') || '' : '';
      cur.frames.push({
        text, wruns, x: +attr(fp, 'x') / TW, y: +attr(fp, 'y') / TW, w: +attr(fp, 'w') / TW,
        size: get('sz') ? +attr(get('sz'), 'val') / 2 : 10, bold: !!b && attr(b, 'val') !== '0' && attr(b, 'val') !== 'false',
        family: /courier|mono|consol/i.test(rf) ? 'mono' : /times|georgia|garamond|cambria|serif/i.test(rf) && !/sans/i.test(rf) ? 'serif' : 'sans',
        italic: !!it && attr(it, 'val') !== '0' && attr(it, 'val') !== 'false',
        color: get('color') && /^[0-9A-Fa-f]{6}$/.test(attr(get('color'), 'val') || '') ? attr(get('color'), 'val') : '000000',
        lineH: sp && attr(sp, 'line') ? +attr(sp, 'line') / TW : null, right: jc && attr(jc, 'val') === 'right',
      });
    }
    const sect = pPr && pPr.getElementsByTagNameNS(W, 'sectPr')[0];
    if (sect) closePage(sect);
  }
  if (cur.frames.length || cur.bg) closePage(null);
  const doc = await PDFDocument.create();
  const F = await embedAllFonts(doc);
  for (const pg of pages) {
    const page = doc.addPage([pg.w, pg.h]);
    if (pg.bg && zip.file('word/' + pg.bg)) {
      const img = await doc.embedJpg(await zip.file('word/' + pg.bg).async('uint8array'));
      page.drawImage(img, { x: 0, y: 0, width: pg.w, height: pg.h });
    }
    for (const f of pg.frames) {
      const font = F.pick(f);
      const lineH = f.lineH || f.size * LINE_H;
      const color = rgb(parseInt(f.color.slice(0, 2), 16) / 255, parseInt(f.color.slice(2, 4), 16) / 255, parseInt(f.color.slice(4, 6), 16) / 255);
      if (f.wruns.length > 1 && !f.text.includes('\n')) {
        const top = Math.max(...f.wruns.map((r) => r.size)), lh = f.lineH || top * LINE_H;
        const parts = f.wruns.map((r) => { const fo = F.pick(r); return { ...r, fo, safe: printable(r.t, fo) }; }).filter((r) => r.safe);
        const tw = parts.reduce((a, r) => a + r.fo.widthOfTextAtSize(r.safe, r.size), 0);
        let x = f.right ? f.x + f.w - 3 - tw : f.x;
        for (const r of parts) {
          const c = r.color;
          page.drawText(r.safe, { x, y: pg.h - (f.y + BASE_AT * lh), size: r.size, font: r.fo,
            color: rgb(parseInt(c.slice(0, 2), 16) / 255, parseInt(c.slice(2, 4), 16) / 255, parseInt(c.slice(4, 6), 16) / 255) });
          x += r.fo.widthOfTextAtSize(r.safe, r.size);
        }
        continue;
      }
      f.text.split('\n').forEach((line, i) => {
        const safe = printable(line, font); if (!safe) return;
        const tw = font.widthOfTextAtSize(safe, f.size);
        const x = f.right ? f.x + f.w - 3 - tw : f.x;
        const base = f.y + BASE_AT * lineH + i * lineH;
        page.drawText(safe, { x, y: pg.h - base, size: f.size, font, color });
      });
    }
  }
  hint('Word file turned back into a PDF with the same layout. Edit it here or click Save edited PDF.');
  return await doc.save();
}

/** Any other Word file: shown as pages, then saved through the print dialog (choose "Save as PDF"). */
async function printOtherWord(buf, name) {
  const area = $('printArea'), holder = $('printDoc');
  holder.innerHTML = '';
  area.style.display = 'block';
  $('printTitle').textContent = name;
  busy('Opening the Word file…');
  try { await docx.renderAsync(buf, holder, null, { inWrapper: true, ignoreLastRenderedPageBreak: false }); }
  catch (e) { holder.textContent = 'Could not read this Word file: ' + (e.message || e); }
  finally { busy(''); }
}
$('btnPrintPdf').onclick = () => window.print();
$('btnPrintClose').onclick = () => { $('printArea').style.display = 'none'; $('printDoc').innerHTML = ''; };
// used by the Photos and Merge sections to open their result here
window.dudelyOpen = (files) => openFiles(files);

// Test hooks (only with ?debug in the address)
if (/[?&]debug/.test(location.search)) {
  window.dbgFind = (texts) => texts.map((t) => (S.pages[0].lines || []).find((L) => L.text.includes(t)) || null);
  window.dbgMarks = () => S.pages.flatMap((P) => (P.marks || []).map(({ line, ...m }) => m));
}
