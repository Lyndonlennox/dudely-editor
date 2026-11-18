// Dudely Editor PC: shared PDF helpers used by the tools in pdftools.js.
import * as pdfjsLib from './pdf.min.mjs';
pdfjsLib.GlobalWorkerOptions.workerSrc = './pdf.worker.min.mjs';
export { pdfjsLib };
export const L = window.PDFLib;
const { PDFDocument, StandardFonts, rgb, degrees } = L;

export const $ = (id) => document.getElementById(id);
export const baseName = (n) => n.replace(/\.[^.]+$/, '');
export const isPdf = (f) => f.type === 'application/pdf' || /\.pdf$/i.test(f.name);
export const isImg = (f) => (f.type || '').startsWith('image/') || /\.(jpe?g|png|webp|gif|bmp)$/i.test(f.name);
export const MM = 72 / 25.4;
export const fmtSize = (b) => (b < 1024 ? `${b} B` : b < 1048576 ? `${(b / 1024).toFixed(0)} KB` : `${(b / 1048576).toFixed(2)} MB`);
export const pdfBlob = (bytes) => new Blob([bytes], { type: 'application/pdf' });
export const hexRgb = (h) => rgb(parseInt(h.slice(1, 3), 16) / 255, parseInt(h.slice(3, 5), 16) / 255, parseInt(h.slice(5, 7), 16) / 255);
export const tick = () => new Promise((r) => setTimeout(r, 0));

/** Load with pdf-lib. Tells the person plainly when the file is locked. */
export async function libLoad(bytes, extra = {}) {
  try { return await PDFDocument.load(bytes, { updateMetadata: false, ...extra }); }
  catch (e) {
    if (/encrypt/i.test(e.message)) throw new Error('This PDF is password-protected. Use Unlock first.');
    throw new Error('This PDF could not be read. Try Repair first.');
  }
}
export async function jsLoad(bytes, password) {
  return pdfjsLib.getDocument({ data: bytes.slice(), password, isEvalSupported: false }).promise;
}

/** "1-3, 5, 8-" → sorted page indexes (0-based). Empty text → every page. */
export function pageList(text, n) {
  const t = (text || '').trim();
  if (!t || /^all$/i.test(t)) return [...Array(n).keys()];
  const set = new Set();
  for (const part of t.split(/[,;\s]+/).filter(Boolean)) {
    const m = part.match(/^(\d*)\s*-\s*(\d*)$/);
    let a, b;
    if (m) { a = m[1] ? +m[1] : 1; b = m[2] ? +m[2] : n; } else if (/^\d+$/.test(part)) { a = b = +part; }
    else if (/^odd$/i.test(part)) { for (let i = 0; i < n; i += 2) set.add(i); continue; }
    else if (/^even$/i.test(part)) { for (let i = 1; i < n; i += 2) set.add(i); continue; }
    else throw new Error(`"${part}" is not a page or range. Use numbers like 1-3, 5, 8-.`);
    if (a > b) [a, b] = [b, a];
    for (let i = Math.max(1, a); i <= Math.min(n, b); i++) set.add(i - 1);
  }
  if (!set.size) throw new Error(`No pages match "${t}". This PDF has ${n} page${n > 1 ? 's' : ''}.`);
  return [...set].sort((x, y) => x - y);
}

// ------------------------------------------------------------ rotation-aware page geometry
/** What the reader sees: width, height and rotation of a page. */
export function vis(page) {
  const cb = page.getCropBox(); const r = ((page.getRotation().angle % 360) + 360) % 360;
  return { r, cb, W: r % 180 ? cb.height : cb.width, H: r % 180 ? cb.width : cb.height };
}
/** Point as seen (origin bottom-left of what the reader sees) → page's own coordinates. */
export function toUser(page, vx, vy) {
  const { r, cb } = vis(page); let x, y;
  if (r === 90) { x = cb.width - vy; y = vx; } else if (r === 180) { x = cb.width - vx; y = cb.height - vy; }
  else if (r === 270) { x = vy; y = cb.height - vx; } else { x = vx; y = vy; }
  return { x: x + cb.x, y: y + cb.y };
}
export function rectToUser(page, x0, y0, x1, y1) {
  const a = toUser(page, x0, y0), b = toUser(page, x1, y1);
  return { x: Math.min(a.x, b.x), y: Math.min(a.y, b.y), w: Math.abs(a.x - b.x), h: Math.abs(a.y - b.y) };
}
/** Draw text placed as the reader sees it, whatever the page rotation. angle = extra turn (degrees, anticlockwise). */
export function drawSeen(page, text, vx, vy, o) {
  const { r } = vis(page); const p = toUser(page, vx, vy);
  page.drawText(text, { x: p.x, y: p.y, size: o.size, font: o.font, color: o.color, opacity: o.opacity ?? 1, rotate: degrees(r + (o.angle || 0)) });
}
/** x,y for a text block at one of nine places. pos: tl tc tr ml c mr bl bc br */
export function placeAt(pos, W, H, tw, size, margin) {
  const col = pos[1] === 'l' ? 'l' : pos[1] === 'r' ? 'r' : 'c';
  const row = pos === 'c' ? 'm' : pos[0];
  const x = col === 'l' ? margin : col === 'r' ? W - margin - tw : (W - tw) / 2;
  const y = row === 't' ? H - margin - size * 0.8 : row === 'b' ? margin : (H - size * 0.7) / 2;
  return { x, y };
}
export const POS9 = [['tl', 'Top left'], ['tc', 'Top centre'], ['tr', 'Top right'], ['bl', 'Bottom left'], ['bc', 'Bottom centre'], ['br', 'Bottom right']];

/** Draw a page from another document into box (x,y,w,h), upright and scaled to fit. Returns the scale used. */
export async function placePage(out, target, srcPage, box, fit = 'contain') {
  const { r, cb } = vis(srcPage);
  const emb = await out.embedPage(srcPage, { left: cb.x, bottom: cb.y, right: cb.x + cb.width, top: cb.y + cb.height });
  const vw = r % 180 ? emb.height : emb.width, vh = r % 180 ? emb.width : emb.height;
  const s = fit === 'cover' ? Math.max(box.w / vw, box.h / vh) : Math.min(box.w / vw, box.h / vh);
  const x0 = box.x + (box.w - vw * s) / 2, y0 = box.y + (box.h - vh * s) / 2;
  const w = emb.width * s, h = emb.height * s;
  const at = { 0: [x0, y0], 90: [x0, y0 + w], 180: [x0 + w, y0 + h], 270: [x0 + h, y0] }[r];
  target.drawPage(emb, { x: at[0], y: at[1], xScale: s, yScale: s, rotate: degrees(-r) });
  return s;
}
/** Copy of the document where every page is upright (rotation 0) at the size the reader sees. */
export async function upright(doc) {
  const out = await PDFDocument.create();
  for (const p of doc.getPages()) {
    const { W, H } = vis(p); const np = out.addPage([W, H]);
    await placePage(out, np, p, { x: 0, y: 0, w: W, h: H });
  }
  return PDFDocument.load(await out.save());   // saved once so its pages can be embedded again
}

// ------------------------------------------------------------ pictures of pages (pdf.js)
export async function renderPage(js, i, scale, opts = {}) {
  const page = await js.getPage(i + 1);
  const vp = page.getViewport({ scale });
  const c = document.createElement('canvas'); c.width = Math.max(1, Math.round(vp.width)); c.height = Math.max(1, Math.round(vp.height));
  const g = c.getContext('2d', { willReadFrequently: !!opts.read }); g.fillStyle = '#fff'; g.fillRect(0, 0, c.width, c.height);
  await page.render({ canvasContext: g, viewport: vp, annotationMode: opts.annotations === false ? pdfjsLib.AnnotationMode.DISABLE : pdfjsLib.AnnotationMode.ENABLE }).promise;
  const base = page.getViewport({ scale: 1 });
  page.cleanup();
  return { canvas: c, w: base.width, h: base.height };
}
export const canvasBlob = (c, type = 'image/jpeg', q = 0.85) => new Promise((r) => c.toBlob(r, type, q));
export async function canvasBytes(c, type, q) { return new Uint8Array(await (await canvasBlob(c, type, q)).arrayBuffer()); }
export function grayCanvas(c) {
  const g = c.getContext('2d'); const im = g.getImageData(0, 0, c.width, c.height); const d = im.data;
  for (let i = 0; i < d.length; i += 4) { const v = 0.299 * d[i] + 0.587 * d[i + 1] + 0.114 * d[i + 2]; d[i] = d[i + 1] = d[i + 2] = v; }
  g.putImageData(im, 0, 0); return c;
}
/** Make a PDF where every page is a picture. dpi, quality, gray, pages (indexes). */
export async function rasterPdf(bytes, o, progress) {
  const js = await jsLoad(bytes, o.password); const out = await PDFDocument.create();
  const idx = o.pages || [...Array(js.numPages).keys()];
  for (const [k, i] of idx.entries()) {
    progress?.(`Page ${k + 1} of ${idx.length}…`);
    const { canvas, w, h } = await renderPage(js, i, (o.dpi || 150) / 72, { read: o.gray, annotations: o.annotations });
    if (o.gray) grayCanvas(canvas);
    const img = await out.embedJpg(await canvasBytes(canvas, 'image/jpeg', o.quality || 0.8));
    out.addPage([w, h]).drawImage(img, { x: 0, y: 0, width: w, height: h });
    canvas.width = canvas.height = 0; await tick();
  }
  await js.destroy();
  return out;
}

// ------------------------------------------------------------ text from pages
/** Lines of text on a page: [{y, items:[{x, w, s, size}]}], top to bottom, using what the reader sees. */
export async function pageLines(js, i) {
  const page = await js.getPage(i + 1); const vp = page.getViewport({ scale: 1 });
  const tc = await page.getTextContent(); const items = [];
  for (const it of tc.items) {
    if (!it.str || !it.str.trim()) continue;
    const [a, b, , , e, f] = pdfjsLib.Util.transform(vp.transform, it.transform);
    const size = Math.hypot(a, b) || 10;
    items.push({ x: e, y: f, w: it.width * (vp.scale || 1), s: it.str, size });
  }
  items.sort((p, q) => p.y - q.y || p.x - q.x);
  const lines = [];
  for (const it of items) {
    const L0 = lines[lines.length - 1];
    if (L0 && Math.abs(L0.y - it.y) < Math.max(2, it.size * 0.45)) L0.items.push(it); else lines.push({ y: it.y, items: [it] });
  }
  for (const L1 of lines) L1.items.sort((p, q) => p.x - q.x);
  return { lines, width: vp.width, height: vp.height };
}
export function lineText(L1) {
  let s = '', end = null;
  for (const it of L1.items) {
    if (end !== null) { const gap = it.x - end; if (gap > it.size * 0.15 && !s.endsWith(' ') && !it.s.startsWith(' ')) s += gap > it.size * 2 ? '    ' : ' '; }
    s += it.s; end = it.x + it.w;
  }
  return s.trim();
}

// ------------------------------------------------------------ fonts
export async function fontOf(doc, fam = 'sans', bold = false) {
  const F = { sans: [StandardFonts.Helvetica, StandardFonts.HelveticaBold], serif: [StandardFonts.TimesRoman, StandardFonts.TimesRomanBold], mono: [StandardFonts.Courier, StandardFonts.CourierBold] };
  return doc.embedFont(F[fam][bold ? 1 : 0]);
}
/** Standard fonts only know Western characters; swap anything else so saving never fails. */
export const safeText = (s) => String(s).replace(/[‘’]/g, "'").replace(/[“”]/g, '"').replace(/[–—]/g, '-').replace(/[^\x20-\x7E\xA0-\xFF•€]/g, '?');

// ------------------------------------------------------------ OCR (Tesseract, works offline)
let ocrWorker = null;
export async function ocr() {
  if (!ocrWorker) ocrWorker = await window.Tesseract.createWorker('eng', 1, {
    workerPath: new URL('./worker.min.js', location.href).href,
    corePath: new URL('./', location.href).href,
    langPath: new URL('.', location.href).href.replace(/\/$/, ''),
  });
  return ocrWorker;
}

// ------------------------------------------------------------ zip
export async function zipOf(files) {
  const z = new window.JSZip(); const used = new Set();
  for (const f of files) {
    let n = f.name, k = 2; while (used.has(n.toLowerCase())) n = f.name.replace(/(\.[^.]+)?$/, ` (${k++})$1`);
    used.add(n.toLowerCase()); z.file(n, f.blob);
  }
  return z.generateAsync({ type: 'blob', compression: 'DEFLATE' });
}
export const esc = (s) => String(s).replace(/[&<>"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
export const xmlEsc = (s) => String(s).replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&apos;' }[c])).replace(/[\x00-\x08\x0B\x0C\x0E-\x1F]/g, '');
