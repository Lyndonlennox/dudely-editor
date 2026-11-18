// Dudely Editor PC: the tool list. Each tool takes files and options and gives back new files.
import {
  L, $, baseName, isPdf, isImg, MM, fmtSize, pdfBlob, hexRgb, tick, libLoad, jsLoad, pageList, vis, toUser, rectToUser, drawSeen,
  placeAt, POS9, placePage, upright, renderPage, canvasBlob, canvasBytes, grayCanvas, rasterPdf, pageLines, lineText, fontOf, safeText, ocr, pdfjsLib,
} from './pdfcore.js';
import { xlsxBlob, pptxBlob } from './officefiles.js';
const { PDFDocument, PDFName, PDFHexString, PDFDict, PDFArray, degrees, rgb, PDFTextField, PDFCheckBox, PDFDropdown, PDFOptionList, PDFRadioGroup } = L;
const kindOf = (f) => f instanceof PDFTextField ? 'text' : f instanceof PDFCheckBox ? 'check' : f instanceof PDFRadioGroup ? 'radio' : (f instanceof PDFDropdown || f instanceof PDFOptionList) ? 'list' : '';

export const GROUPS = [
  ['merge', 'Merge & organise', '#0E9F6E'], ['split', 'Split', '#65A30D'], ['edit', 'Edit & sign', '#2563EB'],
  ['compress', 'Compress', '#0891B2'], ['security', 'Security', '#0D9488'], ['from', 'Convert from PDF', '#EA580C'],
  ['to', 'Convert to PDF', '#7C3AED'], ['other', 'Page & file tools', '#DB2777'], ['scans', 'Scans & photos', '#DC2626'],
];
export const TOOLS = [];
const T = (id, def) => TOOLS.push({ id, accept: 'pdf', multi: false, opts: [], ...def });
const out = (name, blob) => ({ name, blob });
const pdfOut = async (doc, name, o = {}) => out(name, pdfBlob(await doc.save({ useObjectStreams: o.compact !== false })));
const SIZE = { a4: [595.28, 841.89], letter: [612, 792], legal: [612, 1008], a3: [841.89, 1190.55], a5: [419.53, 595.28] };
const SIZE_OPTS = [['a4', 'A4'], ['letter', 'Letter'], ['legal', 'Legal'], ['a3', 'A3'], ['a5', 'A5']];
const pagesOpt = (label = 'Pages') => ({ id: 'pages', label, type: 'text', value: '', placeholder: 'All pages, or e.g. 1-3, 5, 8-', hint: 'Leave empty for every page.' });
const fontOpt = { id: 'font', label: 'Font', type: 'select', value: 'sans', options: [['sans', 'Arial / Helvetica'], ['serif', 'Times New Roman'], ['mono', 'Courier']] };

// =============================================================== links to the existing workspaces
T('edit', { name: 'Edit', group: 'edit', icon: 'edit', go: 'edit', desc: 'Change words already on the page, add text, white-out, highlight.' });
T('sign', { name: 'Fill & Sign', group: 'edit', icon: 'sign', go: 'edit', tool: 'pen', desc: 'Write or draw your signature, ticks and crosses on any PDF.' });
T('merge', { name: 'Merge', group: 'merge', icon: 'merge', go: 'merge', desc: 'Join PDFs and photos into one PDF.' });
T('split', { name: 'Split by pages', group: 'split', icon: 'split', go: 'split', desc: 'A new file every few pages, or by page ranges.' });
T('photos', { name: 'Improve photos', group: 'scans', icon: 'photo', go: 'photos', desc: 'Brighten, sharpen and crop document photos.' });
T('pdf2word', { name: 'PDF to Word', group: 'from', icon: 'word', go: 'edit', then: 'word', desc: 'Editable Word file with the same layout.' });
T('word2pdf', { name: 'Word to PDF', group: 'to', icon: 'word', go: 'edit', accept: 'docx', desc: 'Turn a Word file into a PDF.' });

// =============================================================== page grid tools (pages view)
T('organize', { name: 'Organise', group: 'merge', icon: 'grid', pages: 'organize', desc: 'Reorder, turn, copy, add blank and remove pages.' });
T('delete', { name: 'Delete pages', group: 'edit', icon: 'trash', pages: 'delete', desc: 'Pick pages to remove and save the rest.' });
T('extract', { name: 'Extract pages', group: 'split', icon: 'extract', pages: 'extract', desc: 'Pick pages and save them as a new PDF.' });
T('rotate', { name: 'Rotate', group: 'other', icon: 'rotate', pages: 'rotate', desc: 'Turn some or all pages and save.' });
T('forms', { name: 'Create forms', group: 'edit', icon: 'form', forms: true, desc: 'Draw text boxes, tick boxes and drop-down lists people can fill in.' });

// =============================================================== MERGE
T('mix', {
  name: 'Alternate & mix', group: 'merge', icon: 'mix', multi: true, min: 2,
  desc: 'Take pages in turn from two or more PDFs, for example front and back scans.',
  opts: [{ id: 'take', label: 'Pages to take from each file in turn', type: 'number', value: 1, min: 1, max: 50 },
    { id: 'rev', label: 'Reverse the second file (back sides scanned last page first)', type: 'check', value: true }],
  async run(files, o) {
    const docs = []; for (const f of files) docs.push(await libLoad(f.bytes));
    const res = await PDFDocument.create();
    const queues = docs.map((d, i) => { const ix = d.getPageIndices(); return i === 1 && o.rev ? ix.reverse() : ix; });
    const take = Math.max(1, +o.take || 1);
    while (queues.some((q) => q.length)) {
      for (const [i, q] of queues.entries()) {
        const part = q.splice(0, take); if (!part.length) continue;
        (await res.copyPages(docs[i], part)).forEach((p) => res.addPage(p));
      }
    }
    return [await pdfOut(res, `${baseName(files[0].name)}_mixed.pdf`)];
  },
});

// =============================================================== SPLIT
T('half', {
  name: 'Split in half', group: 'split', icon: 'half', desc: 'Cut each page into two, for scanned double pages.',
  opts: [{ id: 'dir', label: 'Cut', type: 'select', value: 'v', options: [['v', 'Down the middle (left | right)'], ['h', 'Across the middle (top / bottom)']] },
    { id: 'rtl', label: 'Right half first (right-to-left books)', type: 'check', value: false },
    { id: 'skip', label: 'Leave the first and last page whole', type: 'check', value: false }],
  async run([f], o) {
    const src = await upright(await libLoad(f.bytes)); const res = await PDFDocument.create(); const n = src.getPageCount();
    for (const [i, p] of src.getPages().entries()) {
      const { width: w, height: h } = p.getSize();
      if (o.skip && (i === 0 || i === n - 1)) { const [c] = await res.copyPages(src, [i]); res.addPage(c); continue; }
      let boxes = o.dir === 'v' ? [[0, 0, w / 2, h], [w / 2, 0, w, h]] : [[0, h / 2, w, h], [0, 0, w, h / 2]];
      if (o.rtl && o.dir === 'v') boxes = boxes.reverse();
      for (const [l, b, r, t] of boxes) {
        const e = await res.embedPage(p, { left: l, bottom: b, right: r, top: t });
        res.addPage([r - l, t - b]).drawPage(e, { x: 0, y: 0 });
      }
    }
    return [await pdfOut(res, `${baseName(f.name)}_halves.pdf`)];
  },
});
T('splitsize', {
  name: 'Split by size', group: 'split', icon: 'scale', desc: 'Cut a large PDF into files no bigger than a size you choose, for email.',
  opts: [{ id: 'mb', label: 'Largest file size (MB)', type: 'number', value: 5, min: 0.2, max: 500, step: 0.1 }],
  async run([f], o, ctx) {
    const src = await libLoad(f.bytes); const n = src.getPageCount(); const limit = Math.max(0.1, +o.mb) * 1048576;
    const build = async (ix) => { const d = await PDFDocument.create(); (await d.copyPages(src, ix)).forEach((p) => d.addPage(p)); return d.save({ useObjectStreams: true }); };
    const outs = []; let start = 0;
    while (start < n) {
      ctx.progress(`Working out file ${outs.length + 1} (from page ${start + 1})…`);
      let lo = start + 1, hi = n, best = start + 1, bestBytes = null;
      while (lo <= hi) { const mid = (lo + hi) >> 1; const b = await build(range(start, mid)); if (b.length <= limit || mid === start + 1) { best = mid; bestBytes = b; lo = mid + 1; } else hi = mid - 1; }
      if (!bestBytes) bestBytes = await build(range(start, best));
      outs.push(out(`${baseName(f.name)}_part${outs.length + 1}_p${start + 1}-${best}.pdf`, pdfBlob(bestBytes)));
      start = best;
    }
    ctx.note(outs.length === 1 ? 'The whole PDF already fits in one file of that size.' : `${outs.length} files. A single page bigger than the limit stays as its own file.`);
    return outs;
  },
});
const range = (a, b) => [...Array(b - a).keys()].map((k) => k + a);
T('splitbm', {
  name: 'Split by bookmarks', group: 'split', icon: 'bookmark', desc: 'One file for each bookmark (chapter) in the PDF.',
  opts: [{ id: 'level', label: 'Bookmark level', type: 'select', value: '1', options: [['1', 'Top-level bookmarks only'], ['2', 'Top level and the level below']] }],
  async run([f], o) {
    const js = await jsLoad(f.bytes); const outline = (await js.getOutline()) || [];
    if (!outline.length) throw new Error('This PDF has no bookmarks. Use Split by pages, or add bookmarks with Create bookmarks.');
    const marks = [];
    const walk = async (items, lvl) => {
      for (const it of items) {
        let dest = it.dest; if (typeof dest === 'string') dest = await js.getDestination(dest);
        if (Array.isArray(dest) && dest[0]) { try { const pi = typeof dest[0] === 'number' ? dest[0] : await js.getPageIndex(dest[0]); marks.push({ title: it.title, page: pi }); } catch (e) { /* skip broken */ } }
        if (lvl < +o.level && it.items?.length) await walk(it.items, lvl + 1);
      }
    };
    await walk(outline, 1);
    const starts = [...new Map(marks.sort((a, b) => a.page - b.page).map((m) => [m.page, m])).values()];
    if (starts[0]?.page > 0) starts.unshift({ title: 'Start', page: 0 });
    const src = await libLoad(f.bytes); const outs = [];
    for (const [i, m] of starts.entries()) {
      const end = i + 1 < starts.length ? starts[i + 1].page : js.numPages; const d = await PDFDocument.create();
      (await d.copyPages(src, range(m.page, end))).forEach((p) => d.addPage(p));
      const safe = m.title.replace(/[\\/:*?"<>|]+/g, ' ').trim().slice(0, 60) || `part${i + 1}`;
      outs.push(await pdfOut(d, `${String(i + 1).padStart(2, '0')} ${safe}.pdf`));
    }
    await js.destroy(); return outs;
  },
});
T('splittext', {
  name: 'Split by text', group: 'split', icon: 'search', desc: 'Start a new file at every page that contains words you type, e.g. "Invoice No".',
  opts: [{ id: 'text', label: 'Words to look for', type: 'text', value: '', placeholder: 'e.g. Invoice No', required: true },
    { id: 'name', label: 'Name each file after the words that follow it on that page', type: 'check', value: true }],
  async run([f], o, ctx) {
    const needle = o.text.trim().toLowerCase(); if (!needle) throw new Error('Type the words to look for.');
    const js = await jsLoad(f.bytes); const starts = [];
    for (let i = 0; i < js.numPages; i++) {
      ctx.progress(`Reading page ${i + 1} of ${js.numPages}…`);
      const { lines } = await pageLines(js, i); const txt = lines.map(lineText).join('\n'); const k = txt.toLowerCase().indexOf(needle);
      if (k >= 0) starts.push({ page: i, label: txt.slice(k + needle.length).split('\n')[0].replace(/^[\s:#.\-]+/, '').trim().split(/\s{2,}/)[0].slice(0, 40) });
    }
    if (!starts.length) throw new Error(`"${o.text}" was not found. If the PDF is a scan, run OCR first.`);
    if (starts[0].page > 0) starts.unshift({ page: 0, label: 'start' });
    const src = await libLoad(f.bytes); const outs = [];
    for (const [i, s] of starts.entries()) {
      const end = i + 1 < starts.length ? starts[i + 1].page : js.numPages; const d = await PDFDocument.create();
      (await d.copyPages(src, range(s.page, end))).forEach((p) => d.addPage(p));
      const lbl = o.name && s.label ? ' ' + s.label.replace(/[\\/:*?"<>|]+/g, ' ') : '';
      outs.push(await pdfOut(d, `${baseName(f.name)}_${String(i + 1).padStart(2, '0')}${lbl}.pdf`));
    }
    await js.destroy(); return outs;
  },
});

// =============================================================== EDIT
T('fill', {
  name: 'Fill in a form', group: 'edit', icon: 'fill', desc: 'Type into the boxes of a PDF form and save it, optionally locked.',
  async fields([f]) {
    const d = await libLoad(f.bytes); const fields = d.getForm().getFields();
    if (!fields.length) return [{ type: 'info', label: 'This PDF has no form boxes to fill. Use Fill & Sign to write on it, or Create forms to add boxes.' }];
    const opts = fields.map((fl, i) => {
      const k = kindOf(fl); const name = fl.getName(); const id = 'f' + i;
      if (k === 'text') return { id, field: name, label: name, type: fl.isMultiline() ? 'textarea' : 'text', value: fl.getText() || '' };
      if (k === 'check') return { id, field: name, label: name, type: 'check', value: fl.isChecked() };
      if (k === 'list') return { id, field: name, label: name, type: 'select', value: fl.getSelected()[0] || '', options: [['', '—'], ...fl.getOptions().map((x) => [x, x])] };
      if (k === 'radio') return { id, field: name, label: name, type: 'select', value: fl.getSelected() || '', options: [['', '—'], ...fl.getOptions().map((x) => [x, x])] };
      return null;
    }).filter(Boolean);
    return [...opts, { id: 'flat', label: 'Lock the answers so they cannot be changed (flatten)', type: 'check', value: false }];
  },
  async run([f], o, ctx) {
    const d = await libLoad(f.bytes); const form = d.getForm();
    for (const op of ctx.opts.filter((x) => x.field)) {
      const fl = form.getField(op.field); const v = o[op.id]; const k = kindOf(fl);
      try {
        if (k === 'text') fl.setText(v || undefined);
        else if (k === 'check') v ? fl.check() : fl.uncheck();
        else if (k === 'radio') { if (v) fl.select(v); else fl.clear(); }
        else if (v) fl.select(v); else fl.clear();
      } catch (e) { ctx.note(`Could not fill "${op.field}".`); }
    }
    if (o.flat) form.flatten();
    return [await pdfOut(d, `${baseName(f.name)}_filled.pdf`)];
  },
});

// =============================================================== COMPRESS
T('compress', {
  name: 'Compress', group: 'compress', icon: 'compress', multi: true, desc: 'Make PDFs smaller to send by email or WhatsApp.',
  opts: [{ id: 'level', label: 'How much', type: 'select', value: 'medium', options: [['light', 'Light: tidy the file, text stays text'], ['medium', 'Medium: pages become clear pictures (150 dpi)'], ['strong', 'Strong: smallest file, lower picture quality (100 dpi)']] },
    { id: 'gray', label: 'Also make it black and white (smaller still)', type: 'check', value: false, show: (o) => o.level !== 'light' }],
  async run(files, o, ctx) {
    const outs = [];
    for (const f of files) {
      let bytes;
      if (o.level === 'light') bytes = await (await libLoad(f.bytes)).save({ useObjectStreams: true });
      else bytes = await (await rasterPdf(f.bytes, { dpi: o.level === 'strong' ? 100 : 150, quality: o.level === 'strong' ? 0.55 : 0.7, gray: o.gray }, ctx.progress)).save({ useObjectStreams: true });
      const was = f.bytes.length;
      if (bytes.length >= was) { ctx.note(`${f.name}: already as small as this setting can make it (${fmtSize(was)}). Try a stronger level.`); bytes = f.bytes; }
      else ctx.note(`${f.name}: ${fmtSize(was)} → ${fmtSize(bytes.length)} (${Math.round((1 - bytes.length / was) * 100)}% smaller).`);
      outs.push(out(`${baseName(f.name)}_compressed.pdf`, pdfBlob(bytes)));
    }
    if (o.level !== 'light') ctx.note('Medium and Strong turn pages into pictures, so the words can no longer be selected or searched.');
    return outs;
  },
});

// =============================================================== SECURITY
let SecPDF = null;
const secure = async () => (SecPDF ||= await import('./pdf-lib-secure.mjs'));
T('protect', {
  name: 'Protect', group: 'security', icon: 'lock', desc: 'Add a password. Choose whether people may print, copy or change it.',
  opts: [{ id: 'pw', label: 'Password to open the PDF', type: 'password', value: '', required: true },
    { id: 'pw2', label: 'Type the password again', type: 'password', value: '', required: true },
    { id: 'print', label: 'Allow printing', type: 'check', value: true }, { id: 'copy', label: 'Allow copying text', type: 'check', value: false },
    { id: 'modify', label: 'Allow changes', type: 'check', value: false }],
  async run([f], o) {
    if (!o.pw) throw new Error('Type a password.'); if (o.pw !== o.pw2) throw new Error('The two passwords are different. Type them again.');
    const S = await secure(); const d = await S.PDFDocument.load(f.bytes, { updateMetadata: false });
    const owner = o.pw + '#' + Math.random().toString(36).slice(2, 10);
    d.encrypt({ userPassword: o.pw, ownerPassword: owner, permissions: { printing: o.print ? 'highResolution' : false, copying: !!o.copy, modifying: !!o.modify, annotating: !!o.modify, fillingForms: !!o.modify, contentAccessibility: true, documentAssembly: !!o.modify } });
    return [out(`${baseName(f.name)}_protected.pdf`, pdfBlob(await d.save()))];
  },
});
T('unlock', {
  name: 'Unlock', group: 'security', icon: 'unlock', desc: 'Remove the password and restrictions from a PDF you are allowed to open.',
  opts: [{ id: 'pw', label: 'Password (leave empty if it opens without one)', type: 'password', value: '' }],
  async run([f], o) {
    const S = await secure(); let src;
    try { src = await S.PDFDocument.load(f.bytes, { password: o.pw || '', updateMetadata: false }); }
    catch (e) { if (/password/i.test(e.message)) throw new Error(o.pw ? 'That password is wrong.' : 'This PDF needs its password. Type it above.'); throw e; }
    const d = await S.PDFDocument.create(); (await d.copyPages(src, src.getPageIndices())).forEach((p) => d.addPage(p));
    for (const k of ['Title', 'Author', 'Subject']) { const v = src[`get${k}`]?.(); if (v) d[`set${k}`](v); }
    return [out(`${baseName(f.name)}_unlocked.pdf`, pdfBlob(await d.save({ useObjectStreams: true })))];
  },
});
T('watermark', {
  name: 'Watermark', group: 'security', icon: 'stamp', multi: true, desc: 'Write words like CONFIDENTIAL or COPY across the pages.',
  opts: [{ id: 'text', label: 'Text', type: 'text', value: 'CONFIDENTIAL', required: true },
    { id: 'layout', label: 'Where', type: 'select', value: 'diag', options: [['diag', 'Across the middle, slanted'], ['flat', 'Across the middle, straight'], ['tile', 'Repeated over the page'], ['top', 'Top of the page'], ['bottom', 'Bottom of the page']] },
    { id: 'size', label: 'Size (pt)', type: 'number', value: 60, min: 6, max: 300 }, { id: 'color', label: 'Colour', type: 'color', value: '#DC2626' },
    { id: 'opacity', label: 'See-through (%)', type: 'range', value: 75, min: 0, max: 95 }, fontOpt, pagesOpt()],
  async run(files, o) {
    const outs = [];
    for (const f of files) {
      const d = await libLoad(f.bytes); const font = await fontOf(d, o.font, true); const text = safeText(o.text); const size = +o.size;
      const tw = font.widthOfTextAtSize(text, size); const base = { size, font, color: hexRgb(o.color), opacity: 1 - o.opacity / 100 };
      for (const i of pageList(o.pages, d.getPageCount())) {
        const p = d.getPage(i); const { W, H } = vis(p);
        if (o.layout === 'tile') {
          const a = 35 * Math.PI / 180, stepX = tw * 0.9 + size * 2, stepY = size * 3.2;
          for (let y = -H; y < H * 1.6; y += stepY) for (let x = -W; x < W * 1.4; x += stepX) drawSeen(p, text, x + (y / stepY % 2) * stepX / 2, y, { ...base, angle: 35 });
          void a;
        } else if (o.layout === 'top' || o.layout === 'bottom') {
          drawSeen(p, text, (W - tw) / 2, o.layout === 'top' ? H - size - 18 : 18, base);
        } else {
          const ang = o.layout === 'diag' ? Math.atan2(H, W) : 0; const c = Math.cos(ang), s = Math.sin(ang);
          const fit = Math.min(1, (o.layout === 'diag' ? Math.hypot(W, H) * 0.8 : W * 0.9) / tw); const sz = size * fit, w2 = tw * fit;
          drawSeen(p, text, W / 2 - (w2 / 2) * c + (sz * 0.35) * s, H / 2 - (w2 / 2) * s - (sz * 0.35) * c, { ...base, size: sz, angle: ang * 180 / Math.PI });
        }
      }
      outs.push(await pdfOut(d, `${baseName(f.name)}_watermarked.pdf`));
    }
    return outs;
  },
});
T('flatten', {
  name: 'Flatten', group: 'security', icon: 'layers', desc: 'Lock form answers and comments into the page so they cannot be changed.',
  opts: [{ id: 'mode', label: 'What to flatten', type: 'select', value: 'forms', options: [['forms', 'Form boxes only (text stays text)'], ['all', 'Everything, including comments and drawings (pages become pictures)']] }],
  async run([f], o, ctx) {
    if (o.mode === 'all') return [await pdfOut(await rasterPdf(f.bytes, { dpi: 200, quality: 0.85 }, ctx.progress), `${baseName(f.name)}_flat.pdf`)];
    const d = await libLoad(f.bytes); const form = d.getForm(); const n = form.getFields().length;
    if (!n) ctx.note('This PDF has no form boxes. To lock comments and drawings, choose "Everything".');
    try { form.flatten(); } catch (e) { form.flatten({ updateFieldAppearances: false }); }
    return [await pdfOut(d, `${baseName(f.name)}_flat.pdf`)];
  },
});

// =============================================================== CONVERT FROM PDF
T('pdf2jpg', {
  name: 'PDF to JPG', group: 'from', icon: 'image', desc: 'Save each page as a picture.',
  opts: [{ id: 'fmt', label: 'Picture type', type: 'select', value: 'jpg', options: [['jpg', 'JPG (smaller)'], ['png', 'PNG (sharpest)']] },
    { id: 'dpi', label: 'Quality', type: 'select', value: '150', options: [['96', 'Screen (96 dpi)'], ['150', 'Good (150 dpi)'], ['200', 'High (200 dpi)'], ['300', 'Print (300 dpi)']] }, pagesOpt()],
  async run([f], o, ctx) {
    const js = await jsLoad(f.bytes); const ix = pageList(o.pages, js.numPages); const outs = []; const png = o.fmt === 'png';
    for (const [k, i] of ix.entries()) {
      ctx.progress(`Page ${k + 1} of ${ix.length}…`);
      const { canvas } = await renderPage(js, i, +o.dpi / 72);
      outs.push(out(`${baseName(f.name)}_page${String(i + 1).padStart(String(js.numPages).length, '0')}.${png ? 'png' : 'jpg'}`, await canvasBlob(canvas, png ? 'image/png' : 'image/jpeg', 0.9)));
      canvas.width = 0; await tick();
    }
    await js.destroy(); return outs;
  },
});
T('pdf2text', {
  name: 'PDF to Text', group: 'from', icon: 'text', desc: 'Copy all the words out into a plain text file.',
  opts: [{ id: 'marks', label: 'Mark where each page starts', type: 'check', value: true }],
  async run([f], o, ctx) {
    const js = await jsLoad(f.bytes); const parts = []; let words = 0;
    for (let i = 0; i < js.numPages; i++) {
      ctx.progress(`Reading page ${i + 1} of ${js.numPages}…`);
      const t = (await pageLines(js, i)).lines.map(lineText).join('\r\n'); words += t.split(/\s+/).filter(Boolean).length;
      parts.push((o.marks ? `----- Page ${i + 1} -----\r\n` : '') + t);
    }
    await js.destroy();
    if (!words) ctx.note('No text was found. This looks like a scan: run OCR first, then try again.');
    return [out(`${baseName(f.name)}.txt`, new Blob(['﻿' + parts.join('\r\n\r\n')], { type: 'text/plain;charset=utf-8' }))];
  },
});
T('pdf2excel', {
  name: 'PDF to Excel', group: 'from', icon: 'table', desc: 'Put the tables and text of each page into Excel rows and columns.',
  opts: [{ id: 'sheets', label: 'Sheets', type: 'select', value: 'each', options: [['each', 'One sheet per page'], ['one', 'All pages on one sheet']] }],
  async run([f], o, ctx) {
    const js = await jsLoad(f.bytes); const sheets = []; let all = [];
    for (let i = 0; i < js.numPages; i++) {
      ctx.progress(`Reading page ${i + 1} of ${js.numPages}…`);
      const rows = tableRows(await pageLines(js, i));
      if (o.sheets === 'each') sheets.push({ name: `Page ${i + 1}`, rows }); else all = all.concat(i ? [[], [`Page ${i + 1}`]] : [[`Page 1`]], rows);
    }
    await js.destroy();
    if (o.sheets !== 'each') sheets.push({ name: 'All pages', rows: all });
    if (!sheets.some((s) => s.rows.length > 1)) ctx.note('Very little text was found. If this is a scan, run OCR first.');
    return [out(`${baseName(f.name)}.xlsx`, await xlsxBlob(sheets))];
  },
});
/** Lines → rows of cells. Words far apart become separate cells; cells line up into shared columns. */
function tableRows({ lines }) {
  const rowsCells = lines.map((L1) => {
    const cells = []; let cur = null;
    const parts = [];
    for (const it of L1.items) {
      if (!/\S\s{2,}\S/.test(it.s)) { parts.push(it); continue; }
      const cw = it.w / Math.max(1, it.s.length); let k = 0;
      for (const m of it.s.matchAll(/\S+(?: \S+)*/g)) { parts.push({ ...it, x: it.x + m.index * cw, w: m[0].length * cw, s: m[0] }); k++; }
    }
    for (const it of parts) {
      if (cur && it.x - (cur.x1) < Math.max(it.size * 1.2, 6)) { cur.s += (it.x - cur.x1 > it.size * 0.15 ? ' ' : '') + it.s; cur.x1 = it.x + it.w; }
      else { cur = { x0: it.x, x1: it.x + it.w, s: it.s }; cells.push(cur); }
    }
    return cells;
  });
  const starts = rowsCells.flat().map((c) => c.x0).sort((a, b) => a - b); const cols = [];
  for (const x of starts) { const c = cols[cols.length - 1]; if (c && x - c.max < 9) { c.max = x; } else cols.push({ min: x, max: x }); }
  return rowsCells.map((cells) => {
    const row = []; for (const c of cells) {
      let k = cols.findIndex((col) => c.x0 >= col.min - 1 && c.x0 <= col.max + 1); if (k < 0) k = row.length;
      while (row[k] !== undefined) k++; row[k] = c.s.trim();
    }
    for (let k = 0; k < row.length; k++) if (row[k] === undefined) row[k] = '';
    return row;
  });
}
T('pdf2ppt', {
  name: 'PDF to PowerPoint', group: 'from', icon: 'slides', desc: 'Each page becomes a slide, ready to present.',
  opts: [{ id: 'dpi', label: 'Picture quality', type: 'select', value: '150', options: [['110', 'Normal'], ['150', 'Good'], ['220', 'High']] }],
  async run([f], o, ctx) {
    const js = await jsLoad(f.bytes); const slides = [];
    for (let i = 0; i < js.numPages; i++) {
      ctx.progress(`Page ${i + 1} of ${js.numPages}…`);
      const { canvas, w, h } = await renderPage(js, i, +o.dpi / 72); slides.push({ jpg: await canvasBytes(canvas, 'image/jpeg', 0.85), w, h }); canvas.width = 0; await tick();
    }
    await js.destroy();
    ctx.note('Each slide is a picture of the page. Add text boxes in PowerPoint to change words.');
    return [out(`${baseName(f.name)}.pptx`, await pptxBlob(slides))];
  },
});

// =============================================================== CONVERT TO PDF
T('jpg2pdf', {
  name: 'JPG to PDF', group: 'to', icon: 'image', accept: 'image', multi: true, desc: 'Put photos and pictures into one PDF, one per page.',
  opts: [{ id: 'size', label: 'Page size', type: 'select', value: 'fit', options: [['fit', 'Same as each picture'], ...SIZE_OPTS] },
    { id: 'orient', label: 'Page direction', type: 'select', value: 'auto', options: [['auto', 'Match each picture'], ['p', 'Portrait'], ['l', 'Landscape']], show: (o) => o.size !== 'fit' },
    { id: 'margin', label: 'Margin', type: 'select', value: '0', options: [['0', 'None'], ['10', 'Small (10 mm)'], ['20', 'Large (20 mm)']] }],
  async run(files, o, ctx) {
    const d = await PDFDocument.create();
    for (const [k, f] of files.entries()) {
      ctx.progress(`Picture ${k + 1} of ${files.length}…`);
      const bmp = await createImageBitmap(new Blob([f.bytes], { type: f.type || 'image/jpeg' }));
      const sc = Math.min(1, 3000 / Math.max(bmp.width, bmp.height)); const c = document.createElement('canvas');
      c.width = Math.round(bmp.width * sc); c.height = Math.round(bmp.height * sc); const g = c.getContext('2d'); g.fillStyle = '#fff'; g.fillRect(0, 0, c.width, c.height); g.drawImage(bmp, 0, 0, c.width, c.height);
      const png = /png/i.test(f.type || f.name); const img = png ? await d.embedPng(await canvasBytes(c, 'image/png')) : await d.embedJpg(await canvasBytes(c, 'image/jpeg', 0.9));
      const m = +o.margin * MM; let pw, ph;
      if (o.size === 'fit') { pw = 595.28; ph = pw * c.height / c.width; pw += 2 * m; ph += 2 * m; }
      else { [pw, ph] = SIZE[o.size]; const land = o.orient === 'l' || (o.orient === 'auto' && c.width > c.height); if (land) [pw, ph] = [ph, pw]; }
      const s = Math.min((pw - 2 * m) / c.width, (ph - 2 * m) / c.height);
      d.addPage([pw, ph]).drawImage(img, { x: (pw - c.width * s) / 2, y: (ph - c.height * s) / 2, width: c.width * s, height: c.height * s });
      bmp.close?.(); c.width = 0;
    }
    return [await pdfOut(d, `${baseName(files[0].name)}${files.length > 1 ? '_and_more' : ''}.pdf`)];
  },
});
T('html2pdf', {
  name: 'HTML to PDF', group: 'to', icon: 'code', accept: 'html', desc: 'Turn a saved web page (.html) into a PDF.',
  async run([f], o, ctx) {
    const html = new TextDecoder().decode(f.bytes);
    const fr = document.createElement('iframe'); fr.style.cssText = 'position:fixed;right:0;bottom:0;width:0;height:0;border:0';
    fr.srcdoc = html; document.body.appendChild(fr);
    await new Promise((r) => { fr.onload = r; setTimeout(r, 4000); });
    try { fr.contentDocument.title = baseName(f.name); } catch (e) { /* ignore */ }
    fr.contentWindow.focus(); fr.contentWindow.print();
    setTimeout(() => fr.remove(), 60000);
    ctx.note('In the print window, choose "Save as PDF" (or "Microsoft Print to PDF") as the printer, then Save. Pictures linked from the internet only show when you are online.');
    return [];
  },
});

// =============================================================== OTHER
T('pagenum', {
  name: 'Page numbers', group: 'other', icon: 'hash', multi: true, desc: 'Add page numbers in the corner or centre of each page.',
  opts: [{ id: 'pos', label: 'Where', type: 'select', value: 'bc', options: POS9 },
    { id: 'fmt', label: 'Style', type: 'select', value: '{n}', options: [['{n}', '1'], ['Page {n}', 'Page 1'], ['Page {n} of {total}', 'Page 1 of 9'], ['{n} / {total}', '1 / 9'], ['- {n} -', '- 1 -']] },
    { id: 'start', label: 'First number', type: 'number', value: 1, min: 0, max: 99999 }, { id: 'from', label: 'Start numbering on page', type: 'number', value: 1, min: 1, max: 9999, hint: 'Use 2 to leave a cover page without a number.' },
    { id: 'size', label: 'Size (pt)', type: 'number', value: 10, min: 5, max: 48 }, { id: 'margin', label: 'Distance from edge (mm)', type: 'number', value: 10, min: 2, max: 60 }, fontOpt],
  async run(files, o) {
    const outs = [];
    for (const f of files) {
      const d = await libLoad(f.bytes); const font = await fontOf(d, o.font); const n = d.getPageCount(); const from = Math.max(1, +o.from) - 1; const total = n - from + (+o.start) - 1;
      d.getPages().forEach((p, i) => {
        if (i < from) return; const t = safeText(o.fmt.replace('{n}', i - from + (+o.start)).replace('{total}', total));
        const { W, H } = vis(p); const tw = font.widthOfTextAtSize(t, +o.size); const at = placeAt(o.pos, W, H, tw, +o.size, +o.margin * MM);
        drawSeen(p, t, at.x, at.y, { size: +o.size, font, color: rgb(0, 0, 0) });
      });
      outs.push(await pdfOut(d, `${baseName(f.name)}_numbered.pdf`));
    }
    return outs;
  },
});
T('headfoot', {
  name: 'Header & footer', group: 'other', icon: 'header', multi: true, desc: 'Add a line of text at the top or bottom of every page, with page numbers and date.',
  opts: [{ id: 'hl', label: 'Header left', type: 'text', value: '' }, { id: 'hc', label: 'Header centre', type: 'text', value: '' }, { id: 'hr', label: 'Header right', type: 'text', value: '' },
    { id: 'fl', label: 'Footer left', type: 'text', value: '{file}' }, { id: 'fc', label: 'Footer centre', type: 'text', value: '' }, { id: 'fr', label: 'Footer right', type: 'text', value: 'Page {n} of {total}' },
    { id: 'info', type: 'info', label: 'You can use {n} page number, {total} number of pages, {date} today, {file} file name.' },
    { id: 'size', label: 'Size (pt)', type: 'number', value: 9, min: 5, max: 36 }, { id: 'margin', label: 'Distance from edge (mm)', type: 'number', value: 10, min: 2, max: 50 },
    { id: 'color', label: 'Colour', type: 'color', value: '#333333' }, fontOpt, { id: 'skip', label: 'Leave the first page without header and footer', type: 'check', value: false }],
  async run(files, o) {
    const outs = []; const date = new Date().toLocaleDateString('en-GB', { day: 'numeric', month: 'short', year: 'numeric' });
    for (const f of files) {
      const d = await libLoad(f.bytes); const font = await fontOf(d, o.font); const n = d.getPageCount(); const size = +o.size, m = +o.margin * MM;
      d.getPages().forEach((p, i) => {
        if (o.skip && i === 0) return; const { W, H } = vis(p);
        for (const [key, pos] of [['hl', 'tl'], ['hc', 'tc'], ['hr', 'tr'], ['fl', 'bl'], ['fc', 'bc'], ['fr', 'br']]) {
          if (!o[key]?.trim()) continue;
          const t = safeText(o[key].replace(/\{n\}/g, i + 1).replace(/\{total\}/g, n).replace(/\{date\}/g, date).replace(/\{file\}/g, baseName(f.name)));
          const tw = font.widthOfTextAtSize(t, size); const at = placeAt(pos, W, H, tw, size, m);
          drawSeen(p, t, at.x, at.y, { size, font, color: hexRgb(o.color) });
        }
      });
      outs.push(await pdfOut(d, `${baseName(f.name)}_header_footer.pdf`));
    }
    return outs;
  },
});
T('bates', {
  name: 'Bates numbering', group: 'other', icon: 'bates', multi: true, desc: 'Stamp every page with a running reference like CASE-000001, across several files.',
  opts: [{ id: 'prefix', label: 'Before the number', type: 'text', value: 'DOC-' }, { id: 'start', label: 'First number', type: 'number', value: 1, min: 0, max: 99999999 },
    { id: 'digits', label: 'Digits', type: 'number', value: 6, min: 1, max: 12 }, { id: 'suffix', label: 'After the number', type: 'text', value: '' },
    { id: 'pos', label: 'Where', type: 'select', value: 'br', options: POS9 }, { id: 'size', label: 'Size (pt)', type: 'number', value: 9, min: 5, max: 36 },
    { id: 'color', label: 'Colour', type: 'color', value: '#000000' }],
  async run(files, o, ctx) {
    let k = +o.start; const outs = [];
    for (const f of files) {
      const d = await libLoad(f.bytes); const font = await fontOf(d, 'sans', true); const first = k;
      for (const p of d.getPages()) {
        const t = safeText(`${o.prefix}${String(k++).padStart(+o.digits, '0')}${o.suffix}`); const { W, H } = vis(p);
        const tw = font.widthOfTextAtSize(t, +o.size); const at = placeAt(o.pos, W, H, tw, +o.size, 10 * MM);
        drawSeen(p, t, at.x, at.y, { size: +o.size, font, color: hexRgb(o.color) });
      }
      ctx.note(`${f.name}: ${o.prefix}${String(first).padStart(+o.digits, '0')}${o.suffix} to ${o.prefix}${String(k - 1).padStart(+o.digits, '0')}${o.suffix}`);
      outs.push(await pdfOut(d, `${baseName(f.name)}_bates.pdf`));
    }
    return outs;
  },
});
T('bookmarks', {
  name: 'Create bookmarks', group: 'other', icon: 'bookmark', desc: 'Add a clickable table of contents in the side panel of PDF readers.',
  async fields([f]) {
    const js = await jsLoad(f.bytes); const lines = [];
    for (let i = 0; i < Math.min(js.numPages, 200); i++) { const t = (await pageLines(js, i)).lines.map(lineText).find((s) => s.length > 2) || `Page ${i + 1}`; lines.push(`${t.slice(0, 70).replace(/\|/g, '/')} | ${i + 1}`); }
    await js.destroy();
    return [{ id: 'list', label: 'Bookmarks: one per line, as  Title | page', type: 'textarea', rows: 12, value: lines.join('\n'), hint: 'Delete the lines you do not want and change the titles. Each title opens its page.' },
      { id: 'open', label: 'Show the bookmarks panel when the PDF opens', type: 'check', value: true }];
  },
  async run([f], o) {
    const d = await libLoad(f.bytes); const pages = d.getPages(); const items = [];
    for (const line of (o.list || '').split('\n')) {
      if (!line.trim()) continue; const m = line.match(/^(.*?)\s*\|\s*(\d+)\s*$/);
      if (!m) throw new Error(`"${line}" needs a page number after a | sign, like:  Introduction | 3`);
      const pg = +m[2]; if (pg < 1 || pg > pages.length) throw new Error(`Page ${pg} does not exist (the PDF has ${pages.length} pages).`);
      items.push({ title: m[1].trim() || `Page ${pg}`, page: pg - 1 });
    }
    if (!items.length) throw new Error('Add at least one bookmark.');
    const ctx = d.context; const root = ctx.nextRef(); const refs = items.map(() => ctx.nextRef());
    items.forEach((it, i) => {
      const dict = ctx.obj({ Title: PDFHexString.fromText(it.title), Parent: root, Dest: [pages[it.page].ref, 'XYZ', null, null, null] });
      if (i > 0) dict.set(PDFName.of('Prev'), refs[i - 1]); if (i < items.length - 1) dict.set(PDFName.of('Next'), refs[i + 1]);
      ctx.assign(refs[i], dict);
    });
    ctx.assign(root, ctx.obj({ Type: 'Outlines', First: refs[0], Last: refs[refs.length - 1], Count: items.length }));
    d.catalog.set(PDFName.of('Outlines'), root); if (o.open) d.catalog.set(PDFName.of('PageMode'), PDFName.of('UseOutlines'));
    return [await pdfOut(d, `${baseName(f.name)}_bookmarked.pdf`)];
  },
});
T('crop', {
  name: 'Crop', group: 'other', icon: 'crop', desc: 'Cut away margins: automatically, or by the amount you choose.',
  opts: [{ id: 'mode', label: 'How', type: 'select', value: 'auto', options: [['auto', 'Automatically trim empty white margins'], ['mm', 'Cut the margins I type']] },
    { id: 'pad', label: 'Keep a border of (mm)', type: 'number', value: 3, min: 0, max: 30, show: (o) => o.mode === 'auto' },
    { id: 't', label: 'Top (mm)', type: 'number', value: 10, min: 0, max: 300, show: (o) => o.mode === 'mm' }, { id: 'b', label: 'Bottom (mm)', type: 'number', value: 10, min: 0, max: 300, show: (o) => o.mode === 'mm' },
    { id: 'l', label: 'Left (mm)', type: 'number', value: 10, min: 0, max: 300, show: (o) => o.mode === 'mm' }, { id: 'r', label: 'Right (mm)', type: 'number', value: 10, min: 0, max: 300, show: (o) => o.mode === 'mm' }, pagesOpt()],
  async run([f], o, ctx) {
    const d = await libLoad(f.bytes); const js = o.mode === 'auto' ? await jsLoad(f.bytes) : null; const ix = pageList(o.pages, d.getPageCount());
    for (const [k, i] of ix.entries()) {
      const p = d.getPage(i); const { W, H } = vis(p); let x0, y0, x1, y1;
      if (o.mode === 'mm') { x0 = o.l * MM; x1 = W - o.r * MM; y0 = o.b * MM; y1 = H - o.t * MM; }
      else {
        ctx.progress(`Finding the edges of page ${k + 1} of ${ix.length}…`);
        const sc = 1.2; const { canvas } = await renderPage(js, i, sc, { read: true }); const box = inkBox(canvas); canvas.width = 0;
        if (!box) continue; const pad = o.pad * MM;
        x0 = box.x0 / sc - pad; x1 = box.x1 / sc + pad; y1 = H - box.y0 / sc + pad; y0 = H - box.y1 / sc - pad;
      }
      x0 = Math.max(0, x0); y0 = Math.max(0, y0); x1 = Math.min(W, x1); y1 = Math.min(H, y1);
      if (x1 - x0 < 20 || y1 - y0 < 20) throw new Error(`Those margins leave nothing of page ${i + 1}. Use smaller numbers.`);
      const u = rectToUser(p, x0, y0, x1, y1); p.setCropBox(u.x, u.y, u.w, u.h); p.setMediaBox(u.x, u.y, u.w, u.h);
    }
    await js?.destroy();
    return [await pdfOut(d, `${baseName(f.name)}_cropped.pdf`)];
  },
});
function inkBox(c) {
  const g = c.getContext('2d'); const { data, width: w, height: h } = g.getImageData(0, 0, c.width, c.height);
  let x0 = w, y0 = h, x1 = -1, y1 = -1;
  for (let y = 0; y < h; y += 1) for (let x = 0; x < w; x += 1) { const k = (y * w + x) * 4; if (data[k] < 235 || data[k + 1] < 235 || data[k + 2] < 235) { if (x < x0) x0 = x; if (x > x1) x1 = x; if (y < y0) y0 = y; if (y > y1) y1 = y; } }
  return x1 < 0 ? null : { x0, y0, x1: x1 + 1, y1: y1 + 1 };
}
T('meta', {
  name: 'Edit metadata', group: 'other', icon: 'info', desc: 'Change the title, author, subject and keywords stored in the PDF.',
  async fields([f]) {
    const d = await libLoad(f.bytes);
    return [['title', 'Title', d.getTitle()], ['author', 'Author', d.getAuthor()], ['subject', 'Subject', d.getSubject()], ['keywords', 'Keywords (separate with commas)', d.getKeywords()],
      ['creator', 'Made with (creator)', d.getCreator()], ['producer', 'PDF producer', d.getProducer()]].map(([id, label, value]) => ({ id, label, type: 'text', value: value || '' }));
  },
  async run([f], o) {
    const d = await libLoad(f.bytes);
    d.setTitle(o.title || ''); d.setAuthor(o.author || ''); d.setSubject(o.subject || ''); d.setKeywords((o.keywords || '').split(',').map((s) => s.trim()).filter(Boolean));
    d.setCreator(o.creator || ''); d.setProducer(o.producer || ''); d.setModificationDate(new Date());
    return [await pdfOut(d, f.name)];
  },
});
T('images', {
  name: 'Extract images', group: 'other', icon: 'images', desc: 'Save every photo and picture inside the PDF as its own file.',
  opts: [{ id: 'min', label: 'Skip pictures smaller than (pixels)', type: 'number', value: 40, min: 1, max: 2000 }],
  async run([f], o, ctx) {
    const js = await jsLoad(f.bytes); const outs = []; const seen = new Set(); const OPS = pdfjsLib.OPS;
    for (let i = 0; i < js.numPages; i++) {
      ctx.progress(`Looking at page ${i + 1} of ${js.numPages}…`);
      const page = await js.getPage(i + 1); const ops = await page.getOperatorList(); let k = 0;
      for (let j = 0; j < ops.fnArray.length; j++) {
        const fn = ops.fnArray[j]; if (fn !== OPS.paintImageXObject && fn !== OPS.paintInlineImageXObject && fn !== OPS.paintImageXObjectRepeat) continue;
        const a = ops.argsArray[j][0]; let img = a;
        if (typeof a === 'string') { if (seen.has(a)) continue; seen.add(a); img = await getObj(a.startsWith('g_') ? page.commonObjs : page.objs, a); }
        if (!img || img.width < +o.min || img.height < +o.min) continue;
        const c = imgCanvas(img); if (!c) continue; k++;
        outs.push(out(`${baseName(f.name)}_p${i + 1}_${k}.png`, await canvasBlob(c, 'image/png'))); c.width = 0;
      }
      page.cleanup();
    }
    await js.destroy();
    if (!outs.length) throw new Error('No pictures were found inside this PDF. To save whole pages as pictures, use PDF to JPG.');
    return outs;
  },
});
const getObj = (store, name) => new Promise((res) => { try { store.get(name, res); } catch (e) { res(null); } setTimeout(() => res(null), 5000); });
function imgCanvas(img) {
  const c = document.createElement('canvas'); c.width = img.width; c.height = img.height; const g = c.getContext('2d');
  if (img.bitmap) { g.drawImage(img.bitmap, 0, 0); return c; }
  if (!img.data) return null;
  const id = g.createImageData(img.width, img.height); const d = id.data, s = img.data; const n = img.width * img.height;
  if (img.kind === 3) d.set(s.subarray(0, d.length));
  else if (img.kind === 2) for (let i = 0, j = 0; i < n; i++, j += 3) { d[i * 4] = s[j]; d[i * 4 + 1] = s[j + 1]; d[i * 4 + 2] = s[j + 2]; d[i * 4 + 3] = 255; }
  else if (img.kind === 1) { const row = Math.ceil(img.width / 8); for (let y = 0; y < img.height; y++) for (let x = 0; x < img.width; x++) { const bit = (s[y * row + (x >> 3)] >> (7 - (x & 7))) & 1; const v = bit ? 255 : 0; const i = (y * img.width + x) * 4; d[i] = d[i + 1] = d[i + 2] = v; d[i + 3] = 255; } }
  else return null;
  g.putImageData(id, 0, 0); return c;
}
T('flip', {
  name: 'Flip', group: 'other', icon: 'flip', desc: 'Mirror pages left-to-right or upside down, for scans made the wrong way.',
  opts: [{ id: 'dir', label: 'Flip', type: 'select', value: 'h', options: [['h', 'Left to right (mirror)'], ['v', 'Upside down']] }, pagesOpt()],
  async run([f], o) {
    const src = await upright(await libLoad(f.bytes)); const res = await PDFDocument.create(); const sel = new Set(pageList(o.pages, src.getPageCount()));
    for (const [i, p] of src.getPages().entries()) {
      const { width: w, height: h } = p.getSize(); const e = await res.embedPage(p); const np = res.addPage([w, h]);
      if (!sel.has(i)) np.drawPage(e, { x: 0, y: 0 }); else if (o.dir === 'h') np.drawPage(e, { x: w, y: 0, xScale: -1, yScale: 1 }); else np.drawPage(e, { x: 0, y: h, xScale: 1, yScale: -1 });
    }
    return [await pdfOut(res, `${baseName(f.name)}_flipped.pdf`)];
  },
});
T('gray', {
  name: 'Grayscale', group: 'other', icon: 'contrast', desc: 'Turn colour pages into black and white for cheaper printing.',
  opts: [{ id: 'dpi', label: 'Quality', type: 'select', value: '150', options: [['110', 'Normal (smaller file)'], ['150', 'Good'], ['220', 'High (for printing)']] }],
  async run([f], o, ctx) {
    ctx.note('Pages become black-and-white pictures, so words can no longer be selected. Run OCR afterwards if you need that.');
    return [await pdfOut(await rasterPdf(f.bytes, { dpi: +o.dpi, quality: 0.8, gray: true }, ctx.progress), `${baseName(f.name)}_grayscale.pdf`)];
  },
});
T('nup', {
  name: 'N-up', group: 'other', icon: 'nup', desc: 'Print several pages on one sheet to save paper.',
  opts: [{ id: 'n', label: 'Pages per sheet', type: 'select', value: '2', options: [['2', '2'], ['4', '4'], ['6', '6'], ['8', '8'], ['9', '9'], ['16', '16']] },
    { id: 'size', label: 'Sheet size', type: 'select', value: 'a4', options: SIZE_OPTS }, { id: 'border', label: 'Draw a thin line around each page', type: 'check', value: true }],
  async run([f], o) {
    const src = await libLoad(f.bytes); const pages = src.getPages(); const n = +o.n; const res = await PDFDocument.create();
    const v0 = vis(pages[0]); let best = null;
    for (const land of [false, true]) for (let cols = 1; cols <= n; cols++) {
      if (n % cols) continue; const rows = n / cols; let [W, H] = SIZE[o.size]; if (land) [W, H] = [H, W];
      const s = Math.min((W - 30) / cols / v0.W, (H - 30) / rows / v0.H); if (!best || s > best.s) best = { s, cols, rows, W, H };
    }
    const { cols, rows, W, H } = best; const cw = (W - 30) / cols, ch = (H - 30) / rows;
    for (let i = 0; i < pages.length; i += n) {
      const sheet = res.addPage([W, H]);
      for (let k = 0; k < n && i + k < pages.length; k++) {
        const c = k % cols, r = Math.floor(k / cols); const box = { x: 15 + c * cw + 4, y: H - 15 - (r + 1) * ch + 4, w: cw - 8, h: ch - 8 };
        const s = await placePage(res, sheet, pages[i + k], box);
        if (o.border) { const { W: pw, H: ph } = vis(pages[i + k]); sheet.drawRectangle({ x: box.x + (box.w - pw * s) / 2, y: box.y + (box.h - ph * s) / 2, width: pw * s, height: ph * s, borderColor: rgb(0.7, 0.7, 0.7), borderWidth: 0.5 }); }
      }
    }
    return [await pdfOut(res, `${baseName(f.name)}_${n}-up.pdf`)];
  },
});
T('rename', {
  name: 'Rename', group: 'other', icon: 'rename', desc: 'Save a copy of a file with a new name.', accept: 'any',
  async fields([f]) { return [{ id: 'name', label: 'New name', type: 'text', value: baseName(f.name), required: true }]; },
  async run([f], o) {
    const ext = (f.name.match(/\.[^.]+$/) || [''])[0]; const n = (o.name || '').replace(/[\\/:*?"<>|]+/g, ' ').trim();
    if (!n) throw new Error('Type a new name.');
    return [out(n.toLowerCase().endsWith(ext.toLowerCase()) ? n : n + ext, new Blob([f.bytes], { type: f.type }))];
  },
});
T('repair', {
  name: 'Repair', group: 'other', icon: 'wrench', desc: 'Rescue a damaged PDF that will not open or shows errors.',
  async run([f], o, ctx) {
    try {
      const src = await PDFDocument.load(f.bytes, { ignoreEncryption: true, throwOnInvalidObject: false, updateMetadata: false });
      const d = await PDFDocument.create(); (await d.copyPages(src, src.getPageIndices())).forEach((p) => d.addPage(p));
      if (!d.getPageCount()) throw new Error('no pages');
      ctx.note(`Rebuilt the file structure: ${d.getPageCount()} page${d.getPageCount() > 1 ? 's' : ''} kept, text stays text.`);
      return [await pdfOut(d, `${baseName(f.name)}_repaired.pdf`)];
    } catch (e) {
      ctx.note('The structure was too damaged to rebuild, so the pages were saved as pictures of what could be read.');
      return [await pdfOut(await rasterPdf(f.bytes, { dpi: 150, quality: 0.85 }, ctx.progress), `${baseName(f.name)}_repaired.pdf`)];
    }
  },
});
T('resize', {
  name: 'Resize', group: 'other', icon: 'resize', desc: 'Change the paper size, for example Letter to A4, keeping everything in proportion.',
  opts: [{ id: 'size', label: 'New size', type: 'select', value: 'a4', options: SIZE_OPTS },
    { id: 'orient', label: 'Direction', type: 'select', value: 'auto', options: [['auto', 'Keep each page’s direction'], ['p', 'Portrait'], ['l', 'Landscape']] },
    { id: 'margin', label: 'Margin (mm)', type: 'number', value: 0, min: 0, max: 50 }],
  async run([f], o) {
    const src = await libLoad(f.bytes); const res = await PDFDocument.create(); const m = +o.margin * MM;
    for (const p of src.getPages()) {
      let [W, H] = SIZE[o.size]; const v = vis(p); const land = o.orient === 'l' || (o.orient === 'auto' && v.W > v.H); if (land) [W, H] = [H, W];
      await placePage(res, res.addPage([W, H]), p, { x: m, y: m, w: W - 2 * m, h: H - 2 * m });
    }
    return [await pdfOut(res, `${baseName(f.name)}_${o.size}.pdf`)];
  },
});
T('unannot', {
  name: 'Remove annotations', group: 'other', icon: 'eraser', desc: 'Take out comments, highlights, sticky notes and drawings.',
  opts: [{ id: 'mode', label: 'Remove', type: 'select', value: 'notes', options: [['notes', 'Comments, highlights and drawings (keep links and form boxes)'], ['all', 'Everything, including links and form boxes']] }],
  async run([f], o, ctx) {
    const d = await libLoad(f.bytes); let removed = 0;
    for (const p of d.getPages()) {
      const arr = p.node.lookup(PDFName.of('Annots')); if (!(arr instanceof PDFArray)) continue; const keep = [];
      for (let i = 0; i < arr.size(); i++) {
        const a = arr.lookup(i, PDFDict); const st = a?.get(PDFName.of('Subtype'))?.toString();
        if (o.mode === 'notes' && (st === '/Link' || st === '/Widget')) keep.push(arr.get(i)); else removed++;
      }
      if (keep.length) p.node.set(PDFName.of('Annots'), d.context.obj(keep)); else p.node.delete(PDFName.of('Annots'));
    }
    if (o.mode === 'all') d.catalog.delete(PDFName.of('AcroForm'));
    ctx.note(removed ? `${removed} item${removed > 1 ? 's' : ''} removed.` : 'Nothing to remove: this PDF has no comments or drawings.');
    return [await pdfOut(d, `${baseName(f.name)}_clean.pdf`)];
  },
});

// =============================================================== SCANS
T('ocr', {
  name: 'OCR', group: 'scans', icon: 'ocr', desc: 'Make a scanned PDF searchable: words can be found, selected and copied.',
  opts: [{ id: 'dpi', label: 'Reading quality', type: 'select', value: '200', options: [['150', 'Fast'], ['200', 'Good'], ['300', 'Best (slow)']] }, pagesOpt()],
  async run([f], o, ctx) {
    const js = await jsLoad(f.bytes); const src = await libLoad(f.bytes); const res = await PDFDocument.create();
    const sel = new Set(pageList(o.pages, js.numPages)); ctx.progress('Getting the text reader ready…'); const w = await ocr(); let words = 0;
    for (let i = 0; i < js.numPages; i++) {
      if (!sel.has(i)) { const [c] = await res.copyPages(src, [i]); res.addPage(c); continue; }
      ctx.progress(`Reading page ${i + 1} of ${js.numPages}…`);
      const { canvas, w: pw, h: ph } = await renderPage(js, i, +o.dpi / 72);
      const { data } = await w.recognize(canvas, {}, { pdf: true, text: true });
      words += (data.text || '').split(/\s+/).filter(Boolean).length;
      const tp = await PDFDocument.load(new Uint8Array(data.pdf)); const np = res.addPage([pw, ph]);
      const e = await res.embedPage(tp.getPage(0)); np.drawPage(e, { x: 0, y: 0, width: pw, height: ph });
      canvas.width = 0; await tick();
    }
    await js.destroy(); ctx.note(`${words} words found. Use Ctrl+F in your PDF reader to search them.`);
    return [await pdfOut(res, `${baseName(f.name)}_searchable.pdf`)];
  },
});
T('deskew', {
  name: 'Deskew', group: 'scans', icon: 'deskew', desc: 'Straighten pages that were scanned or photographed at a slant.',
  opts: [{ id: 'dpi', label: 'Quality', type: 'select', value: '200', options: [['150', 'Normal'], ['200', 'Good'], ['300', 'High (slow)']] }],
  async run([f], o, ctx) {
    const js = await jsLoad(f.bytes); const res = await PDFDocument.create(); ctx.progress('Getting the page reader ready…'); const w = await ocr(); const fixes = [];
    for (let i = 0; i < js.numPages; i++) {
      ctx.progress(`Straightening page ${i + 1} of ${js.numPages}…`);
      const { canvas, w: pw, h: ph } = await renderPage(js, i, +o.dpi / 72);
      const small = document.createElement('canvas'); const k = Math.min(1, 1400 / Math.max(canvas.width, canvas.height)); small.width = canvas.width * k; small.height = canvas.height * k; small.getContext('2d').drawImage(canvas, 0, 0, small.width, small.height);
      let ang = 0; try { const { data } = await w.recognize(small, { rotateAuto: true }, { text: false }); ang = data.rotateRadians || 0; } catch (e) { ang = 0; }
      let c = canvas;
      if (Math.abs(ang) > 0.002 && Math.abs(ang) < 0.35) {
        c = document.createElement('canvas'); c.width = canvas.width; c.height = canvas.height; const g = c.getContext('2d');
        g.fillStyle = '#fff'; g.fillRect(0, 0, c.width, c.height); g.translate(c.width / 2, c.height / 2); g.rotate(ang); g.drawImage(canvas, -canvas.width / 2, -canvas.height / 2);
        fixes.push(`${i + 1} (${(ang * 180 / Math.PI).toFixed(1)}°)`);
      }
      const img = await res.embedJpg(await canvasBytes(c, 'image/jpeg', 0.85)); res.addPage([pw, ph]).drawImage(img, { x: 0, y: 0, width: pw, height: ph });
      canvas.width = 0; c.width = 0; small.width = 0; await tick();
    }
    await js.destroy();
    ctx.note(fixes.length ? `Straightened page${fixes.length > 1 ? 's' : ''} ${fixes.join(', ')}.` : 'All pages were already straight.');
    return [await pdfOut(res, `${baseName(f.name)}_straight.pdf`)];
  },
});

// workflows are not offered; everything else above
export const byId = Object.fromEntries(TOOLS.map((t) => [t.id, t]));
export { pageList };
