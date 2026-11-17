// Keeps the editor working offline once it has been opened.
const CACHE = 'dudely-editor-pc-v5';
const FILES = ['./', 'index.html', 'app.js', 'tools.js', 'logo.png', 'manifest.webmanifest', 'icon-192.png', 'icon-512.png',
  'pdf.min.mjs', 'pdf.worker.min.mjs', 'pdf-lib.min.js', 'jszip.min.js', 'docx-preview.min.js', 'tesseract.min.js', 'worker.min.js',
  'tesseract-core-simd-lstm.wasm.js', 'tesseract-core-lstm.wasm.js', 'eng.traineddata.gz'];
self.addEventListener('install', (e) => { e.waitUntil(caches.open(CACHE).then((c) => c.addAll(FILES)).then(() => self.skipWaiting())); });
self.addEventListener('activate', (e) => { e.waitUntil(caches.keys().then((ks) => Promise.all(ks.filter((k) => k !== CACHE).map((k) => caches.delete(k)))).then(() => self.clients.claim())); });
self.addEventListener('fetch', (e) => {
  if (e.request.method !== 'GET') return;
  e.respondWith(caches.match(e.request, { ignoreSearch: true }).then((hit) => hit || fetch(e.request)));
});
