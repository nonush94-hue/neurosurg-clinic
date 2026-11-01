// Cache per l'uso offline: rete prima per la pagina, cache prima per il resto
const C = 'neurosurg-exam-v1';
const CORE = ['./', './index.html', './manifest.webmanifest', './icon-512.png'];
self.addEventListener('install', e => { e.waitUntil(caches.open(C).then(c => c.addAll(CORE))); self.skipWaiting(); });
self.addEventListener('activate', e => {
  e.waitUntil(caches.keys().then(ks => Promise.all(ks.filter(k => k !== C).map(k => caches.delete(k)))));
  self.clients.claim();
});
self.addEventListener('fetch', e => {
  if (e.request.method !== 'GET') return;
  const put = res => { const cp = res.clone(); caches.open(C).then(c => c.put(e.request, cp)); return res; };
  if (e.request.mode === 'navigate') {
    e.respondWith(fetch(e.request).then(put).catch(() => caches.match('./index.html')));
    return;
  }
  e.respondWith(caches.match(e.request).then(r => r || fetch(e.request).then(put)));
});
