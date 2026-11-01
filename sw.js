// Cache per l'uso offline: rete prima per i file dell'app (così gli aggiornamenti arrivano subito), cache prima per i font
const C = 'neurosurg-clinic-v4';
const CORE = ['./', './index.html', './app.js', './fields.js', './xlsx.js', './manifest.webmanifest', './icon-512.png'];
self.addEventListener('install', e => { e.waitUntil(caches.open(C).then(c => c.addAll(CORE))); self.skipWaiting(); });
self.addEventListener('activate', e => {
  e.waitUntil(caches.keys().then(ks => Promise.all(ks.filter(k => k !== C).map(k => caches.delete(k)))));
  self.clients.claim();
});
self.addEventListener('fetch', e => {
  if (e.request.method !== 'GET') return;
  const put = res => { if (res.ok || res.type === 'opaque') { const cp = res.clone(); caches.open(C).then(c => c.put(e.request, cp)); } return res; };
  if (new URL(e.request.url).origin === location.origin) {
    e.respondWith(fetch(e.request).then(put).catch(() =>
      caches.match(e.request).then(r => r || (e.request.mode === 'navigate' ? caches.match('./index.html') : Response.error()))));
    return;
  }
  e.respondWith(caches.match(e.request).then(r => r || fetch(e.request).then(put)));
});
