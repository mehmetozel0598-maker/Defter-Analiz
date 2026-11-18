// Mali Tablo — çevrimdışı önbellek (diğer uygulamalarla çakışmayan benzersiz önbellek adı)
const CACHE = 'mali-tablo-yevmiye-bilanco-v2';
const FILES = ['./', './index.html', './style.css', './core.js', './app.js', './xlsx.full.min.js',
  './manifest.webmanifest', './icon-192.png', './icon-512.png', './icon-maskable-512.png'];
self.addEventListener('install', e => { e.waitUntil(caches.open(CACHE).then(c => c.addAll(FILES))); self.skipWaiting(); });
self.addEventListener('activate', e => {
  e.waitUntil(caches.keys().then(ks => Promise.all(ks.filter(k => k.startsWith('mali-tablo-yevmiye-bilanco-') && k !== CACHE).map(k => caches.delete(k)))));
  self.clients.claim();
});
self.addEventListener('fetch', e => {
  if (e.request.method !== 'GET') return;
  e.respondWith(caches.match(e.request, { ignoreSearch: true }).then(r => r || fetch(e.request)));
});
