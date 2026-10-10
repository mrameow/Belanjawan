/* Service worker: simpan "cangkerang" app supaya terbuka laju & tanpa talian.
 * Panggilan API ke Google (silang-asal) tidak disentuh. Naikkan VERSION bila mengubah fail app. */
const VERSION = 'v1';
const CACHE = 'belanjawan-' + VERSION;
const SHELL = ['./', 'index.html', 'manifest.webmanifest', 'icons/icon-192.png', 'icons/icon-512.png', 'icons/icon-maskable-512.png'];

self.addEventListener('install', e => {
  e.waitUntil(caches.open(CACHE).then(c => c.addAll(SHELL)).then(() => self.skipWaiting()));
});

self.addEventListener('activate', e => {
  e.waitUntil(caches.keys().then(ks => Promise.all(ks.filter(k => k !== CACHE).map(k => caches.delete(k)))).then(() => self.clients.claim()));
});

self.addEventListener('fetch', e => {
  const req = e.request, url = new URL(req.url);
  if (req.method !== 'GET' || url.origin !== location.origin) return; // API & lain-lain: terus ke rangkaian
  // rangkaian dahulu (supaya kemas kini app sampai), jatuh ke simpanan jika tiada talian
  e.respondWith(fetch(req).then(res => {
    if (res.ok) { const copy = res.clone(); caches.open(CACHE).then(c => c.put(req, copy)); }
    return res;
  }).catch(() => caches.match(req).then(r => r || caches.match('./'))));
});
