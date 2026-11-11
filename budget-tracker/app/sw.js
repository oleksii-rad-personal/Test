// Кеш оболонки застосунку: працює офлайн, оновлюється у фоні.
const CACHE = 'bt-shell-v3';
const SHELL = ['./', 'index.html', 'app.css', 'app.js', 'manifest.webmanifest', 'manifest.ios.webmanifest',
  'icons/icon-192.png', 'icons/icon-512.png', 'icons/apple-touch-icon.png'];

self.addEventListener('install', e => {
  e.waitUntil(caches.open(CACHE).then(c => c.addAll(SHELL)).then(() => self.skipWaiting()));
});
self.addEventListener('activate', e => {
  e.waitUntil(caches.keys().then(keys => Promise.all(keys.filter(k => k !== CACHE).map(k => caches.delete(k))))
    .then(() => self.clients.claim()));
});
self.addEventListener('fetch', e => {
  const url = new URL(e.request.url);
  if (e.request.method !== 'GET' || url.origin !== self.location.origin) return; // запити до таблиці — напряму
  e.respondWith(caches.open(CACHE).then(async cache => {
    const cached = await cache.match(e.request, { ignoreSearch: true });
    const fresh = fetch(e.request).then(r => { if (r.ok) cache.put(e.request.url.split('?')[0], r.clone()); return r; }).catch(() => cached);
    return cached || fresh;
  }));
});
