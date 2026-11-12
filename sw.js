// Service Worker：缓存全部页面文件，断网可用；不向任何外部地址发请求
const VERSION = 'v1.0.1-1';
const CACHE = 'kaoqin-' + VERSION;
const FILES = [
  './', './index.html', './styles.css', './manifest.webmanifest',
  './src/app.js', './src/db.js', './src/pin.js', './src/ui.js',
  './src/core/util.js', './src/core/jie.js', './src/core/dates.js', './src/core/store.js', './src/core/roster.js', './src/core/export.js', './src/core/backup.js', './src/core/pinyin.js',
  './vendor/xlsx.mjs', './icons/icon-192.png', './icons/icon-512.png', './icons/apple-touch-icon.png',
];
self.addEventListener('install', (e) => {
  e.waitUntil(caches.open(CACHE).then((c) => c.addAll(FILES.map((f) => new Request(f, { cache: 'reload' })))).then(() => self.skipWaiting()));
});
self.addEventListener('activate', (e) => {
  e.waitUntil(caches.keys().then((keys) => Promise.all(keys.filter((k) => k.startsWith('kaoqin-') && k !== CACHE).map((k) => caches.delete(k))))
    .then(() => self.clients.claim()));
});
self.addEventListener('fetch', (e) => {
  const url = new URL(e.request.url);
  if (e.request.method !== 'GET' || url.origin !== self.location.origin) return;
  e.respondWith(caches.match(e.request, { ignoreSearch: true }).then((hit) => hit || fetch(e.request).then((res) => {
    if (res.ok) { const copy = res.clone(); caches.open(CACHE).then((c) => c.put(e.request, copy)); }
    return res;
  }).catch(() => caches.match('./index.html'))));
});
