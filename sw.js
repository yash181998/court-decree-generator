// Caches the whole app so it keeps working with no internet connection.
// Bump CACHE when any file below changes, so phones pick up the new version.
const CACHE = 'decree-v4';

const ASSETS = [
  '.',
  'index.html',
  'style.css',
  'manifest.webmanifest',
  'icon.svg',
  'icon-maskable.svg',
  'js/jszip.min.js',
  'js/base-parts.js',
  'js/notice-template.js',
  'js/ooxml.js',
  'js/common.js',
  'js/mc.js',
  'js/os.js',
  'js/docx.js',
  'js/measure.js',
  'js/address.js',
  'js/notice.js',
  'js/notice-docx.js',
  'js/notice-ui.js',
  'js/app.js',
];

self.addEventListener('install', (event) => {
  event.waitUntil(caches.open(CACHE).then((cache) => cache.addAll(ASSETS)));
  self.skipWaiting();
});

self.addEventListener('activate', (event) => {
  event.waitUntil(
    caches
      .keys()
      .then((keys) => Promise.all(keys.filter((k) => k !== CACHE).map((k) => caches.delete(k))))
      .then(() => self.clients.claim())
  );
});

self.addEventListener('fetch', (event) => {
  if (event.request.method !== 'GET') return;
  event.respondWith(
    // no-store: a plain network-first fetch can still be answered from the
    // browser's own HTTP cache and silently serve yesterday's JS. Force a real
    // round trip so a redeploy is never masked by a stale cached response.
    fetch(event.request, { cache: 'no-store' })
      .then((response) => {
        const copy = response.clone();
        caches.open(CACHE).then((cache) => cache.put(event.request, copy));
        return response;
      })
      .catch(() => caches.match(event.request).then((hit) => hit || caches.match('index.html')))
  );
});
