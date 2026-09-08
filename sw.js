// Caches the whole app so it keeps working with no internet connection.
// Bump CACHE when any file below changes, so phones pick up the new version.
const CACHE = 'decree-v2';

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
    // Prefer the network so a redeploy is picked up, fall back to cache offline.
    fetch(event.request)
      .then((response) => {
        const copy = response.clone();
        caches.open(CACHE).then((cache) => cache.put(event.request, copy));
        return response;
      })
      .catch(() => caches.match(event.request).then((hit) => hit || caches.match('index.html')))
  );
});
