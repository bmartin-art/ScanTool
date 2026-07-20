/* Duct Tracker service worker.
   Bump VERSION whenever you want to force every device to drop its cached shell.
   Day-to-day you don't need to: navigations are network-first, so a fresh
   index.html on GitHub Pages loads automatically whenever the tablet is online. */
const VERSION = 'duct-tracker-v1';

// App shell — cached on install so the app opens instantly and survives a flaky
// shop wifi signal. The data itself always comes from the network (see below).
const SHELL = [
  './',
  './index.html',
  './manifest.webmanifest',
  'https://cdnjs.cloudflare.com/ajax/libs/qrcodejs/1.0.0/qrcode.min.js',
  'https://fonts.googleapis.com/css2?family=Archivo:wght@500;600;700;800&family=JetBrains+Mono:wght@500;700&display=swap',
];

self.addEventListener('install', (e) => {
  self.skipWaiting();
  e.waitUntil(caches.open(VERSION).then((c) => Promise.allSettled(SHELL.map((u) => c.add(u)))));
});

self.addEventListener('activate', (e) => {
  e.waitUntil(
    caches.keys().then((keys) => Promise.all(keys.filter((k) => k !== VERSION).map((k) => caches.delete(k))))
      .then(() => self.clients.claim())
  );
});

self.addEventListener('fetch', (e) => {
  const req = e.request;
  if (req.method !== 'GET') return;                       // never touch PUT/DELETE (data writes)
  const url = new URL(req.url);

  // Data API (the Cloudflare Worker) — always live, never cached.
  if (url.hostname.endsWith('workers.dev')) return;

  // App navigations — network-first so new deploys show up, cache as offline fallback.
  if (req.mode === 'navigate') {
    e.respondWith(
      fetch(req).then((res) => { caches.open(VERSION).then((c) => c.put('./index.html', res.clone())); return res; })
        .catch(() => caches.match('./index.html').then((r) => r || caches.match('./')))
    );
    return;
  }

  // Static assets (fonts, QR lib) — cache-first, fill the cache on first online hit.
  e.respondWith(
    caches.match(req).then((hit) => hit || fetch(req).then((res) => {
      const copy = res.clone();
      caches.open(VERSION).then((c) => c.put(req, copy)).catch(() => {});
      return res;
    }).catch(() => hit))
  );
});
