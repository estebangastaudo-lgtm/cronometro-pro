const CACHE_PREFIX = 'enduro-solo-';
const CACHE_NAME = 'enduro-solo-v0.3.2';
const INDEX_FALLBACK = './index.html';

const APP_SHELL = [
  './',
  './index.html',
  './css/style.css?v=solo-v0.3',
  './js/app.js?v=solo-v0.3.2',
  './js/lib/jszip.min.js',
  './audio/enduro_braap.mp3',
  './manifest.json',
  './splash.png',
  './apple-touch-icon-solo.png',
  './icon-solo-192.png',
  './icon-solo-512.png'
];

self.addEventListener('install', event => {
  event.waitUntil((async () => {
    // addAll is atomic: if any essential resource fails, this worker does
    // not activate and the previously installed version remains available.
    const cache = await caches.open(CACHE_NAME);
    await cache.addAll(APP_SHELL);
    await self.skipWaiting();
  })());
});

self.addEventListener('activate', event => {
  event.waitUntil((async () => {
    const keys = await caches.keys();
    await Promise.all(
      keys
        .filter(key => key.startsWith(CACHE_PREFIX) && key !== CACHE_NAME)
        .map(key => caches.delete(key))
    );
    await self.clients.claim();
  })());
});

self.addEventListener('fetch', event => {
  const request = event.request;
  if (request.method !== 'GET') return;

  event.respondWith((async () => {
    const cache = await caches.open(CACHE_NAME);

    // Navigations always use the complete shell belonging to this SW version.
    // This avoids mixing a new index with old JS/CSS and still opens offline.
    if (request.mode === 'navigate') {
      const cachedIndex = await cache.match(INDEX_FALLBACK) || await cache.match('./');
      if (cachedIndex) return cachedIndex;
    } else {
      const cached = await cache.match(request);
      if (cached) return cached;
    }

    try {
      const response = await fetch(request);
      const requestUrl = new URL(request.url);
      if (requestUrl.origin === self.location.origin && response && response.ok) {
        event.waitUntil(cache.put(request, response.clone()).catch(() => {}));
      }
      return response;
    } catch (error) {
      if (request.mode === 'navigate') {
        const fallback = await cache.match(INDEX_FALLBACK) || await cache.match('./');
        if (fallback) return fallback;
      }
      throw error;
    }
  })());
});
