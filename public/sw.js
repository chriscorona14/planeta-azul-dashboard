const CACHE_NAME = 'dashboard-financiero-v1';
const ASSETS_TO_CACHE = [
  '/',
  '/index.html',
  '/manifest.json',
  '/icon.svg',
  '/main.js',
  '/financialEngine.js',
  '/resumenComercialEngine.js',
  '/costoUnitarioEngine.js',
  '/worker.js',
  '/script.js'
];

self.addEventListener('install', event => {
  event.waitUntil(
    caches.open(CACHE_NAME)
      .then(cache => {
        // Cache core assets safely (don't fail the whole install if one fails)
        console.log('[Service Worker] Caching core assets');
        return Promise.allSettled(
          ASSETS_TO_CACHE.map(url => cache.add(url))
        );
      })
      .then(() => self.skipWaiting())
  );
});

self.addEventListener('activate', event => {
  console.log('[Service Worker] Activating...');
  event.waitUntil(
    caches.keys().then(cacheNames => {
      return Promise.all(
        cacheNames.map(cacheName => {
          if (cacheName !== CACHE_NAME) {
            console.log('[Service Worker] Deleting old cache:', cacheName);
            return caches.delete(cacheName);
          }
        })
      );
    }).then(() => self.clients.claim())
  );
});

self.addEventListener('fetch', event => {
  // Only handle GET requests
  if (event.request.method !== 'GET') return;

  // Stale-while-revalidate strategy for reliable fast loading
  // Useful for thick apps processing heavy files locally
  event.respondWith(
    caches.match(event.request).then(cachedResponse => {
      const fetchPromise = fetch(event.request).then(networkResponse => {
        // Cache valid responses (including opaque responses from CDNs)
        if (networkResponse && (networkResponse.status === 200 || networkResponse.type === 'opaque')) {
          const responseToCache = networkResponse.clone();
          caches.open(CACHE_NAME).then(cache => {
            cache.put(event.request, responseToCache);
          });
        }
        return networkResponse;
      }).catch(err => {
        console.warn('[Service Worker] Fetch failed, relying on cache', err);
      });

      return cachedResponse || fetchPromise;
    })
  );
});
