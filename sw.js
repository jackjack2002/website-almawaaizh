const CACHE_NAME = 'Mawaaizh-admin-v14'; // Naik ke v10
const coreAssets = [
  '/admin.html',
  '/manifest.json',
  '/assets/logo_new.png',
  '/assets/logo_new.png'
];

// 1. Install & Cache File Inti Saja
self.addEventListener('install', event => {
  event.waitUntil(
    caches.open(CACHE_NAME).then(cache => {
      return cache.addAll(coreAssets);
    })
  );
  self.skipWaiting();
});

// 2. Hapus Cache Versi Lama
self.addEventListener('activate', event => {
  event.waitUntil(
    caches.keys().then(cacheNames => {
      return Promise.all(
        cacheNames.filter(name => name !== CACHE_NAME).map(name => caches.delete(name))
      );
    })
  );
  self.clients.claim();
});

// 3. Dynamic Cache
self.addEventListener('fetch', event => {
  if (event.request.url.includes('script.google.com') || event.request.method !== 'GET') {
    return;
  }

  event.respondWith(
    caches.match(event.request).then(cachedResponse => {
      if (cachedResponse) return cachedResponse;

      return fetch(event.request).then(networkResponse => {
        if (!networkResponse || networkResponse.status !== 200 || networkResponse.type !== 'basic') {
          return networkResponse;
        }

        let responseToCache = networkResponse.clone();
        caches.open(CACHE_NAME).then(cache => {
          cache.put(event.request, responseToCache);
        });

        return networkResponse;
      });
    }).catch(() => {})
  );
});
