// ============================================================
//  Athlete Grocery PWA — Service Worker
//  Provides full offline support via Cache-First strategy
// ============================================================

const CACHE_NAME    = 'athlete-grocery-v1';
const OFFLINE_URLS  = [
  './',
  './index.html',
  './manifest.json',
  'https://cdn.tailwindcss.com'
];

// ── Install: pre-cache shell ──────────────────────────────────
self.addEventListener('install', event => {
  event.waitUntil(
    caches.open(CACHE_NAME).then(cache => {
      // Cache what we can; ignore failures for CDN resources
      return Promise.allSettled(
        OFFLINE_URLS.map(url =>
          cache.add(url).catch(() => console.warn('[SW] Could not cache:', url))
        )
      );
    }).then(() => self.skipWaiting())
  );
});

// ── Activate: purge old caches ────────────────────────────────
self.addEventListener('activate', event => {
  event.waitUntil(
    caches.keys().then(keys =>
      Promise.all(
        keys
          .filter(key => key !== CACHE_NAME)
          .map(key => caches.delete(key))
      )
    ).then(() => self.clients.claim())
  );
});

// ── Fetch: Cache-First with network fallback ──────────────────
self.addEventListener('fetch', event => {
  // Skip cross-origin requests to the Apps Script API (always needs network)
  const url = new URL(event.request.url);
  if (url.hostname.includes('script.google.com')) {
    event.respondWith(
      fetch(event.request).catch(() =>
        new Response(JSON.stringify({ error: 'Offline — réseau indisponible' }), {
          headers: { 'Content-Type': 'application/json' }
        })
      )
    );
    return;
  }

  // Cache-first for everything else
  event.respondWith(
    caches.match(event.request).then(cached => {
      if (cached) return cached;

      return fetch(event.request).then(response => {
        // Only cache successful same-origin or CDN responses
        if (
          response.ok &&
          (url.origin === self.location.origin ||
           url.hostname === 'cdn.tailwindcss.com')
        ) {
          const clone = response.clone();
          caches.open(CACHE_NAME).then(cache => cache.put(event.request, clone));
        }
        return response;
      }).catch(() => {
        // If no network and no cache, return blank for navigation
        if (event.request.mode === 'navigate') {
          return caches.match('./index.html');
        }
        return new Response('', { status: 408 });
      });
    })
  );
});

// ── Background Sync (optional future use) ────────────────────
self.addEventListener('sync', event => {
  if (event.tag === 'sync-grocery-list') {
    console.log('[SW] Background sync triggered');
    // Actual sync is handled in the app via the Synchroniser button
  }
});
