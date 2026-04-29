const CACHE_NAME = "quran-mac-word-v1";

// Install: cache the app shell immediately
self.addEventListener("install", (event) => {
  self.skipWaiting();
  event.waitUntil(
    caches.open(CACHE_NAME).then((cache) => {
      // The app shell will be cached on first fetch via the fetch handler.
      // No explicit list needed since webpack chunk names are hashed.
      return cache;
    })
  );
});

// Activate: clean old caches
self.addEventListener("activate", (event) => {
  event.waitUntil(
    caches.keys().then((keys) =>
      Promise.all(
        keys
          .filter((key) => key !== CACHE_NAME)
          .map((key) => caches.delete(key))
      )
    ).then(() => self.clients.claim())
  );
});

// Fetch: cache-first for same-origin assets, network-only for external
self.addEventListener("fetch", (event) => {
  const url = new URL(event.request.url);

  // Only cache same-origin requests (our own assets)
  if (url.origin !== self.location.origin) {
    return;
  }

  // Skip non-GET requests
  if (event.request.method !== "GET") {
    return;
  }

  event.respondWith(
    caches.open(CACHE_NAME).then((cache) =>
      cache.match(event.request).then((cached) => {
        if (cached) {
          // Return cached, but also update cache in background (stale-while-revalidate)
          const fetchPromise = fetch(event.request)
            .then((response) => {
              if (response.ok) {
                cache.put(event.request, response.clone());
              }
              return response;
            })
            .catch(() => cached);
          return cached;
        }
        // Not cached yet: fetch, cache, return
        return fetch(event.request).then((response) => {
          if (response.ok) {
            cache.put(event.request, response.clone());
          }
          return response;
        });
      })
    )
  );
});
