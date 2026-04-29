const CACHE_NAME = "quran-mac-word-v2";
const API_CACHE_NAME = "quran-api-cache-v1";

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

// Activate: clean old caches (keep current app + API caches)
self.addEventListener("activate", (event) => {
  const keepCaches = [CACHE_NAME, API_CACHE_NAME];
  event.waitUntil(
    caches.keys().then((keys) =>
      Promise.all(
        keys
          .filter((key) => !keepCaches.includes(key))
          .map((key) => caches.delete(key))
      )
    ).then(() => self.clients.claim())
  );
});

// Fetch handler
self.addEventListener("fetch", (event) => {
  const url = new URL(event.request.url);

  // Skip non-GET requests
  if (event.request.method !== "GET") {
    return;
  }

  // Cross-origin: cache quran.com API translation responses (cache-first, immutable data)
  if (url.hostname === "api.quran.com" && url.pathname.startsWith("/api/v4/quran/translations/")) {
    event.respondWith(
      caches.open(API_CACHE_NAME).then((cache) =>
        cache.match(event.request).then((cached) => {
          if (cached) return cached;
          return fetch(event.request).then((response) => {
            if (response.ok) {
              cache.put(event.request, response.clone());
            }
            return response;
          });
        })
      )
    );
    return;
  }

  // Only cache same-origin requests (our own assets)
  if (url.origin !== self.location.origin) {
    return;
  }

  // Same-origin: stale-while-revalidate
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
