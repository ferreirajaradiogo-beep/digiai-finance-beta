const CACHE_NAME = "notafacil-beta-pwa-v3";
const APP_SHELL = [
  "/offline",
  "/manifest.webmanifest?v=beta-3",
  "/static_v3/style_v3.css?v=beta-3",
  "/static_v3/app.js?v=beta-3",
  "/static_v3/icons/icon-app.png",
  "/static_v3/icons/icon-maskable.png",
  "/static_v3/icons/favicon.png",
  "/static_v3/icons/company-logo.png"
];

self.addEventListener("install", (event) => {
  event.waitUntil(
    caches.open(CACHE_NAME).then((cache) => cache.addAll(APP_SHELL)).then(() => self.skipWaiting())
  );
});

self.addEventListener("activate", (event) => {
  event.waitUntil(
    caches.keys().then((keys) =>
      Promise.all(keys.filter((key) => key !== CACHE_NAME).map((key) => caches.delete(key)))
    ).then(() => self.clients.claim())
  );
});

self.addEventListener("fetch", (event) => {
  if (event.request.method !== "GET") return;

  if (event.request.mode === "navigate") {
    event.respondWith(
      fetch(event.request)
        .then((response) => response)
        .catch(() => caches.match(event.request).then((response) => response || caches.match("/offline")))
    );
    return;
  }

  event.respondWith(
    caches.match(event.request).then((cached) => {
      if (cached) return cached;
      return fetch(event.request).then((response) => {
        const copy = response.clone();
        caches.open(CACHE_NAME).then((cache) => cache.put(event.request, copy));
        return response;
      });
    })
  );
});
