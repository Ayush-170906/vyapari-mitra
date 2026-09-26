// Minimal offline cache so Vyapari Mitra opens even with a bad connection —
// exactly the kind of merchant a Sarvam AI copilot needs to work for.
const CACHE_NAME = "vyapari-mitra-v1";
const ASSETS = ["./index.html", "./manifest.json", "./paytm-logo.png"];

self.addEventListener("install", (event) => {
  event.waitUntil(caches.open(CACHE_NAME).then((cache) => cache.addAll(ASSETS)));
  self.skipWaiting();
});

self.addEventListener("activate", (event) => {
  event.waitUntil(
    caches.keys().then((keys) =>
      Promise.all(keys.filter((k) => k !== CACHE_NAME).map((k) => caches.delete(k)))
    )
  );
  self.clients.claim();
});

self.addEventListener("fetch", (event) => {
  // Network-first for the Sarvam AI API itself, cache-first for the app shell
  if (event.request.url.includes("api.sarvam.ai")) return;
  event.respondWith(
    caches.match(event.request).then((cached) => cached || fetch(event.request))
  );
});
