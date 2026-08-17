const CACHE_NAME = "tracker-lavoro-cafe-v9-2-stopwatch-stats";
const APP_SHELL = [
  "./working-tracker.html",
  "./kurorei-chill.png",
  "./manifest.webmanifest",
  "./icons/favicon-32.png",
  "./icons/apple-touch-icon.png",
  "./icons/icon-192.png",
  "./icons/icon-512.png",
  "./icons/icon-maskable-512.png"
];

self.addEventListener("install", (event) => {
  event.waitUntil(
    caches.open(CACHE_NAME)
      .then((cache) => cache.addAll(APP_SHELL))
      .then(() => self.skipWaiting())
  );
});

self.addEventListener("activate", (event) => {
  event.waitUntil(
    caches.keys()
      .then((keys) => Promise.all(keys.filter((key) => key !== CACHE_NAME).map((key) => caches.delete(key))))
      .then(() => self.clients.claim())
  );
});

self.addEventListener("message", (event) => {
  if (event.data && event.data.type === "SKIP_WAITING") self.skipWaiting();
});

self.addEventListener("fetch", (event) => {
  if (event.request.method !== "GET") return;

  const url = new URL(event.request.url);

  // I dati di Google Sheets devono arrivare sempre dalla rete.
  // Non memorizzare mai le risposte di Apps Script nella cache della PWA.
  if (url.hostname === "script.google.com" || url.hostname === "script.googleusercontent.com") {
    event.respondWith(fetch(event.request, { cache: "no-store" }));
    return;
  }

  if (url.origin === self.location.origin) {
    // Per la pagina principale usa prima la rete: evita di mostrare una
    // vecchia versione GitHub Pages per un intero avvio della PWA.
    if (event.request.mode === "navigate" || url.pathname.endsWith("/working-tracker.html")) {
      event.respondWith(
        fetch(event.request, { cache: "no-store" })
          .then((response) => {
            if (response && response.status === 200) {
              const copy = response.clone();
              caches.open(CACHE_NAME).then((cache) => cache.put("./working-tracker.html", copy));
            }
            return response;
          })
          .catch(() => caches.match("./working-tracker.html"))
      );
      return;
    }
    event.respondWith(
      caches.match(event.request).then((cached) => {
        const network = fetch(event.request)
          .then((response) => {
            if (response && response.status === 200) {
              const copy = response.clone();
              caches.open(CACHE_NAME).then((cache) => cache.put(event.request, copy));
            }
            return response;
          })
          .catch(() => cached || caches.match("./working-tracker.html"));
        return cached || network;
      })
    );
    return;
  }

  event.respondWith(
    caches.match(event.request).then((cached) => {
      if (cached) return cached;
      return fetch(event.request).then((response) => {
        if (response && (response.status === 200 || response.type === "opaque")) {
          const copy = response.clone();
          caches.open(CACHE_NAME).then((cache) => cache.put(event.request, copy));
        }
        return response;
      });
    })
  );
});
