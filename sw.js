/* Sallee Load Board — minimal service worker.
   Purpose: make the board installable as a PWA (Add to Home Screen on iOS)
   and let the app shell still open when offline. Deliberately does NOT
   cache or intercept config.js, bundle.js, orders.json, or anything
   cross-origin (Microsoft Graph/SharePoint, the MSAL CDN script) — those
   must always go to the network so sign-in, live data, and future code
   updates keep working exactly as before. Matches the same pattern used
   on the Sallee Dispatch Board. */

const CACHE_NAME = "sallee-load-board-shell-v1";
const APP_SHELL = [
  "./",
  "./index.html",
  "./manifest.json",
  "./icon-192.png",
  "./icon-512.png",
  "./icon-maskable-192.png",
  "./icon-maskable-512.png",
];

self.addEventListener("install", (event) => {
  event.waitUntil(
    caches
      .open(CACHE_NAME)
      .then((cache) => cache.addAll(APP_SHELL))
      .then(() => self.skipWaiting())
  );
});

self.addEventListener("activate", (event) => {
  event.waitUntil(
    caches
      .keys()
      .then((names) =>
        Promise.all(names.filter((n) => n !== CACHE_NAME).map((n) => caches.delete(n)))
      )
      .then(() => self.clients.claim())
  );
});

self.addEventListener("fetch", (event) => {
  const req = event.request;
  if (req.method !== "GET") return;

  const url = new URL(req.url);
  if (url.origin !== location.origin) return; // never touch cross-origin requests

  // Live/changing files must always hit the network, never be served stale.
  if (
    url.pathname.endsWith("/config.js") ||
    url.pathname.endsWith("/bundle.js") ||
    url.pathname.includes("/data/")
  ) {
    return;
  }

  if (req.mode === "navigate") {
    // Loading the app itself: prefer the network (fresh code), fall back to
    // the cached shell only when offline.
    event.respondWith(fetch(req).catch(() => caches.match("./index.html")));
    return;
  }

  // Static app-shell assets (icons, manifest): cache-first.
  event.respondWith(caches.match(req).then((cached) => cached || fetch(req)));
});
