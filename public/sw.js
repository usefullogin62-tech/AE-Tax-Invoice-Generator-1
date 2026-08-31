// Minimal service worker — its only job is to exist with a fetch handler
// so browsers consider this site "installable" (Add to Home Screen / Add
// to Desktop). It passes every request straight through to the network.
self.addEventListener("install", () => {
  self.skipWaiting();
});

self.addEventListener("activate", (event) => {
  event.waitUntil(self.clients.claim());
});

self.addEventListener("fetch", (event) => {
  event.respondWith(fetch(event.request));
});
