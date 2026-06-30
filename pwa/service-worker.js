const CACHE_NAME = "project-intelligence-hub-v1";

self.addEventListener("install", event => {
  self.skipWaiting();
});

self.addEventListener("activate", event => {
  event.waitUntil(self.clients.claim());
});

self.addEventListener("fetch", event => {
  event.respondWith(
    fetch(event.request).catch(() => {
      return new Response(
        "Project Intelligence Hub is offline. Please reconnect to the internet and try again.",
        { headers: { "Content-Type": "text/plain" } }
      );
    })
  );
});
