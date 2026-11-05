// RCI Marriage: service worker.
// Keeps a copy of the app so it opens offline and can be installed, but always
// asks the server for the newest version first, so updates show up right away.
// Couple data and sign-in always go to Microsoft live; they are never cached.
const CACHE = "rci-marriage-v4";
const SHELL = ["./", "index.html", "manifest.webmanifest", "config.js", "live.js", "src/data.js", "vendor/msal-browser.min.js",
  "icons/icon-192.png", "icons/icon-512.png", "icons/icon-maskable-512.png",
  "icons/apple-touch-icon.png", "icons/favicon.png"];

self.addEventListener("install", e => {
  e.waitUntil(caches.open(CACHE)
    .then(c => c.addAll(SHELL.map(u => new Request(u, { cache: "reload" }))))
    .then(() => self.skipWaiting()));
});

self.addEventListener("activate", e => {
  e.waitUntil(caches.keys()
    .then(keys => Promise.all(keys.filter(k => k.startsWith("rci-marriage-") && k !== CACHE).map(k => caches.delete(k))))
    .then(() => self.clients.claim()));
});

function newestFirst(request, cacheKey){
  return fetch(request.url, { cache: "no-cache", credentials: "same-origin" })
    .then(res => {
      if (res.ok){ const copy = res.clone(); caches.open(CACHE).then(c => c.put(cacheKey || request, copy)); }
      return res;
    })
    .catch(() => caches.match(cacheKey || request).then(hit => hit || (request.mode === "navigate" ? caches.match("index.html") : Response.error())));
}

self.addEventListener("fetch", e => {
  const url = new URL(e.request.url);
  if (e.request.method !== "GET" || url.origin !== self.location.origin) return;
  if (e.request.mode === "navigate" || /\/(index\.html)?$/.test(url.pathname)){ e.respondWith(newestFirst(e.request, "index.html")); return; }
  e.respondWith(newestFirst(e.request, url.pathname));
});
