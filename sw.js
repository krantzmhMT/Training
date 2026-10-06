// Offline support for the app shell. Data lives in localStorage, not here.
//
// Strategy (changed in v3):
//   - The page itself (index.html) is NETWORK-FIRST: when you're online you always
//     get the latest deploy; the cached copy is only a fallback for offline.
//   - Icons and the manifest are CACHE-FIRST: they almost never change.
// Before v3 everything was cache-first, so a new index.html only reached the phone
// if this file's CACHE name was bumped. Network-first for the page removes that trap.
const CACHE = "morning-coach-v3";
const ASSETS = ["./", "./index.html", "./manifest.webmanifest", "./icon-192.png", "./icon-512.png"];

// cache:"reload" skips the browser's HTTP cache. GitHub Pages serves files with a
// 10-minute cache header, so without it a fresh install could store the OLD page.
self.addEventListener("install", e => e.waitUntil(
  caches.open(CACHE)
    .then(c => c.addAll(ASSETS.map(u => new Request(u, { cache: "reload" }))))
    .then(() => self.skipWaiting())
));

// skipWaiting + claim: the new worker takes over immediately instead of waiting
// for every window of the app to close.
self.addEventListener("activate", e => e.waitUntil(
  caches.keys()
    .then(keys => Promise.all(keys.filter(k => k !== CACHE).map(k => caches.delete(k))))
    .then(() => self.clients.claim())
));

self.addEventListener("fetch", e => {
  const req = e.request;
  if (req.method !== "GET") return;
  if (new URL(req.url).origin !== self.location.origin) return;   // fonts, API: browser handles them

  if (req.mode === "navigate") {
    e.respondWith(
      fetch(req, { cache: "no-cache" })
        .then(res => {
          if (res.ok) {                       // never let an error page replace the good copy
            const copy = res.clone();
            caches.open(CACHE).then(c => c.put("./index.html", copy));
          }
          return res;
        })
        .catch(() => caches.match("./index.html"))
    );
    return;
  }
  e.respondWith(caches.match(req).then(hit => hit || fetch(req)));
});
