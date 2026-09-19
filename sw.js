// Minimal offline cache: app shell only. Data lives in localStorage, not here.
const CACHE = "morning-coach-v2";
const ASSETS = ["./", "./index.html", "./manifest.webmanifest", "./icon-192.png", "./icon-512.png"];
// skipWaiting + claim: without these a new version sits in "waiting" until every
// tab of the app is closed. In a standalone PWA that means the user has to fully
// kill the app to see an update — the classic "I pushed, why is it still old?"
self.addEventListener("install", e => e.waitUntil(
  caches.open(CACHE).then(c => c.addAll(ASSETS)).then(() => self.skipWaiting())
));
self.addEventListener("activate", e => e.waitUntil(
  caches.keys()
    .then(keys => Promise.all(keys.filter(k => k !== CACHE).map(k => caches.delete(k))))
    .then(() => self.clients.claim())
));
self.addEventListener("fetch", e => {
  if (e.request.method !== "GET") return;
  e.respondWith(caches.match(e.request).then(hit => hit || fetch(e.request)));
});
