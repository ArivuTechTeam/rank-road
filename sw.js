// Rank Road – offline support. Bump VERSION whenever you redeploy changed files.
const VERSION = "rankroad-v6";
const CORE = ["./","index.html","manifest.webmanifest","assets/icon-192.png","assets/icon-512.png",
  "foundation/","inter-group1/","inter-group2/","inter-both/"];
const NEVER_CACHE = /(^|\.)supabase\.co$|^cloudflare-dns\.com$|^script\.google\.com$|^script\.googleusercontent\.com$|(^|\.)clarity\.ms$/;
self.addEventListener("install", e => { e.waitUntil(caches.open(VERSION).then(c => c.addAll(CORE)).then(() => self.skipWaiting())); });
self.addEventListener("activate", e => { e.waitUntil(caches.keys().then(ks => Promise.all(ks.filter(k => k !== VERSION).map(k => caches.delete(k)))).then(() => self.clients.claim())); });
self.addEventListener("fetch", e => {
  if (e.request.method !== "GET" || NEVER_CACHE.test(new URL(e.request.url).hostname)) return;
  e.respondWith(fetch(e.request).then(r => { const copy = r.clone(); caches.open(VERSION).then(c => c.put(e.request, copy)); return r; })
    .catch(() => caches.match(e.request).then(m => m || caches.match("index.html"))));
});
