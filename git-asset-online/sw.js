// Service worker: simpan aplikasi di perangkat supaya bisa dibuka tanpa internet.
// Naikkan angka VERSI setiap kali file aplikasi diubah agar semua perangkat memuat versi terbaru.
const VERSI = "v9";
const CACHE = "git-asset-" + VERSI;
const FILE = ["./", "index.html", "css/style.css", "js/app.js", "vendor/xlsx.full.min.js", "vendor/cryptofb.js", "manifest.webmanifest", "icons/icon-192.png", "icons/icon-512.png"];
self.addEventListener("install", e => { e.waitUntil(caches.open(CACHE).then(c => c.addAll(FILE)).then(() => self.skipWaiting())); });
self.addEventListener("activate", e => { e.waitUntil(caches.keys().then(ks => Promise.all(ks.filter(k => k !== CACHE).map(k => caches.delete(k)))).then(() => self.clients.claim())); });
self.addEventListener("fetch", e => {
  if (new URL(e.request.url).pathname.endsWith("/data.js")) { e.respondWith(fetch(e.request).then(r => { if (r.ok) caches.open(CACHE).then(c => c.put(e.request, r.clone())); return r; }).catch(() => caches.match(e.request, { ignoreSearch: true }))); return; }
  if (e.request.method !== "GET" || new URL(e.request.url).origin !== location.origin) return;
  e.respondWith(caches.match(e.request, { ignoreSearch: true }).then(hit => {
    const net = fetch(e.request).then(r => { if (r.ok) caches.open(CACHE).then(c => c.put(e.request, r.clone())); return r; }).catch(() => hit);
    return hit || net;
  }));
});
