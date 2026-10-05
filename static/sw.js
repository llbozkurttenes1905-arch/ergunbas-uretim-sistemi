// ERGUNBAS Üretim Sistemi - Minimal Service Worker v2.4.2
// Amaç: "Ana Ekrana Ekle" (PWA) yüklenebilirliğini sağlamak.
// Tüm eski cache'leri temizler ve canlı ağı kullanır.

self.addEventListener('install', (event) => {
  self.skipWaiting();
});

self.addEventListener('activate', (event) => {
  event.waitUntil(
    caches.keys().then((keys) => Promise.all(keys.map((k) => caches.delete(k))))
      .then(() => self.clients.claim())
  );
});

self.addEventListener('fetch', (event) => {
  // Her isteği doğrudan ağdan getir; no-store ile önbellekleme yapma
  event.respondWith(fetch(event.request, { cache: 'no-store' }));
});
