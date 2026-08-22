// ERGUNBAS Üretim Sistemi - Minimal Service Worker
// Amaç: "Ana Ekrana Ekle" (PWA) yüklenebilirliğini sağlamak.
// Veri tazeliği kritik olduğu için API isteklerini ÖNBELLEĞE ALMIYORUZ,
// sadece ağdan geçiriyoruz (pass-through).

self.addEventListener('install', (event) => {
  self.skipWaiting();
});

self.addEventListener('activate', (event) => {
  event.waitUntil(self.clients.claim());
});

self.addEventListener('fetch', (event) => {
  // Her isteği doğrudan ağdan getir; herhangi bir önbellekleme yapma.
  event.respondWith(fetch(event.request));
});
