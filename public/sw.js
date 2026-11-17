// SW Pembersih Legacy — Sistem Eviden Rekonsiliasi
// Tujuan khusus: menghapus service worker & cache lama dari aplikasi lain
// (contoh: PWA "BOP Konstruksi" yang pernah berjalan di origin yang sama,
// cache bop-konstruksi-v1) lalu membatalkan registrasi SW ini sendiri.
// Dengan begitu halaman selalu diambil langsung dari server, bukan cache.

self.addEventListener('install', () => {
  self.skipWaiting();
});

self.addEventListener('activate', event => {
  event.waitUntil(
    caches.keys()
      .then(keys => Promise.all(keys.map(k => caches.delete(k))))
      .then(() => self.registration.unregister())
  );
  self.clients.claim();
});

self.addEventListener('fetch', event => {
  event.respondWith(fetch(event.request));
});