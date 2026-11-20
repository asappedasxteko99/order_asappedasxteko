const CACHE_NAME = 'asappedas-pwa-v1';

self.addEventListener('install', (event) => {
  self.skipWaiting();
});

self.addEventListener('activate', (event) => {
  event.waitUntil(clients.claim());
});

self.addEventListener('fetch', (event) => {
  // Membiarkan request berjalan normal (network-first) agar data Apps Script selalu terbaru
  event.respondWith(fetch(event.request));
});
