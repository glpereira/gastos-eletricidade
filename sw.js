// Guarda a aplicação em cache para funcionar sem rede.
// Ao publicar uma nova versão, altere VERSAO.
const VERSAO = 'v1';
const FICHEIROS = [
  './', 'index.html', 'css/app.css', 'js/app.js', 'js/core.js', 'js/store.js',
  'vendor/chart.umd.js', 'vendor/xlsx.full.min.js', 'manifest.webmanifest',
  'icons/icon-180.png', 'icons/icon-192.png', 'icons/icon-512.png',
];

self.addEventListener('install', (e) => {
  e.waitUntil(caches.open(VERSAO).then((c) => c.addAll(FICHEIROS)).then(() => self.skipWaiting()));
});

self.addEventListener('activate', (e) => {
  e.waitUntil(
    caches.keys()
      .then((ks) => Promise.all(ks.filter((k) => k !== VERSAO).map((k) => caches.delete(k))))
      .then(() => self.clients.claim()),
  );
});

// Rede primeiro (para receber atualizações); cache quando não há rede.
self.addEventListener('fetch', (e) => {
  if (e.request.method !== 'GET' || new URL(e.request.url).origin !== location.origin) return;
  e.respondWith(
    fetch(e.request)
      .then((r) => {
        const copia = r.clone();
        caches.open(VERSAO).then((c) => c.put(e.request, copia));
        return r;
      })
      .catch(() => caches.match(e.request, { ignoreSearch: true })),
  );
});
