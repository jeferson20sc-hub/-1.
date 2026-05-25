// Service worker: cache offline + atualização rápida
const CACHE = 'lis-v2';
const CORE = ['./', './index.html', './manifest.json', './icon-192.png', './icon-512.png'];

self.addEventListener('install', e => {
  self.skipWaiting();
  e.waitUntil(caches.open(CACHE).then(c => c.addAll(CORE).catch(() => {})));
});

self.addEventListener('activate', e => {
  e.waitUntil(
    caches.keys().then(keys => Promise.all(keys.filter(k => k !== CACHE).map(k => caches.delete(k))))
      .then(() => self.clients.claim())
  );
});

self.addEventListener('fetch', e => {
  const req = e.request;
  if (req.method !== 'GET') return;
  const url = new URL(req.url);
  // Não interceptar chamadas para o Power Automate (precisam ir direto à rede).
  if (url.host.includes('powerplatform.com') || url.host.includes('logic.azure.com')) return;
  // Stale-while-revalidate para os arquivos do próprio site.
  if (url.origin === self.location.origin) {
    e.respondWith(
      caches.match(req).then(cached => {
        const fetcher = fetch(req).then(res => {
          if (res && res.ok) {
            const copy = res.clone();
            caches.open(CACHE).then(c => c.put(req, copy));
          }
          return res;
        }).catch(() => cached);
        return cached || fetcher;
      })
    );
  }
});
