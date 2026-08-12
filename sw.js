// Service worker mínimo: solo cachea el "cascarón" estático de la app
// (para que el formulario abra aunque no haya señal) y habilita que Chrome/
// Android ofrezcan "Instalar app". No cachea nada del Apps Script ni la
// base de clientes: esos siempre se piden en vivo.
const CACHE = 'registro-visitas-v1';
const SHELL = ['./', './index.html', './admin.html', './manifest.json', './icon.svg'];

self.addEventListener('install', e => {
  e.waitUntil(caches.open(CACHE).then(c => c.addAll(SHELL)));
  self.skipWaiting();
});

self.addEventListener('activate', e => {
  e.waitUntil(
    caches.keys().then(keys => Promise.all(keys.filter(k => k !== CACHE).map(k => caches.delete(k))))
  );
  self.clients.claim();
});

self.addEventListener('fetch', e => {
  const url = new URL(e.request.url);
  if (url.origin !== location.origin) return; // deja pasar Apps Script, CSV, CDNs, mapas
  e.respondWith(
    caches.match(e.request).then(cached => {
      const network = fetch(e.request).then(res => {
        caches.open(CACHE).then(c => c.put(e.request, res.clone()));
        return res;
      }).catch(() => cached);
      return cached || network;
    })
  );
});
