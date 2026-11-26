const prefix = `travel-planner:${self.registration.scope}:`;
const cacheName = `${prefix}v1`;
const shell = [
  './index.html', './styles.css', './app.js', './itinerary.js', './manifest.webmanifest',
  './vendor/xlsx.min.js', './icons/192.png', './icons/512.png', './icons/apple.png'
].map(path => new URL(path, self.registration.scope).href);

self.addEventListener('install', event => {
  event.waitUntil(caches.open(cacheName).then(cache =>
    cache.addAll(shell.map(url => new Request(url, { cache: 'reload' })))
  ));
});
self.addEventListener('message', event => {
  if (event.data === 'activate-update') self.skipWaiting();
});
self.addEventListener('activate', event => {
  event.waitUntil((async () => {
    const keys = await caches.keys();
    await Promise.all(keys.filter(key => key.startsWith(prefix) && key !== cacheName).map(key => caches.delete(key)));
    await self.clients.claim();
  })());
});
self.addEventListener('fetch', event => {
  const { request } = event;
  if (request.method !== 'GET' || !request.url.startsWith(self.registration.scope)) return;
  const url = request.mode === 'navigate' ? shell[0] : request.url;
  if (!shell.includes(url)) return;
  event.respondWith(caches.open(cacheName).then(async cache => (await cache.match(url)) || fetch(request)));
});
