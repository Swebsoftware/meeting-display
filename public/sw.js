const CACHE = 'office-shell-v3';
self.addEventListener('install', event => { self.skipWaiting(); event.waitUntil(caches.open(CACHE).then(cache => cache.addAll(['/', '/icon.svg', '/manifest.webmanifest']))); });
self.addEventListener('activate', event => { event.waitUntil(Promise.all([self.clients.claim(), caches.keys().then(keys => Promise.all(keys.filter(key => key !== CACHE).map(key => caches.delete(key))))])); });
self.addEventListener('fetch', event => {
  const url = new URL(event.request.url);
  if (url.origin !== location.origin || (url.pathname.startsWith('/api/') || url.pathname==='/connect') || event.request.method !== 'GET') return;
  if (url.pathname.startsWith('/display/')) { event.respondWith(fetch(event.request).catch(async () => await caches.match('/') || new Response('Offline. Reconnect to load the display.', {status:503}))); return; }
  event.respondWith(fetch(event.request).then(response => { if(response.ok) { const copy = response.clone(); event.waitUntil(caches.open(CACHE).then(cache => cache.put(event.request, copy))); } return response; }).catch(async () => (await caches.match(event.request)) || (event.request.mode === 'navigate' ? await caches.match('/') : undefined) || new Response('Offline. Reconnect to load the display.', {status:503})));
});
