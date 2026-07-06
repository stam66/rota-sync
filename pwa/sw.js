// App shell cache. Navigations are network-first so a new deploy is picked
// up on the very next launch (no force-quit/reboot dance); the cache serves
// only when offline. Static assets are cache-first with background refresh.
const CACHE = 'rota-shell-v2';
const SHELL = ['./', 'index.html', 'manifest.webmanifest', 'icon-192.png', 'icon-512.png'];

self.addEventListener('install', function (e) {
  e.waitUntil(caches.open(CACHE).then(function (c) { return c.addAll(SHELL); }));
  self.skipWaiting();
});

self.addEventListener('activate', function (e) {
  e.waitUntil(caches.keys().then(function (keys) {
    return Promise.all(keys.filter(function (k) { return k !== CACHE; })
      .map(function (k) { return caches.delete(k); }));
  }));
  self.clients.claim();
});

self.addEventListener('fetch', function (e) {
  const url = new URL(e.request.url);
  if (url.origin !== location.origin) return; // feed fetches go straight to network

  const isShell = e.request.mode === 'navigate' ||
    url.pathname.endsWith('/index.html') || url.pathname.endsWith('/');

  if (isShell) {
    // Network-first: always serve the latest deploy when online.
    e.respondWith(
      fetch(e.request).then(function (res) {
        if (res.ok) caches.open(CACHE).then(function (c) { c.put(e.request, res.clone()); });
        return res;
      }).catch(function () { return caches.match(e.request); })
    );
    return;
  }

  // Assets: cache-first, refresh in background.
  e.respondWith(
    caches.match(e.request).then(function (hit) {
      const net = fetch(e.request).then(function (res) {
        if (res.ok) caches.open(CACHE).then(function (c) { c.put(e.request, res.clone()); });
        return res;
      }).catch(function () { return hit; });
      return hit || net;
    })
  );
});
