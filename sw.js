/* Service worker. Build 86917f7de6 is a hash of the whole site, so every release gets a new worker and a new cache.
   - hashed css/js/images (file name contains the content hash): cache first, they never change
   - pages (HTML): stale-while-revalidate, the revalidation asks the server with If-None-Match (cache: 'no-cache'),
     so it ignores the 10-minute max-age that GitHub Pages sets
   - everything else (PDF, sitemap, other origins, non-GET, Range): not touched */
var VERSION = '86917f7de6';
var CACHE = 'site-' + VERSION;
var PRECACHE = ["/", "/en/", "/styles.88445d70.css", "/script.552efde0.js", "/favicon.svg"];
var HASHED = /\.[0-9a-f]{8}\.(css|js|webp|svg|png)$/;

self.addEventListener('install', function (e) {
  e.waitUntil(caches.open(CACHE).then(function (c) {
    return Promise.all(PRECACHE.map(function (u) { return c.add(new Request(u, { cache: 'reload' })); }));
  }).then(function () { return self.skipWaiting(); }));
});

self.addEventListener('activate', function (e) {
  e.waitUntil(caches.keys().then(function (keys) {
    return Promise.all(keys.filter(function (k) { return k.indexOf('site-') === 0 && k !== CACHE; }).map(function (k) { return caches.delete(k); }));
  }).then(function () { return self.clients.claim(); }));
});

function isPage(url) { return url.pathname.slice(-1) === '/' || /\.html$/.test(url.pathname); }

self.addEventListener('fetch', function (e) {
  var req = e.request;
  if (req.method !== 'GET' || req.headers.has('range')) return;
  var url = new URL(req.url);
  if (url.origin !== self.location.origin || url.pathname === '/sw.js') return;

  if (HASHED.test(url.pathname)) {
    e.respondWith(caches.open(CACHE).then(function (c) {
      return c.match(req.url).then(function (hit) {
        return hit || fetch(req).then(function (res) {
          if (res.ok && res.status === 200) c.put(req.url, res.clone());
          return res;
        });
      });
    }));
    return;
  }

  if ((req.mode === 'navigate' || req.destination === 'document') && isPage(url)) {
    e.respondWith(caches.open(CACHE).then(function (c) {
      return c.match(req.url).then(function (hit) {
        var net = fetch(req.url, { cache: 'no-cache', credentials: 'same-origin' }).then(function (res) {
          if (res.ok && res.status === 200 && !res.redirected) c.put(req.url, res.clone());
          return res;
        });
        if (hit) { e.waitUntil(net.catch(function () {})); return hit; }
        return net.catch(function () {
          return c.match(url.pathname.indexOf('/en/') === 0 ? '/en/' : '/').then(function (r) { return r || Response.error(); });
        });
      });
    }));
  }
});
