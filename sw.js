/* Prizm service worker — offline support + installability.
 *
 * Design goals:
 *  - App keeps opening with no network (cached shell).
 *  - New deploys still roll out: the HTML is fetched network-first, so an online
 *    launch always gets the latest build and falls back to cache only when offline.
 *  - The user's data paths are NEVER cached or intercepted: all Google endpoints
 *    (Apps Script web app, Sheets API, Identity Services) go straight to the network.
 *
 * IMPORTANT: keep CACHE_VERSION in lockstep with APP_VERSION in app.js and bump
 * it on every release so old caches are purged. (It was stuck at 'opus-v1' across
 * releases, so returning users could be served stale cached assets.)
 */
const CACHE_VERSION = 'prizm-v4.1';
const CORE_ASSETS = [
  './',
  './index.html',
  './styles.css',
  './app.js',
  './logo-mark.png',
  './icon-192.png',
  './icon-512.png',
  './icon-maskable-512.png',
];

// Cross-origin static assets we're happy to cache (fonts + charting lib).
// These are safe to serve stale; everything else cross-origin is left alone.
const CACHEABLE_HOSTS = [
  'fonts.googleapis.com',
  'fonts.gstatic.com',
  'cdnjs.cloudflare.com',
];

// Hosts that must always hit the live network (user data + auth). Never cached.
const BYPASS_HOSTS = [
  'script.google.com',
  'script.googleusercontent.com',
  'sheets.googleapis.com',
  'www.googleapis.com',
  'accounts.google.com',
  'apis.google.com',
  'oauth2.googleapis.com',
];

self.addEventListener('install', event => {
  event.waitUntil(
    caches.open(CACHE_VERSION)
      .then(cache => cache.addAll(CORE_ASSETS))
      .then(() => self.skipWaiting())
      .catch(() => self.skipWaiting())
  );
});

self.addEventListener('activate', event => {
  event.waitUntil(
    caches.keys()
      .then(keys => Promise.all(keys.filter(k => k !== CACHE_VERSION).map(k => caches.delete(k))))
      .then(() => self.clients.claim())
  );
});

self.addEventListener('message', event => {
  if (event.data === 'skipWaiting') self.skipWaiting();
});

self.addEventListener('fetch', event => {
  const req = event.request;
  if (req.method !== 'GET') return; // never touch writes (Sheet updates use GET too — see below)

  let url;
  try { url = new URL(req.url); } catch (e) { return; }

  // 1) Anything that talks to Google (data/auth) always goes to the live network.
  if (BYPASS_HOSTS.some(h => url.hostname === h || url.hostname.endsWith('.' + h))) return;

  const sameOrigin = url.origin === self.location.origin;

  // The Prizm Sheet backend is a script.google.com/macros/... URL (already bypassed
  // above). Extra guard: never cache anything carrying a query string (Prizm's Sheet
  // reads/writes are GETs with ?body=... / ?_ts=...), so live data is never served stale.
  if (url.search) return;

  // 2) App navigations / same-origin HTML → network-first, fall back to cache offline.
  if (req.mode === 'navigate' || (sameOrigin && req.destination === 'document')) {
    event.respondWith(
      fetch(req)
        .then(res => {
          const copy = res.clone();
          caches.open(CACHE_VERSION).then(c => c.put('./index.html', copy)).catch(() => {});
          return res;
        })
        .catch(() => caches.match('./index.html').then(r => r || caches.match('./')))
    );
    return;
  }

  // 3) Same-origin code (app.js / styles.css) → network-first so a new build shows
  //    on the next launch; fall back to the cached copy only when offline.
  if (sameOrigin && /\.(?:js|css)$/i.test(url.pathname)) {
    event.respondWith(
      fetch(req).then(res => {
        if (res && res.ok) { const copy = res.clone(); caches.open(CACHE_VERSION).then(c => c.put(req, copy)).catch(() => {}); }
        return res;
      }).catch(() => caches.match(req))
    );
    return;
  }

  // 4) Other same-origin static (icons) + whitelisted CDN static → cache-first w/ refresh.
  if (sameOrigin || CACHEABLE_HOSTS.some(h => url.hostname === h)) {
    event.respondWith(
      caches.match(req).then(cached => {
        const network = fetch(req).then(res => {
          if (res && (res.ok || res.type === 'opaque')) {
            const copy = res.clone();
            caches.open(CACHE_VERSION).then(c => c.put(req, copy)).catch(() => {});
          }
          return res;
        }).catch(() => cached);
        return cached || network;
      })
    );
    return;
  }

  // Everything else: default browser handling (no caching).
});
