const CACHE = 'tj-v16';
const SHELL = ['./', './index.html', './manifest.json', './icon.svg'];
// Keep in sync with the <script> tags in index.html. config.js is generated at
// deploy time, so it's best-effort too (it may not exist in a local checkout).
const OPTIONAL = [
  './config.js',
  'https://www.gstatic.com/firebasejs/10.14.1/firebase-app-compat.js',
  'https://www.gstatic.com/firebasejs/10.14.1/firebase-auth-compat.js',
  'https://www.gstatic.com/firebasejs/10.14.1/firebase-firestore-compat.js',
];
// Live APIs: Apps Script, Firestore, Firebase Auth / Google sign-in — never cached.
// (Firestore keeps its own offline copy in IndexedDB.)
const NETWORK_ONLY = ['script.google.com', 'googleapis.com', 'firebaseapp.com', 'accounts.google.com'];

self.addEventListener('install', e => {
  e.waitUntil(caches.open(CACHE).then(c =>
    // Best-effort so a CDN hiccup or missing config can't break the install;
    // the fetch handler below also fills these in on first use.
    c.addAll(SHELL).then(() => Promise.all(OPTIONAL.map(u => c.add(u).catch(() => {}))))
  ));
  self.skipWaiting();
});

self.addEventListener('activate', e => {
  e.waitUntil(
    caches.keys().then(keys =>
      // tj-img holds portfolio images the page saved for offline use — keep it.
      Promise.all(keys.filter(k => k !== CACHE && k !== 'tj-img').map(k => caches.delete(k)))
    )
  );
  self.clients.claim();
});

function putInCache(req, res) {
  if (res.ok) { const copy = res.clone(); caches.open(CACHE).then(c => c.put(req, copy)); }
  return res;
}

self.addEventListener('fetch', e => {
  const req = e.request;
  const url = req.url;
  if (req.method !== 'GET' || NETWORK_ONLY.some(h => url.includes(h))) return;

  // The page itself: network-first so new versions show up, cached copy when offline.
  if (req.mode === 'navigate') {
    e.respondWith(
      fetch(req).then(res => putInCache('./index.html', res))
        .catch(() => caches.match('./index.html'))
    );
    return;
  }
  // Generated config and the Apps Script source: same, so a changed repository
  // variable or script update is picked up on the next load.
  const path = url.split('?')[0];
  if (url.startsWith(self.location.origin) && (path.endsWith('/config.js') || path.endsWith('/script.gs'))) {
    e.respondWith(fetch(req).then(res => putInCache(req, res)).catch(() => caches.match(req)));
    return;
  }

  // Everything else: cache-first. Same-origin files and the versioned SDK are
  // stored as they're fetched so they're there next time we're offline.
  e.respondWith(
    caches.match(req).then(cached => cached || fetch(req).then(res =>
      (url.startsWith(self.location.origin) || url.includes('gstatic.com/firebasejs/'))
        ? putInCache(req, res) : res
    ))
  );
});
