// Offline support: the app shell and lessons are cached as they are used.
// The student list is always fetched fresh so new logins work straight away.
const CACHE = 'sheetah-v1';
const SHELL = ['./', 'index.html', 'css/app.css', 'css/grid.css', 'vendor/jszip.min.js', 'js/main.js', 'js/app.js', 'js/engine.js', 'js/grid.js', 'js/checker.js', 'js/xlsxcheck.js', 'js/xlsxwrite.js', 'js/store.js', 'js/srs.js', 'js/auth.js', 'js/content.js', 'js/ui.js',
  'js/views/home.js', 'js/views/lesson.js', 'js/views/review.js', 'js/views/study.js', 'js/views/tests.js', 'js/views/library.js', 'js/views/account.js', 'content/catalog.json', 'content/c1/course.json', 'content/library.json', 'manifest.webmanifest', 'img/icon-192.png'];
self.addEventListener('install', e => { e.waitUntil(caches.open(CACHE).then(c => c.addAll(SHELL)).then(() => self.skipWaiting())); });
self.addEventListener('activate', e => { e.waitUntil(caches.keys().then(ks => Promise.all(ks.filter(k => k !== CACHE).map(k => caches.delete(k)))).then(() => self.clients.claim())); });
self.addEventListener('fetch', e => {
  const u = new URL(e.request.url); if (e.request.method !== 'GET') return;
  if (u.pathname.endsWith('/data/students.json') || u.hostname === 'api.github.com') return; // always live
  if (u.origin !== location.origin && !/fonts\.(googleapis|gstatic)\.com$/.test(u.hostname)) return;
  // network first for code so updates arrive; fall back to cache offline
  e.respondWith(fetch(e.request).then(r => { if (r.ok) { const cp = r.clone(); caches.open(CACHE).then(c => c.put(e.request, cp)); } return r; }).catch(() => caches.match(e.request)));
});
