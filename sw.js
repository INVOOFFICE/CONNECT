const CACHE_NAME = 'yaro-admin-v1';
const ASSETS = [
    './', './index.html', './styles/main.css', './scripts/config.js', './scripts/api.js', './scripts/app.js', './scripts/bootstrap.js',
    './assets/pwa.png', './config/manifest.json', './components/splash.html', './sections/app-shell.html', './components/bottom-nav.html',
    './components/car-modal.html', './components/reservation-modal.html', './components/sync-modal.html', './components/settings-modal.html',
    './components/toast-container.html', './components/loading-overlay.html', './components/install-prompt.html'
];

self.addEventListener('install', (event) => {
    event.waitUntil(caches.open(CACHE_NAME).then((cache) => cache.addAll(ASSETS)));
});

self.addEventListener('activate', (event) => {
    event.waitUntil(caches.keys().then((keys) => Promise.all(keys.filter((key) => key !== CACHE_NAME).map((key) => caches.delete(key)))));
});

self.addEventListener('fetch', (event) => {
    if (event.request.method !== 'GET') return;

    if (event.request.mode === 'navigate') {
        event.respondWith(
            caches.match('/index.html').then((cached) => cached || fetch('/index.html'))
        );
        return;
    }

    event.respondWith(caches.match(event.request).then((cached) => cached || fetch(event.request)));
});
