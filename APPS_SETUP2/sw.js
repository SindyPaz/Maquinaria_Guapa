// Service Worker - Labores Maquinaria PWA
// Versión: cambia este número para forzar actualización en todos los dispositivos
var CACHE_VERSION = 'labores-v1.2';

var APP_SHELL = [
  './APP2_Labores.html',
  './manifest.json',
  './icon-192.png',
  './icon-512.png',
  './icon-apple.png',
  './logo_tractor.png'
];

// ── INSTALAR: guarda el shell en caché ──────────────────────────────────────
self.addEventListener('install', function(e) {
  e.waitUntil(
    caches.open(CACHE_VERSION).then(function(cache) {
      return cache.addAll(APP_SHELL);
    }).then(function() {
      return self.skipWaiting();
    })
  );
});

// ── ACTIVAR: borra cachés viejas ────────────────────────────────────────────
self.addEventListener('activate', function(e) {
  e.waitUntil(
    caches.keys().then(function(keys) {
      return Promise.all(
        keys.filter(function(k) { return k !== CACHE_VERSION; })
            .map(function(k) { return caches.delete(k); })
      );
    }).then(function() {
      return self.clients.claim();
    })
  );
});

// ── FETCH: Network-first para sync con Google, Cache-first para el shell ────
self.addEventListener('fetch', function(e) {
  var url = e.request.url;

  // Las llamadas al API de Google siempre van a la red (sin caché)
  if (url.includes('script.google.com')) {
    e.respondWith(fetch(e.request).catch(function() {
      return new Response(JSON.stringify({ ok: false, error: 'Sin conexion' }),
        { headers: { 'Content-Type': 'application/json' } });
    }));
    return;
  }

  // Para todo lo demás: Cache-first (funciona offline)
  e.respondWith(
    caches.match(e.request).then(function(cached) {
      if (cached) return cached;
      return fetch(e.request).then(function(response) {
        // Guarda en caché si es una respuesta válida
        if (response && response.status === 200 && response.type === 'basic') {
          var clone = response.clone();
          caches.open(CACHE_VERSION).then(function(cache) {
            cache.put(e.request, clone);
          });
        }
        return response;
      }).catch(function() {
        // Si falla la red y no hay caché, devuelve el HTML principal
        return caches.match('./APP2_Labores.html');
      });
    })
  );
});

// ── MENSAJE: forzar actualización desde la app ──────────────────────────────
self.addEventListener('message', function(e) {
  if (e.data && e.data.action === 'skipWaiting') {
    self.skipWaiting();
  }
});
