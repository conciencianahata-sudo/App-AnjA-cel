/* Guarda los archivos de la app para que abra rápido y funcione sin conexión.
   Nunca guarda datos de usuarios ni respuestas de Google: solo archivos de la propia página. */
const VERSION = 'anja-v2';
const ARCHIVOS = ['./', 'index.html', 'app.js', 'core.js', 'drive.js', 'preguntas.js', 'manifest.webmanifest',
  'vendor/exceljs.min.js', 'vendor/jszip.min.js', 'icons/arbol.png', 'icons/icon-192.png', 'icons/icon-512.png'];

self.addEventListener('install', (e) => { e.waitUntil(caches.open(VERSION).then((c) => c.addAll(ARCHIVOS)).then(() => self.skipWaiting())); });
self.addEventListener('activate', (e) => {
  e.waitUntil(caches.keys().then((k) => Promise.all(k.filter((n) => n !== VERSION).map((n) => caches.delete(n)))).then(() => self.clients.claim()));
});
self.addEventListener('fetch', (e) => {
  const u = new URL(e.request.url);
  if (e.request.method !== 'GET' || u.origin !== location.origin) return;
  if (u.pathname.endsWith('/config.js')) return;                              // la configuración siempre se pide a internet (nunca una copia vieja)     // Google y Drive van directo a internet
  e.respondWith(caches.open(VERSION).then(async (c) => {
    const guardado = await c.match(e.request, { ignoreSearch: true });
    const red = fetch(e.request).then((r) => { if (r.ok) c.put(e.request, r.clone()); return r; }).catch(() => guardado);
    return guardado || red;                                                   // rápido desde la copia; se actualiza en segundo plano
  }));
});
