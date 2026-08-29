// ─── File Protection · Service Worker ────────────────────────────────────────
// Estrategia:
//   • index.html      → Network-first (siempre busca la versión más nueva)
//   • Assets estáticos → Cache-first con fallback a network
//   • APIs Google      → Network-only (nunca cachear)

const CACHE_VERSION = "fp-remitos-v4";
const STATIC_CACHE  = `${CACHE_VERSION}-static`;

// Solo cachear assets que no cambian (librerías externas + fuentes)
const STATIC_ASSETS = [
  "https://fonts.googleapis.com/css2?family=Syne:wght@400;600;700;800&family=DM+Sans:wght@300;400;500&display=swap",
  "https://unpkg.com/react@18/umd/react.production.min.js",
  "https://unpkg.com/react-dom@18/umd/react-dom.production.min.js",
  "https://unpkg.com/@babel/standalone/babel.min.js",
];

// ── INSTALL: cachear solo assets estáticos ───────────────────────────────────
self.addEventListener("install", e => {
  e.waitUntil(
    caches.open(STATIC_CACHE)
      .then(c => c.addAll(STATIC_ASSETS))
      .catch(() => {}) // no fallar si algún asset no carga
  );
  self.skipWaiting(); // activar inmediatamente sin esperar a que cierren las tabs
});

// ── ACTIVATE: limpiar cachés viejos ──────────────────────────────────────────
self.addEventListener("activate", e => {
  e.waitUntil(
    caches.keys().then(keys =>
      Promise.all(
        keys
          .filter(k => k !== STATIC_CACHE)
          .map(k => caches.delete(k))
      )
    )
  );
  self.clients.claim(); // tomar control de todas las tabs abiertas
});

// ── FETCH: estrategia según tipo de recurso ───────────────────────────────────
self.addEventListener("fetch", e => {
  const url = new URL(e.request.url);

  // 1. APIs de Google → siempre network, nunca cachear
  if (
    url.hostname.includes("script.google.com") ||
    url.hostname.includes("googleapis.com") ||
    url.hostname.includes("accounts.google.com")
  ) {
    e.respondWith(
      fetch(e.request).catch(() =>
        new Response(
          JSON.stringify({ error: "Sin conexión" }),
          { headers: { "Content-Type": "application/json" } }
        )
      )
    );
    return;
  }

  // 2. index.html y manifest.json → Network-first
  //    Intenta la red primero; si falla usa caché como fallback
  if (
    url.pathname === "/" ||
    url.pathname.endsWith("/index.html") ||
    url.pathname.endsWith("/manifest.json") ||
    url.pathname.endsWith("/sw.js")
  ) {
    e.respondWith(
      fetch(e.request)
        .then(res => {
          // Guardar la versión nueva en caché
          if (res.ok) {
            const clone = res.clone();
            caches.open(STATIC_CACHE).then(c => c.put(e.request, clone));
          }
          return res;
        })
        .catch(() => caches.match(e.request)) // offline: usar caché
    );
    return;
  }

  // 3. Assets estáticos (React, Babel, fuentes) → Cache-first
  e.respondWith(
    caches.match(e.request).then(cached => {
      if (cached) return cached;
      return fetch(e.request).then(res => {
        if (res.ok && e.request.method === "GET") {
          const clone = res.clone();
          caches.open(STATIC_CACHE).then(c => c.put(e.request, clone));
        }
        return res;
      });
    })
  );
});
