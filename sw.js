/* Service worker "rete prima" (network first): ad ogni apertura l'app scarica da GitHub l'ultima versione dei file;
   se la rete manca o è lenta (oltre 4 secondi) usa la copia salvata, quindi funziona anche in aereo.
   NESSUN numero di versione da cambiare a mano: basta sostituire i file su GitHub. */
const CACHE = 'assenze-app';            // nome stabile: non va mai cambiato
const CORE = ['./', './index.html', './styles.css', './engine.js', './xlsx.js', './modello.xlsx', './app.js', './manifest.webmanifest',
  './icon-192.png', './icon-512.png', './icon-maskable-512.png', './apple-touch-icon.png', './favicon.png'];
const SDK_HOST = 'https://www.gstatic.com';
const SDK_PATH = '/firebasejs/';        // SDK Firebase dalla CDN Google: file con versione nell'indirizzo, quindi immutabili
const NETWORK_TIMEOUT = 4000;           // soglia normale: oltre questo tempo si usa la copia salvata
const SLOW_TIMEOUT = 700;               // dopo un timeout la rete è considerata lenta: le richieste successive aspettano poco
let slowUntil = 0;

self.addEventListener('install', (e) => {
  e.waitUntil(precache().then(() => self.skipWaiting()));
});

async function precache() {
  const cache = await caches.open(CACHE);
  // 1) i file dell'app, sempre freschi dalla rete (mai dalla cache HTTP del browser)
  await Promise.all(CORE.map(async (u) => {
    try { const r = await fetch(new Request(u, { cache: 'reload' })); if (r.ok) await cache.put(u, r); } catch (err) { /* si riempie alla prima apertura */ }
  }));
  // 2) l'SDK Firebase: la versione si legge da index.html, così non va ripetuta qui
  try {
    const page = await cache.match('./index.html');
    const m = page && /firebasejs\/(\d+\.\d+\.\d+)\//.exec(await page.text());
    if (m) {
      await Promise.all(['app', 'auth', 'firestore'].map(async (n) => {
        const url = SDK_HOST + SDK_PATH + m[1] + '/firebase-' + n + '.js';
        try { const r = await fetch(url, { mode: 'cors' }); if (r.ok) await cache.put(url, r); } catch (err) { /* ritenta a runtime */ }
      }));
    }
  } catch (err) { /* ignora */ }
}

self.addEventListener('activate', (e) => {
  e.waitUntil((async () => {
    // elimina le vecchie cache a numero (v1.x); poi prende subito il controllo delle finestre aperte
    const old = (await caches.keys()).filter((k) => k !== CACHE);
    await Promise.all(old.map((k) => caches.delete(k)));
    await self.clients.claim();
    if (old.length) {                    // aggiornamento da una versione precedente: ricarica le finestre aperte, una volta sola
      const wins = await self.clients.matchAll({ type: 'window' });
      wins.forEach((w) => { try { w.navigate(w.url); } catch (err) { /* ignora */ } });
    }
  })());
});

self.addEventListener('fetch', (e) => {
  const req = e.request;
  if (req.method !== 'GET') return;
  const url = new URL(req.url);
  if (url.origin === self.location.origin) { e.respondWith(networkFirst(e)); return; }
  if (url.origin === SDK_HOST && url.pathname.startsWith(SDK_PATH)) { e.respondWith(cacheFirst(req)); return; }
  // tutto il resto (accesso Google, Firestore, API) passa direttamente dalla rete
});

function cacheKey(req) { return req.mode === 'navigate' ? './index.html' : req; }

function networkFirst(e) {
  const req = e.request, key = cacheKey(req);
  const fresh = req.mode === 'navigate' ? new Request(req.url, { cache: 'no-cache' }) : new Request(req, { cache: 'no-cache' });
  const network = fetch(fresh).then(async (res) => {
    if (res && res.ok) { const c = await caches.open(CACHE); await c.put(key, res.clone()); }
    return res;
  });
  e.waitUntil(network.catch(() => {}));  // se la rete è lenta, finisce di scaricare comunque e aggiorna la cache
  const fromCache = async () => (await caches.open(CACHE)).match(key, { ignoreSearch: true });
  return new Promise((resolve) => {
    let done = false;
    const finish = (r) => { if (!done && r) { done = true; resolve(r); } };
    const wait = Date.now() < slowUntil ? SLOW_TIMEOUT : NETWORK_TIMEOUT;
    const timer = setTimeout(async () => { slowUntil = Date.now() + 60000; finish(await fromCache()); }, wait);
    network.then((res) => { clearTimeout(timer); finish(res); },
      async () => { clearTimeout(timer); finish((await fromCache()) || Response.error()); });
  });
}

async function cacheFirst(req) {
  const cache = await caches.open(CACHE);
  const hit = await cache.match(req);
  if (hit) return hit;
  const res = await fetch(req);
  if (res && res.ok) cache.put(req, res.clone());
  return res;
}
