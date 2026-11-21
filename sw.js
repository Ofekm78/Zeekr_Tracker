/*
  Zeekr electricity tracker - offline cache for the app's own files (main page).

  Goal: open the installed app instantly instead of downloading the page from
  the site on every launch.

  How it works ("stale while revalidate"):
    - Every file of this page (index.html, app.js, styles.css, icons, manifest)
      is answered at once from the phone's cache when it is there.
    - At the same time the file is fetched from the site in the background and
      the cache is refreshed, so a new version shows on the NEXT launch.
    - The first launch after installing fills the cache from the network.

  Only same-origin GET requests inside this folder are handled, except the
  trial copy in beta/, which has its own sw.js and cache. Microsoft
  sign-in, OneDrive and Google Fonts go straight to the network, untouched,
  so measurements are never served from this cache.

  Raise CACHE_VERSION to throw away everything cached by an older version.
*/
const CACHE_VERSION = 1;
const CACHE = "zk-main-v" + CACHE_VERSION;
const SCOPE = self.registration.scope; // e.g. https://<site>/Zeekr_Tracker/
const SKIP = "beta/";

self.addEventListener("install", (event) => {
  // Cache the page itself right away so the next launch is instant.
  event.waitUntil(caches.open(CACHE).then((cache) => cache.add(SCOPE)).catch(() => {}));
  self.skipWaiting();
});

self.addEventListener("activate", (event) => {
  event.waitUntil((async () => {
    const names = await caches.keys();
    await Promise.all(names.filter((n) => n.startsWith("zk-main-") && n !== CACHE).map((n) => caches.delete(n)));
    await self.clients.claim();
  })());
});

self.addEventListener("fetch", (event) => {
  const request = event.request;
  if (request.method !== "GET" || !request.url.startsWith(SCOPE)) return; // network as usual
  if (SKIP && request.url.startsWith(SCOPE + SKIP)) return; // another copy with its own cache

  // Opening the page (including the return from Microsoft sign-in with
  // ?code=... in the address) is always index.html, stored under SCOPE.
  const path = request.url.split(/[?#]/)[0];
  const isPage = request.mode === "navigate";
  if (isPage && path !== SCOPE && path !== SCOPE + "index.html") return;
  const key = isPage ? SCOPE : request;

  event.respondWith((async () => {
    const cache = await caches.open(CACHE);
    const cached = await cache.match(key);
    const fresh = fetch(isPage ? SCOPE : request, { cache: "no-cache" })
      .then((response) => {
        if (response.ok) cache.put(key, response.clone());
        return response;
      });
    if (cached) {
      event.waitUntil(fresh.catch(() => {})); // refresh in the background
      return cached;
    }
    return fresh;
  })());
});
