/* word_docx_editor_pdf Service Worker：App shell 快取，讓第一次載入後可以離線使用。
   改版時把 VERSION 加一（或改日期），activate 會把舊快取清掉。 */
const VERSION = '2026-09-27-1';
const CACHE = 'wde-' + VERSION;
const ASSETS = ['./', './index.html', './vendor/mammoth.browser.min.js', './vendor/html2pdf.bundle.min.js'];

self.addEventListener('install', (e) => {
  e.waitUntil(caches.open(CACHE).then((c) => c.addAll(ASSETS)).then(() => self.skipWaiting()));
});

self.addEventListener('activate', (e) => {
  e.waitUntil((async () => {
    for (const k of await caches.keys()) if (k.startsWith('wde-') && k !== CACHE) await caches.delete(k);
    await self.clients.claim();
  })());
});

self.addEventListener('fetch', (e) => {
  const req = e.request;
  if (req.method !== 'GET') return;
  const url = new URL(req.url);
  if (url.origin !== self.location.origin) return;

  // 頁面本身與更新紀錄要拿最新的（「有新版本」提示靠 changelog.json 比對）
  const fresh = req.mode === 'navigate' || url.pathname.endsWith('/index.html') || url.pathname.endsWith('changelog.json');
  if (fresh) {
    e.respondWith((async () => {
      try {
        const res = await fetch(req);
        const c = await caches.open(CACHE);
        c.put(req, res.clone());
        return res;
      } catch {
        return (await caches.match(req)) || (await caches.match('./index.html')) || Response.error();
      }
    })());
    return;
  }

  // 其餘（函式庫等）cache-first，背景更新
  e.respondWith((async () => {
    const hit = await caches.match(req);
    if (hit) {
      fetch(req).then((res) => { if (res.ok) caches.open(CACHE).then((c) => c.put(req, res)); }).catch(() => {});
      return hit;
    }
    try {
      const res = await fetch(req);
      if (res.ok) caches.open(CACHE).then((c) => c.put(req, res.clone()));
      return res;
    } catch { return Response.error(); }
  })());
});
