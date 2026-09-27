/* word_docx_editor_pdf Service Worker：App shell 快取，讓第一次載入後可以離線使用。
   版本不必手動維護：install 時會讀 changelog.json 的版本當快取名稱，頁面載入後也會把
   目前版本 postMessage 進來；版本不同就重抓 shell 並把舊快取刪掉。 */
const VERSION = 'unsigned';          // 拿不到 changelog.json 時用的備援版本
let CACHE = 'wde-' + VERSION;
const ASSETS = ['./', './index.html', './vendor/mammoth.browser.min.js', './vendor/html2pdf.bundle.min.js'];

const cacheName = (v) => 'wde-' + String(v).replace(/[^0-9A-Za-z._-]/g, '-');

// 安裝時先問 changelog.json 目前是哪一版，這樣就不必手動維護版本
async function currentVersion() {
  try {
    const r = await fetch('./changelog.json', { cache: 'no-store' });
    const log = await r.json();
    if (log && log[0] && log[0].version) return log[0].version;
  } catch {}
  return VERSION;
}

self.addEventListener('install', (e) => {
  e.waitUntil((async () => {
    CACHE = cacheName(await currentVersion());
    await caches.open(CACHE).then((c) => c.addAll(ASSETS));
    await self.skipWaiting();
  })());
});

// 頁面載入後會把 changelog.json 的版本送上來；版本變了就換一個快取並重抓 shell
self.addEventListener('message', (e) => {
  const v = e.data && e.data.version;
  if (!v) return;
  const want = cacheName(v);
  if (want === CACHE) return;
  e.waitUntil((async () => {
    const fresh = await caches.open(want);
    await fresh.addAll(ASSETS);
    const old = CACHE;
    CACHE = want;
    if (old && old !== want) await caches.delete(old);
    await self.clients.claim();
  })());
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
