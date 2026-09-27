/* word_docx_editor_pdf Service Worker：App shell 快取，讓第一次載入後可以離線使用。
   版本不必手動維護：install 時會讀 changelog.json 的版本當快取名稱，頁面載入後也會把
   目前版本 postMessage 進來；版本不同就重抓 shell 並把舊快取刪掉。 */
const VERSION = 'unsigned';          // 拿不到 changelog.json 時用的備援版本
let CACHE = 'wde-' + VERSION;
const ASSETS = ['./', './index.html', './docx-fidelity.js', './manifest.webmanifest',
  './vendor/mammoth.browser.min.js', './vendor/html2pdf.bundle.min.js',
  './icons/icon-192.png', './icons/icon-512.png', './icons/apple-touch-icon.png', './icons/favicon-32.png'];

// SW 閒置被回收後重啟，CACHE 會回到預設值；沿用已存在的 shell 快取，免得寫進錯的快取或整包重抓
const adopt = caches.keys().then((ks) => { const k = ks.find((x) => x.startsWith('wde-')); if (k && CACHE === 'wde-' + VERSION) CACHE = k; });

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
  e.waitUntil((async () => {
    await adopt;
    if (want === CACHE) return;
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
  // 自家程式（頁面、更新紀錄、docx 保版面模組）走 network-first，改版才不會拿到舊的；
  // vendor/ 的第三方函式庫維持 cache-first（版本換了才重抓）
  const fresh = req.mode === 'navigate' || url.pathname.endsWith('/index.html')
    || url.pathname.endsWith('changelog.json') || url.pathname.endsWith('/docx-fidelity.js');
  if (fresh) {
    e.respondWith((async () => {
      try {
        const res = await fetch(req);
        if (res.ok) { await adopt; (await caches.open(CACHE)).put(req, res.clone()); }   // 404／500 不要蓋掉好的快取
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
      adopt.then(() => fetch(req)).then((res) => { if (res.ok) caches.open(CACHE).then((c) => c.put(req, res)); }).catch(() => {});
      return hit;
    }
    try {
      const res = await fetch(req);
      await adopt;
      if (res.ok) caches.open(CACHE).then((c) => c.put(req, res.clone()));
      return res;
    } catch { return Response.error(); }
  })());
});
