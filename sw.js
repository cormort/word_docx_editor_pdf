/* word_docx_editor_pdf Service Worker：App shell 快取，讓第一次載入後可以離線使用。
   版本不必手動維護：install 時會讀 changelog.json 的版本當快取名稱，頁面載入後也會把
   目前版本 postMessage 進來；版本不同就重抓 shell 並把舊快取刪掉。 */
const VERSION = 'unsigned';          // 完全拿不到 changelog.json、本機又沒有任何快取時才用的備援名稱
let CACHE = 'wde-' + VERSION;
const ASSETS = ['./', './index.html', './docx-fidelity.js', './pdf-export.js', './manifest.webmanifest',
  './changelog.json',
  './vendor/mammoth.browser.min.js', './vendor/html2pdf.bundle.min.js', './vendor/pdf-lib.esm.min.js',
  './icons/icon-192.png', './icons/icon-512.png', './icons/icon-192-maskable.png',
  './icons/icon-512-maskable.png', './icons/apple-touch-icon.png', './icons/favicon-32.png',
  './icons/favicon-16.png'];
// 這幾個缺了就沒有離線可用性；圖示之類少了只影響主畫面圖案，不該讓整個安裝失敗
const ESSENTIAL = ['./index.html', './docx-fidelity.js', './pdf-export.js',
  './vendor/mammoth.browser.min.js', './vendor/html2pdf.bundle.min.js', './vendor/pdf-lib.esm.min.js'];

const cacheName = (v) => 'wde-' + String(v).replace(/[^0-9A-Za-z._-]/g, '-');

// SW 閒置被回收後重啟，CACHE 會回到預設值；沿用已存在的 shell 快取（名稱最新的那個），
// 免得寫進錯的快取或整包重抓。無痕模式／儲存空間被關掉時 caches 會直接丟錯，不能讓它炸掉整支 SW
const adopt = Promise.resolve()
  .then(() => caches.keys())
  .catch(() => [])
  .then((ks) => {
    const all = ks.filter((x) => x.startsWith('wde-')).sort();   // 版本字串照時間排序，最後一個最新
    if (all.length && CACHE === 'wde-' + VERSION) CACHE = all[all.length - 1];
  });

// 安裝時先問 changelog.json 目前是哪一版，這樣就不必手動維護版本
async function currentVersion() {
  try {
    const r = await fetch('./changelog.json', { cache: 'no-store' });
    const log = await r.json();
    if (log && log[0] && log[0].version) return log[0].version;
  } catch {}
  return null;
}

// 逐檔加入：addAll 只要有一個檔案 404 就整包失敗，離線快取會永遠建不起來而且沒有任何訊息
async function fillCache(name) {
  const c = await caches.open(name);
  const rs = await Promise.allSettled(ASSETS.map((u) => c.add(new Request(u, { cache: 'reload' }))));
  const failed = ASSETS.filter((u, i) => rs[i].status === 'rejected');
  if (failed.length) console.warn('[sw] 這些檔案沒能放進離線快取：' + failed.join('、'));
  return { ok: !failed.some((u) => ESSENTIAL.includes(u)), failed };
}

self.addEventListener('install', (e) => {
  e.waitUntil((async () => {
    await adopt;
    const before = CACHE;
    const v = await currentVersion();
    if (v) CACHE = cacheName(v);            // 讀不到版本時沿用現有快取名稱，不要發明一個 wde-unsigned 蓋掉好的快取
    const isNew = CACHE !== before;
    let r = { ok: false };
    try {
      r = await fillCache(CACHE);
    } catch (err) {
      console.warn('[sw] 離線快取建立失敗：' + err);
    }
    if (!r.ok) {
      if (isNew) await caches.delete(CACHE).catch(() => {});   // 只清掉這次新建的半成品，不動原本可用的快取
      throw new Error('[sw] 離線快取不完整，保留舊版 service worker');
    }
    await self.skipWaiting();
  })());
});

// 頁面載入後會把 changelog.json 的版本送上來；版本變了就換一個快取並重抓 shell。
// 只接受「跟 SW 自己讀到的版本一致」的訊息：不然任何一個分頁（或版本回滾）都能叫 SW 砍掉目前這份快取
self.addEventListener('message', (e) => {
  const v = e.data && e.data.version;
  if (!v) return;
  e.waitUntil((async () => {
    await adopt;
    const want = cacheName(v);
    if (want === CACHE) return;
    if (want !== cacheName(await currentVersion())) return;
    let r = { ok: false };
    try {
      r = await fillCache(want);
    } catch (err) {
      console.warn('[sw] 離線快取更新失敗：' + err);
    }
    if (!r.ok) {
      await caches.delete(want).catch(() => {});
      return;                                  // 新的建不起來就繼續用舊的，至少離線還開得起來
    }
    CACHE = want;
    await self.clients.claim();
  })());
});

self.addEventListener('activate', (e) => {
  e.waitUntil((async () => {
    await adopt;
    const keys = await caches.keys().catch(() => []);
    for (const k of keys) if (k.startsWith('wde-') && k !== CACHE) await caches.delete(k).catch(() => {});
    await self.clients.claim();
  })());
});

self.addEventListener('fetch', (e) => {
  const req = e.request;
  if (req.method !== 'GET') return;
  const url = new URL(req.url);
  if (url.origin !== self.location.origin) return;

  // 頁面本身與更新紀錄要拿最新的（「有新版本」提示靠 changelog.json 比對）
  // 自家程式（頁面、更新紀錄、docx 保版面模組、PDF 模組）走 network-first，改版才不會拿到舊的；
  // vendor/ 的第三方函式庫維持 cache-first（版本換了才重抓）
  const fresh = req.mode === 'navigate' || url.pathname.endsWith('/index.html')
    || url.pathname.endsWith('changelog.json') || url.pathname.endsWith('/docx-fidelity.js') || url.pathname.endsWith('/pdf-export.js');
  if (fresh) {
    e.respondWith((async () => {
      try {
        const res = await fetch(req);
        if (res.ok) {
          // 404／500 不要蓋掉好的快取；快取寫入失敗（空間不足）也不該讓頁面拿不到回應
          e.waitUntil((async () => {
            try {
              await adopt;
              const c = await caches.open(CACHE);
              await c.put(req, res.clone());
            } catch {}
          })());
        }
        return res;
      } catch {
        // 離線：先找回原本的請求，再退回 app shell。caches 本身丟錯時也要有東西回傳
        try {
          return (await caches.match(req)) || (await caches.match('./index.html')) || Response.error();
        } catch {
          return Response.error();
        }
      }
    })());
    return;
  }

  // 其餘（函式庫等）cache-first，背景更新
  e.respondWith((async () => {
    let hit = null;
    try {
      hit = await caches.match(req);
    } catch {}
    if (hit) {
      e.waitUntil((async () => {
        try {
          await adopt;
          const res = await fetch(req);
          if (res.ok) (await caches.open(CACHE)).put(req, res);
        } catch {}
      })());
      return hit;
    }
    try {
      const res = await fetch(req);
      if (res.ok) {
        e.waitUntil((async () => {
          try {
            await adopt;
            (await caches.open(CACHE)).put(req, res.clone());
          } catch {}
        })());
      }
      return res;
    } catch {
      return Response.error();
    }
  })());
});
