/**
 * みらいコンパス PWA - Service Worker
 * ==============================================================================
 * インストール可能なPWAにするための必須ファイルです。
 * シェル（この外側ページ）の資産のみをキャッシュします。
 * 本体（GAS上のみらいコンパス）は iframe 内で常にネットワークから読み込まれるため、
 * ここではキャッシュしません（常に最新のアプリが表示されます）。
 *
 * この Service Worker は localStorage を一切操作しません。
 */

/*
 * 【最重要】activate では自アプリ以外のキャッシュを削除しない。
 *   gigayama.github.io は数十個のアプリが同一オリジンを共有しているため、
 *   CACHE_PREFIX で始まるキャッシュだけを掃除する。
 *   以前はここで caches.keys() の結果を全部消していた。そのため
 *   このアプリを開くたびに、同じ端末に入っている他の GIGA アプリの
 *   キャッシュまで巻き添えで消え、それらがオフラインで起動しなくなっていた。
 */
const CACHE_PREFIX = 'mirai-compass-shell-';
const APP_VERSION = 'v3';   // ← リリースごとに必ず上げる
const CACHE_VERSION = CACHE_PREFIX + APP_VERSION;

// キャッシュするシェル資産（すべて相対パス = GitHub Pages のサブパス配信に対応）
const SHELL_ASSETS = [
  './',
  './index.html',
  './app.js',
  './install-hook.js',
  './offline.html',
  './manifest.webmanifest',
  './icons/icon-192.png',
  './icons/icon-512.png',
  './icons/icon-maskable-192.png',
  './icons/icon-maskable-512.png',
  './icons/apple-touch-icon.png'
];

// インストール時：シェル資産を事前キャッシュ
self.addEventListener('install', (event) => {
  event.waitUntil((async () => {
    const cache = await caches.open(CACHE_VERSION);
    // 1本でも失敗すると addAll 全体が落ちるため、個別に入れる
    await Promise.all(SHELL_ASSETS.map((u) =>
      cache.add(new Request(u, { cache: 'reload' }))
        .catch((err) => console.warn('[sw] precache skipped', u, err))
    ));

    // ★ ここでは skipWaiting しない。
    //   以前はここで呼んでいたため、版を上げると利用者が何も押していないのに
    //   新しい版へ入れ替わっていた（実測：3秒放置で waiting: false、旧キャッシュも消滅）。
    //   先生が児童に説明している最中に画面が入れ替わると混乱するので、
    //   画面側で「さいしんに する」を押してもらってから切り替える。
  })());
});

// 有効化時：古いバージョンのキャッシュを削除
self.addEventListener('activate', (event) => {
  event.waitUntil((async () => {
    const keys = await caches.keys();
    await Promise.all(keys
      // ← 自アプリ接頭辞のものだけを削除する。ここを外すと
      //    同一オリジンの他アプリを巻き添えにする。
      .filter((k) => k.startsWith(CACHE_PREFIX) && k !== CACHE_VERSION)
      .map((k) => caches.delete(k)));
    await self.clients.claim();
  })());
});

// フェッチ時：同一オリジンのシェル資産のみ扱う
// （GAS・Google系のクロスオリジン通信には一切関与しない）
self.addEventListener('fetch', (event) => {
  const req = event.request;
  if (req.method !== 'GET') return;
  const url = new URL(req.url);
  if (url.origin !== self.location.origin) return;

  // 画面遷移は network-first。更新をすぐ届け、圏外ならキャッシュ、
  // それも無ければ offline.html を出す（「壊れた」と思わせないため）。
  if (req.mode === 'navigate') {
    event.respondWith((async () => {
      try {
        return await fetch(req);
      } catch {
        return (await caches.match('./index.html'))
            || (await caches.match('./offline.html'))
            || Response.error();
      }
    })());
    return;
  }

  // 静的ファイルは cache-first + 裏で更新（校内Wi-Fiが混んでいても即表示）
  event.respondWith((async () => {
    const cached = await caches.match(req);
    const fetching = fetch(req).then((res) => {
      if (res && res.ok) {
        const clone = res.clone();
        caches.open(CACHE_VERSION).then((c) => c.put(req, clone));
      }
      return res;
    }).catch(() => cached);
    return cached || fetching;
  })());
});

// 画面側で「さいしんに する」が押されたときだけ切り替える
self.addEventListener('message', (event) => {
  if (event.data && event.data.type === 'SKIP_WAITING') self.skipWaiting();
});
