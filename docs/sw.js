/**
 * みらいコンパス PWA - Service Worker
 * ==============================================================================
 * インストール可能なPWAにするための必須ファイルです。
 * シェル（この外側ページ）の資産のみをキャッシュします。
 * 本体（GAS上のみらいコンパス）は iframe 内で常にネットワークから読み込まれるため、
 * ここではキャッシュしません（常に最新のアプリが表示されます）。
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
const APP_VERSION = 'v2';   // ← リリースごとに必ず上げる
const CACHE_VERSION = CACHE_PREFIX + APP_VERSION;

// キャッシュするシェル資産（すべて相対パス = GitHub Pages のサブパス配信に対応）
const SHELL_ASSETS = [
  './',
  './index.html',
  './manifest.webmanifest',
  './icons/icon-192.png',
  './icons/icon-512.png',
  './icons/icon-maskable-192.png',
  './icons/icon-maskable-512.png',
  './icons/apple-touch-icon.png'
];

// インストール時：シェル資産を事前キャッシュ
self.addEventListener('install', (event) => {
  event.waitUntil(
    caches.open(CACHE_VERSION)
      .then((cache) => cache.addAll(SHELL_ASSETS))
      .then(() => self.skipWaiting())
  );
});

// 有効化時：古いバージョンのキャッシュを削除
self.addEventListener('activate', (event) => {
  event.waitUntil(
    caches.keys()
      .then((keys) => Promise.all(
        keys
          // ← 自アプリ接頭辞のものだけを削除する。ここを外すと
          //    同一オリジンの他アプリを巻き添えにする。
          .filter((k) => k.startsWith(CACHE_PREFIX) && k !== CACHE_VERSION)
          .map((k) => caches.delete(k))
      ))
      .then(() => self.clients.claim())
  );
});

// フェッチ時：同一オリジンのシェル資産のみ「キャッシュ優先 + 裏で更新」
// （GAS・Google系のクロスオリジン通信には一切関与しない）
self.addEventListener('fetch', (event) => {
  const url = new URL(event.request.url);
  if (url.origin !== self.location.origin) return;
  if (event.request.method !== 'GET') return;

  event.respondWith(
    caches.match(event.request).then((cached) => {
      const fetchAndUpdate = fetch(event.request)
        .then((res) => {
          if (res && res.ok) {
            const clone = res.clone();
            caches.open(CACHE_VERSION).then((cache) => cache.put(event.request, clone));
          }
          return res;
        })
        .catch(() => cached); // オフライン時はキャッシュで応答
      return cached || fetchAndUpdate;
    })
  );
});
