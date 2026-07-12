/**
 * みらいコンパス PWA - Service Worker
 * ==============================================================================
 * インストール可能なPWAにするための必須ファイルです。
 * シェル（この外側ページ）の資産のみをキャッシュします。
 * 本体（GAS上のみらいコンパス）は iframe 内で常にネットワークから読み込まれるため、
 * ここではキャッシュしません（常に最新のアプリが表示されます）。
 */

const CACHE_VERSION = 'mirai-compass-shell-v1';

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
        keys.filter((k) => k !== CACHE_VERSION).map((k) => caches.delete(k))
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
