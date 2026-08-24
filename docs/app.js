/**
 * みらいコンパス シェル（GitHub Pages 側）
 * ==============================================================================
 * GAS のウェブアプリは iframe の中で動くため、そのままでは PWA にできません。
 * そこで GitHub Pages 側にこの外側ページを置き、ここを PWA にしています。
 *
 * ★ このファイルは index.html から外に出してあります。
 *   以前は index.html の中にインラインで書かれ、ボタンも onclick= でした。
 *   その形のままだと CSP（script-src 'self'）を入れた瞬間にアプリが起動しません。
 *   'unsafe-inline' を足して逃げると CSP を入れた意味がほとんど無くなるので、
 *   外部ファイルへ切り出し、onclick= は addEventListener に繋ぎ替えています。
 */
'use strict';

var LS_KEY = 'mirai_compass_gas_url';

// ---------------------------------------------------------------------------
// URLの取得・保存
// ---------------------------------------------------------------------------

function getStoredUrl() {
  try { return localStorage.getItem(LS_KEY) || ''; } catch (e) { return ''; }
}

function isValidGasUrl(url) {
  // GASウェブアプリの正規URLのみ許可（他サイトの埋め込みを防止）
  return /^https:\/\/script\.google\.com\/(a\/[^/]+\/)?macros\/[^\s]+\/exec(\?.*)?$/.test(url);
}

function saveUrl() {
  var input = document.getElementById('gas-url');
  var errEl = document.getElementById('setup-error');
  var url = input.value.trim();

  if (!isValidGasUrl(url)) {
    errEl.textContent = 'URLの形式が正しくありません。https://script.google.com/macros/.../exec の形のURLを入力してください。';
    errEl.style.display = 'block';
    return;
  }
  errEl.style.display = 'none';
  try { localStorage.setItem(LS_KEY, url); } catch (e) {}
  showApp(url);
}

function clearUrl() {
  if (!confirm('保存されているURL設定を削除しますか？')) return;
  // ★ localStorage.clear() は使わない。同じオリジンに他の GIGA アプリが
  //   入っているため、全部消すとそれらの設定まで巻き添えになる。
  try { localStorage.removeItem(LS_KEY); } catch (e) {}
  location.href = location.pathname; // クエリを消してリロード
}

// ---------------------------------------------------------------------------
// 画面切り替え
// ---------------------------------------------------------------------------

// ★ 表示の切り替えは <body> のクラス1つで行う。
//   このページは、URL を覚えていない端末では「導入案内のページ」として読まれ、
//   覚えている端末では全画面のアプリになる。案内はスクロールできないと読めず、
//   アプリは固定されていないと iPad で下端が動くので、CSS 側で分けている
//   （app-mode のときだけ 100dvh + overflow:hidden）。
function showApp(url) {
  var frame = document.getElementById('app-frame');
  if (frame.src !== url) frame.src = url;
  document.body.classList.add('app-mode');
}

function showSetup(prefillUrl) {
  var stored = prefillUrl || getStoredUrl();
  if (stored) {
    document.getElementById('gas-url').value = stored;
    document.getElementById('extra-actions').style.display = 'block';
    document.getElementById('link-open-tab').href = stored;
  }
  var wasApp = document.body.classList.contains('app-mode');
  document.body.classList.remove('app-mode');
  // アプリから設定を開いたときは、案内の先頭ではなく入力欄へ連れていく
  if (wasApp) {
    var start = document.getElementById('start');
    if (start && start.scrollIntoView) start.scrollIntoView();
  }
}

// ---------------------------------------------------------------------------
// 隠し設定ホットスポット（左上すみを2.5秒以内に5回タップ）
// ---------------------------------------------------------------------------

(function () {
  var taps = [];
  var hotspot = document.getElementById('settings-hotspot');
  if (!hotspot) return;
  hotspot.addEventListener('click', function () {
    var now = Date.now();
    taps = taps.filter(function (t) { return now - t < 2500; });
    taps.push(now);
    if (taps.length >= 5) {
      taps = [];
      showSetup();
    }
  });
})();

// ---------------------------------------------------------------------------
// オフライン表示
// ---------------------------------------------------------------------------

function updateOnlineStatus() {
  document.getElementById('offline-bar').style.display = navigator.onLine ? 'none' : 'block';
}
window.addEventListener('online', updateOnlineStatus);
window.addEventListener('offline', updateOnlineStatus);

// ---------------------------------------------------------------------------
// ボタンの結線（onclick= を使わない）
// ---------------------------------------------------------------------------

document.getElementById('btn-save').addEventListener('click', saveUrl);
document.getElementById('btn-clear').addEventListener('click', clearUrl);

// ---------------------------------------------------------------------------
// インストール
//   ★ 案内できるときだけボタンを出す。
//     出せないボタンを置いておくと「押しても何も起きない」と言われる。
// ---------------------------------------------------------------------------

var installBtn = document.getElementById('btn-install');

function isStandalone() {
  return matchMedia('(display-mode: standalone)').matches || navigator.standalone === true;
}

function refreshInstallButton() {
  if (!installBtn) return;
  var canPrompt = !!window.__pwaInstallPrompt;
  installBtn.style.display = (canPrompt && !isStandalone()) ? 'inline-flex' : 'none';
}

window.addEventListener('pwa-install-available', refreshInstallButton);
window.addEventListener('pwa-installed', refreshInstallButton);
refreshInstallButton();

if (installBtn) {
  installBtn.addEventListener('click', function () {
    var p = window.__pwaInstallPrompt;
    if (!p) return;
    p.prompt();
    p.userChoice.finally(function () {
      window.__pwaInstallPrompt = null;
      refreshInstallButton();
    });
  });
}

// iPhone / iPad は beforeinstallprompt が無いので、手順を案内する
var iosHint = document.getElementById('ios-hint');
if (iosHint) {
  var isIos = /iPad|iPhone|iPod/.test(navigator.userAgent) ||
    (navigator.platform === 'MacIntel' && navigator.maxTouchPoints > 1);
  if (isIos && !isStandalone()) iosHint.style.display = 'block';
}

// ---------------------------------------------------------------------------
// 起動処理
// ---------------------------------------------------------------------------

(function init() {
  updateOnlineStatus();

  var params = new URLSearchParams(location.search);

  // ?app=<GASのURL> でURLを配布・事前設定できる（例: 先生がQRコードで配る）
  var presetUrl = params.get('app');
  if (presetUrl && isValidGasUrl(presetUrl)) {
    try { localStorage.setItem(LS_KEY, presetUrl); } catch (e) {}
    // クエリを消してリロード（インストール後のstart_urlと一致させる）
    location.replace(location.pathname);
    return;
  }

  // ?settings=1 で設定画面を強制表示
  if (params.get('settings') === '1') {
    showSetup();
    return;
  }

  var stored = getStoredUrl();
  if (stored && isValidGasUrl(stored)) {
    showApp(stored);
  } else {
    showSetup();
  }
})();

// ---------------------------------------------------------------------------
// Service Worker の登録と、更新の案内
// ---------------------------------------------------------------------------

// ★ controllerchange は、はじめて開いたときにも飛んでくる。
//   activate の clients.claim() でページが管理下に入るためである。
//   これを素直に受けると「初回訪問が必ず1回リロードされる」ことになり、
//   先生が入力しかけていた URL が消える。
//   「もともと管理下だったか」で分ける直し方は別の形で壊れる
//   （入れた直後に更新を押した場合、切り替わったのに読み込み直されない）。
//   見るべきは【利用者が押したかどうか】だけ。
var userAskedUpdate = false;
var reloading = false;

function showUpdateBar(worker) {
  var bar = document.getElementById('update-bar');
  var btn = document.getElementById('btn-update');
  if (!bar || !btn) return;
  bar.hidden = false;
  btn.addEventListener('click', function () {
    userAskedUpdate = true;
    bar.hidden = true;
    worker.postMessage({ type: 'SKIP_WAITING' });
  }, { once: true });
}

function startServiceWorker() {
  navigator.serviceWorker.register('sw.js').then(function (registration) {
    navigator.serviceWorker.addEventListener('controllerchange', function () {
      if (!userAskedUpdate || reloading) return;
      reloading = true;
      location.reload();
    });

    registration.addEventListener('updatefound', function () {
      var sw = registration.installing;
      if (!sw) return;
      sw.addEventListener('statechange', function () {
        // controller が居る＝初回インストールではなく更新。
        // 初回で通知すると「入れた直後に更新があります」と出て混乱する。
        if (sw.state === 'installed' && navigator.serviceWorker.controller) showUpdateBar(sw);
      });
    });

    // 前回のうちに入っていた場合も拾う
    if (registration.waiting && navigator.serviceWorker.controller) showUpdateBar(registration.waiting);
  }).catch(function (e) {
    console.warn('Service Worker registration failed:', e);
  });
}

if ('serviceWorker' in navigator) {
  // ★ 'load' を待つだけだと、すでに load が済んでいる場合にリスナーが
  //   二度と呼ばれず、Service Worker が登録されないままになる。
  //   済んでいるならその場で走らせる。
  if (document.readyState === 'complete') startServiceWorker();
  else window.addEventListener('load', startServiceWorker, { once: true });
}
