/*
 * インストールの合図を「いちばん先に」受け取る。
 *
 * Chrome は条件が揃うと即座に beforeinstallprompt を出す。
 * ほかの読み込みより後にこの登録を書くと合図を取りこぼし、
 * 通信が遅い端末で「インストール」ボタンが出なくなる。
 *
 * CSP に 'unsafe-inline' を足さずに済むよう、インラインではなく
 * 小さな外部ファイルにして <head> の先頭で同期読み込みする。
 */
(function () {
  window.__pwaInstallPrompt = null;

  window.addEventListener('beforeinstallprompt', function (e) {
    // 既定のミニ情報バーを止めて、アプリ側の好きなタイミングで出す
    e.preventDefault();
    window.__pwaInstallPrompt = e;
    window.dispatchEvent(new Event('pwa-install-available'));
  });

  window.addEventListener('appinstalled', function () {
    window.__pwaInstallPrompt = null;
    window.dispatchEvent(new Event('pwa-installed'));
  });
})();
