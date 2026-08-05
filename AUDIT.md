# 🧭 みらいコンパス — GIGA Standard v5 監査

- 監査日: 2026-08-05
- 対象コミット: `main`
- アーキテクチャ: **C+型**（GAS ウェブアプリ本体 + `docs/` の GitHub Pages シェル）
- 監査者: GIGA Standard v5 Rollout Engineer

> **「未計測」は ✅ ではない。** 測っていないものは「未計測」と書いてある。
> 数字はすべて実ブラウザ（Chromium 1366×768 / DPR 2 / 320×568）での実測値である。

---

## 0. どうやって測ったか

GAS の本番（`script.google.com`）へは作業環境から到達できない。
しかし **表示は測れる**（v5 §7-3）。GAS が返す画面は `index.html` に
`include('css') / include('js_core') / include('js_student') / include('js_teacher')`
を貼り合わせたものなので、同じ貼り合わせを手元で行った。

1. `include()` を実体に置き換える
2. `code.gs` の `doGet` にある `addMetaTag('viewport', …)` を反映する
   （**サーバー側の処理なので、`index.html` を組み立てただけでは再現されない**）
3. `google.script.run` をダミーに差し替え、**戻り値の見本**を与えて本編まで進める
4. CDN 資産は npm から同じ版を取り、jsDelivr と同じパスに並べた**検査用の複製**へ向ける
   （§7-4：塞がれたまま測ると Bootstrap が当たらない素の HTML を測ることになり、数字が全部でたらめになる）
   - ミラーには `Access-Control-Allow-Origin: *` を付けた
   - **Google Fonts はわざと塞いだまま**にした。フィルタリングされた学校と同じ状態で測るため

**歩いた画面は 18 画面。** 児童：なまえ入力／きょうの学習／計画をつくる／ポートフォリオ／タイマー／
ギャラリーウォーク。先生：LIVE／一覧／座席／まとめ／設計／進行／設定／単元情報ほか。
加えて 320×568 で 2 画面。

### 測っていないもの（重要）

| 測っていないもの | なぜ |
|---|---|
| 本番 GAS の動作 | `script.google.com` へ到達できない。デプロイ・差分確認ができない |
| OAuth スコープ | **`appsscript.json` がリポジトリに含まれていない**。宣言されているスコープを読めない |
| サーバー側の実際の権限挙動 | 同上。5段ガードの不備は**コードから読み取った**もので、実行して確かめてはいない |
| `gigayama.github.io` 上での PWA 実挙動 | プロキシ 403 で到達できない。PWA は `docs/` をローカル配信して測った |
| キーボードのみでの全機能到達（F3） | 今回は走査していない |
| 印刷プレビュー | 印刷 CSS が存在しないため、比較対象がない |

---

## A. 法務・配布

| # | 項目 | 判定 | 実測 |
|---|---|:--:|---|
| A1 | LICENSE 実ファイル | ❌ | 無い |
| A2 | .gitignore | ❌ | 無い |
| A3 | dependabot.yml | ❌ | `.github/` ディレクトリ自体が無い |
| A4 | README.md / MANUAL.md / AUDIT.md | △ | README.md のみ（14節）。MANUAL.md・AUDIT.md は無い |
| A5 | CI（`pull_request` でも動く） | ❌ | ワークフローが無い |

---

## B. セキュリティ

| # | 項目 | 判定 | 実測 |
|---|---|:--:|---|
| B1 | CSP（入れたうえで動作確認済み） | ❌ | GAS 本体・`docs/` シェルとも `Content-Security-Policy` が無い |
| B2 | 秘密情報・IDの直書きなし | △ | 秘密は `PropertiesService` 経由で正しい。ただし `code.gs:79` の `setFaviconUrl` に Google Drive のファイル ID が直書き（公開ファイルなので秘密ではないが、直書きではある） |
| B3 | OAuthスコープ最小 | **未計測** | `appsscript.json` がリポジトリに無く、宣言スコープを読めない。コード上は `SpreadsheetApp.openById`（7箇所）・`PropertiesService`・`Utilities` のみで、`DriveApp` / `MailApp` / `GmailApp` / `UrlFetchApp` は**不使用** |
| B4 | postMessage の宛先が `*` でない | ✅ | `postMessage` を使っていない（0件） |
| B5 | サーバー側5段ガード | ❌ | **後述。いちばん重い指摘** |
| B6 | CDN から取る実行コードが 0 | ❌ | **6本**。後述 |
| B7 | 残る外部資産に SRI と版の固定 | ❌ | SRI **0件**。うち2本は版が浮いている |

### B5 — 教員用 API に本人確認が無い

`code.gs` の関数は `google.script.run` からそのまま呼べる。
教員専用の操作にサーバー側の確認が一切入っていない。

| 関数 | 何ができるか |
|---|---|
| `changeTeacherPassword(newPass)` | **教員パスワードを誰でも書き換えられる** |
| `saveClassRoster(classId, list)` | 名簿の入れ替え |
| `deleteUnitTask(taskId)` / `archiveUnitData(unitId, …)` | 単元・課題の削除 |
| `savePortfolioFeedback(…)` / `saveAllPortfolios(…)` | 先生からの評価の書き換え |
| `initSystem()` | **パスワードを `admin` に戻す** |

画面側の `attemptTeacherLogin()` は `verifyPassword()` の戻り値 `{authenticated:true/false}` で
UI を出し分けているだけである。v5 Phase 4-2 の通り、**フロントの出し分けは防御ではない。**
初期パスワードは `admin` 固定で、初期設定画面に大きく表示される（`js_core.html:270`）。

**これは「直し方が1つに決まらない」問題である。**
デプロイの実行者設定（アクセスユーザー／アプリアカウント）と、
児童が Google アカウントを持つかどうかで、正しい形が変わる。
本番で確かめられない以上、**このロールアウトでは変更しない。**
v5 §停止条件（アーキテクチャの変更／本番で確かめられない）に該当するため、
**別 PR での提案に留める。**

### B6 — CDN から取る実行コードが 6本ある（**最優先**）

```
index.html:23   bootstrap@5.3.0/dist/css/bootstrap.min.css       232.9 KB
index.html:25   bootstrap-icons@1.11.1/font/bootstrap-icons.css   98.3 KB (+ woff2 130.6 KB)
index.html:113  bootstrap@5.3.0/dist/js/bootstrap.bundle.min.js   80.4 KB
index.html:114  sweetalert2@11                       ← 版が浮いている  79.2 KB
index.html:115  chart.js                             ← 版指定が無い   208.5 KB
index.html:116  sortablejs@1.15.0/Sortable.min.js                  44.1 KB
```

**塞がれた状態で実際に開いて測った。**（`cdn.jsdelivr.net` を全部落として起動）

| | 通ったとき | 塞がれたとき |
|---|---|---|
| `window.bootstrap` | あり | **undefined** |
| `window.Swal` | あり | **undefined**（呼び出し **39箇所**） |
| `window.Chart` | あり | **undefined** |
| `window.Sortable` | あり | **undefined**（計画づくりの D&D **3箇所**） |
| スタイルシート枚数 | 5枚 | **4枚**（Bootstrap が当たっていない） |
| ローディング画面 | 消える | **消えないまま残る** |
| 画面 | 通常 | **素の HTML。`d-none`（13箇所）が効かず、児童画面と先生用ボタンが同時に出る** |

つまり学校のフィルタリング下では、**白い画面ではなく「崩れた画面が半分動く」**という、
より分かりにくい壊れ方をする。児童からは「壊れている」としか見えず、
**原因はアプリの外にあるので先生が調べても分からない。**

### B7 — 版が浮いている2本

```html
<script src="https://cdn.jsdelivr.net/npm/sweetalert2@11"></script>  <!-- ❌ メジャー内で勝手に上がる -->
<script src="https://cdn.jsdelivr.net/npm/chart.js"></script>        <!-- ❌ 版指定が無い。メジャーも上がる -->
```

npm で解決すると、いま `sweetalert2@11` は **11.26.25**、`chart.js` は **4.5.1** になる。
中身が変わるので SRI を付けられない。**ある日突然壊れる形である。**

---

## C. 堅牢性

| # | 項目 | 判定 | 実測 |
|---|---|:--:|---|
| C1 | LockService + try/finally（GAS） | ✅ | 書き込み系 12関数すべてに `LockService.getScriptLock()` + `finally { lock.releaseLock(); }` |
| C2 | 自動復旧（シート再生成） | ✅ | `checkAndFixSheets(ss)` が `getData()` の冒頭で毎回シート構造を点検・再生成 |
| C3 | pagehide で記録確定 | ❌ | `pagehide` / `beforeunload` / `visibilitychange` **0件**。Chromebook のタブ破棄で打ちかけが消える |
| C4 | 通信失敗時のリトライと明示 | ❌ | `withFailureHandler(handleError)` で止まるだけ。再試行なし。10秒ポーリングの失敗は `console.error` のみで画面に出ない |
| C5 | localStorage.clear() を使っていない | ✅ | 0件。`removeItem` で個別に消している |

---

## D. 表示（Part I §2）

| # | 項目 | 判定 | 実測 |
|---|---|:--:|---|
| D1 | viewport に viewport-fit=cover（GASは code.gs も） | ❌ | `index.html:9` と `code.gs:76` の**両方**に無い |
| D2 | 100dvh を使用 | ❌ | `css.html:41` `height: 100vh` / `js_student.html:615` `calc(100vh - 150px)` |
| D3 | safe-area-inset を適用 | △ | **1箇所のみ**（`css.html:100` footer の `padding-bottom`）。左右パディング・上部は無い |
| D4 | clamp() による fluid type | ❌ | **0件**。`rem`/`px` 固定のみ |
| D5 | Canvas に devicePixelRatio 補正（上限2） | ❌ | `js_student.html:1062` の `getContext('2d')` は Chart.js に渡すだけ。`devicePixelRatio` の記述 **0件**。3倍端末で 9倍の面積を描き、メモリ4GBの Chromebook でタブが落ちうる |
| D6 | 320px 幅で横スクロールが出ない | ✅ | **`scrollWidth 320 = clientWidth 320`。横スクロール 0件** |
| D7 | 画像に width/height、150KB以下 | △ | `<img>` は `docs/index.html:153` の1つのみで `width`/`height` 属性が無い（CSS 指定）。150KB 超の画像は **0件**（最大 `icon-512.png` 48.5KB） |
| D8 | **コントラスト 4.5:1 以上** | ❌ | **36件**。後述 |
| D9 | **タップ領域 44px 以上** | ❌ | **62件**。後述 |
| D10 | prefers-reduced-motion 対応 | ❌ | **0件**。`pulse` / `heartBeat` / `flash` / `float` / `cuteSpin` が無限ループで回り続ける |
| D11 | forced-colors 対応 | ❌ | **0件** |
| D12 | 提示モード | ❌ | 無い。一斉授業（先生の LIVE 画面を電子黒板に映す）で使うアプリなので必要 |
| D13 | 印刷CSS | ❌ | `@media print` **0件**。まとめ一覧・ポートフォリオは印刷したい画面である |
| D14 | 拡大を禁止していない | ✅ | `user-scalable=no` / `maximum-scale` **0件**（`index.html` / `code.gs` とも） |

### D8 — コントラスト 36件（全画面走査）

比の悪い順。**背景は白ではなく `--bg-body: #f0f2f5` なので、v5 §2-8 の白地の表より少しずつ低く出る。**

| 比 | 必要 | 出どころ | 色 | 面 | どこ |
|---:|---:|---|---|---|---|
| **1.04** | 4.5 | **`rt`（ふりがな）** | `#6c757d` | `#0d6efd`（青ボタン） | 「かんりょう」「かいし」 |
| 1.61 | 3 | SweetAlert2 既定 `.swal2-close` | `#ccc` | `#fff` | ポップアップの × |
| **1.96** | 4.5 | Bootstrap 既定 `.btn-info` | `#fff` | `#0dcaf0` | 「集中」ボタン |
| 4.01 | 4.5 | Bootstrap 既定 `.text-primary` | `#0d6efd` | `#f0f2f5` | 児童名・「計画」「課題」・戻る（5件） |
| 4.04 | 4.5 | Bootstrap 既定 `.btn-outline-success` | `#198754` | `#f0f2f5` | 「ふりかえり」 |
| 4.18 | 4.5 | `--text-muted` / `.text-secondary` | `#6c757d` | `#f0f2f5` | バッジ・時数・カテゴリ名ほか（**24件**） |
| 4.38 | 4.5 | `--text-muted` | `#6c757d` | `#fbf5ff` | マイタスクの「15分」 |
| 4.45 | 4.5 | `--text-muted` | `#6c757d` | `#f8f9fa` | ギャラリーの空セル `-` |

#### いちばん重いのは `rt`（ふりがな）の 1.04

```css
/* css.html:51 */
rt { font-size: 0.6em; color: var(--text-muted); }   /* = #6c757d 決め打ち */
```

`#6c757d` を青ボタン `#0d6efd` の上に載せると **比 1.04**。ほぼ見えない。
v5 §4 が挙げている実例（1.28 / 1.47）より悪い。

**ふりがなが必要なのは低学年の児童である。**
つまり**いちばん読めなくて困る人が、いちばん読めない**という形になっている。

該当は `<button class="btn btn-primary">` の中の `<ruby>` で、
「完了」「保存する」「開始」の3箇所（`js_student.html`）。
**1か所ずつ潰すのではなく、色のついた面ではまとめて `inherit` させるのが正しい。**

#### アプリ固有の配色よりフレームワークの既定色が先

`css.html` は `--color-primary: #1a73e8` を定義しているが、
**Bootstrap の `.text-primary` / `.btn-info` / `.btn-outline-success` を上書きしていない。**
そのため画面に出ているのは Bootstrap の既定色（`#0d6efd` / `#0dcaf0` / `#198754`）である。
v5 §2-8 の通り、**アプリ固有の配色を疑う前に、フレームワークの既定色を疑う。**

なお `#1a73e8` 自体も v5 §2-8 の表にある通り 4.27 で、白地でも白抜きでも基準に届いていない。

### D9 — タップ領域 62件

`::after` の当たり判定込みで実測した（疑似要素による拡張は**0件**なので、見たままの寸法）。

代表例（小さい順）:

| 寸法 | 要素 | 何のボタンか |
|---|---|---|
| 16×16 | `input.form-check-input` | 「情報を保存する」のチェックボックス |
| 16×24 | `i.bi-x` | 計画から課題を外す × |
| 24×24 | `button.btn-close` | タイマーを閉じる |
| 31×31 | `button.btn-outline-secondary` | 単元情報 ⓘ |
| 32×31 | 3種 | 設計画面の丸ボタン群 |
| 49×16 | `a` GIGA山 | フッターのリンク（v5 Phase 3 が名指ししている箇所） |
| 84×18 | `a` 教員用ログイン | フッター |
| 46〜60×31 | `button.btn-sm` ×6 | 先生のビュー切替（LIVE／一覧／座席／まとめ／設計／進行） |
| 62〜74×33 | `button.btn-sm` | タイマーの開始／リセット |
| 102×40 | `button.nav-link` ×3 | ポートフォリオのタブ |

**ボタン自体を大きくすると、詰めて組んであるツールバー（先生のビュー切替6個）で折り返しが起きる。**
v5 §2-9 の通り、**疑似要素で当たり判定だけを広げる**のが正しい。
チェックボックスは疑似要素を持てないので、囲みの `<label>` 側で確保する。

---

## E. PWA（Part I §3）— `docs/` シェル

| # | 項目 | 判定 | 実測 |
|---|---|:--:|---|
| E1 | manifest の id/scope/start_url がリポジトリ名絶対パス | ❌ | 3つとも `"./"` のまま |
| E2 | アイコン4種 + **透明を含まない** apple-touch-icon | ❌ | 4種は揃っている。**`apple-touch-icon.png` は全体の 3.22% が完全透明、四隅ボックスの 62.09% が透明** |
| E3 | beforeinstallprompt を head 最上部で捕捉（外部ファイル） | ❌ | `install-hook.js` が無い。`<head>` 内に `<script>` **0本**（`firstScriptIndexInHead: -1`） |
| E4 | インストールボタン | ❌ | 無い（画面内に「インストール」の語 0件） |
| E5 | sw.js が自アプリ接頭辞のキャッシュのみ削除 | ✅ | **実測。別名キャッシュ2件（`other-app-static-v1` / `another-giga-app-v3`）を置いて版を上げ、両方とも残ることを確認** |
| E6 | sw.js が localStorage に触れていない | ✅ | 0件 |
| E7 | 更新通知（押すまで切り替わらない） | ❌ | **実測。`sw.js` の実バイトを変えて `update()` → 3秒放置したところ `waiting: false` / 旧キャッシュ `…-v2` は消滅済み。`install` の中の `skipWaiting()`（`docs/sw.js:39`）が原因。更新の案内も無い** |
| E8 | 初回訪問で勝手にリロードしない | ✅ | **実測。まっさらな状態で1回開き、`framenavigated` = 1回** |
| E9 | Service Worker が実際に登録されている | ✅ | **実測。`getRegistration()` → `active: true` / scope 一致** |
| E10 | offline.html | ❌ | 無い。本体キャッシュを消して圏外にすると、応答はブラウザの HTTP キャッシュ頼みで不定 |
| E11 | APP_VERSION を今回のリリース値に更新した | — | 現在 `v2`。今回の修正で上げる |
| E12 | maskable のセーフゾーン外の中身 0.2% 以下 | ✅ | **実測 0.00%**（192／512 とも）。下地 `rgb(26,115,232)` が端まで伸びている。透明 0%。**このリポジトリの maskable は正しい** |
| E13 | iOS の「ホーム画面に追加」手順を MANUAL に記載 | ❌ | MANUAL.md が無い。README に PWA 節はあるが iOS の手順は無い |

### E2 — apple-touch-icon の透明（実測）

| ファイル | 完全透明 | 半透明 | 四隅ボックスの透明率 |
|---|---:|---:|---:|
| `apple-touch-icon.png` (180) | **3.22%** | 1.23% | **62.09%** |
| `icon-192.png` | 3.24% | 1.21% | 63.52% |
| `icon-512.png` | 3.77% | 0.50% | 62.66% |
| `icon-maskable-192.png` | 0% | 0% | 0% |
| `icon-maskable-512.png` | 0% | 0% | 0% |

角丸の外側が透明なままである。**iOS は透明部分を黒で埋めるため、
ホーム画面でアイコンの四隅だけが黒く出る。**

### E7 — 更新が押す前に切り替わる（実測手順と結果）

```
1. docs/ を複製してローカル配信し、開く      → active SW あり / caches: [mirai-compass-shell-v2]
2. sw.js の APP_VERSION を v2 → v99 に書き換える（実バイトを変える）
3. 何も押さずに registration.update() し、3秒放置
   結果: waiting = false   ← 押していないのに切り替わっている
         caches = [mirai-compass-shell-v99]   ← 旧版は既に消えている
         更新の案内は画面に出ていない
```

原因は `docs/sw.js:35-41`。`install` の中で `self.skipWaiting()` を呼んでいる。

---

## F. アクセシビリティ・性能

| # | 項目 | 判定 | 実測 |
|---|---|:--:|---|
| F1 | alt / aria-label / aria-live / role="alert" | ❌ | `aria-label` 3件・`aria-live` **0件**・`role="alert"` 1件・`alt` 1件・`aria-hidden` 1件。アイコンのみのボタン（ⓘ・×・丸ボタン群）に名前が無い |
| F2 | モーダルに role="dialog"・Esc で閉じる | ❌ | `role="dialog"` **0件**。`Escape` の処理 **0件**。モーダルは Bootstrap の `data-bs-toggle`（7箇所）任せで、CDN が塞がれると開かない |
| F3 | キーボードのみで全機能に到達 | **未計測** | 今回走査していない |
| F4 | rt の色を決め打ちしていない | ❌ | `css.html:51` で決め打ち。**比 1.04** |
| F5 | 初回JS 300KB以下 | ❌ | CDN の JS だけで **412.2 KB**（gzip前）。CSS を足すと 743.4 KB、アイコン woff2 を足すと 874.0 KB |
| F6 | 1ファイル 5,000行 / 400KB 以内 | ✅ | 最大 `js_teacher.html` 1,740行 / 79.9KB |

---

## G. 学習ログ（学習系）

| # | 項目 | 判定 | 実測 |
|---|---|:--:|---|
| G1 | study.v1 準拠・個人情報を持たない | ❌ | `localStorage['study.records.v1']` **未対応**（0件）。学習記録はスプレッドシートにのみ保存 |
| G2 | 中断記録・5分ルール | ❌ | `status:"aborted"` に相当する記録なし |

> ただし本アプリは**記録をクラスで共有する**設計（C+型）であり、
> `study.v1` は端末内・アプリ間共有の学習ログである。**目的が異なる。**
> 導入するかどうかは機能追加の判断になるので、**このロールアウトの対象外とし、提案に留める。**

---

## 直す順番（v5 Part III の修正フェーズに割り付け）

| フェーズ | 中身 | 破壊リスク |
|---|---|---|
| **P0** | LICENSE / .gitignore / dependabot.yml / CI を作る | 無し |
| **P0.5** | **CDN 実行コード6本の自己ホスト化**（`vendor*.html`）。塞がれた状態で画面が出ることを実測 | 中（アーキテクチャ変更。マージは人間に委ねる） |
| **P1** | 表示・PWA。`100dvh` → `viewport-fit=cover` → safe-area → DPR補正 → fluid type → タップ44px → reduced-motion/forced-colors → **コントラスト** → **`rt` の色** → PWA一式 → CSP | 高（CSP がいちばん壊す） |
| **P2** | 画像の `width`/`height`、`loading="lazy"` | 低 |
| **P3** | MANUAL.md、README の不足節 | 無し |
| **P4** | 品質ゲート（`scripts/check-project.mjs`）。**わざと壊して通ることを確認する** | 無し |

### このロールアウトで**やらない**もの（提案に留める）

| 項目 | 理由（v5 §停止条件） |
|---|---|
| **B5 教員用 API の5段ガード** | 本番で確かめられない。デプロイ設定と児童アカウントの有無で正しい形が変わる |
| **B3 OAuth スコープ** | `appsscript.json` がリポジトリに無く、現状を読めない。外して間違えると全教員で認可が通らなくなり授業が止まる |
| **G1 study.v1 の導入** | 機能追加であり、本アプリの記録設計と目的が異なる |
| **巨大ファイルの分割** | 分割案の合意が先（v5 P3） |

---

## 人間に決めてほしいこと

1. **P0.5（CDN 自己ホスト化）を今回やるか。**
   6本すべてを `vendor*.html` に取り込むと、GAS が返す HTML が **約 874KB** 増える。
   §8 の「初回 JS 300KB 以下」は**どのみち満たせない**（Chart.js だけで 208KB）。
   選択肢は「全部取り込む」「Chart.js を落として取り込む」「版固定＋SRI だけにする」の3つ。
2. **B5（教員 API の本人確認）を別 PR で提案するだけでよいか。**
   いま `changeTeacherPassword` は誰でも呼べる状態である。
3. `appsscript.json` をリポジトリに入れてよいか（入れないとスコープを監査できない）。
