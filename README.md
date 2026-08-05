# **🧭 みらいコンパス**

**「みらいコンパス」は、公立小学校での「自由進度学習」を強力に支援するGoogle Apps Script (GAS) ウェブアプリケーションです。**

子どもたちが自ら学習計画を立て、実行し、振り返るサイクルをデジタル化。先生はリアルタイムで子どもたちの学びを見取り、適切なサポートを行うことができます。GIGAスクール構想の1人1台端末（Chromebook等）での動作に最適化されています。

## **✨ 特徴**

### **👦 児童用機能 (Student Mode)**

* **学習計画の作成**: ドラッグ＆ドロップで直感的に時間割ごとのタスク計画を作成。  
* **今日の学習**: タイムライン形式で進捗を可視化。スタンプ機能で達成感を演出。  
* **ふりかえり & ポートフォリオ**: 毎時間の学びを記録し、単元を通しての成長をレーダーチャートで可視化。  
* **みんなの計画 (ギャラリーウォーク)**: 友達の計画を参照し、学び合いを促進（匿名/記名切り替え可）。  
* **自分用タイマー**: 個別学習に便利なカウントアップ/ダウンタイマー。  
* **SOS / 集中モード**: 先生へステータスをリアルタイム通知。

### **👩‍🏫 先生用機能 (Teacher Mode)**

* **LIVEモニタリング**: クラス全員の学習状況（タスク、進捗、SOS）をリアルタイムで把握。  
* **座席表ビュー**: 実際の教室配置に合わせてカードを配置・監視。  
* **ヒートマップ**: クラス全体の進捗状況を一覧で俯瞰。  
* **単元設計 (AIアシスト)**: 生成AI（ChatGPT/Gemini）で作成した授業案JSONのインポートに対応。  
* **名簿管理 (Ver 2.1 NEW)**: Excel/CSVからのコピペ登録、出席番号管理、転出児童の非表示設定。  
* **外部連携**: 「みらいパスポート」等の外部アプリへ学習データを連携。

## **🛠 ファイル構成**

プロジェクトは以下のファイルで構成されています。

```
/
├── code.gs             # サーバーサイドロジック (DB操作、APIエンドポイント)
├── index.html          # エントリーポイント (ライブラリ読み込み、ファイル結合)
├── css.html            # スタイルシート (デザイン定義)
├── js_core.html        # 共通ロジック (設定、データ管理、通信)
├── js_student.html     # 児童用画面ロジック
├── js_teacher.html     # 先生用画面ロジック
│
├── vendor_css.html     # ⚠️ 生成物：Bootstrap CSS
├── vendor_icons.html   # ⚠️ 生成物：使用中の Bootstrap Icons だけを SVG 化したもの
├── vendor_js.html      # ⚠️ 生成物：Bootstrap JS / SweetAlert2 / Chart.js / Sortable
│
├── package.json        # ライブラリの版を固定している（原本）
├── tools/
│   ├── build-vendor.mjs    # vendor_*.html を作る
│   ├── build-icons.mjs     # 透明を含まない apple-touch-icon を作る
│   └── verify-generated.mjs# 生成物が原本と食い違っていないか確かめる
├── scripts/check-project.mjs / quality.config.json   # 品質ゲート
├── tests/              # 中核ロジックのテスト
│
└── docs/               # PWAシェル (GitHub Pages用)
    ├── index.html            # シェル本体 (GAS本体を全画面表示)
    ├── app.js                # シェルのロジック（CSP のため外部ファイル）
    ├── install-hook.js       # beforeinstallprompt の捕捉（head 最上部）
    ├── offline.html          # 圏外のときに出る画面
    ├── manifest.webmanifest  # PWAマニフェスト
    ├── sw.js                 # Service Worker
    └── icons/                # アプリアイコン
```

### ⚠️ 編集してよいファイル / してはいけないファイル

| ファイル | 編集してよいか |
|---|---|
| `code.gs` / `index.html` / `css.html` / `js_*.html` / `docs/*` | **ここを直す** |
| `package.json` / `tools/*.mjs` | **ここを直す**（ライブラリの版や組み立て方） |
| `vendor_css.html` / `vendor_icons.html` / `vendor_js.html` | **手で編集しない**（生成物） |

**原本を直したら、必ず `npm run build` を走らせてから push してください。**
忘れると、ビルドも静的解析も通るのにリポジトリの中身だけが古いまま残ります。
CI（`npm run ci`）がこの食い違いを検出して落ちます。

```bash
npm ci          # ライブラリを版どおりに入れる
npm run build   # vendor_*.html を作り直す
npm run check   # 品質ゲート（GIGA Standard v5 の検査）
npm test        # 中核ロジックのテスト
npm run ci      # 上の全部（CI と同じもの）
```

## **🚀 インストール & デプロイ手順**

このアプリはGoogle Apps Scriptとして動作します。サーバーの契約は不要です。

1. **プロジェクトの作成**:  
   * [Google Apps Script](https://script.google.com/) にアクセスし、「新しいプロジェクト」を作成します。  
2. **生成物を作る**:
   * 手元で `npm ci && npm run build` を実行します。
     `vendor_css.html` / `vendor_icons.html` / `vendor_js.html` が作られます。
     これらはライブラリ本体で、**GAS に貼るファイルに含まれます**。
3. **ファイルの作成**:
   * エディタ上で以下の **9つ** のファイルを作成し、それぞれのコードを貼り付けます。
     `code.gs` / `index.html` / `css.html` / `js_core.html` / `js_student.html` /
     `js_teacher.html` / `vendor_css.html` / `vendor_icons.html` / `vendor_js.html`
   * ※ `code.gs` 以外のファイルは、拡張子を `.html` として作成してください。
   * ※ `vendor_*.html` は生成物です。GAS のエディタ上で編集しないでください。
     ライブラリを更新するときは、手元で `npm run build` し直して貼り替えます。
4. **デプロイ**:  
   * 右上の「デプロイ」ボタン \> 「新しいデプロイ」を選択。  
   * **種類の選択**: 「ウェブアプリ」  
   * **次のユーザーとして実行**: 「自分」  
   * **アクセスできるユーザー**: 「全員」（または「Googleアカウントを持つ全員」）  
   * 「デプロイ」をクリックし、発行された **ウェブアプリURL** をコピーします。  
5. **初回セットアップ**:  
   * 発行されたURLにアクセスします。  
   * 「初期設定を開始する」ボタンが表示されるのでクリックします（Googleドライブにデータベース用スプレッドシートが自動生成されます）。  
   * 先生用ログイン（初期パスワード: admin）で入り、名簿や単元を登録してください。

## **📱 PWA対応（Chromeから「アプリ」としてインストール）**

本リポジトリの docs/ フォルダには、みらいコンパスを **PWA（プログレッシブウェブアプリ）** としてインストールできるようにするための「シェルアプリ」が含まれています。Chromebook等で通常のアプリのように全画面・専用アイコンで起動できます。

> **なぜシェルが必要？**
> GASのウェブアプリはGoogleのサンドボックス（iframe）内で配信されるため、GAS単体ではPWAのインストール条件（manifest / Service Worker）を満たせません。そこでGitHub Pages上に配置した軽量なシェルページがPWAとして振る舞い、その中でGAS本体を全画面表示します。アプリ本体は常にGASから最新版が読み込まれます。

### **セットアップ手順（管理者・先生）**

1. **GitHub Pagesを有効化**:
   * このリポジトリの Settings > Pages を開く
   * Source: 「Deploy from a branch」/ Branch: main / フォルダ: /docs を選択して Save
   * 数分後に `https://<ユーザー名>.github.io/MIRAI-Compass/` が公開されます
2. **初期設定**:
   * 公開されたURLをChromeで開き、GASのウェブアプリURL（`.../exec`）を入力して「アプリをはじめる」
3. **インストール**:
   * アドレスバー右端の「インストール」アイコン、またはChromeメニュー >「保存と共有」>「ページをアプリとしてインストール」をクリック
   * デスクトップ/シェルフにアイコンが追加され、独立ウィンドウで起動できるようになります

### **児童端末への一括配布**

URLパラメータでGASのURLを事前設定できます。以下のリンクをQRコードやクラスルームで配布すると、児童は開くだけで設定完了です。

```
https://<ユーザー名>.github.io/MIRAI-Compass/?app=<GASウェブアプリURL>
```

### **設定の変更**

* 画面**左上すみを素早く5回タップ**すると設定画面が開きます（URL変更・リセット・新しいタブで開く）
* または `?settings=1` を付けてアクセス

### **注意事項**

* GAS側のデプロイ設定で「アクセスできるユーザー: **全員**」を推奨します。「Googleアカウントを持つ全員」の場合、iframe内ではGoogleログイン画面が開けないため、表示できない時は設定画面の「新しいタブで開く」から一度ログインしてください。

## **💻 技術スタック**

* **Backend**: Google Apps Script (GAS)  
* **Database**: Google Spreadsheet  
* **Frontend**: HTML5, CSS3, JavaScript (ES6)  
* **Libraries（すべて自己ホスト。CDN から取る実行コードは 0 バイト）**:
  * [Bootstrap 5](https://getbootstrap.com/) 5.3.0 (UI Framework)
  * [Bootstrap Icons](https://icons.getbootstrap.com/) 1.11.1 — 使用中の53種類のみ SVG マスク化（229KB → 32KB）
  * [SweetAlert2](https://sweetalert2.github.io/) 11.26.25 (Modals)
  * [Chart.js](https://www.chartjs.org/) 4.5.1 (Data Visualization)
  * [Sortable.js](https://sortablejs.github.io/Sortable/) 1.15.0 (Drag & Drop)
  * [Google Fonts](https://fonts.google.com/) (Zen Maru Gothic) — **これだけは外部から読む**

> **なぜ実行コードを自己ホストするのか。**
> 学校のネットワークは `cdn.jsdelivr.net` を塞いでいることがあります。
> 塞がれた状態でこのアプリを開くと、白い画面ではなく
> **「Bootstrap が当たっていない素の HTML が半分だけ動く」**という壊れ方をしました。
> ローディング画面が消えず、`d-none` が効かないので児童画面と先生用ボタンが
> 同時に出て、`Swal` / `Chart` / `Sortable` がすべて `undefined` になります。
> 児童からは「壊れている」としか見えず、原因がアプリの外にあるので
> 先生が調べても分かりません。
>
> Google Fonts だけは外部のままにしています。届かなくても**字の形が変わるだけ**で
> アプリは動くからです（`css.html` で端末側の日本語フォントを後ろに並べてあります）。
> 日本語フォントを自己ホストすると初回転送が数MBになり、
> 校内Wi-Fiで40人が同時に開くという、いちばん避けたい状況を自分で作ります。

## **🔒 セキュリティ設計と、いま分かっている課題**

| 項目 | 状態 |
|---|---|
| 秘密情報 | `PropertiesService`（スクリプトプロパティ）に保存。コードに直書きしていない |
| 排他制御 | 書き込み系 12関数すべてに `LockService` + `try...finally` |
| CSP | `docs/` シェルには適用済み（`script-src 'self'`、インライン無し）。GAS 本体は後述 |
| 外部からの実行コード | 0 バイト |

### ⚠️ 教員用 API に本人確認がありません（未修正）

`changeTeacherPassword` / `initSystem` / `saveClassRoster` / `deleteUnitTask` /
`savePortfolioFeedback` などの教員向け関数は、`google.script.run` から直接呼べます。
画面側の `attemptTeacherLogin()` は `verifyPassword()` の戻り値で UI を出し分けて
いるだけで、**フロントの出し分けは防御になりません。**

直し方はデプロイの実行者設定（アクセスユーザー／アプリアカウント）と、
児童が Google アカウントを持つかどうかで変わります。本番で確かめられないまま
変えると全教員で認可が通らなくなり授業が止まるため、**このロールアウトでは
変更していません。** 詳細と対処案は `AUDIT.md` の B5 を参照してください。

### GAS 本体に CSP を入れていない理由

GAS は `.gs` と `.html` しか置けず、JavaScript は `include()` でインラインの
`<script>` として埋め込まれます。つまり **`script-src 'self'` は構造上成立しません。**
`'unsafe-inline'` を足せば入れられますが、それでは CSP を入れた意味がほとんど
無くなります。代わりに、**外部から読む実行コードを 0 にする**ことで
攻撃面を減らしました。`docs/` シェル側には本物の CSP が入っています。

## **🤝 外部連携について**

本アプリは、外部の学習記録アプリ（通称：みらいパスポート）との連携機能を備えています。

* **Direct DB Write**: 先生が作成した単元計画を、相手先のスプレッドシートに直接書き込んで連携します。  
* **Launcher**: タスクカードからパラメータ付きURLで外部アプリを起動します。  
* **Status Sync**: 外部アプリからのSOS通知等をWebhook (doPost) で受信します。

## **📝 ライセンス**

This project is licensed under the MIT [License](https://www.google.com/search?q=LICENSE) - see the LICENSE file for details.

© 2026 みらいコンパス Project
