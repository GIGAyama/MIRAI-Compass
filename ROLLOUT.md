# ロールアウト記録 — MIRAI-Compass

GIGA Standard v5 の適用記録。
**他のリポジトリにも効く知見を中心に書く。** 1本の問題で終わらせないため。

- 実施日: 2026-08-05
- 型: **C+型**（GAS ウェブアプリ本体 + `docs/` の GitHub Pages シェル）
- ブランチ: `claude/rollout-iibj1c`
- 詳しい実測値: `AUDIT.md`

---

## 何をしたか

| フェーズ | 中身 | 状態 |
|---|---|:--:|
| P0 | LICENSE / .gitignore / dependabot.yml | ✅ |
| P0.5 | **CDN 実行コード6本の自己ホスト化** | ✅ |
| P1 | 表示（コントラスト36→0・タップ62→0）・PWA・提示モード・印刷 | ✅ |
| P2 | 画像（すでに基準内。測って据え置き） | ✅ |
| P3 | MANUAL.md / README | ✅ |
| P4 | 品質ゲート36件 + CI + テスト。**わざと壊して確認済み** | ✅ |
| — | 教員用 API の本人確認（B5） | **未修正・別PRで提案** |

---

## 🔁 他のリポジトリにも効く知見

### 1. GAS でも表示は実測できる（C型・C+型すべてに効く）

「GAS だから測れない」は正しくない。`script.google.com` へ到達できなくても、
`include()` を手元で貼り合わせれば表示は測れる。**この1本で 36件のコントラスト
違反と 62件のタップ違反が見つかった。**

貼り合わせのときに必ず要るもの:

- `code.gs` の `doGet` にある `addMetaTag('viewport', …)` を反映する。
  **これはサーバー側の処理なので、`index.html` を組み立てただけでは再現されない。**
  「指定が無い」と「悪い値が入っている」を取り違える。
- `google.script.run` のダミーに**戻り値の見本**を与える。
  見本が無いと初期設定画面から先に進めず、本編を1画面も測れない。
  このリポジトリでは `getAppInitialData` / `getData` / `getStudentProgress` の
  3つに見本を与えるだけで、児童6画面・先生9画面まで歩けた。

横断で使えるよう、貼り合わせ・ダミー・測定器は同じ形で書いてある。

### 2. 「白い画面」ではなく「崩れた画面が半分動く」ことがある

v5 §6 は CDN が塞がれると「画面が白いまま何も出ない」と書いている。
**実際に測ると、このリポジトリではもっと分かりにくい壊れ方をした。**

```
Bootstrap CSS が当たらない → d-none が効かない
  → 児童画面と先生用ボタンが同時に出る
  → ローディング画面が消えないまま残る
  → Swal / Chart / Sortable は undefined なので押すと何も起きない
```

**「白い画面が出るか」で判定すると見逃す。** 見るべきは次の3つ。

```javascript
// CDN を落とした状態で
!!window.bootstrap && !!window.Swal && !!window.Chart   // ライブラリが居るか
document.styleSheets.length                            // スタイルシートの枚数
getComputedStyle(document.getElementById('loading')).display  // ローディングが消えたか
```

### 3. Bootstrap Icons はまるごと持たなくていい（**229KB → 32KB**）

`bootstrap-icons.css` は 98KB、woff2 は 131KB ある。
**実際に使っているのは数十種類しかない。**
使用中のアイコンだけ SVG を取り出して `mask-image` にすれば、
`<i class="bi bi-x">` というマークアップを1文字も変えずに置き換えられる。
`background-color: currentColor` なので `text-danger` などの色指定もそのまま効く。

```css
.bi { display:inline-block; width:1em; height:1em; vertical-align:-.125em;
      background-color:currentColor;
      -webkit-mask:var(--bi-i) center/contain no-repeat; mask:var(--bi-i) center/contain no-repeat }
.bi-compass-fill { --bi-i: url("data:image/svg+xml,…") }
```

**このとき、存在しないアイコン名が見つかることがある。**
このリポジトリでは `bi-pencil-ruler` が bootstrap-icons 1.11.1 に存在せず、
先生の「設計」ボタンのアイコンがずっと空白だった。
フォント方式だと「何も描かれない」だけなので誰も気づかない。
**SVG 化すると、ビルド時に `missing` として名前が出る。**

```bash
# 全リポジトリで使用中アイコンの存在確認をするなら
grep -oh "bi bi-[a-z0-9-]*" $(git ls-files '*.html') | sed 's/bi bi-//' | sort -u
```

### 4. ふりがな（`rt`）は「決め打ちしているか」ではなく「継がせているか」で見る

v5 §4 の指摘どおりだったが、**このリポジトリの比は 1.04** で、
標準が挙げている実例（1.28 / 1.47）より悪かった。
`rt { color: var(--text-muted) }` のように**変数経由**だったため、
`grep "rt {" | grep "#"` のような探し方では見つからない。

```bash
# 変数経由も拾う
grep -n -A3 "^\s*rt\s*{" $(git ls-files '*.html' '*.css')
```

直すときは**1か所ずつ潰さない。** 色のついた面でまとめて継がせる。

### 5. 濃くする一括置換は、淡い地の上の文字を壊す

v5 §2-8 は「濃い面の上の薄い文字」が壊れると書いている。
**このリポジトリで壊れたのは逆方向だった。**

`--bs-*-text-emphasis` を `--color-*-d` で上書きしたところ、
`.alert-info`（淡い水色の地）の文字が **4.34 に悪化**した。
`--bs-*-text-emphasis` は Bootstrap が「淡い地の上に載せる文字」用に
用意している変数で、**既定値がすでに十分濃い**（`#055160` など）。

**上書きしてよいのは `.text-*` ユーティリティとボタン変数だけ。**
`--bs-*-text-emphasis` は触らない。

### 6. Playwright の `setOffline` は Service Worker 内の `fetch` を止めない

`context.setOffline(true)` してキャッシュを消しても、SW の `fetch(req)` が
**成功してしまう**ため、`offline.html` が出るかを確かめられない。
キャッシュを消したのに素通りするので「offline.html が要らない」と誤読しかねない。

**圏外はサーバーを実際に落として作る。**

```javascript
await new Promise(r => { srv.close(r); for (const c of sockets) c.destroy(); });
```

### 7. コントラスト測定器の絵文字除外を、符号位置の範囲で書かない

絵文字を除外するつもりで `/[‼-㊙…]/` のような範囲を書くと、
**U+3040〜U+30FF（ひらがな・カタカナ）が範囲に入る。**
児童向け画面の文字はほとんどひらがななので、**まるごと測れなくなる。**

このリポジトリでは、この誤りのせいで最初 **36件中 19件しか見えていなかった**
（`rt` の 1.04 も見えていなかった）。

```javascript
// ✅ 「絵文字を含む」ではなく「絵文字と記号だけでできている」ときだけ外す
const PICTO = /\p{Extended_Pictographic}/u;
const isOnlyEmoji = (s) => PICTO.test(s) &&
  !/[\p{Letter}\p{Number}]/u.test(s.replace(/\p{Extended_Pictographic}/gu, ''));
```

**測定器を疑う手がかり**：児童向けアプリなのに違反が「先生向けの英数字だけ」に
偏っていたら、ひらがなを取りこぼしている。

### 8. 検査コードの `stripComments` が URL の `/*` を誤読する

品質ゲートを書くとき、判定前にコメントを落とすのは v5 §P4 のとおり正しい。
**ただし URL を先に伏せないと壊れる。**

`https://*.googleusercontent.com` の `/*` をブロックコメントの開始と読み、
**次の `*/` までを丸ごと消す。** 消えた範囲に検査対象が入っていると、
その検査は「壊したのに通る」状態になる。

これは `--self-test`（わざと壊す）でしか見つからない。
**共通の検査に手を入れたら、正本にも反映すること。**

### 9. 品質ゲートの自己診断には「壊していない検査」の検出を入れる

`--self-test` で36件中35件を壊して確かめても、**書き忘れた1件は素通りする。**
壊し方を書いていない検査が残っていたら自己診断そのものを落とすようにした。
実際、これで `F6_FILE_SIZE` の確認漏れが見つかった。

また、**壊し方が効いていないことがある。**

- `clamp(` → `notclamp(` は、検査の正規表現 `/clamp\s*\(/` に**まだ一致する**
- 1ファイルだけ壊しても、他ファイルの件数で閾値を割らないことがある

`patch()` は置換が no-op なら例外を投げるようにしてあるが、
**「置換はされたが検査は落ちない」**は自己診断の結果を読まないと分からない。

---

## 🔎 フリート横断で数えるべきもの（このリポジトリで見つかった形）

```bash
# ① ふりがなの色（変数経由も拾う）
grep -ln -A3 "^\s*rt\s*{" $(git ls-files '*.html' '*.css')

# ② install の中の skipWaiting（押す前に切り替わる）
grep -l "skipWaiting" $(git ls-files '*sw.js') | \
  xargs -I{} sh -c 'awk "/addEventListener..install/,/addEventListener..activate/" {} | grep -q skipWaiting && echo {}'

# ③ apple-touch-icon の透明（iOS で四隅が黒くなる）
#    node scripts/lib/png-alpha.mjs 相当で画素を数える。目視では分からない。

# ④ manifest の id/scope/start_url が "./" のまま
grep -l '"id": "\./"' $(git ls-files '*manifest.webmanifest')

# ⑤ code.gs 側の viewport（index.html だけ直しても効かない）
grep -n "addMetaTag..viewport" $(git ls-files '*.gs')

# ⑥ 版が浮いている CDN 参照（@11 や版指定なし）
grep -n "cdn.jsdelivr.net/npm/[a-z0-9-]*\(@[0-9]*\)\?[\"']" $(git ls-files '*.html')

# ⑦ bootstrap-icons をまるごと読んでいる（229KB）
grep -l "bootstrap-icons.*\.css" $(git ls-files '*.html')
```

**②③④⑦は、このリポジトリで実際に見つかった。**
とくに **⑦は使用アイコン数さえ数えれば効果が読める**ので、
横断で先に回す価値が高い。

---

## 人間に決めてほしいこと（未決）

1. **B5：教員用 API の本人確認をどう入れるか。**
   `changeTeacherPassword` / `initSystem` などがいまも誰でも呼べる。
   サーバー側トークン方式か `Session.getActiveUser()` 方式か、
   デプロイ設定と児童アカウントの有無で決まる。
2. **`appsscript.json` をリポジトリに入れるか。**
   入れないと OAuth スコープを監査できない（いまは「未計測」のまま）。
3. **Chart.js（208KB）を残すか。**
   残す限り §8 の「初回 JS 300KB 以下」は満たせない。
   レーダーチャート1箇所のみなので、自前 SVG にすれば 204KB になり基準内に入る。
4. **C3 / C4（`pagehide` での確定保存、通信失敗時のリトライ）をいつ入れるか。**
   どちらも書き込み回数・`LockService` の保持区間に影響するため別 PR にした。
5. **GAS 本体の CSP。**
   `include()` で JS がインラインになるため `script-src 'self'` は構造上成立しない。
   nonce 方式は本番で確かめられないと入れられない。

---

## 作業環境の制約（引き継ぎ）

- `script.google.com` へ到達できない → **GAS 本番の動作確認・デプロイはできない**
- `gigayama.github.io` へ到達できない → **PWA の本番確認はできない**（`docs/` をローカル配信して測った）
- `cdn.jsdelivr.net` / `fonts.googleapis.com` へ出られない
  → **学校のフィルタリングと同じ状態で測れる（利点として使った）**
- npm レジストリへは出られる → 実バイトを取得してローカル控え・自己ホスト化が可能
