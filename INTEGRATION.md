# みらいコンパス × みらいパスポート 統合設計

このドキュメントは、2つのアプリ（みらいコンパス＝自由進度学習の計画・見取り／
みらいパスポート＝AIワークシートの作成・提出・添削）を **最終的に単一プロジェクトへ
統合する** ための設計と、そこへ至る移行手順をまとめたものです。
両リポジトリに**バイト単位で同一のコピー**を置いています（`INTEGRATION.md`）。片方だけ直さないでください。

---

## 1. なぜ統合するのか

2つは同じ「1クラス・40人・GIGA端末」を対象にした姉妹アプリで、対象児童・単元・
教員が完全に重なります。いま両者は別 GAS プロジェクト・別スプレッドシートで動き、
次の3経路だけで細く繋がっています。

| 経路 | 実装 | 方向 | 現状の弱点 |
|---|---|---|---|
| ① 状態同期（SOS/集中） | Passport `syncToCompass` → Compass `doPost` | Passport → Compass | 送りっぱなし。無認証だった（本改修でトークン化） |
| ② 単元計画の受け渡し | Compass が Passport の DB（`ImportQueue`シート）へ**直接書き込み** | Compass → Passport | 相手のDBスキーマに密結合。最も壊れやすい |
| ③ 児童の起動・同定 | Compass が `?studentId=&studentName=` を付けて Passport を起動 | Compass → Passport | `studentId` に児童名を素で渡している |

統合すると、この3経路は**すべて内部関数呼び出しに置き換わり**、DB は1つ、認証は1回、
児童IDは1体系になります。まずはその土台（共通契約）を両アプリに敷きました。

---

## 2. すでに敷いた共通契約（本改修で導入）

統合の前提となる「両アプリで同一のもの」を先に入れました。**バイト単位で同一**に保ってください。

### 2.1 `MiraiShared`（クライアント / `js_core.html`）
- 安全描画ヘルパー: `escHtml` / `escUrl`（`javascript:` 等を遮断、`data:image/` は許可）/ `escJs`
- 共通enum（統合後の唯一の語彙）と正規化関数:
  - **在席モード** `normal | focus | sos` … `normPresence(v)`
  - **学習進捗** `todo | doing | done` … `normProgress(v)`
  - **提出状態** `draft | submitted | graded` … `normSubmission(v)`

### 2.2 `MiraiAuth`（サーバー / `code.gs`）
- `CacheService` に置く短命トークン（6時間）による**先生の本人確認**。
- `issueToken()` / `isValid(token)` / `requireTeacher(token)` / `revoke(token)`。
- **デプロイ設定（実行者・アクセス範囲）に依存しない**ため、本番で認可が原因で授業が
  止まりません。ここが従来「本番で確かめられないから触れない」とされていた B5 の回避策です。

> この2ブロックは両リポジトリで `diff` が空になる状態を維持します。統合時はそのまま
> 共有ファイル（GAS ライブラリ or 貼り付け用 `shared.gs` / `shared.html`）へ昇格させます。

---

## 3. 認可モデル（現状）

| 呼び出し | 認可 |
|---|---|
| 児童向け関数（読み書きとも） | 認証なし（`mode=student` は認証ではない、という設計は据え置き） |
| 先生専用関数 | **サーバー先頭で `MiraiAuth.requireTeacher(token)`**。クライアントは Compass=`runTeacher(...)` / Passport=`Server.callTeacher(...)` でトークンを自動前置 |
| 先生トークンの発行 | Compass=教員パスワード照合 `verifyPassword` 成功時／Passport=教員あいことば `verifyTeacherPass` 成功時 |
| 初期化 | 初回（DB未作成）のみ無認証、既存システムの再実行はトークン必須 |
| 外部 Webhook（Compass `doPost`） | `INTEGRATION_TOKEN` の一致時のみ書き込み（フェイルクローズ） |

**残る限界（統合で解消予定）**: 児童どうしの IDOR（他児童の提出物・ポートフォリオを
`studentId` の差し替えで読む）は、児童に認証が無いため現状のトークンでは塞げません。
統合時に「Compass 発行の署名付き児童トークン」を Passport が検証する形にすると解決します
（→ 5章 フェーズ2）。本改修では最低限、**児童向けレスポンスから他児童のメールアドレス・
シートIDを除去**しました。

---

## 4. 連携トークンの設定手順（②③が動くために必要／移行中の一時運用）

Compass の `doPost` を認可した結果、Passport → Compass の状態同期は
**連携キーを設定するまで届きません（安全側の停止）**。次の1回だけ設定してください。

1. **Compass**（先生モード）→ 設定 → **「パスポート連携キーを表示」** で値をコピー
   （このキーはシステム初期化時に自動発行される `INTEGRATION_TOKEN` です）。
2. **Passport**（先生設定 ⚙）→ **「Compass 連携キー」** に貼り付けて保存。
3. あわせて Passport の「Compass URL」に Compass の `/exec` URL が入っていることを確認。

これで SOS・集中・状態通知が再び Compass の LIVE 画面へ届きます。
**統合後はこの手順ごと不要**になります（同一プロジェクト内の関数呼び出しになるため）。

---

## 5. 単一プロジェクトへの統合ロードマップ

事故を減らす順に並べています。各フェーズは独立して価値が出ます。

### フェーズ0（完了）— 共通契約を敷く
- `MiraiShared` / `MiraiAuth` を両アプリへ導入（本改修）。
- 保存型XSS・データ破損・無認可の主要な穴を先に塞ぐ（統合すると露出面が倍になるため）。

### フェーズ1 — 児童IDと課題IDの統一（統合の核心）
- **児童ID**: 名簿（Compass `StudentRoster`）を正本にし、`studentId` を発番する。
  以後、Compass `LiveStatus`・Passport `Responses` の主キーをこの ID に寄せる。
  いまは Compass=児童名、Passport=`manabi_sid`/メールと**3体系**が混在している。
- **課題ID**: 本改修で Passport の `t1,t2…` 使い回しを**単元スコープの一意ID**へ移行済み。
  統合時は Compass の `T…` と名前空間を1つに統べる。Compass で課題を削除したら
  Passport の対応ワークシートも消える「削除の伝播」を定義する。

### フェーズ2 — 連携をAPIに一本化し、DB直書きをやめる
- ② `ImportQueue` への**直接書き込みを廃止**し、トークン付きの受け口（doPost or 内部関数）に統一。
- ③ 児童同定を「Compass 発行の署名付きトークン」に変更 → Passport がサーバー側で検証。
  これで児童IDORも閉じる。
- 状態語彙・提出語彙は 2.1 の共通enumに統一（受け側で `||` を並べている箇所を一掃）。

### フェーズ3 — DBとシェルの統合
- スプレッドシートを1つに統合（または明確に分割し、相互参照を内部APIに限定）。
- `docs/` PWAシェルを1枚に集約（`manifest` の `id/scope/start_url`、`sw.js` の
  `CACHE_PREFIX` を統合アプリ用に確定。`gigayama.github.io` は同一オリジン共有のため必須）。
- `appsscript.json` を両アプリともリポジトリに入れ、OAuth スコープを統合設計する
  （現状どちらもリポジトリに無く、スコープを監査できない）。

### フェーズ4 — 品質ゲートの正本化
- `scripts/lib/project-quality.mjs`（フリート共通の正本）を確定し、両リポジトリで
  バイト単位共有。**「テンプレートリテラルへの生挿入禁止」「先生専用関数の未ガード検出」**
  といった、今回のクラスのバグを機械的に捕まえる検査を足す（現ゲートはXSS・認可を見ていない）。

---

## 6. 統合時に決めるべき論点（人間の判断が要る）

| 論点 | 選択肢 |
|---|---|
| ポートフォリオの正本 | Compass `Portfolios`＋`DailyReflections` か、Passport `Responses.reflectionText` か |
| AI連携 | Compass の「手貼りJSON」方式 か、Passport の「Gemini API サーバー呼び出し」方式へ寄せるか |
| 教員認証 | あいことば＋トークン（現状）か、`Session.getActiveUser()` ベースへ格上げか |
| デプロイ単位 | 1プロジェクトに完全統合（OAuth・6分制限を共有）か、疎結合APIを維持か |

---

© 2026 みらいコンパス / みらいパスポート Project
