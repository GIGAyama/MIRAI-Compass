/**
 * GIGA Standard v5 Part I の検査。
 *
 * ⚠️ 検査を書くときの落とし穴（実際に踏んだもの）
 *   1. 「消す式」を正規表現で追うと (k) => caches.delete(k) を見落とす。
 *      見るべきは「消す式があるか」ではなく「startsWith で絞る式があるか」。
 *   2. 「localStorage は操作しない」という【注意書き】に反応して誤検知する。
 *      判定の前にコメントを落とす。
 *   3. @supports not (height: 100dvh) { … 100vh } を 100vh の単独使用と誤判定する。
 *      前方も見る。
 *
 * これらは「わざと壊して通ることを確認する」ことでしか見つからない。
 * scripts/check-project.mjs --self-test がその確認を行う。
 */
import { readFileSync, existsSync, statSync } from 'node:fs';
import { join } from 'node:path';
import { scanPngAlpha } from './png-alpha.mjs';

/**
 * コメントを落とす。判定を注意書きに反応させないため。
 *
 * ⚠️ URL を先に伏せてからコメントを落とす。
 *   `https://*.googleusercontent.com` の `/*` をブロックコメントの開始と読んで
 *   しまい、そこから次の `*&#47;` までを丸ごと消していた。
 *   実際、CSP の frame-ancestors 検査が「壊したのに通る」状態になっていた
 *   （消された範囲に検査対象が入っていたため）。自己診断で見つかった。
 */
export function stripComments(src, kind) {
  const urls = [];
  let s = src.replace(/\b[a-z][a-z0-9+.-]*:\/\/[^\s"'<>)]+/gi, (m) => {
    urls.push(m);
    return `@@URL${urls.length - 1}@@`;
  });

  if (kind === 'html') s = s.replace(/<!--[\s\S]*?-->/g, '');
  s = s.replace(/\/\*[\s\S]*?\*\//g, '');
  s = s.replace(/(^|[^"'`\\])\/\/[^\n]*/g, '$1');

  return s.replace(/@@URL(\d+)@@/g, (_, i) => urls[i]);
}

const read = (root, f) => readFileSync(join(root, f), 'utf8');
const has = (root, f) => existsSync(join(root, f));

/**
 * トップレベル関数 `function name(...) { ... }` の本文（おおよそ）を返す。
 * ブレース対応ではなく「次のトップレベル function まで」で切る。
 * 文字列や JSON テンプレート内の { } に振り回されないための割り切り。
 * requireTeacher は各関数の先頭で呼ぶ約束なので、多少後ろを削っても判定は保てる。
 */
export function fnSlice(src, name) {
  const m = src.match(new RegExp('function\\s+' + name + '\\s*\\('));
  if (!m) return null;
  const rest = src.slice(m.index + m[0].length);
  const next = rest.search(/\nfunction\s/);
  return next < 0 ? rest : rest.slice(0, next);
}

/**
 * 各検査は { id, title, run(ctx) } を持ち、run は
 * { ok: boolean, detail: string } を返す。
 */
export function buildChecks(cfg) {
  const APP_HTML = cfg.appHtmlFiles;
  const SHELL = cfg.shellDir;
  const REPO = cfg.repoName;

  return [
    // ---------------- A. 法務・配布 ----------------
    {
      id: 'A1_LICENSE', title: 'LICENSE 実ファイルがある',
      run: ({ root }) => ({ ok: has(root, 'LICENSE'), detail: has(root, 'LICENSE') ? 'あり' : 'LICENSE が無い' }),
    },
    {
      id: 'A2_GITIGNORE', title: '.gitignore に node_modules と秘密ファイル',
      run: ({ root }) => {
        if (!has(root, '.gitignore')) return { ok: false, detail: '.gitignore が無い' };
        const s = read(root, '.gitignore');
        const need = ['node_modules', '.clasp.json', '.env'];
        const miss = need.filter((n) => !s.includes(n));
        return { ok: miss.length === 0, detail: miss.length ? '不足: ' + miss.join(', ') : 'あり' };
      },
    },
    {
      id: 'A3_DEPENDABOT', title: 'dependabot.yml がある',
      run: ({ root }) => ({ ok: has(root, '.github/dependabot.yml'), detail: has(root, '.github/dependabot.yml') ? 'あり' : '無い' }),
    },
    {
      id: 'A4_DOCS', title: 'README / MANUAL / AUDIT がある',
      run: ({ root }) => {
        const miss = ['README.md', 'MANUAL.md', 'AUDIT.md'].filter((f) => !has(root, f));
        return { ok: miss.length === 0, detail: miss.length ? '不足: ' + miss.join(', ') : 'あり' };
      },
    },
    {
      id: 'A5_CI_ON_PR', title: 'CI が pull_request でも動く',
      run: ({ root }) => {
        const f = '.github/workflows/ci.yml';
        if (!has(root, f)) return { ok: false, detail: 'ワークフローが無い' };
        const s = read(root, f);
        return { ok: /pull_request/.test(s), detail: /pull_request/.test(s) ? 'あり' : 'push だけでは PR の時点で落ちていることに気づけない' };
      },
    },

    // ---------------- B. セキュリティ・依存 ----------------
    {
      id: 'B6_NO_CDN_EXEC', title: 'CDN から取る実行コードが 0',
      run: ({ root }) => {
        const bad = [];
        for (const f of [...APP_HTML, ...cfg.shellHtmlFiles]) {
          if (!has(root, f)) continue;
          const s = stripComments(read(root, f), 'html');
          const re = /<(?:script|link)\b[^>]*?(?:src|href)\s*=\s*["'](https?:\/\/[^"']+)["'][^>]*>/gi;
          for (const m of s.matchAll(re)) {
            const url = m[1];
            // Web フォントは「見た目だけ」の依存なので許す（届かなくても動く）
            if (/fonts\.googleapis\.com|fonts\.gstatic\.com/.test(url)) continue;
            bad.push(`${f}: ${url}`);
          }
        }
        return { ok: bad.length === 0, detail: bad.length ? bad.join('\n      ') : '0 バイト' };
      },
    },
    {
      id: 'B6b_NO_BROWSER_BABEL', title: 'ブラウザ内 Babel / Tailwind CDN を使っていない',
      run: ({ root }) => {
        const bad = [];
        for (const f of [...APP_HTML, ...cfg.shellHtmlFiles]) {
          if (!has(root, f)) continue;
          const s = stripComments(read(root, f), 'html');
          if (/babel\/standalone/.test(s)) bad.push(`${f}: @babel/standalone`);
          if (/cdn\.tailwindcss\.com/.test(s)) bad.push(`${f}: cdn.tailwindcss.com`);
          if (/type\s*=\s*["']text\/babel["']/.test(s)) bad.push(`${f}: type="text/babel"`);
        }
        return { ok: bad.length === 0, detail: bad.length ? bad.join('\n      ') : '無し' };
      },
    },
    {
      id: 'B2_NO_SECRETS', title: '秘密情報の直書きが無い',
      run: ({ root }) => {
        const bad = [];
        for (const f of cfg.sourceFiles) {
          if (!has(root, f)) continue;
          const s = stripComments(read(root, f), f.endsWith('.html') ? 'html' : 'js');
          if (/\bAIza[0-9A-Za-z_-]{35}\b/.test(s)) bad.push(`${f}: Google API キーらしき文字列`);
          if (/\bAKIA[0-9A-Z]{16}\b/.test(s)) bad.push(`${f}: AWS アクセスキーらしき文字列`);
          if (/(password|passwd|secret|token)\s*[:=]\s*["'][^"']{8,}["']/i.test(s)) {
            bad.push(`${f}: パスワードらしき直書き`);
          }
        }
        return { ok: bad.length === 0, detail: bad.length ? bad.join('\n      ') : '無し（値は転記しない）' };
      },
    },
    {
      id: 'B4_POSTMESSAGE', title: 'postMessage の宛先が * でない',
      run: ({ root }) => {
        const bad = [];
        for (const f of cfg.sourceFiles) {
          if (!has(root, f)) continue;
          const s = stripComments(read(root, f), f.endsWith('.html') ? 'html' : 'js');
          if (/postMessage\s*\([^)]*,\s*['"]\*['"]\s*\)/.test(s)) bad.push(f);
        }
        return { ok: bad.length === 0, detail: bad.length ? bad.join(', ') : '該当なし' };
      },
    },
    {
      id: 'B1_CSP_SHELL', title: 'シェルに CSP があり unsafe-inline を script に足していない',
      run: ({ root }) => {
        const f = `${SHELL}/index.html`;
        if (!has(root, f)) return { ok: false, detail: `${f} が無い` };
        const s = read(root, f);
        const m = s.match(/http-equiv=["']Content-Security-Policy["'][^>]*content=["']([\s\S]*?)["']\s*>/i);
        if (!m) return { ok: false, detail: 'CSP が無い' };
        const csp = m[1];
        const scriptSrc = (csp.match(/script-src([^;]*)/) || [])[1] || '';
        if (/unsafe-inline|unsafe-eval/.test(scriptSrc)) {
          return { ok: false, detail: `script-src に unsafe-inline / unsafe-eval がある: ${scriptSrc.trim()}` };
        }
        return { ok: true, detail: 'あり（script-src はインラインを許していない）' };
      },
    },
    {
      id: 'B1b_NO_FRAME_ANCESTORS_META', title: 'frame-ancestors を <meta> に書いていない',
      run: ({ root }) => {
        const f = `${SHELL}/index.html`;
        if (!has(root, f)) return { ok: false, detail: `${f} が無い` };
        const s = stripComments(read(root, f), 'html');
        const m = s.match(/http-equiv=["']Content-Security-Policy["'][^>]*content=["']([\s\S]*?)["']\s*>/i);
        const bad = m && /frame-ancestors/.test(m[1]);
        return { ok: !bad, detail: bad ? '<meta> の frame-ancestors は無視され、警告が出るだけ' : '書いていない' };
      },
    },
    {
      id: 'B1c_NO_INLINE_HANDLER_IN_SHELL', title: 'シェルにインライン script / onclick が無い',
      run: ({ root }) => {
        const f = `${SHELL}/index.html`;
        if (!has(root, f)) return { ok: false, detail: `${f} が無い` };
        const s = stripComments(read(root, f), 'html');
        const bad = [];
        // src を持たない <script> は中身がインライン
        for (const m of s.matchAll(/<script\b([^>]*)>/gi)) {
          if (!/\bsrc\s*=/.test(m[1])) bad.push('インラインの <script>');
        }
        for (const m of s.matchAll(/\bon[a-z]+\s*=\s*["']/gi)) bad.push(`インラインハンドラ ${m[0].trim()}`);
        return { ok: bad.length === 0, detail: bad.length ? [...new Set(bad)].join(', ') + '（CSP を入れると黙って動かなくなる）' : '無し' };
      },
    },

    // ---------------- SEC. サーバー認可（getActiveUser 移行の砦） ----------------
    {
      id: 'SEC_PRIVILEGED_FN_GUARDED', title: '先生専用関数がサーバー側で requireTeacher している',
      run: ({ root }) => {
        const f = 'code.gs';
        if (!has(root, f)) return { ok: false, detail: 'code.gs が無い' };
        const s = stripComments(read(root, f), 'js');
        // 画面側の出し分けは防御にならない。google.script.run から直接呼べる
        // 教員向け関数は、本文の先頭でサーバーが本人確認していなければならない。
        const fns = [
          'saveWorksheetToDB', 'saveFeedback', 'batchSaveFeedback', 'generateSingleWorksheet',
          'generateRubricAI', 'generateBatchComments', 'getDashboardData', 'getTaskSubmissions',
          'getSubmissionDetail', 'saveAiConfig', 'createWorksheetsForUnit', 'saveClassRoster',
          'deleteUnitTask', 'createNewUnit', 'importUnitJson', 'archiveUnitData',
        ];
        const bad = [];
        for (const fn of fns) {
          const body = fnSlice(s, fn);
          if (body === null) { bad.push(`${fn}（関数が無い）`); continue; }
          if (!/MiraiAuth\.requireTeacher\s*\(/.test(body)) bad.push(`${fn}（requireTeacher が無い）`);
        }
        return { ok: bad.length === 0, detail: bad.length ? '未ガード: ' + bad.join(', ') : `${fns.length} 関数すべてに requireTeacher あり` };
      },
    },
    {
      id: 'SEC_NO_PASSWORD_AUTH', title: 'パスワード認証を復活させていない（getActiveUser 移行の回帰防止）',
      run: ({ root }) => {
        const f = 'code.gs';
        if (!has(root, f)) return { ok: false, detail: 'code.gs が無い' };
        const s = stripComments(read(root, f), 'js');
        const bad = [];
        if (/\bTEACHER_PASS\b/.test(s)) bad.push('TEACHER_PASS');
        if (/\bverifyPassword\s*\(/.test(s)) bad.push('verifyPassword(');
        return { ok: bad.length === 0, detail: bad.length ? '復活している: ' + bad.join(', ') + '（認可は Session.getActiveUser 由来のメール許可制のみ）' : '無し（メール許可制のみ）' };
      },
    },

    // ---------------- C. 堅牢性 ----------------
    {
      id: 'C5_NO_LS_CLEAR', title: 'localStorage.clear() を使っていない',
      run: ({ root }) => {
        const bad = [];
        for (const f of cfg.sourceFiles) {
          if (!has(root, f)) continue;
          const s = stripComments(read(root, f), f.endsWith('.html') ? 'html' : 'js');
          if (/localStorage\s*\.\s*clear\s*\(/.test(s)) bad.push(f);
        }
        return { ok: bad.length === 0, detail: bad.length ? bad.join(', ') + '（同一オリジンの他アプリを巻き添えにする）' : '使っていない' };
      },
    },

    // ---------------- D. 表示 ----------------
    {
      id: 'D1_VIEWPORT_FIT', title: 'viewport-fit=cover（GAS は code.gs も）',
      run: ({ root }) => {
        const bad = [];
        for (const f of cfg.viewportFiles) {
          if (!has(root, f)) continue;
          const s = stripComments(read(root, f), f.endsWith('.html') ? 'html' : 'js');
          const declares = f.endsWith('.gs')
            ? /addMetaTag\s*\(\s*['"]viewport['"]/.test(s)
            : /<meta[^>]+name=["']viewport["']/i.test(s);
          if (!declares) continue;
          if (!/viewport-fit\s*=\s*cover/.test(s)) bad.push(f);
        }
        return { ok: bad.length === 0, detail: bad.length ? '不足: ' + bad.join(', ') + '（片方だけでは安全領域が使えるようにならない）' : '両方にあり' };
      },
    },
    {
      id: 'D14_NO_USER_SCALABLE', title: '拡大を禁止していない',
      run: ({ root }) => {
        const bad = [];
        for (const f of cfg.viewportFiles) {
          if (!has(root, f)) continue;
          const s = stripComments(read(root, f), f.endsWith('.html') ? 'html' : 'js');
          if (/user-scalable\s*=\s*no|maximum-scale\s*=\s*1(\.0)?\b/.test(s)) bad.push(f);
        }
        return { ok: bad.length === 0, detail: bad.length ? bad.join(', ') + '（見えづらい子が拡大できなくなる）' : '禁止していない' };
      },
    },
    {
      id: 'D2_DVH', title: '100vh の単独使用が無い',
      run: ({ root }) => {
        const bad = [];
        for (const f of cfg.styleFiles) {
          if (!has(root, f)) continue;
          const s = stripComments(read(root, f), 'html');
          const lines = s.split('\n');
          lines.forEach((line, i) => {
            if (!/\b100vh\b/.test(line)) return;
            // ★ @supports not (height: 100dvh) { … 100vh } はフォールバックなので正しい。
            //    その行だけを見ると誤検知するので、前方 8行も見る。
            const ctx = lines.slice(Math.max(0, i - 8), i + 1).join('\n');
            if (/@supports\s+not\s*\(\s*height\s*:\s*100dvh/.test(ctx)) return;
            if (/\b100dvh\b/.test(line)) return;
            bad.push(`${f}:${i + 1}: ${line.trim().slice(0, 70)}`);
          });
        }
        return { ok: bad.length === 0, detail: bad.length ? bad.join('\n      ') : '@supports フォールバックのみ' };
      },
    },
    {
      id: 'D3_SAFE_AREA', title: 'safe-area-inset を使っている',
      run: ({ root }) => {
        let n = 0;
        for (const f of cfg.styleFiles) {
          if (!has(root, f)) continue;
          n += (read(root, f).match(/safe-area-inset/g) || []).length;
        }
        return { ok: n >= 4, detail: `${n} 箇所（上下左右の4方向を想定）` };
      },
    },
    {
      id: 'D4_FLUID_TYPE', title: 'clamp() による fluid type',
      run: ({ root }) => {
        let n = 0;
        for (const f of cfg.styleFiles) {
          if (!has(root, f)) continue;
          n += (read(root, f).match(/(?:^|[^a-zA-Z0-9_-])clamp\s*\(/g) || []).length;
        }
        return { ok: n > 0, detail: `${n} 箇所` };
      },
    },
    {
      id: 'D5_CANVAS_DPR', title: 'Canvas に devicePixelRatio 補正（上限 2）',
      run: ({ root }) => {
        const uses = [], fixes = [];
        for (const f of cfg.sourceFiles) {
          if (!has(root, f)) continue;
          const s = stripComments(read(root, f), f.endsWith('.html') ? 'html' : 'js');
          if (/getContext\s*\(\s*['"]2d['"]/.test(s)) uses.push(f);
          if (/Math\.min\s*\(\s*(window\.)?devicePixelRatio[^,]*,\s*2\s*\)/.test(s)) fixes.push(f);
        }
        if (uses.length === 0) return { ok: true, detail: 'Canvas を使っていない' };
        const ok = fixes.length > 0;
        return { ok, detail: ok ? `補正あり: ${fixes.join(', ')}` : `getContext('2d') はあるが DPR 補正が無い: ${uses.join(', ')}` };
      },
    },
    {
      id: 'D10_REDUCED_MOTION', title: 'prefers-reduced-motion（.01ms であって 0 でない）',
      run: ({ root }) => {
        for (const f of cfg.styleFiles) {
          if (!has(root, f)) continue;
          const s = read(root, f);
          if (!/prefers-reduced-motion/.test(s)) continue;
          const block = s.slice(s.indexOf('prefers-reduced-motion'), s.indexOf('prefers-reduced-motion') + 500);
          // ★ 0 にすると fill-mode: forwards が壊れ、fadeIn 系が opacity:0 のまま消える
          if (/animation-duration\s*:\s*0m?s\b/.test(block)) {
            return { ok: false, detail: `${f}: animation-duration が 0。fill-mode: forwards が壊れて中身が消える` };
          }
          return { ok: true, detail: `${f} にあり（.01ms）` };
        }
        return { ok: false, detail: 'prefers-reduced-motion が無い' };
      },
    },
    {
      id: 'D11_FORCED_COLORS', title: 'forced-colors 対応',
      run: ({ root }) => {
        const found = cfg.styleFiles.filter((f) => has(root, f) && /forced-colors\s*:\s*active/.test(read(root, f)));
        return { ok: found.length > 0, detail: found.length ? found.join(', ') : '無い' };
      },
    },
    {
      id: 'D13_PRINT_CSS', title: '印刷 CSS がある',
      run: ({ root }) => {
        const found = cfg.styleFiles.filter((f) => has(root, f) && /@media\s+print/.test(read(root, f)));
        return { ok: found.length > 0, detail: found.length ? found.join(', ') : '無い' };
      },
    },
    {
      id: 'F4_RT_COLOR', title: 'rt（ふりがな）の色を決め打ちしていない',
      run: ({ root }) => {
        for (const f of cfg.styleFiles) {
          if (!has(root, f)) continue;
          const s = stripComments(read(root, f), 'html');
          if (!/(^|[\s,}])rt\s*\{/m.test(s)) continue;
          // 色を決めること自体は良い（白地の既定値）。
          // 見るべきは「色のついた面で inherit させているか」。
          const inherits = /rt\s*\{[^}]*\}|[^{}]*\brt\s*\{\s*color\s*:\s*inherit/s.test(s) &&
            /rt\s*\{\s*color\s*:\s*inherit/.test(s);
          if (!inherits) {
            return { ok: false, detail: `${f}: rt に色を決め打ちしているが、色のついた面で inherit させていない（青ボタンの上で比 1.04 になる）` };
          }
          return { ok: true, detail: `${f}: 色のついた面では inherit させている` };
        }
        return { ok: true, detail: 'rt の指定が無い' };
      },
    },
    {
      id: 'F6_FILE_SIZE', title: '1ファイル 5,000行 / 400KB 以内',
      run: ({ root }) => {
        const bad = [];
        for (const f of cfg.sourceFiles) {
          if (!has(root, f)) continue;
          const s = read(root, f);
          const lines = s.split('\n').length;
          const kb = statSync(join(root, f)).size / 1024;
          if (lines > 5000 || kb > 400) bad.push(`${f}: ${lines}行 / ${kb.toFixed(1)}KB`);
        }
        return { ok: bad.length === 0, detail: bad.length ? bad.join(', ') : '全ファイル基準内' };
      },
    },

    // ---------------- E. PWA ----------------
    {
      id: 'E1_MANIFEST_PATHS', title: 'manifest の id/scope/start_url が配信場所と合っている',
      run: ({ root }) => {
        const f = `${SHELL}/manifest.webmanifest`;
        if (!has(root, f)) return { ok: false, detail: 'manifest が無い' };
        const m = JSON.parse(read(root, f));
        // 独自ドメイン（CNAME あり）ではアプリはドメイン直下に置かれるので "/"。
        // 旧構成（gigayama.github.io/MIRAI-Compass/）のリポジトリ名の絶対パスを
        // 残すと、scope がページの URL を含まなくなって manifest ごと無視され、
        // PWA としてインストールできなくなる。実際にその状態で残っていた。
        const hasCname = has(root, 'CNAME') || has(root, `${SHELL}/CNAME`);
        const want = hasCname ? '/' : `/${REPO}/`;
        const bad = ['id', 'start_url', 'scope'].filter((k) => m[k] !== want);
        return {
          ok: bad.length === 0,
          detail: bad.length
            ? bad.map((k) => `${k}=${JSON.stringify(m[k])} （期待: "${want}"）`).join(', ')
            : `3つとも ${want}`,
        };
      },
    },
    {
      id: 'E2_APPLE_ICON_OPAQUE', title: 'apple-touch-icon に透明が無い',
      run: ({ root }) => {
        const f = `${SHELL}/icons/apple-touch-icon.png`;
        if (!has(root, f)) return { ok: false, detail: 'apple-touch-icon.png が無い' };
        const r = scanPngAlpha(readFileSync(join(root, f)));
        if (r.unsupported) return { ok: false, detail: '形式を解けなかった（未計測。✅ にはしない）' };
        const pct = (r.transparentPixels / r.totalPixels * 100).toFixed(2);
        const cpct = r.cornerTotal ? (r.cornerTransparent / r.cornerTotal * 100).toFixed(2) : '0.00';
        return {
          ok: r.transparentPixels === 0,
          detail: r.transparentPixels === 0
            ? `透明 0 画素（${r.width}×${r.height}）`
            : `透明 ${pct}% / 四隅ボックス ${cpct}% … iOS がそこを黒で塗る`,
        };
      },
    },
    {
      id: 'E3_INSTALL_HOOK', title: 'beforeinstallprompt を head 最上部の外部ファイルで捕捉',
      run: ({ root }) => {
        const hookFile = `${SHELL}/install-hook.js`;
        if (!has(root, hookFile)) return { ok: false, detail: 'install-hook.js が無い' };
        if (!/beforeinstallprompt/.test(read(root, hookFile))) {
          return { ok: false, detail: 'install-hook.js が beforeinstallprompt を捕捉していない' };
        }
        const s = read(root, `${SHELL}/index.html`);
        const head = s.slice(0, s.search(/<\/head>/i));
        const scripts = [...head.matchAll(/<script\b[^>]*src\s*=\s*["']([^"']+)["']/gi)].map((m) => m[1]);
        const ok = scripts.length > 0 && /install-hook\.js$/.test(scripts[0]);
        return { ok, detail: ok ? 'head の最初の script' : `head の script 順: ${scripts.join(', ') || '(無し)'}` };
      },
    },
    {
      id: 'E5_SW_CACHE_SCOPE', title: 'sw.js が自アプリ接頭辞のキャッシュだけ削除',
      run: ({ root }) => {
        const f = `${SHELL}/sw.js`;
        if (!has(root, f)) return { ok: false, detail: 'sw.js が無い' };
        const s = stripComments(read(root, f), 'js');
        if (!/caches\s*\.\s*keys\s*\(/.test(s)) return { ok: true, detail: 'caches.keys() を使っていない' };
        // ★「消す式」を追わない。「startsWith で絞る式があるか」を見る。
        //   削除式を正規表現で追うと (k) => caches.delete(k) を見落とす。
        const ok = /\.\s*filter\s*\([\s\S]{0,200}?startsWith\s*\(/.test(s);
        return { ok, detail: ok ? 'startsWith で自アプリ分に絞っている' : 'caches.keys() の結果を絞らずに消している（他アプリを巻き添えにする）' };
      },
    },
    {
      id: 'E6_SW_NO_LOCALSTORAGE', title: 'sw.js が localStorage に触れていない',
      run: ({ root }) => {
        const f = `${SHELL}/sw.js`;
        if (!has(root, f)) return { ok: false, detail: 'sw.js が無い' };
        // ★ 判定の前にコメントを落とす。
        //   「localStorage は操作しない」という注意書きに反応してしまうため。
        const s = stripComments(read(root, f), 'js');
        const bad = /\blocalStorage\b/.test(s);
        return { ok: !bad, detail: bad ? 'localStorage を操作している' : '触れていない' };
      },
    },
    {
      id: 'E7_SW_NO_SKIP_WAITING_IN_INSTALL', title: 'install の中で skipWaiting していない',
      run: ({ root }) => {
        const f = `${SHELL}/sw.js`;
        if (!has(root, f)) return { ok: false, detail: 'sw.js が無い' };
        const s = stripComments(read(root, f), 'js');
        const i = s.search(/addEventListener\s*\(\s*['"]install['"]/);
        if (i < 0) return { ok: false, detail: 'install ハンドラが無い' };
        const j = s.search(/addEventListener\s*\(\s*['"](activate|fetch|message)['"]/);
        const block = s.slice(i, j > i ? j : s.length);
        const bad = /skipWaiting\s*\(/.test(block);
        return { ok: !bad, detail: bad ? 'install で skipWaiting している（押す前に切り替わる）' : 'していない' };
      },
    },
    {
      id: 'E7b_UPDATE_PROMPT', title: '更新は利用者が押したときだけ受ける',
      run: ({ root }) => {
        const f = `${SHELL}/app.js`;
        if (!has(root, f)) return { ok: false, detail: `${f} が無い` };
        const s = stripComments(read(root, f), 'js');
        if (!/controllerchange/.test(s)) return { ok: false, detail: 'controllerchange を見ていない' };
        // ★ controllerchange は初回訪問でも飛んでくる。
        //   「押したか」のフラグで守っていなければ、初回が必ず1回リロードされる。
        const i = s.indexOf('controllerchange');
        const block = s.slice(i, i + 400);
        const ok = /if\s*\(\s*!\s*\w*[Aa]sk\w*/.test(block) || /userAskedUpdate/.test(block);
        return { ok, detail: ok ? '押したときだけ受けている' : '無条件に受けている（初回訪問が必ず1回リロードされる）' };
      },
    },
    {
      id: 'E9_SW_REGISTER_READYSTATE', title: 'Service Worker の登録に readyState の分岐がある',
      run: ({ root }) => {
        const f = `${SHELL}/app.js`;
        if (!has(root, f)) return { ok: false, detail: `${f} が無い` };
        const s = stripComments(read(root, f), 'js');
        if (!/serviceWorker\s*\.\s*register/.test(s)) return { ok: false, detail: '登録していない' };
        const ok = /readyState\s*===?\s*['"]complete['"]/.test(s);
        return { ok, detail: ok ? 'あり' : "load を待つだけだと、すでに load 済みのとき二度と呼ばれない" };
      },
    },
    {
      id: 'E10_OFFLINE_HTML', title: 'offline.html があり、外部資産にも JS にも頼らない',
      run: ({ root }) => {
        const f = `${SHELL}/offline.html`;
        if (!has(root, f)) return { ok: false, detail: '無い' };
        const s = stripComments(read(root, f), 'html');
        const bad = [];
        if (/<script/i.test(s)) bad.push('<script> がある');
        if (/(src|href)\s*=\s*["']https?:\/\//i.test(s)) bad.push('外部資産を読んでいる');
        if (/\bon[a-z]+\s*=\s*["']/i.test(s)) bad.push('インラインハンドラがある（CSP に引っかかる）');
        return { ok: bad.length === 0, detail: bad.length ? bad.join(', ') : '自前で完結している' };
      },
    },
    {
      id: 'E11_APP_VERSION', title: 'APP_VERSION が sw.js にある',
      run: ({ root }) => {
        const f = `${SHELL}/sw.js`;
        if (!has(root, f)) return { ok: false, detail: 'sw.js が無い' };
        const m = read(root, f).match(/APP_VERSION\s*=\s*['"]([^'"]+)['"]/);
        return { ok: !!m, detail: m ? m[1] : '無い' };
      },
    },
    {
      id: 'E12_MASKABLE_SAFEZONE', title: 'maskable アイコンに透明が無い（下地が端まで伸びている）',
      run: ({ root }) => {
        const bad = [];
        for (const n of ['icon-maskable-192.png', 'icon-maskable-512.png']) {
          const f = `${SHELL}/icons/${n}`;
          if (!has(root, f)) { bad.push(`${n}: 無い`); continue; }
          const r = scanPngAlpha(readFileSync(join(root, f)));
          if (r.unsupported) { bad.push(`${n}: 未計測`); continue; }
          if (r.transparentPixels > 0) {
            bad.push(`${n}: 透明 ${(r.transparentPixels / r.totalPixels * 100).toFixed(2)}%（切り抜きの内側が余白色で埋まり、縮んで見える）`);
          }
        }
        return { ok: bad.length === 0, detail: bad.length ? bad.join(', ') : '2種とも透明 0 画素' };
      },
    },

    // ---------------- 生成物 ----------------
    {
      id: 'G1_GENERATED_MARKED', title: '生成物に「手で編集しない」と書いてある',
      run: ({ root }) => {
        const bad = [];
        for (const f of cfg.generatedFiles) {
          if (!has(root, f)) { bad.push(`${f}: 無い（npm run build を実行）`); continue; }
          const head = read(root, f).slice(0, 600);
          if (!/手で編集しない/.test(head)) bad.push(`${f}: 注意書きが無い`);
        }
        return { ok: bad.length === 0, detail: bad.length ? bad.join(', ') : '全て記載あり' };
      },
    },
  ];
}
