/**
 * このリポジトリだけの検査。
 *
 * 共通の検査は正本（GIGAyama.github.io/standards/lib/giga-v5-checks.mjs）が
 * 受け持つ。ここに残すのは、正本に対応するものが無いものだけである。
 *
 * ⚠️ なぜ 10 件も残るのか（2026-08-23 に実測して決めた）
 *
 *   このリポジトリは「GitHub Pages の入口シェル（docs/）」と
 *   「GAS 本体（code.gs と *.html の画面）」の2つでできている。
 *   正本が見るのは前者だけである。GAS のテンプレート（index.html が
 *   css.html / js_*.html を include する形）は、正本の想定に無い。
 *
 *   移行のとき、GAS 側だけを12通りに壊して正本に当ててみた。
 *   **11 通りが素通りした。** 正本を入れて「38/38 通過」と出ていても、
 *   先生と児童が実際に使う画面は1行も見ていない状態になる。
 *   だからここに残す。数が多いのは手抜きではなく、実測の結果である。
 *
 *   1件（秘密の直書き）だけは tools/check-secrets.mjs が別に見ているので
 *   ここには置かない（あちらは無ければコマンドごと失敗する作り）。
 *
 * ⚠️ 検査そのものが壊れていないかは check-project.mjs --self-test が確かめる。
 */
import { existsSync, readFileSync, statSync } from 'node:fs';
import { join } from 'node:path';

/**
 * コメントを落とす。判定を注意書きに反応させないため。
 *
 * ⚠️ URL を先に伏せてからコメントを落とす。
 *   `https://*.googleusercontent.com` の `/*` をブロックコメントの開始と読んで
 *   しまい、そこから次の `*&#47;` までを丸ごと消していた。実際、CSP の
 *   frame-ancestors 検査が「壊したのに通る」状態になっていた（消された範囲に
 *   検査対象が入っていたため）。自己診断で見つかった欠陥である。
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

const read = (root, f) => readFileSync(join(root, f), 'utf8');
const has = (root, f) => existsSync(join(root, f));
const kindOf = (f) => (f.endsWith('.html') ? 'html' : 'js');

/**
 * @param {object} cfg quality.config.json の local セクション
 * @returns {{id:string,title:string,run:(ctx:{root:string})=>{ok:boolean,detail:string}}[]}
 */
export function buildLocalChecks(cfg) {
  return [
    // ---------------- サーバー側の認可（GAS の砦） ----------------
    //
    // 画面側の出し分けは防御にならない。google.script.run は、画面に
    // ボタンが出ていなくても関数名さえ分かれば直接呼べる。
    {
      id: 'SEC_PRIVILEGED_FN_GUARDED',
      title: '先生専用関数がサーバー側で requireTeacher している',
      run: ({ root }) => {
        const f = 'code.gs';
        if (!has(root, f)) return { ok: false, detail: 'code.gs が無い' };
        const s = stripComments(read(root, f), 'js');
        const fns = [
          'saveWorksheetToDB', 'saveFeedback', 'batchSaveFeedback', 'generateSingleWorksheet',
          'generateRubricAI', 'generateBatchComments', 'getDashboardData', 'getTaskSubmissions',
          'getSubmissionDetail', 'saveAiConfig', 'createWorksheetsForUnit', 'saveClassRoster',
          'deleteUnitTask', 'createNewUnit', 'importUnitJson', 'archiveUnitData',
          // スプレッドシートのコピーで配る形にしたときに足した、先生専用の API。
          // google.script.run はトップレベル関数を誰でも呼べるので、1つ抜けると境界が破れる。
          // （メニューから呼ぶ showSheetCheck / repairSheetsFromMenu はここに入れない。
          //   あちらは先に SpreadsheetApp.getUi() を取って、画面が無い文脈で止める作り）
          'getSetupStatus', 'getDatabaseHealth', 'repairDatabase',
        ];
        const bad = [];
        for (const fn of fns) {
          const body = fnSlice(s, fn);
          if (body === null) { bad.push(`${fn}（関数が無い）`); continue; }
          if (!/MiraiAuth\.requireTeacher\s*\(/.test(body)) bad.push(`${fn}（requireTeacher が無い）`);
        }
        return {
          ok: bad.length === 0,
          detail: bad.length ? '未ガード: ' + bad.join(', ') : `${fns.length} 関数すべてに requireTeacher あり`,
        };
      },
    },
    {
      id: 'SEC_NO_PASSWORD_AUTH',
      title: 'パスワード認証を復活させていない（getActiveUser 移行の回帰防止）',
      run: ({ root }) => {
        const f = 'code.gs';
        if (!has(root, f)) return { ok: false, detail: 'code.gs が無い' };
        const s = stripComments(read(root, f), 'js');
        const bad = [];
        if (/\bTEACHER_PASS\b/.test(s)) bad.push('TEACHER_PASS');
        if (/\bverifyPassword\s*\(/.test(s)) bad.push('verifyPassword(');
        return {
          ok: bad.length === 0,
          detail: bad.length
            ? '復活している: ' + bad.join(', ') + '（認可は Session.getActiveUser 由来のメール許可制のみ）'
            : '無し（メール許可制のみ）',
        };
      },
    },

    // ---------------- GAS 側の堅牢性 ----------------
    {
      id: 'GAS_NO_POSTMESSAGE_STAR',
      title: 'GAS 側の postMessage の宛先が * でない',
      run: ({ root }) => {
        const bad = [];
        for (const f of cfg.sourceFiles) {
          if (!has(root, f)) continue;
          const s = stripComments(read(root, f), kindOf(f));
          if (/postMessage\s*\([^)]*,\s*['"]\*['"]\s*\)/.test(s)) bad.push(f);
        }
        return { ok: bad.length === 0, detail: bad.length ? bad.join(', ') + '（どのページに届いてもよい、という意味になる）' : '宛先を決めている' };
      },
    },
    {
      id: 'GAS_NO_LS_CLEAR',
      title: 'GAS 側が localStorage.clear() を使っていない',
      run: ({ root }) => {
        const bad = [];
        for (const f of cfg.sourceFiles) {
          if (!has(root, f)) continue;
          const s = stripComments(read(root, f), kindOf(f));
          if (/localStorage\s*\.\s*clear\s*\(/.test(s)) bad.push(f);
        }
        return { ok: bad.length === 0, detail: bad.length ? bad.join(', ') + '（同一オリジンの他アプリを巻き添えにする）' : '使っていない' };
      },
    },
    {
      id: 'GAS_FILE_SIZE',
      title: 'GAS 側の1ファイルが 5,000行 / 400KB 以内',
      run: ({ root }) => {
        const bad = [];
        for (const f of cfg.sourceFiles) {
          if (!has(root, f)) continue;
          const lines = read(root, f).split('\n').length;
          const kb = statSync(join(root, f)).size / 1024;
          if (lines > 5000 || kb > 400) bad.push(`${f}: ${lines}行 / ${kb.toFixed(1)}KB`);
        }
        return { ok: bad.length === 0, detail: bad.length ? bad.join(', ') : '全ファイル基準内' };
      },
    },

    // ---------------- GAS 側の表示 ----------------
    //
    // ⚠️ GAS の viewport は code.gs の addMetaTag('viewport', …) で決まる。
    //    HTML を読んでも書いていないので、正本の D_VIEWPORT では見つけられない。
    {
      id: 'GAS_VIEWPORT_FIT',
      title: 'GAS 側の viewport に viewport-fit=cover がある（code.gs も見る）',
      run: ({ root }) => {
        const bad = [];
        for (const f of cfg.viewportFiles) {
          if (!has(root, f)) continue;
          const s = stripComments(read(root, f), kindOf(f));
          const declares = f.endsWith('.gs')
            ? /addMetaTag\s*\(\s*['"]viewport['"]/.test(s)
            : /<meta[^>]+name=["']viewport["']/i.test(s);
          if (!declares) continue;
          if (!/viewport-fit\s*=\s*cover/.test(s)) bad.push(f);
        }
        return { ok: bad.length === 0, detail: bad.length ? bad.join(', ') + '（ノッチのある端末で下端が隠れる）' : '全て cover' };
      },
    },
    {
      id: 'GAS_NO_USER_SCALABLE',
      title: 'GAS 側が拡大を禁止していない',
      run: ({ root }) => {
        const bad = [];
        for (const f of cfg.viewportFiles) {
          if (!has(root, f)) continue;
          const s = stripComments(read(root, f), kindOf(f));
          if (/user-scalable\s*=\s*no|maximum-scale\s*=\s*1(\.0)?\b/.test(s)) bad.push(f);
        }
        return { ok: bad.length === 0, detail: bad.length ? bad.join(', ') + '（見えづらい子が拡大できなくなる）' : '禁止していない' };
      },
    },
    {
      id: 'GAS_CANVAS_DPR',
      title: 'GAS 側の Canvas に devicePixelRatio 補正（上限 2）',
      run: ({ root }) => {
        const uses = [], fixes = [];
        for (const f of cfg.sourceFiles) {
          if (!has(root, f)) continue;
          const s = stripComments(read(root, f), kindOf(f));
          if (/getContext\s*\(\s*['"]2d['"]/.test(s)) uses.push(f);
          if (/Math\.min\s*\(\s*(window\.)?devicePixelRatio[^,]*,\s*2\s*\)/.test(s)) fixes.push(f);
        }
        if (uses.length === 0) return { ok: true, detail: 'Canvas を使っていない' };
        const ok = fixes.length > 0;
        return { ok, detail: ok ? `補正あり: ${fixes.join(', ')}` : `getContext('2d') はあるが DPR 補正が無い: ${uses.join(', ')}` };
      },
    },
    {
      id: 'GAS_PRINT_CSS',
      title: '印刷 CSS がある（ワークシートを紙で配るため）',
      run: ({ root }) => {
        const found = cfg.styleFiles.filter((f) => has(root, f) && /@media\s+print/.test(read(root, f)));
        return { ok: found.length > 0, detail: found.length ? found.join(', ') : '無い' };
      },
    },

    // ---------------- 生成物 ----------------
    {
      id: 'GENERATED_MARKED',
      title: '生成物に「手で編集しない」と書いてある',
      run: ({ root }) => {
        const bad = [];
        for (const f of cfg.generatedFiles) {
          if (!has(root, f)) { bad.push(`${f}: 無い（npm run build を実行）`); continue; }
          if (!/手で編集しない/.test(read(root, f).slice(0, 600))) bad.push(`${f}: 注意書きが無い`);
        }
        return { ok: bad.length === 0, detail: bad.length ? bad.join(', ') : '全て記載あり' };
      },
    },
  ];
}
