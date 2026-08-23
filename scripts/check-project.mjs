#!/usr/bin/env node
/**
 * 品質ゲート（GIGA Standard v5）。
 *
 * 構成
 *   scripts/lib/giga-v5-checks.mjs … 正本の写し（GIGAyama.github.io/standards/lib/）。
 *                                    直接いじらない。直すなら正本を直して配る
 *                                    （ずれたら CI の drift ジョブが赤くする）
 *   scripts/lib/local-checks.mjs   … 正本に行き先が無い、このリポジトリだけの検査
 *   scripts/check-project.mjs      … 両者を合成する（このファイル）
 *
 * ⚠️ ローカルが 10 件もあるのは、このリポジトリが「Pages の入口シェル（docs/）」と
 *    「GAS 本体（code.gs と *.html の画面）」の2つでできているためである。
 *    正本が見るのは前者だけで、GAS 側だけを12通りに壊して当てたところ
 *    11 通りが素通りした（2026-08-23 実測）。理由は local-checks.mjs の冒頭に書いた。
 *
 * ★ ゲートは必ず「わざと壊して」通ることを確認する。
 *   「0件でした」だけでは、検査が動いているのか何も見ていないのか区別できない。
 *     node scripts/check-project.mjs --self-test
 *
 * 使い方
 *   node scripts/check-project.mjs             検査する
 *   node scripts/check-project.mjs --self-test 検査そのものが動くか確かめる
 */
import { readFileSync } from 'node:fs';
import { join, dirname } from 'node:path';
import { fileURLToPath } from 'node:url';
import { runGigaChecks } from './lib/giga-v5-checks.mjs';
import { buildLocalChecks } from './lib/local-checks.mjs';

const ROOT = join(dirname(fileURLToPath(import.meta.url)), '..');

/**
 * 設定は「いま見ている木」から読む。
 *
 * ⚠️ 外側の定数にしてはいけない。--self-test は木ごと写して壊すので、
 *    写しの quality.config.json を壊しても効かず、「壊したのに落ちない」
 *    検査ができてしまう（100マス計算で実際に起きた）。
 */
const configOf = (root) => JSON.parse(readFileSync(join(root, 'quality.config.json'), 'utf8'));

const GREEN = '\x1b[32m', RED = '\x1b[31m', DIM = '\x1b[2m', OFF = '\x1b[0m';

// かつてここで、フリート共通の正本 scripts/lib/project-quality.mjs を
// 「あれば足す、無ければ Part I の検査だけ」で読んでいた。外した理由
// （2026-08-22 に実測）:
//
//   ・その正本は一度も取り込まれず、**何の知らせも出さないまま**素通り
//     していた。含まれていた秘密の直書きの検査も働いていなかった。
//   ・しかも取り込めば動く、というものでもなかった。この枝は
//     m.buildChecks を探すが、艦隊にある8本のコピーはどれもその名前を
//     export していない（6本が runQualityChecks、1本が run）。実際に
//     gamification のコピーを置いて走らせても、検査は 38 件のまま
//     1件も増えなかった。
//
// 秘密の直書きは tools/check-secrets.mjs が見る（#26 で入れた）。あちらは
// 丸ごと1ファイルで完結し、無ければコマンドごと失敗するので、
// 「取り込み忘れたまま緑」にはならない。
async function run(root) {
  const cfg = configOf(root);
  const results = [];

  // ① 正本の検査。{id, ok, detail[], severity, skip} を、ここの形に読みかえる。
  //    理由つきで飛ばしたものは title に「スキップ: …」が付いて残るので、
  //    黙って消えることはない。
  for (const r of runGigaChecks(root, cfg.standard)) {
    results.push({ id: r.id, title: r.title, ok: r.ok, detail: r.detail.join(' / ') });
  }

  // ② このリポジトリだけの検査（GAS 本体を見る）。
  for (const c of buildLocalChecks(cfg.local)) {
    let r;
    try { r = c.run({ root }); }
    catch (e) { r = { ok: false, detail: '検査が例外で落ちた: ' + String(e).slice(0, 160) }; }
    results.push({ ...c, ...r });
  }
  return results;
}

function report(results) {
  let failed = 0;
  for (const r of results) {
    if (r.ok) {
      console.log(`${GREEN}✅${OFF} ${r.id.padEnd(32)} ${r.title}`);
      if (r.detail) console.log(`   ${DIM}${r.detail}${OFF}`);
    } else {
      failed++;
      console.log(`${RED}❌${OFF} ${r.id.padEnd(32)} ${r.title}`);
      console.log(`      ${RED}${r.detail}${OFF}`);
    }
  }
  console.log(`\n${results.length - failed} / ${results.length} 通過`);
  return failed;
}

// ---------------------------------------------------------------------------
// --self-test … わざと壊して、検査が本当に見ているかを確かめる
// ---------------------------------------------------------------------------
async function selfTest() {
  const { mkdtempSync, cpSync, writeFileSync, readFileSync: rf, rmSync } = await import('node:fs');
  const { tmpdir } = await import('node:os');

  // 各検査を落とすはずの「壊し方」。ここに書いていない検査は自己診断の対象外。
  // 各検査を落とすはずの「壊し方」。ここに書いていない検査は下で赤になる
  // （壊して確かめていない検査は、動いている保証がないため）。
  //
  // ⚠️ 正本（38件）とローカル（10件）の両方を並べる。同じ観点でも、
  //    シェル側と GAS 側では別の検査が見ているので、両方を壊すこと。
  //    片方だけ壊して満足すると、もう片方が何も見ていなくても気づけない。
  const BREAKAGES = [
    // ---- そろえておくもの（正本）
    ['A_LICENSE', (d) => rmSync(join(d, 'LICENSE'))],
    ['A_GITIGNORE', (d) => writeFileSync(join(d, '.gitignore'), 'dist/\n')],
    ['A_DEPENDABOT', (d) => rmSync(join(d, '.github/dependabot.yml'))],
    ['A_DOCS', (d) => rmSync(join(d, 'MANUAL.md'))],
    ['A_CI_ON_PR', (d) => writeFileSync(join(d, '.github/workflows/ci.yml'), 'on:\n  push:\n    branches: [main]\njobs: {}\n')],

    // ---- 差しこまれたコードを止める（正本・シェル）
    ['B_NO_CDN_CODE', (d) => patch(d, 'docs/index.html', (s) => s.replace('</head>', '<script src="https://cdn.jsdelivr.net/npm/chart.js"></script></head>'))],
    ['B_NO_SECRETS', (d) => patch(d, 'docs/app.js', (s) => 'const apiKey = "AIzaSyD0123456789abcdefghijklmnopqrstuv";\n' + s)],
    ['B_CSP', (d) => patch(d, 'docs/index.html', (s) => s.replace("script-src 'self';", "script-src 'self' 'unsafe-inline';"))],
    ['B_NO_INLINE_SCRIPT', (d) => patch(d, 'docs/index.html', (s) => s.replace('</body>', '<script>window.x = 1;</script></body>'))],

    // ---- 打ちかけを消さない（正本・シェル）
    ['C_NO_LS_CLEAR', (d) => patch(d, 'docs/app.js', (s) => s.replace('localStorage.removeItem(LS_KEY)', 'localStorage.clear()'))],
    ['C_NO_POSTMESSAGE_STAR', (d) => patch(d, 'docs/app.js', (s) => s + '\nfunction __t(w) { w.postMessage({ a: 1 }, "*"); }\n')],
    // ⚠️ C_PAGEHIDE は quality.config.json の skips で理由つきで飛ばしている。
    //    壊し方は「その宣言を外すこと」。こうしておくと、飛ばしているから
    //    通っているのか、何も見ていないから通っているのかを区別できる。
    ['C_PAGEHIDE', (d) => patchJson(d, 'quality.config.json', (j) => { j.standard.skips = []; })],

    // ---- 表示（正本・シェルと GAS の CSS の両方を見ている）
    ['D_VIEWPORT', (d) => patch(d, 'docs/index.html', (s) => s.replace(', viewport-fit=cover', ''))],
    // ⚠️ css.html には 100vh が5か所ある。どれも @supports not (height: 100dvh) の
    //    正しいひかえなので、そこを書きかえても落ちない（実際そう書いて素通りした）。
    //    ひかえの外に、素の 100vh を1つ足すのが正しい壊し方である。
    ['D_DVH', (d) => patch(d, 'css.html', (s) => s.replace('</style>', '.__selftest { height: 100vh; }\n</style>'))],
    ['D_SAFE_AREA', (d) => { for (const f of ['css.html', 'docs/index.html', 'docs/offline.html']) patch(d, f, (s) => s.replace(/safe-area-inset/g, 'SAFEAREAINSET')); }],
    ['D_FLUID_TYPE', (d) => { for (const f of ['css.html', 'docs/index.html', 'docs/offline.html']) patch(d, f, (s) => s.replace(/clamp\(/g, 'min(')); }],
    // ⚠️ 正本が見るのはシェル側（docs/app.js）だけである。GAS の js_student.html を
    //    壊しても正本は素通りする（そちらは GAS_CANVAS_DPR が見る）。
    //    シェル側に、補正の無い Canvas をわざと1つ置いて確かめる。
    ['D_CANVAS_DPR', (d) => patch(d, 'docs/app.js', (s) => s + '\nfunction __t() { document.createElement("canvas").getContext("2d"); }\n')],
    ['D_REDUCED_MOTION', (d) => patch(d, 'css.html', (s) => s.replace('animation-duration: .01ms !important;', 'animation-duration: 0ms !important;'))],
    ['D_FORCED_COLORS', (d) => { for (const f of ['css.html', 'docs/index.html', 'docs/offline.html']) patch(d, f, (s) => s.replace(/forced-colors: active/g, 'forced-colours: active')); }],
    // ⚠️ rt に色を足すだけでは落ちない。正本は「色のついた面で継がせる手当てが
    //    CSS のどこかにあるか」も見るので、css.html の
    //      button rt, .btn rt, [class*="bg-"] rt … { color: inherit; }
    //    が身代わりになる。**手当てごと外して**はじめて落ちる。
    //    これは「ふりがなの色を固定し、色のついた面の逃げ道も消した」という、
    //    実際に起きうる後戻りそのものである（青ボタンの上で比 1.04 になる）。
    ['D_RT_COLOR', (d) => patch(d, 'css.html', (s) => s
      .replace('button rt,', '.__removed rt,')
      .replace('[class*="bg-"] rt,', '.__removed2 rt,')
      .replace('[class*="btn-"] rt,', '.__removed3 rt,')
      .replace('[class*="text-white"] rt,', '.__removed4 rt,')
      .replace('.skill-mini-badge rt { color: inherit; }', '.skill-mini-badge rt { color: #6c757d; }')
      .replace('rt {\n  font-size: 0.6em;', 'rt {\n  font-size: 0.6em;\n  color: #6c757d;'))],

    // ---- 配信の形（正本）
    ['E_MANIFEST_ID', (d) => patch(d, 'docs/manifest.webmanifest', (s) => s.replace(/"id": "[^"]*"/, '"id": "/MIRAI-Compass/"'))],
    ['E_CNAME', (d) => patch(d, 'docs/CNAME', (s) => s + '\nextra.example.com\n')],
    ['E_STALE_REPO_PATH', (d) => patch(d, 'docs/app.js', (s) => s + '\nconst __u = "/MIRAI-Compass/app.js";\n')],
    ['E_ICONS', (d) => cpSync(join(d, 'docs/icons/icon-192.png'), join(d, 'docs/icons/apple-touch-icon.png'))],
    ['E_MASKABLE_SAFE_ZONE', (d) => cpSync(join(d, 'docs/icons/icon-192.png'), join(d, 'docs/icons/icon-maskable-192.png'))],
    ['E_INSTALL_HOOK', (d) => patch(d, 'docs/index.html', (s) => s.replace('<script src="install-hook.js"></script>', ''))],

    // ---- Service Worker（正本）
    ['E_SW_CACHE_SCOPE', (d) => patch(d, 'docs/sw.js', (s) => s.replace(/\.filter\(\(k\) => k\.startsWith\(CACHE_PREFIX\) && k !== CACHE_VERSION\)/, ''))],
    ['E_SW_NO_LOCALSTORAGE', (d) => patch(d, 'docs/sw.js', (s) => s.replace('const CACHE_PREFIX', 'localStorage.setItem("x","1");\nconst CACHE_PREFIX'))],
    ['E_SW_NO_SKIP_WAITING_ON_INSTALL', (d) => patch(d, 'docs/sw.js', (s) => s.replace('const cache = await caches.open(CACHE_VERSION);', 'self.skipWaiting();\n    const cache = await caches.open(CACHE_VERSION);'))],
    ['E_SW_UPDATE_PROMPT', (d) => patch(d, 'docs/app.js', (s) => s.replace('if (!userAskedUpdate || reloading) return;', ''))],
    ['E_SW_REGISTER_READYSTATE', (d) => patch(d, 'docs/app.js', (s) => s.replace("if (document.readyState === 'complete') startServiceWorker();\n  else ", ''))],
    ['E_SW_VERSION_GENERATED', (d) => patch(d, 'docs/sw.js', (s) => s.replace(/const APP_VERSION = '([^']*)'; \/\* __APP_VERSION__ \*\//, "const APP_VERSION = '$1';"))],
    ['E_OFFLINE_HTML', (d) => patch(d, 'docs/offline.html', (s) => s.replace('</body>', '<script>console.log(1)</script></body>'))],
    ['E_SW_PRECACHE_OFFLINE', (d) => patch(d, 'docs/sw.js', (s) => s.replace("  './offline.html',\n", ''))],

    // ---- 重さと画像（正本）
    ['F_FILE_SIZE', (d) => patch(d, 'docs/app.js', (s) => s + '\n' + '// ながすぎ\n'.repeat(5200))],
    ['F_IMG_SIZE', (d) => writeFileSync(join(d, 'docs/icons/icon-512.png'), Buffer.concat([rf(join(d, 'docs/icons/icon-512.png')), Buffer.alloc(200 * 1024, 7)]))],
    ['F_IMG_DIMENSIONS', (d) => patch(d, 'docs/index.html', (s) => s.replace(/(<img\b[^>]*?)\s+width="[^"]*"\s+height="[^"]*"/, '$1'))],
    // ⚠️ このリポジトリに <label for> は1つも無い。つまり素の状態では
    //    「見るものが無いから通っている」。マウスでしか押せないボタンの形を
    //    わざと1つ置いて、検査が働くことを確かめる。
    ['F_LABEL_FOR_TABBABLE', (d) => patch(d, 'docs/index.html', (s) => s.replace('</body>',
      '<label class="btn" for="__t">えらぶ</label><input type="file" id="__t" hidden>\n</body>'))],

    // ---- GAS 本体（ローカル）。正本はここを1行も見ない
    ['SEC_PRIVILEGED_FN_GUARDED', (d) => patch(d, 'code.gs', (s) => s.replace(
      'MiraiAuth.requireTeacher();\n    var props = PropertiesService.getScriptProperties();',
      'var props = PropertiesService.getScriptProperties();'))],
    ['SEC_NO_PASSWORD_AUTH', (d) => patch(d, 'code.gs', (s) => s.replace(
      'var MiraiAuth = (function () {',
      'var TEACHER_PASS = ScriptProps.get("pw"); function verifyPassword(p) { return p === TEACHER_PASS; }\nvar MiraiAuth = (function () {'))],
    ['GAS_NO_POSTMESSAGE_STAR', (d) => patch(d, 'js_core.html', (s) => s.replace('<script>', '<script>\nwindow.parent.postMessage({a:1}, "*");'))],
    ['GAS_NO_LS_CLEAR', (d) => patch(d, 'js_core.html', (s) => s.replace('<script>', '<script>\nlocalStorage.clear();'))],
    ['GAS_FILE_SIZE', (d) => patch(d, 'js_core.html', (s) => s + '\n' + '// ながすぎ\n'.repeat(5200))],
    ['GAS_VIEWPORT_FIT', (d) => patch(d, 'code.gs', (s) => s.replace(", viewport-fit=cover'", "'"))],
    ['GAS_NO_USER_SCALABLE', (d) => patch(d, 'index.html', (s) => s.replace('viewport-fit=cover">', 'viewport-fit=cover, user-scalable=no">'))],
    ['GAS_CANVAS_DPR', (d) => patch(d, 'js_student.html', (s) => s.replace(/Math\.min\(window\.devicePixelRatio \|\| 1, 2\)/, '1'))],
    ['GAS_PRINT_CSS', (d) => patch(d, 'css.html', (s) => s.replace(/@media print/g, '@media screenonly'))],
    ['GENERATED_MARKED', (d) => patch(d, 'vendor_css.html', (s) => s.replace('手で編集しない', ''))],
  ];

  /** JSON を読んで書きかえる（設定そのものを壊す変異に使う） */
  function patchJson(dir, file, fn) {
    const p = join(dir, file);
    const j = JSON.parse(rf(p, 'utf8'));
    fn(j);
    writeFileSync(p, JSON.stringify(j, null, 2) + '\n');
  }

  function patch(dir, file, fn) {
    const p = join(dir, file);
    const before = rf(p, 'utf8');
    const after = fn(before);
    if (after === before) throw new Error(`壊し方が効いていない: ${file}`);
    writeFileSync(p, after);
  }

  // まず、素の状態で全部通ることを確かめる
  const baseline = await run(ROOT);
  const baseFails = baseline.filter((r) => !r.ok);
  if (baseFails.length) {
    console.log(`${RED}自己診断の前提が崩れています。素の状態で ${baseFails.length} 件落ちています:${OFF}`);
    for (const r of baseFails) console.log(`   ${r.id}: ${r.detail}`);
    return 1;
  }
  console.log(`${DIM}素の状態: ${baseline.length} 件すべて通過。ここから1件ずつわざと壊します。${OFF}\n`);

  let bad = 0;
  const covered = new Set();
  for (const [id, breaker] of BREAKAGES) {
    covered.add(id);
    const dir = mkdtempSync(join(tmpdir(), 'gate-'));
    cpSync(ROOT, dir, {
      recursive: true,
      filter: (src) => !/[/\\](node_modules|\.git)([/\\]|$)/.test(src),
    });
    let verdict;
    try {
      breaker(dir);
      const res = await run(dir);
      const target = res.find((r) => r.id === id);
      if (!target) verdict = `${RED}検査 ${id} が見つからない${OFF}`;
      else if (target.ok) { verdict = `${RED}壊したのに通ってしまった（検査が何も見ていない）${OFF}`; bad++; }
      else {
        // 壊した1件だけが落ちること（巻き添えが出ていないか）
        const alsoFailed = res.filter((r) => !r.ok && r.id !== id).map((r) => r.id);
        verdict = alsoFailed.length
          ? `${GREEN}検出${OFF} ${DIM}(巻き添え: ${alsoFailed.join(', ')})${OFF}`
          : `${GREEN}検出${OFF}`;
      }
    } catch (e) {
      verdict = `${RED}壊す手順が失敗: ${String(e).slice(0, 120)}${OFF}`;
      bad++;
    } finally {
      rmSync(dir, { recursive: true, force: true });
    }
    console.log(`  ${id.padEnd(34)} ${verdict}`);
  }

  const all = (await run(ROOT)).map((r) => r.id);
  const uncovered = all.filter((id) => !covered.has(id));
  if (uncovered.length) {
    console.log(`\n${RED}自己診断していない検査: ${uncovered.join(', ')}${OFF}`);
    console.log('  （壊して確かめていない検査は、動いている保証がありません）');
    bad += uncovered.length;
  }

  console.log(bad === 0
    ? `\n${GREEN}自己診断 OK。${covered.size} 件すべて、壊すと確かに落ちます。${OFF}`
    : `\n${RED}自己診断 NG: ${bad} 件${OFF}`);
  return bad === 0 ? 0 : 1;
}

// ---------------------------------------------------------------------------
const isSelfTest = process.argv.includes('--self-test');
const code = isSelfTest ? await selfTest() : report(await run(ROOT));
process.exit(code === 0 ? 0 : 1);
