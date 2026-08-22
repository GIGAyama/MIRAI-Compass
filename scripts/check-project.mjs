#!/usr/bin/env node
/**
 * 品質ゲート（GIGA Standard v5）。
 *
 * 構成
 *   scripts/lib/project-quality.mjs … フリート共通の正本（このリポジトリにはまだ無い）
 *   scripts/lib/giga-v5-checks.mjs  … Part I の検査（ここに置く）
 *   scripts/check-project.mjs       … 両者を合成する（このファイル）
 *
 * 共通の正本を丸ごと差し替えで受けられるよう、Part I の検査は分けてある。
 * 正本が置かれたら、下の loadShared() がそれを読み込んで検査に足す。
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
import { buildChecks } from './lib/giga-v5-checks.mjs';

const ROOT = join(dirname(fileURLToPath(import.meta.url)), '..');
const cfg = JSON.parse(readFileSync(join(ROOT, 'quality.config.json'), 'utf8'));

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
  const checks = buildChecks(cfg);
  const results = [];
  for (const c of checks) {
    let r;
    try { r = c.run({ root, cfg }); }
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
  const BREAKAGES = [
    ['A1_LICENSE', (d) => rmSync(join(d, 'LICENSE'))],
    ['A2_GITIGNORE', (d) => writeFileSync(join(d, '.gitignore'), 'dist/\n')],
    ['A3_DEPENDABOT', (d) => rmSync(join(d, '.github/dependabot.yml'))],
    ['A4_DOCS', (d) => rmSync(join(d, 'MANUAL.md'))],
    ['A5_CI_ON_PR', (d) => writeFileSync(join(d, '.github/workflows/ci.yml'), 'on:\n  push:\n    branches: [main]\njobs: {}\n')],
    ['B6_NO_CDN_EXEC', (d) => patch(d, 'index.html', (s) => s.replace('</head>', '<script src="https://cdn.jsdelivr.net/npm/chart.js"></script></head>'))],
    ['B6b_NO_BROWSER_BABEL', (d) => patch(d, 'index.html', (s) => s.replace('</head>', '<script src="https://unpkg.com/@babel/standalone/babel.min.js"></script></head>'))],
    ['B2_NO_SECRETS', (d) => patch(d, 'js_core.html', (s) => s.replace('<script>', '<script>\nconst apiKey = "AIzaSyD0123456789abcdefghijklmnopqrstuv";'))],
    ['B4_POSTMESSAGE', (d) => patch(d, 'js_core.html', (s) => s.replace('<script>', '<script>\nwindow.parent.postMessage({a:1}, "*");'))],
    ['B1_CSP_SHELL', (d) => patch(d, 'docs/index.html', (s) => s.replace('script-src \'self\';', 'script-src \'self\' \'unsafe-inline\';'))],
    ['B1b_NO_FRAME_ANCESTORS_META', (d) => patch(d, 'docs/index.html', (s) => s.replace('object-src \'none\';', 'frame-ancestors \'self\'; object-src \'none\';'))],
    ['B1c_NO_INLINE_HANDLER_IN_SHELL', (d) => patch(d, 'docs/index.html', (s) => s.replace('id="btn-save" type="button"', 'id="btn-save" type="button" onclick="saveUrl()"'))],
    ['C5_NO_LS_CLEAR', (d) => patch(d, 'docs/app.js', (s) => s.replace('localStorage.removeItem(LS_KEY)', 'localStorage.clear()'))],
    ['D1_VIEWPORT_FIT', (d) => patch(d, 'code.gs', (s) => s.replace(", viewport-fit=cover'", "'"))],
    ['D14_NO_USER_SCALABLE', (d) => patch(d, 'index.html', (s) => s.replace('viewport-fit=cover">', 'viewport-fit=cover, user-scalable=no">'))],
    ['D2_DVH', (d) => patch(d, 'css.html', (s) => s.replace('height: 100dvh;', 'height: 100vh;'))],
    ['D3_SAFE_AREA', (d) => { for (const f of ['css.html', 'docs/index.html', 'docs/offline.html']) patch(d, f, (s) => s.replace(/safe-area-inset/g, 'SAFEAREAINSET')); }],
    ['D4_FLUID_TYPE', (d) => { patch(d, 'css.html', (s) => s.replace(/clamp\(/g, 'min(')); patch(d, 'docs/index.html', (s) => s.replace(/clamp\(/g, 'min(')); patch(d, 'docs/offline.html', (s) => s.replace(/clamp\(/g, 'min(')); }],
    ['D5_CANVAS_DPR', (d) => patch(d, 'js_student.html', (s) => s.replace(/Math\.min\(window\.devicePixelRatio \|\| 1, 2\)/, '1'))],
    ['D10_REDUCED_MOTION', (d) => patch(d, 'css.html', (s) => s.replace('animation-duration: .01ms !important;', 'animation-duration: 0ms !important;'))],
    ['D11_FORCED_COLORS', (d) => { for (const f of ['css.html', 'docs/index.html', 'docs/offline.html']) patch(d, f, (s) => s.replace(/forced-colors: active/g, 'forced-colours: active')); }],
    ['D13_PRINT_CSS', (d) => patch(d, 'css.html', (s) => s.replace(/@media print/g, '@media screenonly'))],
    ['F4_RT_COLOR', (d) => patch(d, 'css.html', (s) => s.replace(/color: inherit; \}/, 'color: #6c757d; }'))],
    ['E1_MANIFEST_PATHS', (d) => patch(d, 'docs/manifest.webmanifest', (s) => s.replace(/"id": "[^"]*"/, '"id": "/MIRAI-Compass/"'))],
    ['E2_APPLE_ICON_OPAQUE', (d) => cpSync(join(d, 'docs/icons/icon-192.png'), join(d, 'docs/icons/apple-touch-icon.png'))],
    ['E3_INSTALL_HOOK', (d) => patch(d, 'docs/index.html', (s) => s.replace('<script src="install-hook.js"></script>', ''))],
    ['E5_SW_CACHE_SCOPE', (d) => patch(d, 'docs/sw.js', (s) => s.replace(/\.filter\(\(k\) => k\.startsWith\(CACHE_PREFIX\) && k !== CACHE_VERSION\)/, ''))],
    ['E6_SW_NO_LOCALSTORAGE', (d) => patch(d, 'docs/sw.js', (s) => s.replace('const CACHE_PREFIX', 'localStorage.setItem("x","1");\nconst CACHE_PREFIX'))],
    ['E7_SW_NO_SKIP_WAITING_IN_INSTALL', (d) => patch(d, 'docs/sw.js', (s) => s.replace('const cache = await caches.open(CACHE_VERSION);', 'self.skipWaiting();\n    const cache = await caches.open(CACHE_VERSION);'))],
    ['E7b_UPDATE_PROMPT', (d) => patch(d, 'docs/app.js', (s) => s.replace('if (!userAskedUpdate || reloading) return;', ''))],
    ['E9_SW_REGISTER_READYSTATE', (d) => patch(d, 'docs/app.js', (s) => s.replace("if (document.readyState === 'complete') startServiceWorker();\n  else ", ''))],
    ['E10_OFFLINE_HTML', (d) => patch(d, 'docs/offline.html', (s) => s.replace('</body>', '<script>console.log(1)</script></body>'))],
    ['E11_APP_VERSION', (d) => patch(d, 'docs/sw.js', (s) => s.replace(/const APP_VERSION = '[^']*';/, ''))],
    ['E12_MASKABLE_SAFEZONE', (d) => cpSync(join(d, 'docs/icons/icon-192.png'), join(d, 'docs/icons/icon-maskable-192.png'))],
    ['F6_FILE_SIZE', (d) => patch(d, 'js_core.html', (s) => s + '\n' + '// 行数超過の確認用\n'.repeat(5200))],
    ['G1_GENERATED_MARKED', (d) => patch(d, 'vendor_css.html', (s) => s.replace('手で編集しない', ''))],
    ['SEC_PRIVILEGED_FN_GUARDED', (d) => patch(d, 'code.gs', (s) => s.replace(
      'MiraiAuth.requireTeacher();\n    var props = PropertiesService.getScriptProperties();',
      'var props = PropertiesService.getScriptProperties();'))],
    ['SEC_NO_PASSWORD_AUTH', (d) => patch(d, 'code.gs', (s) => s.replace(
      'var MiraiAuth = (function () {',
      'var TEACHER_PASS = ScriptProps.get("pw"); function verifyPassword(p) { return p === TEACHER_PASS; }\nvar MiraiAuth = (function () {'))],
  ];

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
