/**
 * scripts/lib/local-checks.mjs のテスト。
 *
 * 検査そのものが働くかは check-project.mjs --self-test が48通りの変異で
 * 確かめている。ここで見るのは、その土台になる2つの関数である。
 * どちらも過去に実際の欠陥を出した場所なので、知見をテストの形で残す。
 */
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import path from 'node:path';
import { test } from 'node:test';

import { buildLocalChecks, fnSlice, stripComments } from '../scripts/lib/local-checks.mjs';

const ROOT = path.resolve(import.meta.dirname, '..');
const cfg = JSON.parse(readFileSync(path.join(ROOT, 'quality.config.json'), 'utf8')).local;

/* ── stripComments ────────────────────────────────── */

test('注意書きに反応しないよう、コメントを落とす', () => {
  assert.doesNotMatch(stripComments('/* localStorage.clear() は使わない */ const a = 1;', 'js'), /localStorage/);
  assert.doesNotMatch(stripComments('// localStorage.clear() は使わない\nconst b = 2;', 'js'), /localStorage/);
  assert.doesNotMatch(stripComments('<!-- localStorage.clear() は使わない -->', 'html'), /localStorage/);
  assert.match(stripComments('/* めも */ const a = 1;', 'js'), /const a = 1;/);
});

test('URL の // をコメントの始まりと読まない', () => {
  // ⚠️ これは実際に起きた欠陥である。`https://*.googleusercontent.com` の `/*` を
  //    ブロックコメントの開始と読んで、そこから次の `*/` までを丸ごと消していた。
  //    消された範囲に検査対象が入っていたため、CSP の検査が
  //    「壊したのに通る」状態になっていた（自己診断で見つかった）。
  const src = "const a = 'https://*.googleusercontent.com'; const b = 'https://x.example/'; const c = 1;";
  const out = stripComments(src, 'js');
  assert.match(out, /googleusercontent\.com/);
  assert.match(out, /https:\/\/x\.example\//);
  assert.match(out, /const c = 1;/);
});

test('文字列の中の // を消さない', () => {
  assert.match(stripComments("const u = 'https://example.com/x';", 'js'), /https:\/\/example\.com\/x/);
});

/* ── fnSlice ──────────────────────────────────────── */

test('関数の本文を、次のトップレベル関数の手前まで切り出す', () => {
  const src = 'function a() {\n  guard();\n}\nfunction b() {\n  other();\n}\n';
  assert.match(fnSlice(src, 'a'), /guard\(\)/);
  assert.doesNotMatch(fnSlice(src, 'a'), /other\(\)/);
});

test('無い関数を求められたら null を返す（「本文が空」と取りちがえない）', () => {
  // ⚠️ null と '' を混同すると、関数ごと消えているのに
  //    「本文に requireTeacher が無い」ではなく素通りになりかねない。
  assert.equal(fnSlice('function a() {}', 'zzz'), null);
});

test('最後の関数は末尾まで切り出す', () => {
  assert.match(fnSlice('function a() {}\nfunction last() {\n  guard();\n}', 'last'), /guard\(\)/);
});

/* ── 検査の一覧 ───────────────────────────────────── */

test('いまの木では、ローカルの検査がすべて通る', () => {
  const bad = buildLocalChecks(cfg).map((c) => ({ id: c.id, ...c.run({ root: ROOT }) })).filter((r) => !r.ok);
  assert.deepEqual(bad.map((r) => `${r.id}: ${r.detail}`), []);
});

test('ローカルに残したのは、正本に行き先が無い10件だけ', () => {
  // 増やすときは「正本に本当に行き先が無いか」を確かめてからにする。
  // 正本にあるものをここに置くと、正本を直しても届かない場所が増える。
  assert.equal(buildLocalChecks(cfg).length, 10);
});

test('GAS 本体を見る検査が、設定の取りちがえで空振りしない', () => {
  // ⚠️ cfg.sourceFiles / viewportFiles / styleFiles が空や誤りだと、
  //    検査は「見るものが無い」ので静かに通る。実在を確かめておく。
  for (const key of ['sourceFiles', 'viewportFiles', 'styleFiles', 'generatedFiles']) {
    assert.ok(Array.isArray(cfg[key]) && cfg[key].length > 0, `${key} が空`);
  }
  assert.ok(cfg.sourceFiles.includes('code.gs'), 'GAS 本体 code.gs を見ていない');
  assert.ok(cfg.viewportFiles.includes('code.gs'), 'GAS の viewport は code.gs で決まる');
});
