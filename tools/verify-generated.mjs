/**
 * 生成物が原本と食い違っていないか確かめる。
 *
 * vendor_*.html は生成物なので、原本（package.json / tools/build-vendor.mjs）を直したのに
 * `npm run build` を忘れて push すると、リポジトリの中身だけが古いまま残る。
 * ビルドも静的解析も通ってしまい、動かすまで気づけない。
 * CI でここを踏むと、その取りこぼしが PR の時点で止まる。
 */
import { readFileSync, existsSync } from 'node:fs';
import { execFileSync } from 'node:child_process';
import { join, dirname } from 'node:path';
import { fileURLToPath } from 'node:url';
import { createHash } from 'node:crypto';

const ROOT = join(dirname(fileURLToPath(import.meta.url)), '..');
const FILES = ['vendor_css.html', 'vendor_icons.html', 'vendor_js.html'];

const hash = (f) => createHash('sha256').update(readFileSync(f)).digest('hex').slice(0, 16);

const before = {};
for (const f of FILES) {
  const path = join(ROOT, f);
  if (!existsSync(path)) {
    console.error(`❌ ${f} がありません。\`npm run build\` を実行してください。`);
    process.exit(1);
  }
  before[f] = hash(path);
}

execFileSync(process.execPath, [join(ROOT, 'tools', 'build-vendor.mjs')], { stdio: 'pipe' });

let bad = 0;
for (const f of FILES) {
  const after = hash(join(ROOT, f));
  if (before[f] !== after) {
    console.error(`❌ ${f} が原本と食い違っています（${before[f]} → ${after}）。`);
    bad++;
  } else {
    console.log(`✅ ${f} は最新です（${after}）。`);
  }
}

if (bad) {
  console.error('\n原本を直したら `npm run build` を走らせてから push してください。');
  process.exit(1);
}
