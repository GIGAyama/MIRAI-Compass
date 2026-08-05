/**
 * apple-touch-icon.png を作り直す（GIGA Standard v5 §3-2）。
 *
 * なぜ必要か。
 *   iOS は apple-touch-icon の透明部分を黒で塗りつぶす。
 *   いまの apple-touch-icon.png は角丸の外側が透明で、
 *   実測すると四隅ボックスの 62.09% が透明だった。
 *   そのままホーム画面に追加すると、アイコンの四隅だけが黒く出る。
 *
 * 直し方は「透明を含まない専用画像を用意する」。
 * maskable アイコン（下地が端まで伸びている）を土台にし、
 * その上に元のアイコンの絵を重ねる。こうすると角丸の内側の見た目は変わらず、
 * 外側だけがブランド色で埋まる。
 *
 * 使い方: node tools/build-icons.mjs
 * 生成物: docs/icons/apple-touch-icon.png（180×180・透明なし）
 */
import { readFileSync, writeFileSync } from 'node:fs';
import { join, dirname } from 'node:path';
import { fileURLToPath } from 'node:url';
import { chromium } from 'playwright';

const ROOT = join(dirname(fileURLToPath(import.meta.url)), '..');
const ICONS = join(ROOT, 'docs', 'icons');
const SIZE = 180;

const b64 = (f) => readFileSync(join(ICONS, f)).toString('base64');

const browser = await chromium.launch();
const page = await browser.newPage();
await page.goto('about:blank');

const dataUrl = await page.evaluate(async ({ baseB64, artB64, SIZE }) => {
  const load = (b) => new Promise((res, rej) => {
    const i = new Image(); i.onload = () => res(i); i.onerror = rej;
    i.src = 'data:image/png;base64,' + b;
  });
  const base = await load(baseB64);   // maskable（下地が端まである）
  const art = await load(artB64);     // 通常アイコン（角丸の外が透明）

  const cv = document.createElement('canvas');
  cv.width = cv.height = SIZE;
  const ctx = cv.getContext('2d');

  // ① 下地。maskable を引き伸ばして全面に敷く。
  //    単色で塗るとグラデーションと合わず、角丸四角の輪郭が薄い影として残る。
  ctx.drawImage(base, 0, 0, SIZE, SIZE);

  // ② 絵。元のアイコンをそのまま重ねる（角丸の内側は元の見た目のまま）。
  ctx.drawImage(art, 0, 0, SIZE, SIZE);

  // ③ 透明が1画素も残っていないことを確かめる。
  //    残っていたら、そこが iOS で黒くなる。
  const d = ctx.getImageData(0, 0, SIZE, SIZE).data;
  let transparent = 0;
  for (let i = 3; i < d.length; i += 4) if (d[i] < 255) transparent++;

  return { url: cv.toDataURL('image/png'), transparent, total: SIZE * SIZE };
}, { baseB64: b64('icon-maskable-512.png'), artB64: b64('icon-512.png'), SIZE });

await browser.close();

if (dataUrl.transparent > 0) {
  console.error(`❌ 透明が ${dataUrl.transparent} 画素残っています。iOS で黒く出ます。`);
  process.exit(1);
}

const buf = Buffer.from(dataUrl.url.split(',')[1], 'base64');
writeFileSync(join(ICONS, 'apple-touch-icon.png'), buf);
console.log(`✅ apple-touch-icon.png (${SIZE}×${SIZE}) 透明 0 画素 / ${(buf.length / 1024).toFixed(1)} KB`);
