/**
 * 中核ロジックのテスト。
 *
 * GAS の関数も、シェルの関数も、ブラウザや Google の API 無しでは動かせない。
 * そこで「純粋な入出力だけの関数」をソースから取り出して確かめる。
 * こうしておくと、リポジトリの実体が変わればテストも一緒に変わる
 * （テスト用にコピーした別物を検査してしまう事故を避けられる）。
 */
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { join, dirname } from 'node:path';
import { fileURLToPath } from 'node:url';

const ROOT = join(dirname(fileURLToPath(import.meta.url)), '..');

/**
 * ソースから function 定義を1つ取り出して、実行できる形にする。
 * @param {object} scope GAS の API など、その関数が使う外の名前を差し込む
 */
function extractFn(file, name, scope = {}) {
  const src = readFileSync(join(ROOT, file), 'utf8');
  const start = src.indexOf(`function ${name}(`);
  assert.ok(start >= 0, `${file} に function ${name} が無い`);
  let depth = 0, i = src.indexOf('{', start), end = -1;
  for (; i < src.length; i++) {
    if (src[i] === '{') depth++;
    else if (src[i] === '}') { depth--; if (depth === 0) { end = i + 1; break; } }
  }
  assert.ok(end > 0, `${name} の本体を取り出せなかった`);
  const body = src.slice(start, end);
  const keys = Object.keys(scope);
  // eslint-disable-next-line no-new-func
  return new Function(...keys, `${body}; return ${name};`)(...keys.map((k) => scope[k]));
}

/**
 * GAS の Utilities.formatDate の代わり。JST 固定で必要な書式だけ返す。
 * 本物と同じ結果になることが大事なので、タイムゾーンを明示して組み立てる。
 */
const UtilitiesStub = {
  formatDate(date, tz, fmt) {
    assert.equal(tz, 'JST', 'JST 以外が渡されたら、この代役は正しくない');
    const p = new Intl.DateTimeFormat('en-CA', {
      timeZone: 'Asia/Tokyo', year: 'numeric', month: '2-digit', day: '2-digit',
      hour: '2-digit', minute: '2-digit', hour12: false,
    }).formatToParts(date).reduce((a, x) => (a[x.type] = x.value, a), {});
    if (fmt === 'yyyy-MM-dd') return `${p.year}-${p.month}-${p.day}`;
    if (fmt === 'HH:mm') return `${p.hour}:${p.minute}`;
    throw new Error('未対応の書式: ' + fmt);
  },
};

// ---------------------------------------------------------------------------
// シェル：GAS の URL だけを受け付ける
//   ここが緩いと、先生の端末で任意のサイトを iframe に埋め込めてしまう。
// ---------------------------------------------------------------------------
test('isValidGasUrl は GAS のウェブアプリ URL だけを通す', () => {
  const isValidGasUrl = extractFn('docs/app.js', 'isValidGasUrl');

  // 通すべきもの
  assert.equal(isValidGasUrl('https://script.google.com/macros/s/AKfycbx123/exec'), true);
  assert.equal(isValidGasUrl('https://script.google.com/macros/s/AKfycbx123/exec?v=1'), true);
  assert.equal(isValidGasUrl('https://script.google.com/a/example.ed.jp/macros/s/AK1/exec'), true);

  // 弾くべきもの
  assert.equal(isValidGasUrl(''), false, '空文字');
  assert.equal(isValidGasUrl('http://script.google.com/macros/s/AK1/exec'), false, 'http は通さない');
  assert.equal(isValidGasUrl('https://evil.example.com/exec'), false, '別ドメイン');
  assert.equal(isValidGasUrl('https://script.google.com/macros/s/AK1/dev'), false, '/dev は本番ではない');
  assert.equal(isValidGasUrl('javascript:alert(1)'), false, 'javascript: スキーム');
  assert.equal(isValidGasUrl('https://script.google.com.evil.com/macros/s/AK1/exec'), false,
    'ドメインの後ろに別ドメインを足した形');
  assert.equal(isValidGasUrl('https://notscript.google.com/macros/s/AK1/exec'), false,
    'ホスト名の前に文字を足した形');
});

// ---------------------------------------------------------------------------
// GAS：壊れた JSON を受け取っても落ちない
// ---------------------------------------------------------------------------
test('safeJsonParse は壊れた入力でも落ちない', () => {
  const safeJsonParse = extractFn('code.gs', 'safeJsonParse');
  assert.deepEqual(safeJsonParse('{"a":1}'), { a: 1 });
  assert.deepEqual(safeJsonParse(''), {});
  assert.deepEqual(safeJsonParse(null), {});
  assert.deepEqual(safeJsonParse(undefined), {});
  assert.deepEqual(safeJsonParse('こわれている'), {});
  assert.deepEqual(safeJsonParse('{"a":'), {});
});

// ---------------------------------------------------------------------------
// GAS：スプレッドシートが日付文字列を Date に変えてしまう問題への対処
//   ここが崩れると、今日以降の予定の絞り込み（s.date >= todayStr）が壊れ、
//   先生が入れた授業予定が児童の画面に出なくなる。
// ---------------------------------------------------------------------------
test('normalizeDateStr は Date でも文字列でも yyyy-MM-dd を返す', () => {
  const normalizeDateStr = extractFn('code.gs', 'normalizeDateStr', { Utilities: UtilitiesStub });
  assert.equal(normalizeDateStr('2026-08-05'), '2026-08-05');
  assert.equal(normalizeDateStr(''), '');
  assert.equal(normalizeDateStr(null), '');

  const out = normalizeDateStr(new Date('2026-08-05T09:30:00+09:00'));
  assert.match(out, /^\d{4}-\d{2}-\d{2}$/, 'Date は yyyy-MM-dd になる');
  assert.equal(out, '2026-08-05');
});

test('normalizeTimeStr は Date でも文字列でも HH:mm を返す', () => {
  const normalizeTimeStr = extractFn('code.gs', 'normalizeTimeStr', { Utilities: UtilitiesStub });
  assert.equal(normalizeTimeStr('09:30'), '09:30');
  assert.equal(normalizeTimeStr(''), '');
  assert.equal(normalizeTimeStr(null), '');

  const out = normalizeTimeStr(new Date('2026-08-05T09:05:00+09:00'));
  assert.match(out, /^\d{2}:\d{2}$/, 'Date は HH:mm になる');
  assert.equal(out, '09:05');
});

// ---------------------------------------------------------------------------
// 品質ゲートの部品：PNG の透明を数える
//   apple-touch-icon の透明を見落とすと iOS で四隅が黒くなる。
//   数える側が壊れていないことを、作った PNG で確かめる。
// ---------------------------------------------------------------------------
test('scanPngAlpha は透明のある PNG と無い PNG を見分ける', async () => {
  const { scanPngAlpha } = await import('../scripts/lib/png-alpha.mjs');
  const { deflateSync } = await import('node:zlib');

  // 2×2 の RGBA PNG を組み立てる
  const makePng = (alphas) => {
    const W = 2, H = 2;
    const raw = Buffer.alloc(H * (1 + W * 4));
    let p = 0;
    for (let y = 0; y < H; y++) {
      raw[p++] = 0; // filter: none
      for (let x = 0; x < W; x++) {
        raw[p++] = 26; raw[p++] = 115; raw[p++] = 232;
        raw[p++] = alphas[y * W + x];
      }
    }
    const chunk = (type, data) => {
      const len = Buffer.alloc(4); len.writeUInt32BE(data.length);
      const body = Buffer.concat([Buffer.from(type, 'ascii'), data]);
      const crcBuf = Buffer.alloc(4); crcBuf.writeUInt32BE(crc32(body) >>> 0);
      return Buffer.concat([len, body, crcBuf]);
    };
    const ihdr = Buffer.alloc(13);
    ihdr.writeUInt32BE(W, 0); ihdr.writeUInt32BE(H, 4);
    ihdr[8] = 8; ihdr[9] = 6; ihdr[10] = 0; ihdr[11] = 0; ihdr[12] = 0;
    return Buffer.concat([
      Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]),
      chunk('IHDR', ihdr),
      chunk('IDAT', deflateSync(raw)),
      chunk('IEND', Buffer.alloc(0)),
    ]);
  };

  const opaque = scanPngAlpha(makePng([255, 255, 255, 255]));
  assert.equal(opaque.transparentPixels, 0, '不透明な PNG は 0 画素');
  assert.equal(opaque.totalPixels, 4);

  const withHole = scanPngAlpha(makePng([255, 0, 255, 128]));
  assert.equal(withHole.transparentPixels, 2, '完全透明と半透明の2画素を数える');
});

/** PNG のチャンク CRC。テスト用の PNG を組み立てるためだけに使う。 */
function crc32(buf) {
  let c, crc = 0xffffffff;
  for (let n = 0; n < buf.length; n++) {
    c = (crc ^ buf[n]) & 0xff;
    for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
    crc = c ^ (crc >>> 8);
  }
  return crc ^ 0xffffffff;
}
