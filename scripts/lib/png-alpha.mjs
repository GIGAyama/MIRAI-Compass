/**
 * PNG に透明な画素が含まれているかを、実際に中身を展開して数える。
 *
 * なぜ自前で書くか。
 *   apple-touch-icon に透明があると iOS がそこを黒で塗りつぶし、
 *   ホーム画面でアイコンの四隅だけが黒く出る。
 *   これは「purpose を見る」「ファイル名を見る」では絶対に分からない。
 *   画素を数えるしかない。
 *   CI でブラウザを立ち上げずに確かめたいので、依存なしで書いてある。
 */
import { inflateSync } from 'node:zlib';

const SIG = Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]);

function paeth(a, b, c) {
  const p = a + b - c;
  const pa = Math.abs(p - a), pb = Math.abs(p - b), pc = Math.abs(p - c);
  return (pa <= pb && pa <= pc) ? a : (pb <= pc ? b : c);
}

/**
 * @returns {{width:number,height:number,colorType:number,hasAlphaChannel:boolean,
 *            transparentPixels:number,totalPixels:number,cornerTransparent:number,cornerTotal:number}}
 */
export function scanPngAlpha(buf) {
  if (!buf.subarray(0, 8).equals(SIG)) throw new Error('PNG ではありません');

  let off = 8;
  let ihdr = null;
  let hasTRNS = false;
  const idat = [];

  while (off < buf.length) {
    const len = buf.readUInt32BE(off);
    const type = buf.toString('ascii', off + 4, off + 8);
    const data = buf.subarray(off + 8, off + 8 + len);
    if (type === 'IHDR') {
      ihdr = {
        width: data.readUInt32BE(0),
        height: data.readUInt32BE(4),
        bitDepth: data[8],
        colorType: data[9],
        interlace: data[12],
      };
    } else if (type === 'tRNS') {
      hasTRNS = true;
    } else if (type === 'IDAT') {
      idat.push(data);
    } else if (type === 'IEND') {
      break;
    }
    off += 12 + len; // len + type(4) + data + crc(4)
  }
  if (!ihdr) throw new Error('IHDR がありません');

  const { width, height, bitDepth, colorType, interlace } = ihdr;
  const hasAlphaChannel = colorType === 4 || colorType === 6;
  const base = {
    width, height, colorType, hasAlphaChannel,
    transparentPixels: 0, totalPixels: width * height,
    cornerTransparent: 0, cornerTotal: 0,
  };

  // アルファチャンネルも tRNS も無ければ、透明は原理的に存在しない
  if (!hasAlphaChannel && !hasTRNS) return base;

  // パレット + tRNS や、インタレース、16bit は自前で解くと長くなる。
  // 「確かめられなかった」を「透明なし」と取り違えないよう、その旨を返す。
  if (colorType === 3 || interlace !== 0 || bitDepth !== 8) {
    return { ...base, unsupported: true };
  }

  const channels = colorType === 6 ? 4 : (colorType === 4 ? 2 : (colorType === 2 ? 3 : 1));
  const bpp = channels; // bitDepth 8 固定
  const stride = width * bpp;
  const raw = inflateSync(Buffer.concat(idat));

  const prev = Buffer.alloc(stride);
  const cur = Buffer.alloc(stride);
  const cornerBox = Math.max(1, Math.round(Math.min(width, height) * 0.12));
  let p = 0;

  for (let y = 0; y < height; y++) {
    const filter = raw[p++];
    raw.copy(cur, 0, p, p + stride);
    p += stride;

    for (let i = 0; i < stride; i++) {
      const a = i >= bpp ? cur[i - bpp] : 0;
      const b = prev[i];
      const c = i >= bpp ? prev[i - bpp] : 0;
      switch (filter) {
        case 0: break;
        case 1: cur[i] = (cur[i] + a) & 0xff; break;
        case 2: cur[i] = (cur[i] + b) & 0xff; break;
        case 3: cur[i] = (cur[i] + ((a + b) >> 1)) & 0xff; break;
        case 4: cur[i] = (cur[i] + paeth(a, b, c)) & 0xff; break;
        default: throw new Error('未知のフィルタ ' + filter);
      }
    }

    for (let x = 0; x < width; x++) {
      const alpha = cur[x * bpp + (bpp - 1)];
      const inCorner = (x < cornerBox || x >= width - cornerBox) &&
                       (y < cornerBox || y >= height - cornerBox);
      if (inCorner) base.cornerTotal++;
      if (alpha < 255) {
        base.transparentPixels++;
        if (inCorner) base.cornerTransparent++;
      }
    }
    cur.copy(prev);
  }
  return base;
}
