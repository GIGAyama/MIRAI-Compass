/**
 * vendor_*.html を作る（GIGA Standard v5 §6）。
 *
 * なぜ必要か。
 *   学校のネットワークは cdn.jsdelivr.net を塞いでいることがある。
 *   塞がれた状態でこのアプリを開くと、白い画面ではなく
 *   「Bootstrap が当たっていない素の HTML が半分だけ動く」という壊れ方をする。
 *   ローディング画面が消えず、d-none が効かないので児童画面と先生用ボタンが同時に出て、
 *   Swal / Chart / Sortable がすべて undefined になる。
 *   児童からは「壊れている」としか見えず、原因はアプリの外にあるので先生が調べても分からない。
 *
 * そこで、実行コードは1バイトも外から取らない。npm で版を固定し、
 * ここで .html に包み込んで GAS の include() から読む。
 *
 * ⚠️ 生成物（vendor_css.html / vendor_icons.html / vendor_js.html）は手で編集しない。
 *    ライブラリを差し替えるときは package.json を直して `npm run build` を走らせる。
 */
import { readFileSync, writeFileSync, existsSync } from 'node:fs';
import { join, dirname } from 'node:path';
import { fileURLToPath } from 'node:url';

const ROOT = join(dirname(fileURLToPath(import.meta.url)), '..');
const NM = join(ROOT, 'node_modules');

// package.json の exports が umd/dist を公開していない場合があり、
// require.resolve は ERR_PACKAGE_PATH_NOT_EXPORTED で落ちる（v5 §6）。
// パスで直に指定する。
const p = (...s) => join(NM, ...s);
const read = (f) => {
  if (!existsSync(f)) {
    console.error(`\n[build] ${f} がありません。先に \`npm ci\`（または npm install）を実行してください。\n`);
    process.exit(1);
  }
  return readFileSync(f, 'utf8');
};
const kb = (s) => (Buffer.byteLength(s, 'utf8') / 1024).toFixed(1) + ' KB';

const BANNER = (what) => `<!--
  ⚠️ このファイルは tools/build-vendor.mjs が生成しています。手で編集しないでください。
     直す場所は package.json（版）と tools/build-vendor.mjs（組み立て方）です。
     直したら必ず \`npm run build\` を走らせてから push してください。
     中身: ${what}
-->
`;

// ---------------------------------------------------------------------------
// 1. CSS（Bootstrap 本体）
// ---------------------------------------------------------------------------
{
  const css = read(p('bootstrap', 'dist', 'css', 'bootstrap.min.css'));
  // sourceMappingURL が残っていると、開発者ツールが取りに行って 404 を出す
  const body = css.replace(/\/\*#\s*sourceMappingURL=.*?\*\//g, '').trim();
  const out = BANNER('bootstrap.min.css（自己ホスト）') + '<style>\n' + body + '\n</style>\n';
  writeFileSync(join(ROOT, 'vendor_css.html'), out);
  console.log('vendor_css.html  ', kb(out));
}

// ---------------------------------------------------------------------------
// 2. アイコン（使っている分だけ SVG マスクにする）
//
//    bootstrap-icons をまるごと持つと CSS 98KB + woff2 131KB = 229KB になる。
//    実際に使っているのは 53 種類しかないので、その SVG だけを取り出して
//    mask-image にする。マークアップ（<i class="bi bi-compass-fill">）は変えない。
//    background-color: currentColor なので text-danger などの色指定もそのまま効く。
// ---------------------------------------------------------------------------
{
  const srcFiles = ['index.html', 'js_core.html', 'js_student.html', 'js_teacher.html'];
  const used = new Set();
  for (const f of srcFiles) {
    const t = read(join(ROOT, f));
    for (const m of t.matchAll(/\bbi-([a-z0-9-]+)/g)) used.add(m[1]);
  }

  // bootstrap-icons 1.11.1 に存在しない名前の読み替え。
  // pencil-ruler は存在せず、いまは何も描かれていない（先生の「設計」ボタンのアイコンが空）。
  const ALIAS = { 'pencil-ruler': 'rulers' };

  const rules = [];
  const missing = [];
  for (const name of [...used].sort()) {
    const file = p('bootstrap-icons', 'icons', (ALIAS[name] || name) + '.svg');
    if (!existsSync(file)) { missing.push(name); continue; }
    let svg = readFileSync(file, 'utf8')
      .replace(/<\?xml[^>]*\?>/g, '')
      .replace(/<!--[\s\S]*?-->/g, '')
      .replace(/\s+/g, ' ')
      .trim();
    // currentColor で塗るので、SVG 側の色指定は落として fill="currentColor" に統一する
    svg = svg.replace(/\sfill="[^"]*"/g, '').replace('<svg ', '<svg fill="currentColor" ');
    const uri = svg.replace(/"/g, "'").replace(/[<>#%{}|\\^~[\]`]/g, (c) => '%' + c.charCodeAt(0).toString(16).toUpperCase());
    rules.push(`.bi-${name}{--bi-i:url("data:image/svg+xml,${uri}")}`);
  }
  if (missing.length) console.warn('[build] SVG が見つからないアイコン:', missing.join(', '));

  const base = [
    '/* bootstrap-icons のうち、このアプリが実際に使っている分だけ。',
    '   マスク方式なので currentColor / font-size(1em) がそのまま効く。 */',
    '.bi{display:inline-block;width:1em;height:1em;vertical-align:-.125em;flex-shrink:0;',
    'background-color:currentColor;',
    '-webkit-mask:var(--bi-i) center/contain no-repeat;mask:var(--bi-i) center/contain no-repeat}',
    '/* 高コントラストモードでは mask が消えるため、絵が無くても意味が失われないよう',
    '   アイコンだけのボタンには aria-label を付けてある（css.html 側でも補強する）。 */',
    '@media (forced-colors: active){.bi{background-color:CanvasText}}',
  ].join('\n');

  const out = BANNER(`bootstrap-icons から ${rules.length} 個（使用分のみ）`) +
    '<style>\n' + base + '\n' + rules.join('\n') + '\n</style>\n';
  writeFileSync(join(ROOT, 'vendor_icons.html'), out);
  console.log('vendor_icons.html', kb(out), `(${rules.length} icons)`);
}

// ---------------------------------------------------------------------------
// 3. JavaScript（Bootstrap / SweetAlert2 / Chart.js / Sortable）
// ---------------------------------------------------------------------------
{
  const parts = [
    ['bootstrap.bundle.min.js', p('bootstrap', 'dist', 'js', 'bootstrap.bundle.min.js')],
    ['sweetalert2.all.min.js', p('sweetalert2', 'dist', 'sweetalert2.all.min.js')],
    ['chart.umd.js', p('chart.js', 'dist', 'chart.umd.js')],
    ['Sortable.min.js', p('sortablejs', 'Sortable.min.js')],
    // ワークシートの手書き・添削キャンバス（統合で取り込み。学校で CDN が塞がれても動くよう自己ホスト）
    ['fabric.min.js', p('fabric', 'dist', 'fabric.min.js')],
  ];
  let out = BANNER(parts.map(([n]) => n).join(' + '));
  for (const [name, file] of parts) {
    let js = read(file).replace(/\/\/#\s*sourceMappingURL=.*$/gm, '');
    // </script> が文字列中に現れると、そこで <script> が閉じてしまう
    js = js.replace(/<\/script>/gi, '<\\/script>');
    out += `<!-- ${name} -->\n<script>\n${js.trim()}\n</script>\n`;
  }
  writeFileSync(join(ROOT, 'vendor_js.html'), out);
  console.log('vendor_js.html   ', kb(out));
}

console.log('\n[build] 完了。生成物は手で編集しないこと。');
