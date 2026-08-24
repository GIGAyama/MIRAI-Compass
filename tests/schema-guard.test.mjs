/**
 * データベース（スプレッドシート）の作りを見張る仕組みのテスト。
 *
 * みらいコンパスは、スプレッドシートのコピーを先生ごとに配る形になった。
 * 先生の手元のファイルは、誤字を直すついでに列を消したり足したりできる。
 * ワークシート（WS_COL_*）と答案（RS_COL_*）は **列の番号** で読み書きして
 * いるので、1列ずれるだけで児童の手書きと先生の赤ペンが入れかわる。
 * しかも、その事故は画面に何も出ないまま起きる。
 *
 * そこで code.gs の MiraiDb を、Google の API 抜きで動かして確かめる。
 * ソースからそのまま取り出しているので、code.gs を直せばここも一緒に変わる
 * （テスト用にコピーした別物を検査してしまう事故を避けられる）。
 */
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { join, dirname } from 'node:path';
import { fileURLToPath } from 'node:url';

const ROOT = join(dirname(fileURLToPath(import.meta.url)), '..');
const SRC = readFileSync(join(ROOT, 'code.gs'), 'utf8');

/** code.gs から `var 名前 = (function () { ... })();` を1つ取り出す。 */
function extractModule(name, scope) {
  const start = SRC.indexOf(`var ${name} = (function ()`);
  assert.ok(start >= 0, `code.gs に var ${name} = (function () が無い`);
  let depth = 0, end = -1;
  for (let i = SRC.indexOf('{', start); i < SRC.length; i++) {
    if (SRC[i] === '{') depth++;
    else if (SRC[i] === '}') { depth--; if (depth === 0) { end = SRC.indexOf(';', i) + 1; break; } }
  }
  assert.ok(end > 0, `${name} の本体を取り出せなかった`);
  const keys = Object.keys(scope);
  // eslint-disable-next-line no-new-func
  return new Function(...keys, `${SRC.slice(start, end)}; return ${name};`)(...keys.map((k) => scope[k]));
}

/** code.gs の DB_SCHEMA をそのまま読む（設計図の正本はここ1か所）。 */
function readSchema() {
  const start = SRC.indexOf('const DB_SCHEMA = {');
  assert.ok(start >= 0, 'code.gs に DB_SCHEMA が無い');
  let depth = 0, end = -1;
  for (let i = SRC.indexOf('{', start); i < SRC.length; i++) {
    if (SRC[i] === '{') depth++;
    else if (SRC[i] === '}') { depth--; if (depth === 0) { end = i + 1; break; } }
  }
  // eslint-disable-next-line no-new-func
  return new Function(`return ${SRC.slice(SRC.indexOf('{', start), end)};`)();
}

const DB_SCHEMA = readSchema();

// --- Google の API の代役 --------------------------------------------------
// 見出し行と列数だけを持つ、ごく小さなスプレッドシート。
class FakeSheet {
  constructor(name, header = []) {
    this.name = name;
    this.rows = header.length ? [header.slice()] : [];
    this.maxColumns = Math.max(header.length, 1);
    this.frozen = 0;
  }
  getLastRow() { return this.rows.length; }
  getLastColumn() { return this.rows.length ? this.rows[0].length : 0; }
  getMaxColumns() { return this.maxColumns; }
  setFrozenRows(n) { this.frozen = n; return this; }
  insertColumnsAfter(after, count) { this.maxColumns = after + count; return this; }
  appendRow(values) {
    this.rows.push(values.slice());
    this.maxColumns = Math.max(this.maxColumns, values.length);
    return this;
  }
  getRange(row, col, numRows = 1, numCols = 1) {
    const sheet = this;
    return {
      getValues() {
        const out = [];
        for (let r = 0; r < numRows; r++) {
          const line = sheet.rows[row - 1 + r] || [];
          const cells = [];
          for (let c = 0; c < numCols; c++) {
            const v = line[col - 1 + c];
            cells.push(v === undefined ? '' : v);
          }
          out.push(cells);
        }
        return out;
      },
      setValues(values) {
        values.forEach((line, r) => {
          const target = sheet.rows[row - 1 + r] || (sheet.rows[row - 1 + r] = []);
          line.forEach((v, c) => { target[col - 1 + c] = v; });
        });
        sheet.maxColumns = Math.max(sheet.maxColumns, col - 1 + numCols);
        return this;
      },
      setValue(v) { return this.setValues([[v]]); },
      setBackground() { return this; },
      setFontWeight() { return this; },
    };
  }
}

class FakeSpreadsheet {
  constructor(sheets = []) { this.sheets = sheets; }
  getId() { return 'fake-db'; }
  getSheetByName(name) { return this.sheets.find((s) => s.name === name) || null; }
  getSheets() { return this.sheets.slice(); }
  insertSheet(name) { const s = new FakeSheet(name); this.sheets.push(s); return s; }
}

/** 設計どおりのシートを全部そろえたスプレッドシートを作る。 */
function healthy() {
  return new FakeSpreadsheet(
    Object.keys(DB_SCHEMA).map((k) => new FakeSheet(DB_SCHEMA[k].name, DB_SCHEMA[k].headers)));
}

/** キャッシュを持たない（＝毎回きちんと見る）代役。 */
const noCache = { get: () => null, put: () => {}, remove: () => {} };

function loadMiraiDb() {
  return extractModule('MiraiDb', {
    DB_SCHEMA,
    CacheService: { getScriptCache: () => noCache },
    PropertiesService: { getScriptProperties: () => ({ getProperty: () => null, setProperty: () => {} }) },
    SpreadsheetApp: { getActiveSpreadsheet: () => null, openById: () => { throw new Error('使わない'); } },
  });
}

const MiraiDb = loadMiraiDb();

// ---------------------------------------------------------------------------

test('設計どおりのスプレッドシートは「異常なし」と言う', () => {
  assert.deepEqual(MiraiDb.inspect(healthy()), []);
});

test('シートが1枚無いと、名指しで報告し、作って直す', () => {
  const ss = healthy();
  ss.sheets = ss.sheets.filter((s) => s.name !== 'Responses');

  const found = MiraiDb.inspect(ss);
  assert.equal(found.length, 1);
  assert.equal(found[0].sheet, 'Responses');
  assert.equal(found[0].kind, 'シートが無い');
  assert.equal(found[0].fixable, true);

  const report = MiraiDb.repair(ss);
  assert.match(report.done.join('\n'), /Responses/);
  assert.deepEqual(report.left, [], '直したあとは異常なし');
  assert.deepEqual(ss.getSheetByName('Responses').rows[0], DB_SCHEMA.Responses.headers);
});

test('版が上がって列が増えたときは、末尾に足して直す', () => {
  // 旧版のファイル（StudentRoster に studentId がまだ無い）を作る
  const old = DB_SCHEMA.StudentRoster.headers.slice(0, -1);
  const ss = healthy();
  ss.sheets = ss.sheets.map((s) => (s.name === 'StudentRoster' ? new FakeSheet(s.name, old) : s));

  const found = MiraiDb.inspect(ss);
  assert.equal(found.length, 1);
  assert.equal(found[0].kind, '列が足りない');
  assert.equal(found[0].fixable, true);

  MiraiDb.repair(ss);
  assert.deepEqual(
    ss.getSheetByName('StudentRoster').rows[0],
    DB_SCHEMA.StudentRoster.headers,
    '足りなかった studentId が末尾に付く');
  assert.deepEqual(MiraiDb.inspect(ss), []);
});

test('列の並びが入れかわっているときは、報告するだけで直さない', () => {
  // ★ ここがこのテストの本体。見出しだけを設計どおりに書き直すと、
  //   まちがった列に正しいラベルが付き、事故が見えなくなる。
  const shuffled = DB_SCHEMA.Responses.headers.slice();
  const tmp = shuffled[10]; shuffled[10] = shuffled[11]; shuffled[11] = tmp; // 赤ペンと手書きを入れかえる
  const ss = healthy();
  ss.sheets = ss.sheets.map((s) => (s.name === 'Responses' ? new FakeSheet(s.name, shuffled) : s));

  const found = MiraiDb.inspect(ss);
  assert.equal(found.length, 1);
  assert.equal(found[0].kind, '列の並びがちがう');
  assert.equal(found[0].fixable, false, '機械では直さない');
  assert.match(found[0].detail, /11列目/);

  const report = MiraiDb.repair(ss);
  assert.deepEqual(report.done, [], '何もしない');
  assert.deepEqual(
    ss.getSheetByName('Responses').rows[0], shuffled,
    '見出しを書き換えて、ずれを見えなくしてはいけない');
  assert.equal(report.left.length, 1, '報告は残る');
});

test('見出しが別物のときも、書き換えずに場所を知らせる', () => {
  const renamed = DB_SCHEMA.LearningLogs.headers.slice();
  renamed[3] = '課題ID';
  const ss = healthy();
  ss.sheets = ss.sheets.map((s) => (s.name === 'LearningLogs' ? new FakeSheet(s.name, renamed) : s));

  const found = MiraiDb.inspect(ss);
  assert.equal(found.length, 1);
  assert.equal(found[0].kind, '見出しがちがう');
  assert.equal(found[0].fixable, false);
  assert.match(found[0].detail, /4列目/);

  MiraiDb.repair(ss);
  assert.deepEqual(ss.getSheetByName('LearningLogs').rows[0], renamed, '書き換えない');
});

test('見出しのラベルだけが消えたときは、ほかが全部合っていれば書き戻す', () => {
  const holed = DB_SCHEMA.MyTasks.headers.slice();
  holed[2] = '';
  const ss = healthy();
  ss.sheets = ss.sheets.map((s) => (s.name === 'MyTasks' ? new FakeSheet(s.name, holed) : s));

  const found = MiraiDb.inspect(ss);
  assert.equal(found.length, 1);
  assert.equal(found[0].kind, '見出しが空');
  assert.equal(found[0].fixable, true);

  MiraiDb.repair(ss);
  assert.deepEqual(ss.getSheetByName('MyTasks').rows[0], DB_SCHEMA.MyTasks.headers);
});

test('末尾に足しても、手前の並びがずれていれば足さない', () => {
  // 手前がずれているのに末尾へ足すと、ずれたまま列が増えて傷が深くなる。
  const broken = DB_SCHEMA.Portfolios.headers.slice(0, -2);
  broken[1] = '単元';                       // 別名になっている
  const ss = healthy();
  ss.sheets = ss.sheets.map((s) => (s.name === 'Portfolios' ? new FakeSheet(s.name, broken) : s));

  const report = MiraiDb.repair(ss);
  assert.deepEqual(report.done, [], '足さない');
  assert.deepEqual(ss.getSheetByName('Portfolios').rows[0], broken);
  assert.ok(report.left.some((f) => f.kind === '見出しがちがう'), 'ずれは報告する');
});

test('余分な列は、消さずに「読みません」と知らせるだけ', () => {
  const ss = healthy();
  const sheet = ss.getSheetByName('Feedback');
  sheet.appendRow([]);                       // 行を足しても見出しは変えない
  sheet.rows[0].push('先生メモ');
  sheet.maxColumns = sheet.rows[0].length;

  const found = MiraiDb.inspect(ss);
  assert.equal(found.length, 1);
  assert.equal(found[0].kind, '列が多い');
  assert.match(found[0].detail, /先生メモ/);
  assert.equal(found[0].fixable, false);

  MiraiDb.repair(ss);
  assert.ok(ss.getSheetByName('Feedback').rows[0].includes('先生メモ'), '勝手に消さない');
});

test('最初の先生になれるのは、そのファイルを持っている人だけ', () => {
  // コンテナバインドでは、コピーを作った先生だけが所有者・編集者になる。
  // 児童が先に URL を開いても「最初の先生」にはなれない。
  const owned = {
    getOwner: () => ({ getEmail: () => 'Sensei@example.ed.jp' }),
    getEditors: () => [{ getEmail: () => 'kyoutou@example.ed.jp' }],
  };
  assert.equal(MiraiDb.mayBecomeFirstTeacher(owned, 'sensei@example.ed.jp'), true, '所有者（大文字小文字は問わない）');
  assert.equal(MiraiDb.mayBecomeFirstTeacher(owned, 'kyoutou@example.ed.jp'), true, '編集者');
  assert.equal(MiraiDb.mayBecomeFirstTeacher(owned, 'child@example.ed.jp'), false, '児童は通さない');

  // 所有者が取れない環境（共有ドライブなど）では判定できないので止めない。
  const unknown = { getOwner: () => null, getEditors: () => { throw new Error('権限が無い'); } };
  assert.equal(MiraiDb.mayBecomeFirstTeacher(unknown, 'child@example.ed.jp'), true, '判定できないときは前と同じふるまい');

  // メールが取れない（組織外アカウント）ときも止めない。
  assert.equal(MiraiDb.mayBecomeFirstTeacher(owned, ''), true);
});
