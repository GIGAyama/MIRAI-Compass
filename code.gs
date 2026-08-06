/**
 * 🧭 みらいコンパス Ver. 1.0 - サーバーサイドプログラム
 * ==============================================================================
 * このファイルは、Googleスプレッドシート（データベース）とのやり取りを担当します。
 * データの保存、読み出し、初期設定、および「みらいパスポート」との連携機能が含まれています。
 * * Update: HTMLファイルの分割読み込みに対応 (include関数)
 * Update: 排他制御 (LockService) の導入によるデータ整合性の向上
 */

// ==========================================
//  0. システム設定 (Configuration)
// ==========================================

const PROPERTIES = PropertiesService.getScriptProperties();
const APP_TITLE = 'みらいコンパス';

// データベース（スプレッドシート）の設計図
// ※ シート名や列の並び順を定義しています。変更時はここを修正してください。
const DB_SCHEMA = {
  UnitMaster: {
    name: 'UnitMaster',
    headers: ['unitId', 'taskId', 'type', 'title', 'description', 'estTime', 'deletedAt', 'category', 'step', 'textbook', 'tablet', 'print', 'prerequisites', 'format', 'unitInfo', 'totalHours']
  },
  LearningLogs: {
    name: 'LearningLogs',
    headers: ['logId', 'studentId', 'studentName', 'taskId', 'status', 'reflection', 'timestamp', 'classId']
  },
  LiveStatus: {
    name: 'LiveStatus',
    headers: ['studentId', 'studentName', 'currentTask', 'mode', 'lastUpdate', 'currentUnitId', 'currentHour', 'classId', 'x', 'y']
  },
  Feedback: {
    name: 'Feedback',
    // studentId（末尾）: 本人＝Googleメール（サーバー由来）。studentName は表示・先生の名前グループ化用に残す。
    headers: ['feedbackId', 'studentName', 'taskId', 'stamp', 'timestamp', 'classId', 'studentId']
  },
  MyTasks: {
    name: 'MyTasks',
    headers: ['taskId', 'studentName', 'title', 'description', 'estTime', 'created_at', 'unitId', 'classId', 'studentId']
  },
  StudentPlans: {
    name: 'StudentPlans',
    headers: ['studentName', 'unitId', 'planData', 'lastUpdate', 'classId', 'studentId']
  },
  DailyReflections: {
    name: 'DailyReflections',
    headers: ['studentName', 'unitId', 'hour', 'achievement', 'comment', 'teacherCheck', 'timestamp', 'classId', 'skills', 'studentId']
  },
  Portfolios: {
    name: 'Portfolios',
    headers: ['studentName', 'unitId', 'summary', 'lastUpdate', 'classId', 'feedback', 'stamp', 'studentId']
  },
  ClassSchedule: {
    name: 'ClassSchedule',
    headers: ['scheduleId', 'classId', 'date', 'startTime', 'endTime', 'unitId', 'hour', 'message', 'createdAt']
  },
  StudentRoster: {
    name: 'StudentRoster',
    // studentId（末尾・任意）: 児童の Google アカウントのメールアドレス。
    // 入力しておくと、同名児童がいても名簿と学習記録が正確に結びつく。
    // 空欄でも動く（その場合はログイン時の表示名で結びつける従来方式）。
    headers: ['rosterId', 'classId', 'studentNumber', 'studentName', 'isActive', 'updatedAt', 'studentId']
  },
  // 統合: みらいパスポートのワークシート本体（旧・別アプリDBを取り込み）
  Worksheets: {
    name: 'Worksheets',
    headers: ['taskId', 'unitName', 'stepTitle', 'htmlContent', 'lastUpdated', 'jsonSource', 'canvasJson', 'rubricHtml', 'isShared']
  },
  // 統合: 児童のワークシート回答（studentId = Google メール）。
  // feedbackJson（K列）は先生の赤ペン添削キャンバスで、児童の canvasJson（L列）とは別管理。
  Responses: {
    name: 'Responses',
    headers: ['responseId', 'taskId', 'studentId', 'studentName', 'submittedAt', 'canvasImage', 'textContent', 'status', 'feedbackText', 'score', 'feedbackJson', 'canvasJson', 'isPublic', 'reactions', 'reflectionText', 'studentAnswers']
  }
};

// =============================================================
//  児童の正準キー（canonical key）
// =============================================================
// すべての先生向け集計は、この1つのキー規則で児童をまとめる。
//   ・studentId（本人の Google メール）が分かる行 → そのメールがキー
//   ・分からない行（旧データ・名簿のみの児童）    → 'name:' + 表示名
// 児童の自己書き込みは常に studentId 付きなので、同名児童のデータは
// 混ざらない。クライアント側にも同じ規則の studentKeyOf() がある。
function studentKey(sid, name) {
  const s = String(sid || '').trim();
  return s ? s : 'name:' + String(name || '');
}

// === ワークシート/回答シートの列番号（1始まり）: 統合したパスポート機能で使用 =========
var WS_COL_TASK_ID      = 1;  // A列
var WS_COL_UNIT_NAME    = 2;  // B列
var WS_COL_STEP_TITLE   = 3;  // C列
var WS_COL_HTML_CONTENT = 4;  // D列
var WS_COL_LAST_UPDATED = 5;  // E列
var WS_COL_JSON_SOURCE  = 6;  // F列
var WS_COL_CANVAS_JSON  = 7;  // G列
var WS_COL_RUBRIC_HTML  = 8;  // H列
var WS_COL_IS_SHARED    = 9;  // I列
var WS_TOTAL_COLS       = 9;

var RS_COL_RESPONSE_ID  = 1;   // A列
var RS_COL_TASK_ID      = 2;   // B列
var RS_COL_STUDENT_ID   = 3;   // C列
var RS_COL_STUDENT_NAME = 4;   // D列
var RS_COL_SUBMITTED_AT = 5;   // E列
var RS_COL_CANVAS_IMAGE = 6;   // F列
var RS_COL_TEXT_CONTENT = 7;   // G列
var RS_COL_STATUS       = 8;   // H列
var RS_COL_FEEDBACK_TXT = 9;   // I列
var RS_COL_SCORE        = 10;  // J列
var RS_COL_FEEDBACK_JSON= 11;  // K列: 先生の赤ペン添削
var RS_COL_CANVAS_JSON  = 12;  // L列: 児童の手書き
var RS_COL_IS_PUBLIC    = 13;  // M列
var RS_COL_REACTIONS    = 14;  // N列
var RS_COL_REFLECTION   = 15;  // O列
var RS_COL_ANSWERS_JSON = 16;  // P列
var RS_TOTAL_COLS       = 16;

// === MIRAI SHARED (SERVER) 認証・本人確認 =====================================
// 統合方針: 全員が Google アカウントでアクセスする（デプロイ「実行:自分／アクセス:同一Workspace」）。
// サーバーは呼び出し元を Session.getActiveUser().getEmail() で必ず自分で判定し、
// クライアントが申告する ID を一切信用しない。これにより児童のなりすまし・IDOR が構造的に消える。
var MiraiAuth = (function () {
  var TEACHER_KEY = 'TEACHER_EMAILS'; // ScriptProperties に JSON 配列で保持
  function currentEmail() {
    try { return String(Session.getActiveUser().getEmail() || '').trim().toLowerCase(); }
    catch (e) { return ''; }
  }
  function _list() {
    try { return JSON.parse(PropertiesService.getScriptProperties().getProperty(TEACHER_KEY) || '[]'); }
    catch (e) { return []; }
  }
  function _save(list) {
    PropertiesService.getScriptProperties().setProperty(TEACHER_KEY, JSON.stringify(list));
  }
  function teacherEmails() { return _list(); }
  function isTeacher() {
    var me = currentEmail();
    return !!me && _list().indexOf(me) !== -1;
  }
  function requireTeacher() {
    if (!isTeacher()) {
      throw new Error('AUTH_REQUIRED: 先生として登録されたGoogleアカウントでログインしてください。');
    }
  }
  // ログイン済みの本人メールを必須にする（児童データの本人限定アクセス用）。未ログインなら例外。
  function requireUser() {
    var me = currentEmail();
    if (!me) throw new Error('AUTH_REQUIRED: Googleアカウントでログインしてください。');
    return me;
  }
  function addTeacher(email) {
    var e = String(email || '').trim().toLowerCase();
    if (!e) return _list();
    var list = _list();
    if (list.indexOf(e) === -1) { list.push(e); _save(list); }
    return list;
  }
  function removeTeacher(email) {
    var e = String(email || '').trim().toLowerCase();
    var list = _list().filter(function (x) { return x !== e; });
    _save(list);
    return list;
  }
  // 初期化時のブートストラップ: 一覧が空なら、初期化を実行した本人を最初の先生にする。
  function bootstrapFirstTeacher() {
    if (_list().length === 0) {
      var me = currentEmail();
      if (me) _save([me]);
    }
    return _list();
  }
  return {
    currentEmail: currentEmail, requireUser: requireUser,
    isTeacher: isTeacher, requireTeacher: requireTeacher, teacherEmails: teacherEmails,
    addTeacher: addTeacher, removeTeacher: removeTeacher, bootstrapFirstTeacher: bootstrapFirstTeacher
  };
})();

// ==========================================
//  1. 基本機能 & HTML配信 (Core Functions)
// ==========================================

/**
 * アプリにアクセスした時に実行される関数
 */
function doGet(e) {
  try {
    // index.html をテンプレートとして読み込む
    const template = HtmlService.createTemplateFromFile('index');
    template.mode = 'student'; // デフォルトモード
    
    return template.evaluate()
      // ⚠️ ここは index.html の <meta name="viewport"> と【両方】直す必要がある。
      //    addMetaTag はサーバー側の処理なので、index.html だけ直しても
      //    GAS が返す画面には反映されない（逆に、index.html を手元で
      //    組み立てただけではこちらが再現されない）。
      //    viewport-fit=cover が無いと、iPad のノッチ・ホームバーの領域に
      //    背景が伸びず、safe-area-inset も 0 のままになる。
      .addMetaTag('viewport', 'width=device-width, initial-scale=1.0, viewport-fit=cover')
      .setTitle(APP_TITLE)
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .setFaviconUrl('https://drive.google.com/uc?id=1zzJYaALAtpAVIkEG_k5oGoVQu0hIPS7G&.png');
  } catch (error) {
    return HtmlService.createHtmlOutput(`<h2>起動エラー</h2><p>${error.toString()}</p>`);
  }
}

/**
 * 分割されたHTMLファイルを読み込むためのヘルパー関数
 * index.html 内で <?!= include('filename'); ?> のように使います
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ==========================================
//  2. システム初期化・設定 (System Init)
// ==========================================

/**
 * アプリ起動時の初期データ取得
 * スプレッドシートIDがあるかなどを確認します
 */
function getAppInitialData() {
  try {
    const ssId = PROPERTIES.getProperty('SS_ID');
    return createSuccessResponse({
      isInitialized: !!ssId,
      ssUrl: ssId ? `https://docs.google.com/spreadsheets/d/${ssId}/edit` : null
    });
  } catch (e) {
    return createErrorResponse(e);
  }
}

/**
 * システム初期化実行
 * 新しいスプレッドシートを作成し、IDを保存します
 */
function initSystem() {
  try {
    // 既にシステムが存在する場合の再初期化は、先生として登録されていないと拒否する。
    // 未セットアップ（初回）は誰でも通し、その本人を最初の先生にブートストラップする。
    var existing = PROPERTIES.getProperty('SS_ID');
    if (existing && !MiraiAuth.isTeacher()) {
      throw new Error('AUTH_REQUIRED: 先生として登録されたGoogleアカウントでログインしてください。');
    }

    let ssId = PROPERTIES.getProperty('SS_ID');
    let ss;
    if (ssId) {
      ss = SpreadsheetApp.openById(ssId);
    } else {
      ss = SpreadsheetApp.create('みらいコンパス_データベース');
      ssId = ss.getId();
      PROPERTIES.setProperty('SS_ID', ssId);
    }

    // 全シートの存在確認と作成
    checkAndFixSheets(ss);

    // デフォルトの「シート1」があれば削除
    const defaultSheet = ss.getSheetByName('シート1');
    if (defaultSheet) ss.deleteSheet(defaultSheet);

    // 初期化を実行した本人を最初の先生にする（一覧が空のときのみ）。
    MiraiAuth.bootstrapFirstTeacher();
    return createSuccessResponse({ message: 'システムを初期化しました。' });
  } catch (e) {
    return createErrorResponse(e);
  }
}

// 呼び出し元の役割を返す。クライアントはこの結果で先生モードに入れるかを判断する。
function getMyRole() {
  return { email: MiraiAuth.currentEmail(), isTeacher: MiraiAuth.isTeacher(), teacherCount: MiraiAuth.teacherEmails().length };
}
function getTeacherList() {
  MiraiAuth.requireTeacher();
  return createSuccessResponse({ teachers: MiraiAuth.teacherEmails() });
}
function addTeacherEmail(email) {
  MiraiAuth.requireTeacher();
  return createSuccessResponse({ teachers: MiraiAuth.addTeacher(email) });
}
function removeTeacherEmail(email) {
  MiraiAuth.requireTeacher();
  var me = MiraiAuth.currentEmail();
  if (String(email || '').trim().toLowerCase() === me) throw new Error('自分自身は削除できません。');
  return createSuccessResponse({ teachers: MiraiAuth.removeTeacher(email) });
}

// ==========================================
//  3. データ読み込み (Read Data)
// ==========================================

/**
 * 全データの取得（初回ロード用）
 */
function getData() {
  try {
    const ssId = PROPERTIES.getProperty('SS_ID');
    if (!ssId) return createSuccessResponse({ json: JSON.stringify({}) });
    
    const ss = SpreadsheetApp.openById(ssId);
    checkAndFixSheets(ss); // 念のためシート構造チェック

    // 各シートからデータを取得して整形
    const unitData = fetchSheetData(ss, DB_SCHEMA.UnitMaster.name).map(r => ({
      unitId: String(r[0]), taskId: String(r[1]), type: String(r[2]), title: String(r[3]),
      desc: String(r[4]), time: Number(r[5]), category: String(r[7]), step: String(r[8] || ''),
      textbook: String(r[9] || ''), tablet: String(r[10] || ''), print: String(r[11] || ''),
      prerequisites: String(r[12] || '').split(',').filter(x => x), format: String(r[13] || 'student'),
      unitInfo: safeJsonParse(r[14]), totalHours: Number(r[15] || 8)
    }));

    const liveData = fetchSheetData(ss, DB_SCHEMA.LiveStatus.name).map(r => ({
      id: String(r[0]), name: String(r[1]), task: String(r[2]), mode: String(r[3]),
      time: r[4] ? formatDate(r[4]) : '', currentUnitId: String(r[5] || ''),
      currentHour: Number(r[6] || 1), classId: String(r[7] || ''),
      x: Number(r[8]) || 0, y: Number(r[9]) || 0
    }));

    // [改修] 名簿データの取得：出席番号・在籍状況・本人メール（任意）も取得する
    const rosterData = fetchSheetData(ss, DB_SCHEMA.StudentRoster.name).map(r => ({
      classId: String(r[1]),
      name: String(r[3]),
      number: r[2] !== "" ? Number(r[2]) : 999, // 出席番号（空なら末尾へ）
      isActive: (r[4] === true || r[4] === "TRUE" || r[4] === ""), // 在籍状況（空ならTrue扱い）
      studentId: String(r[6] || "") // 本人メール（任意・あれば記録と正確に結びつく）
    })).filter(r => r.name);

    // ==== 以下の集計はすべて studentKey（本人メール優先・無ければ name: 表示名）でまとめる ====
    // 学習ログを集計（最新ステータスのみ）
    const clsProgress = {};
    fetchSheetData(ss, DB_SCHEMA.LearningLogs.name).forEach(r => {
      const key = studentKey(r[1], r[2]); const taskId = String(r[3]); const status = String(r[4]); const reflection = String(r[5] || "");
      if (!clsProgress[key]) clsProgress[key] = {};
      if (!clsProgress[key][taskId]) clsProgress[key][taskId] = { status: '', reflection: '' };
      if (status && status !== 'メモ') clsProgress[key][taskId].status = status;
      if (reflection) clsProgress[key][taskId].reflection = reflection;
    });

    // その他のデータを取得
    const clsFeedback = {};
    fetchSheetData(ss, DB_SCHEMA.Feedback.name).forEach(r => {
      const key = studentKey(r[6], r[1]); const taskId = String(r[2]); const stamp = String(r[3]);
      if (!clsFeedback[key]) clsFeedback[key] = {};
      clsFeedback[key][taskId] = stamp;
    });

    const allMyTasks = {};
    fetchSheetData(ss, DB_SCHEMA.MyTasks.name).forEach(r => {
      const key = studentKey(r[8], r[1]);
      if (!allMyTasks[key]) allMyTasks[key] = [];
      allMyTasks[key].push({
        taskId: String(r[0]), title: String(r[2]), desc: String(r[3]), time: Number(r[4]),
        category: 'マイタスク', type: 'challenge', format: 'student', unitId: String(r[6] || '')
      });
    });

    const clsPlans = {};
    fetchSheetData(ss, DB_SCHEMA.StudentPlans.name).forEach(r => {
      const key = studentKey(r[5], r[0]); const uid = String(r[1]);
      if (!clsPlans[key]) clsPlans[key] = {};
      clsPlans[key][uid] = safeJsonParse(r[2]);
    });

    const clsReflections = {};
    fetchSheetData(ss, DB_SCHEMA.DailyReflections.name).forEach(r => {
      const key = studentKey(r[9], r[0]); const uid = String(r[1]); const hour = String(r[2]);
      if (!clsReflections[key]) clsReflections[key] = {};
      if (!clsReflections[key][uid]) clsReflections[key][uid] = {};
      clsReflections[key][uid][hour] = {
        achievement: r[3], comment: r[4], check: r[5], skills: safeJsonParse(r[8])
      };
    });

    const clsPortfolios = {};
    fetchSheetData(ss, DB_SCHEMA.Portfolios.name).forEach(r => {
      const key = studentKey(r[7], r[0]); const uid = String(r[1]);
      if (!clsPortfolios[key]) clsPortfolios[key] = {};
      clsPortfolios[key][uid] = {
        summary: String(r[2]), feedback: String(r[5] || ""), stamp: String(r[6] || ""),
        studentId: String(r[7] || "") // 先生フィードバックを本人の行に確実に着地させるための本人メール
      };
    });

    // 今日以降のスケジュールのみ取得
    // ※シート上で日付文字列がDate型に自動変換されるため、必ず正規化してから比較する
    const todayStr = Utilities.formatDate(new Date(), "JST", "yyyy-MM-dd");
    const schedules = fetchSheetData(ss, DB_SCHEMA.ClassSchedule.name)
      .map(r => ({
        id: String(r[0]), classId: String(r[1]), date: normalizeDateStr(r[2]),
        startTime: normalizeTimeStr(r[3]), endTime: normalizeTimeStr(r[4]),
        unitId: String(r[5]), hour: String(r[6]), message: String(r[7])
      }))
      .filter(s => s.date >= todayStr);

    // getData は先生用の一括リーダーだが、児童モードの初期ロード（loadMainData）からも同じ関数が呼ばれる。
    // 児童が受け取る生JSONに他児童の「機微データ」を載せない（＝クライアント側の出し分けに頼らない IDOR 対策）。
    // 児童ビューが実際に使う他者データは「ギャラリーウォーク／ひろば」向けの plans・myTasks（匿名切替あり）と
    // 教室の live 状況・roster のみ。progress / feedback / dailyReflections / portfolios は児童ビューでは
    // 自分の分を getStudentProgress から取得しており、他者分は一切参照しないため、非教員には空で返す。
    // 先生（requireTeacher 相当）には従来どおり全データを返す（集計・座席・ヒートマップ・まとめ表示は不変）。
    const isTeacherCaller = MiraiAuth.isTeacher();
    return createSuccessResponse({
      json: JSON.stringify({
        unit: unitData, live: liveData, roster: rosterData,
        progress: isTeacherCaller ? clsProgress : {},
        feedback: isTeacherCaller ? clsFeedback : {},
        myTasks: allMyTasks, plans: clsPlans,
        dailyReflections: isTeacherCaller ? clsReflections : {},
        portfolios: isTeacherCaller ? clsPortfolios : {},
        schedules: schedules
      })
    });
  } catch (e) { return createErrorResponse(e); }
}

/**
 * 軽量データ取得API（リアルタイムダッシュボード用）
 * 全データを取得せず、LiveStatusのみを取得して返すことで高速化
 */
function getLiveStatusSnapshot() {
  try {
    const ssId = PROPERTIES.getProperty('SS_ID');
    if (!ssId) return createSuccessResponse({ live: [] });

    // 40台が10秒間隔で同時にポーリングすると、同じシートの全読みが毎秒4回走る。
    // 8秒だけ結果を共有キャッシュして、クラス全体で「1回の読み」を使い回す。
    // 書き込み側（updateStatus / saveSeatCoordinates）がキャッシュを即破棄するので、
    // SOSなどの変化は次のポーリングで必ず拾える。
    const cache = CacheService.getScriptCache();
    const cacheKey = 'live_snapshot_' + ssId;
    const cached = cache.get(cacheKey);
    if (cached) return createSuccessResponse({ live: JSON.parse(cached) });

    const ss = SpreadsheetApp.openById(ssId);
    const liveData = fetchSheetData(ss, DB_SCHEMA.LiveStatus.name).map(r => ({
      id: String(r[0]), name: String(r[1]), task: String(r[2]), mode: String(r[3]),
      time: r[4] ? formatDate(r[4]) : '', currentUnitId: String(r[5] || ''),
      currentHour: Number(r[6] || 1), classId: String(r[7] || ''),
      x: Number(r[8]) || 0, y: Number(r[9]) || 0
    }));

    try { cache.put(cacheKey, JSON.stringify(liveData), 8); } catch (ignore) { /* 100KB超過時は素通し */ }
    return createSuccessResponse({ live: liveData });
  } catch (e) {
    return createErrorResponse(e);
  }
}

/**
 * ライブスナップショットの共有キャッシュを破棄する（書き込みAPIから呼ぶ）。
 * これにより児童のSOS・状態変化が次のポーリングで確実に反映される。
 */
function invalidateLiveSnapshotCache() {
  const ssId = PROPERTIES.getProperty('SS_ID');
  if (!ssId) return;
  try { CacheService.getScriptCache().remove('live_snapshot_' + ssId); } catch (ignore) { /* キャッシュ不調でも本処理は続行 */ }
}

/**
 * ギャラリーウォーク用の軽量API（クラスの計画とマイタスク）。
 * 「みんなの計画」はこれまでログイン時のスナップショットを表示し続けており、
 * 友だちが計画を変えても児童の画面には反映されなかった。
 * 返す内容は getData が児童にも返している plans / myTasks と同じ範囲
 * （進捗・ふりかえり等の機微データは含まない）。20秒の共有キャッシュ付き。
 */
function getGalleryData() {
  try {
    MiraiAuth.requireUser();
    const ssId = PROPERTIES.getProperty('SS_ID');
    if (!ssId) return createSuccessResponse({ json: JSON.stringify({}) });

    const cache = CacheService.getScriptCache();
    const cacheKey = 'gallery_data_' + ssId;
    const cached = cache.get(cacheKey);
    if (cached) return createSuccessResponse({ json: cached });

    const ss = SpreadsheetApp.openById(ssId);
    // getData と同じ studentKey 規則でまとめる
    const clsPlans = {};
    fetchSheetData(ss, DB_SCHEMA.StudentPlans.name).forEach(r => {
      const key = studentKey(r[5], r[0]); const uid = String(r[1]);
      if (!clsPlans[key]) clsPlans[key] = {};
      clsPlans[key][uid] = safeJsonParse(r[2]);
    });

    const allMyTasks = {};
    fetchSheetData(ss, DB_SCHEMA.MyTasks.name).forEach(r => {
      const key = studentKey(r[8], r[1]);
      if (!allMyTasks[key]) allMyTasks[key] = [];
      allMyTasks[key].push({
        taskId: String(r[0]), title: String(r[2]), desc: String(r[3]), time: Number(r[4]),
        category: 'マイタスク', type: 'challenge', format: 'student', unitId: String(r[6] || '')
      });
    });

    const json = JSON.stringify({ plans: clsPlans, myTasks: allMyTasks });
    try { cache.put(cacheKey, json, 20); } catch (ignore) { /* 100KB超過時は素通し */ }
    return createSuccessResponse({ json: json });
  } catch (e) { return createErrorResponse(e); }
}

/**
 * 先生向け進捗スナップショットAPI（進捗一覧のライブ更新用）。
 * getData の clsProgress / clsFeedback と同じ集計を、この2シートだけ読んで返す。
 * これまで進捗一覧は画面を切り替えるまで更新されず、「授業中の見取り」に
 * 使えなかった。15秒の共有キャッシュ付きで、複数の先生が同時に開いても
 * シート読みは15秒に1回に抑えられる。
 */
function getProgressSnapshot() {
  try {
    MiraiAuth.requireTeacher(); // クラス全員分のデータを返すため教員限定
    const ssId = PROPERTIES.getProperty('SS_ID');
    if (!ssId) return createSuccessResponse({ json: JSON.stringify({}) });

    const cache = CacheService.getScriptCache();
    const cacheKey = 'progress_snapshot_' + ssId;
    const cached = cache.get(cacheKey);
    if (cached) return createSuccessResponse({ json: cached });

    const ss = SpreadsheetApp.openById(ssId);
    // getData と同じ studentKey 規則でまとめる
    const clsProgress = {};
    fetchSheetData(ss, DB_SCHEMA.LearningLogs.name).forEach(r => {
      const key = studentKey(r[1], r[2]); const taskId = String(r[3]); const status = String(r[4]); const reflection = String(r[5] || "");
      if (!clsProgress[key]) clsProgress[key] = {};
      if (!clsProgress[key][taskId]) clsProgress[key][taskId] = { status: '', reflection: '' };
      if (status && status !== 'メモ') clsProgress[key][taskId].status = status;
      if (reflection) clsProgress[key][taskId].reflection = reflection;
    });

    const clsFeedback = {};
    fetchSheetData(ss, DB_SCHEMA.Feedback.name).forEach(r => {
      const key = studentKey(r[6], r[1]); const taskId = String(r[2]); const stamp = String(r[3]);
      if (!clsFeedback[key]) clsFeedback[key] = {};
      clsFeedback[key][taskId] = stamp;
    });

    const json = JSON.stringify({ progress: clsProgress, feedback: clsFeedback });
    try { cache.put(cacheKey, json, 15); } catch (ignore) { /* 100KB超過時は素通し */ }
    return createSuccessResponse({ json: json });
  } catch (e) { return createErrorResponse(e); }
}

/**
 * 児童向け軽量ポーリングAPI。
 * 先生スタンプ（Feedback）と今日以降の時間割（ClassSchedule）だけを返す。
 * getStudentProgress（6シート全読み）を定期実行するのは重すぎるため、
 * 「開いている児童の画面に先生の反応が届く」ことに必要な最小データに絞っている。
 */
function getStudentPulse(studentName) {
  try {
    const sid = MiraiAuth.requireUser();
    const ssId = PROPERTIES.getProperty('SS_ID');
    if (!ssId) return createSuccessResponse({ json: JSON.stringify({}) });
    const ss = SpreadsheetApp.openById(ssId);

    // スタンプの照合ルールは getStudentProgress と同一
    // （studentId がある行は本人メールでだけ照合、無い行は表示名でフォールバック）
    const fbMap = {};
    fetchSheetData(ss, DB_SCHEMA.Feedback.name).forEach(r => {
      const rowSid = String(r[6] || "");
      if (rowSid ? rowSid === sid : r[1] === studentName) fbMap[String(r[2])] = String(r[3]);
    });

    const todayStr = Utilities.formatDate(new Date(), "JST", "yyyy-MM-dd");
    const schedules = fetchSheetData(ss, DB_SCHEMA.ClassSchedule.name)
      .map(r => ({
        id: String(r[0]), classId: String(r[1]), date: normalizeDateStr(r[2]),
        startTime: normalizeTimeStr(r[3]), endTime: normalizeTimeStr(r[4]),
        unitId: String(r[5]), hour: String(r[6]), message: String(r[7])
      }))
      .filter(s => s.date >= todayStr);

    return createSuccessResponse({ json: JSON.stringify({ feedback: fbMap, schedules: schedules }) });
  } catch (e) { return createErrorResponse(e); }
}

/**
 * 児童個人用データの取得（ログイン時）
 */
function getStudentProgress(studentName, classId, currentUnitId) {
  try {
    // 本人＝Googleメール（サーバー由来）。これ自身の読み取りなので、身元は必ずサーバーで確定する。
    // クライアント申告の studentName は表示名としてのみ使い、アクセス範囲の判定には使わない。
    const sid = MiraiAuth.requireUser();
    const ssId = PROPERTIES.getProperty('SS_ID');
    if (!ssId) return createSuccessResponse({ json: JSON.stringify({}) });
    const ss = SpreadsheetApp.openById(ssId);

    // ログイン時に現在の単元・クラス情報をLiveStatusに書き込む（本人行を studentId で特定）
    if (classId || currentUnitId) {
      updateLiveStatusMeta(ss, sid, studentName, classId, currentUnitId);
    }

    // 必要なデータのみ抽出して返す（すべて本人 studentId で自己スコープ）
    const map = {};
    fetchSheetData(ss, DB_SCHEMA.LearningLogs.name).forEach(r => {
      if (r[1] === sid) { // col1 = studentId
        const tid = String(r[3]);
        if (!map[tid]) map[tid] = { status: '', reflection: '' };
        if (r[4] && r[4] !== 'メモ') map[tid].status = r[4];
        if (r[5]) map[tid].reflection = r[5];
      }
    });

    // スタンプの照合ルール：行に studentId が入っていれば本人メールでだけ照合（同名誤配なし）。
    // 入っていない行（旧データ・ID不明時）は従来どおり表示名で照合するフォールバック。
    const fbMap = {};
    fetchSheetData(ss, DB_SCHEMA.Feedback.name).forEach(r => {
      const rowSid = String(r[6] || "");
      if (rowSid ? rowSid === sid : r[1] === studentName) fbMap[String(r[2])] = String(r[3]);
    });

    const myTasks = [];
    fetchSheetData(ss, DB_SCHEMA.MyTasks.name).forEach(r => {
      if (r[8] === sid) { // col8 = studentId
        myTasks.push({
          taskId: String(r[0]), title: String(r[2]), desc: String(r[3]), time: Number(r[4]),
          category: 'マイタスク', type: 'challenge', format: 'student', unitId: String(r[6] || '')
        });
      }
    });

    const plans = {};
    fetchSheetData(ss, DB_SCHEMA.StudentPlans.name).forEach(r => {
      if (r[5] === sid) plans[r[1]] = safeJsonParse(r[2]); // col5 = studentId
    });

    const reflections = {};
    fetchSheetData(ss, DB_SCHEMA.DailyReflections.name).forEach(r => {
      if (r[9] === sid) { // col9 = studentId
        const uid = String(r[1]);
        const hour = String(r[2]);
        if (!reflections[uid]) reflections[uid] = {};
        reflections[uid][hour] = {
          achievement: r[3], comment: r[4], check: r[5], skills: safeJsonParse(r[8])
        };
      }
    });

    let portfolioData = { summary: "", feedback: "", stamp: "" };
    const pSheet = ss.getSheetByName(DB_SCHEMA.Portfolios.name);
    if(pSheet) {
      const pData = pSheet.getDataRange().getValues();
      for(let i=1; i<pData.length; i++) {
        if(pData[i][7] === sid && String(pData[i][1]) === String(currentUnitId)) { // col7 = studentId
          portfolioData = {
            summary: pData[i][2],
            feedback: pData[i][5] || "",
            stamp: pData[i][6] || ""
          };
          break;
        }
      }
    }

    return createSuccessResponse({
      json: JSON.stringify({
        progress: map, feedback: fbMap, myTasks: myTasks,
        portfolio: portfolioData, plans: plans, reflections: reflections,
        // 自分の正準キー。クライアントがクラス集計（studentKey キー）の中から
        // 自分の分を見分けるのに使う（ギャラリーの自分スキップ・自計画の保持など）。
        myKey: studentKey(sid, studentName)
      })
    });
  } catch (e) { return createErrorResponse(e); }
}

/**
 * 内部関数: 児童のLiveStatus（クラス、単元）を更新
 */
function updateLiveStatusMeta(ss, sid, name, classId, unitId) {
  const lock = LockService.getScriptLock();
  if (lock.tryLock(5000)) { // 5秒待機
    try {
      const sheet = ss.getSheetByName(DB_SCHEMA.LiveStatus.name);
      const data = sheet.getDataRange().getValues();
      for (let i = 1; i < data.length; i++) {
        if (data[i][0] === sid) { // col0 = studentId（本人）
          if (unitId) sheet.getRange(i + 1, 6).setValue(unitId);
          if (classId) sheet.getRange(i + 1, 8).setValue(classId);
          if (name) sheet.getRange(i + 1, 2).setValue(name); // 表示名を最新に
          invalidateLiveSnapshotCache();
          return;
        }
      }
    } finally {
      lock.releaseLock();
    }
  }
}

// ==========================================
//  4. データ保存 (Write Data)
// ==========================================

/**
 * 学習状況の更新（完了、途中、メモ、SOSステータスなど）
 */
function updateStatus(studentName, taskId, taskTitle, status, mode, reflection, classId, currentUnitId) {
  const lock = LockService.getScriptLock();
  // 他のAPIと同様、失敗時も {success:false} 形式で返す（throwするとクライアントの成功/失敗判定が不統一になる）
  if (!lock.tryLock(10000)) {
    return createErrorResponse(new Error("サーバーが混み合っています。もう一度お試しください。"));
  }

  try {
    // 本人＝Googleメール（サーバー由来）を身元/UPSERTキーにする。studentName は表示・先生の名前グループ化用。
    const sid = MiraiAuth.requireUser();
    const ss = getSpreadsheet();
    const now = new Date();

    // ログ履歴に追加（studentId 列＝本人メール、studentName 列＝表示名）
    ss.getSheetByName(DB_SCHEMA.LearningLogs.name).appendRow([
      Utilities.getUuid(), sid, studentName, taskId, status, reflection || "", now, classId || ""
    ]);

    // LiveStatus（現在の状態）を更新。行の一致は studentId（本人）で行い、同名衝突を避ける。
    if (status !== 'メモ') {
      const liveSheet = ss.getSheetByName(DB_SCHEMA.LiveStatus.name);
      const data = liveSheet.getDataRange().getValues();
      let rIdx = -1;
      for (let i = 1; i < data.length; i++) {
        if (data[i][0] === sid) { rIdx = i + 1; break; }
      }
      const displayTitle = taskTitle || taskId;

      if (rIdx > 0) {
        // 既存行の col2〜col8 を1回の setValues で更新（往復5回→1回）。
        // 未指定の unitId / classId / currentHour は既存値を保持する。
        const cur = data[rIdx - 1];
        liveSheet.getRange(rIdx, 2, 1, 7).setValues([[
          studentName, displayTitle, mode, now,
          currentUnitId || cur[5] || "", cur[6] || 1, classId || cur[7] || ""
        ]]);
      } else {
        // 新規追加（studentId 列＝本人メール、studentName 列＝表示名）
        liveSheet.appendRow([sid, studentName, displayTitle, mode, now, currentUnitId || "", 1, classId || "", 0, 0]);
      }
      invalidateLiveSnapshotCache();
    }
    return createSuccessResponse();
  } catch (e) { 
    return createErrorResponse(e); 
  } finally {
    lock.releaseLock();
  }
}

// --- 先生用管理機能（ロック推奨） ---

function updateUnitTask(taskId, updateData) {
  const lock = LockService.getScriptLock();
  if (lock.tryLock(5000)) {
    try {
      MiraiAuth.requireTeacher();
      const ss = getSpreadsheet();
      const sheet = ss.getSheetByName(DB_SCHEMA.UnitMaster.name);
      const data = sheet.getDataRange().getValues();
      let rowIndex = -1;

      for (let i = 1; i < data.length; i++) {
        if (String(data[i][1]) === String(taskId)) { rowIndex = i + 1; break; }
      }

      if (rowIndex > 0) {
        if (updateData.title !== undefined) sheet.getRange(rowIndex, 4).setValue(updateData.title);
        if (updateData.description !== undefined) sheet.getRange(rowIndex, 5).setValue(updateData.description);
        if (updateData.estTime !== undefined) sheet.getRange(rowIndex, 6).setValue(updateData.estTime);
        if (updateData.category !== undefined) sheet.getRange(rowIndex, 8).setValue(updateData.category);
        if (updateData.step !== undefined) sheet.getRange(rowIndex, 9).setValue(updateData.step);
        if (updateData.textbook !== undefined) sheet.getRange(rowIndex, 10).setValue(updateData.textbook);
        if (updateData.tablet !== undefined) sheet.getRange(rowIndex, 11).setValue(updateData.tablet);
        if (updateData.print !== undefined) sheet.getRange(rowIndex, 12).setValue(updateData.print);
        if (updateData.format !== undefined) sheet.getRange(rowIndex, 14).setValue(updateData.format);
        if (updateData.type !== undefined) sheet.getRange(rowIndex, 3).setValue(updateData.type);
        
        return createSuccessResponse({ message: '更新しました' });
      } else {
        throw new Error('タスクが見つかりません');
      }
    } catch (e) { return createErrorResponse(e); } finally { lock.releaseLock(); }
  } else {
    return createErrorResponse(new Error("Timeout"));
  }
}

function addUnitTask(unitId, taskData) {
  const lock = LockService.getScriptLock();
  if (lock.tryLock(5000)) {
    try {
      MiraiAuth.requireTeacher();
      const ss = getSpreadsheet();
      const sheet = ss.getSheetByName(DB_SCHEMA.UnitMaster.name);
      // 単元の基本情報を取得するために検索
      const data = sheet.getDataRange().getValues();
      let refRow = null;
      for(let i=1; i<data.length; i++) {
        if(String(data[i][0]) === String(unitId)) { refRow = data[i]; break; }
      }
      if(!refRow) throw new Error('単元が見つかりません');
      
      const taskId = "T" + Utilities.getUuid().substring(0, 8);
      const newRow = [
        unitId, taskId, taskData.type || 'must', taskData.title || '無題',
        taskData.description || '', taskData.estTime || 15, '', 
        taskData.category || 'まなぶ', taskData.step || '',
        taskData.textbook || '', taskData.tablet || '', taskData.print || '',
        '', taskData.format || 'student', refRow[14], refRow[15] 
      ];
      sheet.appendRow(newRow);
      return createSuccessResponse({ taskId: taskId, message: 'タスクを追加しました' });
    } catch(e) { return createErrorResponse(e); } finally { lock.releaseLock(); }
  } else { return createErrorResponse(new Error("Timeout")); }
}

function deleteUnitTask(taskId) {
  const lock = LockService.getScriptLock();
  if (lock.tryLock(5000)) {
    try {
      MiraiAuth.requireTeacher();
      const ss = getSpreadsheet();
      const sheet = ss.getSheetByName(DB_SCHEMA.UnitMaster.name);
      const data = sheet.getDataRange().getValues();
      let rowIndex = -1;
      for (let i = 1; i < data.length; i++) {
        if (String(data[i][1]) === String(taskId)) { rowIndex = i + 1; break; }
      }
      if (rowIndex > 0) {
        sheet.deleteRow(rowIndex);
        return createSuccessResponse({ message: '削除しました' });
      } else {
        throw new Error('タスクが見つかりません');
      }
    } catch (e) { return createErrorResponse(e); } finally { lock.releaseLock(); }
  } else { return createErrorResponse(new Error("Timeout")); }
}

function createNewUnit(unitInfo) {
  const lock = LockService.getScriptLock();
  if (lock.tryLock(5000)) {
    try {
      MiraiAuth.requireTeacher();
      const ss = getSpreadsheet();
      const sheet = ss.getSheetByName(DB_SCHEMA.UnitMaster.name);

      const unitId = "U" + Utilities.formatDate(new Date(), "JST", "yyyyMMddHHmmss");
      const taskId = "T" + Utilities.getUuid().substring(0, 8);

      const infoObj = {
        title: unitInfo.title, subject: unitInfo.subject,
        grade: unitInfo.grade, totalHours: Number(unitInfo.totalHours)
      };

      // 最初のタスク（導入）を自動作成
      const newRow = [
        unitId, taskId, 'must', '【導入】' + unitInfo.title,
        '単元の目標や計画を確認しよう', 15, '', '導入', 'Step 1',
        '', '', '', '', 'teacher', JSON.stringify(infoObj), infoObj.totalHours
      ];
      sheet.appendRow(newRow);
      return createSuccessResponse({ message: '新しい単元を作成しました', unitId: unitId });
    } catch (e) { return createErrorResponse(e); } finally { lock.releaseLock(); }
  } else { return createErrorResponse(new Error("Timeout")); }
}

function updateUnitBasicInfo(unitId, infoData) {
  const lock = LockService.getScriptLock();
  if (lock.tryLock(5000)) {
    try {
      MiraiAuth.requireTeacher();
      const ss = getSpreadsheet();
      const sheet = ss.getSheetByName(DB_SCHEMA.UnitMaster.name);
      const data = sheet.getDataRange().getValues();

      for (let i = 1; i < data.length; i++) {
        if (String(data[i][0]) === String(unitId)) {
          let currentInfo = safeJsonParse(data[i][14]);
          if (infoData.title !== undefined) currentInfo.title = infoData.title;
          if (infoData.subject !== undefined) currentInfo.subject = infoData.subject;
          if (infoData.grade !== undefined) currentInfo.grade = infoData.grade;
          if (infoData.goal !== undefined) currentInfo.goal = infoData.goal;
          if (infoData.description !== undefined) currentInfo.description = infoData.description;

          sheet.getRange(i + 1, 15).setValue(JSON.stringify(currentInfo));
          if (infoData.title !== undefined) sheet.getRange(i + 1, 4).setValue(infoData.title);
        }
      }
      return createSuccessResponse({ message: '単元情報を更新しました' });
    } catch (e) { return createErrorResponse(e); } finally { lock.releaseLock(); }
  } else { return createErrorResponse(new Error("Timeout")); }
}

/**
 * 単元内の児童用タスクについて、ワークシート（Worksheets シート）の器を用意する。
 * 旧・ImportQueue によるアプリ間連携を置き換える、アプリ内部の処理。
 * UnitMaster の当該単元の各タスク（format!=='teacher'）に対し、
 * まだ Worksheets 行が無ければ taskId を共通名前空間として追加する。
 * @param {string} unitId
 * @return {Object} { success:true, created:N, taskIds:[...] }
 */
function createWorksheetsForUnit(unitId) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) return createErrorResponse(new Error('BUSY'));
  try {
    MiraiAuth.requireTeacher();
    const ss = getSpreadsheet();
    const umSheet = ss.getSheetByName(DB_SCHEMA.UnitMaster.name);
    const rows = umSheet.getDataRange().getValues();

    const wsSheet = ss.getSheetByName(DB_SCHEMA.Worksheets.name);
    // 既存 Worksheets の taskId 集合
    const existing = {};
    if (wsSheet.getLastRow() >= 2) {
      const wsIds = wsSheet.getRange(2, WS_COL_TASK_ID, wsSheet.getLastRow() - 1, 1).getValues();
      wsIds.forEach(function (r) { if (r[0] !== '') existing[String(r[0])] = true; });
    }

    // 単元名を unitInfo から求める
    let unitName = '';
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][0]) === String(unitId)) {
        const info = safeJsonParse(rows[i][14]);
        unitName = info.unitName || info.title || String(rows[i][3] || '');
        if (unitName) break;
      }
    }

    const now = new Date();
    const inserts = [];
    const createdTaskIds = [];
    for (let i = 1; i < rows.length; i++) {
      const r = rows[i];
      if (String(r[0]) !== String(unitId)) continue;   // unitId 不一致
      if (String(r[13] || 'student') === 'teacher') continue; // 先生用タスクは対象外
      if (r[6]) continue;                               // deletedAt があれば除外
      const taskId = String(r[1]);
      if (!taskId || existing[taskId]) continue;        // 既に器がある

      const task = {
        taskId: taskId, unitId: String(r[0]), type: String(r[2]),
        title: String(r[3]), description: String(r[4]), format: String(r[13] || 'student')
      };
      inserts.push([
        taskId,                       // A: taskId（共通名前空間）
        unitName || String(r[3] || ''), // B: unitName
        String(r[3] || ''),           // C: stepTitle（タスク名）
        '',                           // D: htmlContent
        now,                          // E: lastUpdated
        JSON.stringify(task),         // F: jsonSource
        '',                           // G: canvasJson
        '',                           // H: rubricHtml
        false                         // I: isShared
      ]);
      existing[taskId] = true;
      createdTaskIds.push(taskId);
    }

    if (inserts.length > 0) {
      wsSheet.getRange(wsSheet.getLastRow() + 1, 1, inserts.length, WS_TOTAL_COLS).setValues(inserts);
    }
    return createSuccessResponse({ created: inserts.length, taskIds: createdTaskIds });
  } catch (e) {
    return createErrorResponse(e);
  } finally {
    lock.releaseLock();
  }
}

function updateUnitTotalHours(unitId, newTotalHours) {
  const lock = LockService.getScriptLock();
  if (lock.tryLock(5000)) {
    try {
      MiraiAuth.requireTeacher();
      const ss = getSpreadsheet();
      const sheet = ss.getSheetByName(DB_SCHEMA.UnitMaster.name);
      const data = sheet.getDataRange().getValues();

      for (let i = 1; i < data.length; i++) {
        if (String(data[i][0]) === String(unitId)) {
          sheet.getRange(i + 1, 16).setValue(newTotalHours);
          let info = safeJsonParse(data[i][14]);
          if (info) {
            info.totalHours = newTotalHours;
            sheet.getRange(i + 1, 15).setValue(JSON.stringify(info));
          }
        }
      }
      return createSuccessResponse({ message: '時数を更新しました' });
    } catch (e) { return createErrorResponse(e); } finally { lock.releaseLock(); }
  } else { return createErrorResponse(new Error("Timeout")); }
}


// [改修] 名簿保存処理をオブジェクト配列対応に変更

function saveClassRoster(classId, studentList) {
  const lock = LockService.getScriptLock();
  if (lock.tryLock(10000)) {
    try {
      MiraiAuth.requireTeacher();
      const ss = getSpreadsheet();
      const sheet = ss.getSheetByName(DB_SCHEMA.StudentRoster.name);
      const data = sheet.getDataRange().getValues();
      
      // 既存の該当クラスデータを削除（逆順ループで安全に削除）
      for (let i = data.length - 1; i >= 1; i--) {
        if (String(data[i][1]) === classId) { sheet.deleteRow(i + 1); }
      }
      
      // studentList: [{name: '...', number: 1, isActive: true, studentId: 'xxx@school.jp'(任意)}, ...]
      // 従来の文字列リストにも対応（後方互換性）
      const rows = studentList.map(s => {
        if (typeof s === 'string') {
          return [Utilities.getUuid(), classId, '', s, true, new Date(), ''];
        } else {
          return [Utilities.getUuid(), classId, s.number, s.name, s.isActive, new Date(), String(s.studentId || '')];
        }
      });
      if (rows.length > 0) {
        sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
      }
      return createSuccessResponse({ message: `${classId}の名簿を更新しました（${rows.length}名）` });
    } catch (e) { return createErrorResponse(e); } finally { lock.releaseLock(); }
  } else { return createErrorResponse(new Error("Timeout")); }
}

function saveSeatCoordinates(coordinates) {
  const lock = LockService.getScriptLock();
  if (lock.tryLock(5000)) {
    try {
      MiraiAuth.requireTeacher();
      const ss = getSpreadsheet();
      const sheet = ss.getSheetByName(DB_SCHEMA.LiveStatus.name);
      const data = sheet.getDataRange().getValues();
      // 行の特定は studentId（本人メール）優先。同名児童がいても座席が混ざらない。
      // studentId が無い座標（名簿のみの児童など）は従来どおり表示名で照合する。
      const sidToRow = new Map();
      const nameToRow = new Map();
      for (let i = 1; i < data.length; i++) {
        if (data[i][0]) sidToRow.set(String(data[i][0]), i + 1);
        if (!nameToRow.has(data[i][1])) nameToRow.set(data[i][1], i + 1);
      }

      coordinates.forEach(c => {
        const rIdx = (c.studentId && sidToRow.get(String(c.studentId))) || nameToRow.get(c.name);
        if (rIdx) { sheet.getRange(rIdx, 9, 1, 2).setValues([[c.x, c.y]]); }
      });
      invalidateLiveSnapshotCache();
      return createSuccessResponse();
    } catch(e) { return createErrorResponse(e); } finally { lock.releaseLock(); }
  } else { return createErrorResponse(new Error("Timeout")); }
}

function saveClassSchedule(scheduleData) {
  const lock = LockService.getScriptLock();
  if (lock.tryLock(5000)) {
    try {
      MiraiAuth.requireTeacher();
      const ss = getSpreadsheet();
      const sheet = ss.getSheetByName(DB_SCHEMA.ClassSchedule.name);
      sheet.appendRow([
        Utilities.getUuid(), scheduleData.classId, scheduleData.date,
        scheduleData.startTime, scheduleData.endTime,
        scheduleData.unitId, scheduleData.hour, scheduleData.message, new Date()
      ]);
      return createSuccessResponse();
    } catch(e) { return createErrorResponse(e); } finally { lock.releaseLock(); }
  } else { return createErrorResponse(new Error("Timeout")); }
}

// --- 生徒のアクション保存（課題追加、計画、振り返り） ---

function addMyTask(studentName, title, desc, time, unitId, classId) {
  const lock = LockService.getScriptLock();
  if (lock.tryLock(5000)) {
    try {
      const sid = MiraiAuth.requireUser();
      const ss = getSpreadsheet();
      const taskId = "MT" + Utilities.getUuid().substring(0, 8);
      // studentId 列（末尾）＝本人メール、studentName 列＝表示名。
      ss.getSheetByName(DB_SCHEMA.MyTasks.name).appendRow([taskId, studentName, title, desc, time, new Date(), unitId || "", classId || "", sid]);
      return createSuccessResponse({ taskId: taskId });
    } catch (e) { return createErrorResponse(e); } finally { lock.releaseLock(); }
  } else { return createErrorResponse(new Error("Timeout")); }
}

function saveStudentPlan(studentName, unitId, planData, classId) {
  const lock = LockService.getScriptLock();
  if (lock.tryLock(5000)) {
    try {
      const sid = MiraiAuth.requireUser();
      const ss = getSpreadsheet();
      const sheet = ss.getSheetByName(DB_SCHEMA.StudentPlans.name);
      const data = sheet.getDataRange().getValues();
      const json = JSON.stringify(planData);
      const now = new Date();

      // UPSERT キーは studentId（本人メール, col6/index5）＋unitId。同名衝突を避け、自分の行だけを更新する。
      let rowIndex = -1;
      for(let i = 1; i < data.length; i++) {
        if(data[i][5] === sid && data[i][1] === unitId) { rowIndex = i + 1; break; }
      }

      if(rowIndex > 0) {
        sheet.getRange(rowIndex, 3, 1, 2).setValues([[json, now]]);
        if(classId) sheet.getRange(rowIndex, 5).setValue(classId);
        // 表示名は最新の申告値に合わせておく（先生の名前グループ化用）。
        sheet.getRange(rowIndex, 1).setValue(studentName);
      } else {
        sheet.appendRow([studentName, unitId, json, now, classId || "", sid]);
      }
      return createSuccessResponse();
    } catch(e) { return createErrorResponse(e); } finally { lock.releaseLock(); }
  } else { return createErrorResponse(new Error("Timeout")); }
}

function saveDailyReflection(studentName, unitId, hour, achievement, comment, classId, skills) {
  const lock = LockService.getScriptLock();
  if (lock.tryLock(5000)) {
    try {
      const sid = MiraiAuth.requireUser();
      const ss = getSpreadsheet();
      const sheet = ss.getSheetByName(DB_SCHEMA.DailyReflections.name);
      const data = sheet.getDataRange().getValues();
      const now = new Date();
      const skillsJson = JSON.stringify(skills || []);

      // UPSERT キーは studentId（本人メール, col10/index9）＋unitId＋hour。
      let rowIndex = -1;
      for(let i = 1; i < data.length; i++) {
        if(data[i][9] === sid && data[i][1] === unitId && String(data[i][2]) === String(hour)) { rowIndex = i + 1; break; }
      }

      if(rowIndex > 0) {
        sheet.getRange(rowIndex, 4, 1, 2).setValues([[achievement, comment]]);
        sheet.getRange(rowIndex, 7).setValue(now);
        if(classId) sheet.getRange(rowIndex, 8).setValue(classId);
        sheet.getRange(rowIndex, 9).setValue(skillsJson);
        sheet.getRange(rowIndex, 1).setValue(studentName); // 表示名を最新に
      } else {
        sheet.appendRow([studentName, unitId, hour, achievement, comment, "", now, classId || "", skillsJson, sid]);
      }
      return createSuccessResponse();
    } catch(e) { return createErrorResponse(e); } finally { lock.releaseLock(); }
  } else { return createErrorResponse(new Error("Timeout")); }
}

function savePortfolio(studentName, unitId, summary, classId) {
  const lock = LockService.getScriptLock();
  if (lock.tryLock(5000)) {
    try {
      const sid = MiraiAuth.requireUser();
      const ss = getSpreadsheet();
      const sheet = ss.getSheetByName(DB_SCHEMA.Portfolios.name);
      const data = sheet.getDataRange().getValues();
      const now = new Date();

      // UPSERT キーは studentId（本人メール, col8/index7）＋unitId。
      let rowIndex = -1;
      for(let i = 1; i < data.length; i++) {
        if(data[i][7] === sid && String(data[i][1]) === String(unitId)) { rowIndex = i + 1; break; }
      }

      if(rowIndex > 0) {
        sheet.getRange(rowIndex, 3, 1, 2).setValues([[summary, now]]);
        if(classId) sheet.getRange(rowIndex, 5).setValue(classId);
        sheet.getRange(rowIndex, 1).setValue(studentName); // 表示名を最新に
      } else {
        // 末尾 studentId 列まで埋める（feedback/stamp は空、studentId はキー）
        sheet.appendRow([studentName, unitId, summary, now, classId || "", "", "", sid]);
      }
      return createSuccessResponse();
    } catch(e) { return createErrorResponse(e); } finally { lock.releaseLock(); }
  } else { return createErrorResponse(new Error("Timeout")); }
}

// --- フィードバック・評価 ---

function sendFeedback(studentName, taskId, stamp, classId, studentId) {
  const lock = LockService.getScriptLock();
  if (lock.tryLock(5000)) {
    try {
      MiraiAuth.requireTeacher();
      const ss = getSpreadsheet();
      // studentId（本人メール）が分かる場合は末尾列に残す。
      // 児童側の読み出しは「studentId があれば studentId でだけ照合」するため、
      // 同名児童がいてもスタンプが他人に誤配されなくなる（無い行は従来どおり名前照合）。
      ss.getSheetByName(DB_SCHEMA.Feedback.name).appendRow([
        Utilities.getUuid(), studentName, taskId, stamp, new Date(), classId || "", String(studentId || "")
      ]);
      return createSuccessResponse();
    } catch (e) { return createErrorResponse(e); } finally { lock.releaseLock(); }
  } else { return createErrorResponse(new Error("Timeout")); }
}

function savePortfolioFeedback(studentName, unitId, feedback, stamp, studentId) {
  const lock = LockService.getScriptLock();
  if (lock.tryLock(5000)) {
    try {
      MiraiAuth.requireTeacher();
      const ss = getSpreadsheet();
      const sheet = ss.getSheetByName(DB_SCHEMA.Portfolios.name);
      const data = sheet.getDataRange().getValues();
      // 本人の studentId が渡されていれば、生徒が書いたその行に確実に着地させる（同名衝突を回避）。
      // 渡されていない場合のみ、従来どおり studentName で照合（フォールバック）。
      const sid = String(studentId || "");
      let rowIndex = -1;
      for(let i = 1; i < data.length; i++) {
        const match = sid
          ? (data[i][7] === sid && String(data[i][1]) === String(unitId))
          : (data[i][0] === studentName && String(data[i][1]) === String(unitId));
        if(match) { rowIndex = i + 1; break; }
      }

      if(rowIndex > 0) {
        sheet.getRange(rowIndex, 6, 1, 2).setValues([[feedback, stamp]]);
      } else {
        const now = new Date();
        // 末尾 studentId 列まで埋める（生徒が未提出でも本人メールを残しておく）
        sheet.appendRow([studentName, unitId, "", now, "", feedback, stamp, sid]);
      }
      return createSuccessResponse();
    } catch(e) { return createErrorResponse(e); } finally { lock.releaseLock(); }
  } else { return createErrorResponse(new Error("Timeout")); }
}

function saveAllPortfolios(feedbackList) {
  const lock = LockService.getScriptLock();
  if (lock.tryLock(10000)) {
    try {
      MiraiAuth.requireTeacher();
      const ss = getSpreadsheet();
      const sheet = ss.getSheetByName(DB_SCHEMA.Portfolios.name);
      const data = sheet.getDataRange().getValues();
      // studentId（本人メール, col7）と studentName（col0）の両方でインデックスを作り、
      // studentId が来ていればそれを優先して本人の行に着地させる（同名衝突を回避）。
      const idMap = new Map();
      const nameMap = new Map();
      for (let i = 1; i < data.length; i++) {
        if (data[i][7]) idMap.set(data[i][7] + "_" + data[i][1], i);
        nameMap.set(data[i][0] + "_" + data[i][1], i);
      }

      const rowsToAppend = [];
      const now = new Date();

      feedbackList.forEach(item => {
        const sid = String(item.studentId || "");
        const idKey = sid + "_" + item.unitId;
        const nameKey = item.studentName + "_" + item.unitId;
        let rowIndex = -1;
        if (sid && idMap.has(idKey)) rowIndex = idMap.get(idKey);
        else if (!sid && nameMap.has(nameKey)) rowIndex = nameMap.get(nameKey);
        if (rowIndex >= 0) {
          data[rowIndex][5] = item.feedback;
          data[rowIndex][6] = item.stamp;
        } else {
          rowsToAppend.push([item.studentName, item.unitId, "", now, "", item.feedback, item.stamp, sid]);
        }
      });

      if (data.length > 1) {
        sheet.getRange(1, 1, data.length, data[0].length).setValues(data);
      }
      if (rowsToAppend.length > 0) {
        sheet.getRange(sheet.getLastRow() + 1, 1, rowsToAppend.length, rowsToAppend[0].length).setValues(rowsToAppend);
      }
      return createSuccessResponse({ message: '一括保存しました' });
    } catch (e) { return createErrorResponse(e); } finally { lock.releaseLock(); }
  } else { return createErrorResponse(new Error("Timeout")); }
}

// ==========================================
//  5. 高度な機能 (Advanced Features)
// ==========================================

// JSONインポート、アーカイブなど

function importUnitJson(jsonStr) {
  const lock = LockService.getScriptLock();
  if (lock.tryLock(10000)) {
    try {
      MiraiAuth.requireTeacher();
      const ssId = PROPERTIES.getProperty('SS_ID');
      if (!ssId) throw new Error('初期設定未完了');
      const ss = SpreadsheetApp.openById(ssId);
      checkAndFixSheets(ss);

      const data = JSON.parse(jsonStr);
      if (!data.unitInfo) data.unitInfo = {};
      if (!data.unitInfo.title && data.unitInfo.unitName) data.unitInfo.title = data.unitInfo.unitName;
      if (data.unitInfo.grade && typeof data.unitInfo.grade !== 'string') data.unitInfo.grade = String(data.unitInfo.grade);
      if (!data.unitInfo.title) data.unitInfo.title = "無題の単元";

      // 単元IDは createNewUnit と同じ14桁（yyyyMMddHHmmss）に、衝突防止のランダム接尾辞を足す。
      // 分単位（yyyyMMddHHmm）だと同一分内の連続インポートでIDが衝突し、無関係なタスクが混ざる不具合があった。
      const rand = Utilities.getUuid().replace(/-/g, '').substring(0, 6);
      const uid = "U" + Utilities.formatDate(new Date(), "JST", "yyyyMMddHHmmss") + rand;
      const uInfoStr = JSON.stringify(data.unitInfo || {});
      const totalHours = data.unitInfo?.totalHours || 8;

      // タスクIDは単元IDを接頭辞にして全体で一意にする。
      // AI出力やテンプレートは "t_01" のような固定IDを使うため、そのまま保存すると
      // 別単元間で衝突し、進捗・スタンプ・ワークシート（taskId で照合）が混線する。
      // prerequisites も同じ対応表で貼り替える。
      const idMap = {};
      (data.tasks || []).forEach((t, i) => {
        const orig = String(t.id || ('t_' + (i + 1)));
        idMap[orig] = uid + '_' + orig;
      });
      const rows = (data.tasks || []).map((t, i) => {
        const orig = String(t.id || ('t_' + (i + 1)));
        const prereqs = (t.prerequisites || []).map(p => idMap[String(p)] || String(p));
        return [
          uid, idMap[orig], t.type || 'must', t.title || '無題', t.description || '', t.estimatedTime || 10, '',
          t.category || 'まなぶ', t.step || '', t.textbook || '', t.tablet || '', t.print || '',
          prereqs.join(','), t.format || 'student', uInfoStr, totalHours
        ];
      });

      if (rows.length > 0) {
        const sheet = ss.getSheetByName(DB_SCHEMA.UnitMaster.name);
        sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
      }
      return createSuccessResponse({ count: rows.length });
    } catch (e) { return createErrorResponse(e); } finally { lock.releaseLock(); }
  } else { return createErrorResponse(new Error("Timeout")); }
}

function archiveUnitData(unitId, unitTitle) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) { return createErrorResponse(new Error("Timeout")); }
  try {
    MiraiAuth.requireTeacher();
    const ss = getSpreadsheet();
    const archiveName = `アーカイブ_${unitTitle || unitId}_${Utilities.formatDate(new Date(), 'JST', 'yyyyMMdd')}`;
    const archiveSs = SpreadsheetApp.create(archiveName);
    
    const targets = [
      { key: 'MyTasks', colUnitId: 6 }, { key: 'StudentPlans', colUnitId: 1 },
      { key: 'DailyReflections', colUnitId: 1 }, { key: 'Portfolios', colUnitId: 1 }
    ];
    let movedCount = 0;
    
    targets.forEach(t => {
      const sheet = ss.getSheetByName(DB_SCHEMA[t.key].name);
      if(!sheet) return;
      const data = sheet.getDataRange().getValues();
      const headers = data[0];
      const rows = data.slice(1);
      const toArchive = []; const toKeep = [];
      
      rows.forEach(row => {
        if (String(row[t.colUnitId]) === String(unitId)) { toArchive.push(row); } 
        else { toKeep.push(row); }
      });
      
      if (toArchive.length > 0) {
        let archSheet = archiveSs.getSheetByName(DB_SCHEMA[t.key].name);
        if (!archSheet) { archSheet = archiveSs.insertSheet(DB_SCHEMA[t.key].name); archSheet.appendRow(headers); }
        archSheet.getRange(archSheet.getLastRow() + 1, 1, toArchive.length, toArchive[0].length).setValues(toArchive);
        
        sheet.clearContents();
        sheet.appendRow(headers);
        if (toKeep.length > 0) { sheet.getRange(2, 1, toKeep.length, toKeep[0].length).setValues(toKeep); }
        movedCount += toArchive.length;
      }
    });

    // UnitMasterのアーカイブ
    const uSheet = ss.getSheetByName(DB_SCHEMA.UnitMaster.name);
    const uData = uSheet.getDataRange().getValues();
    const uKeep = []; const uArch = [];
    uData.slice(1).forEach(row => {
      if (String(row[0]) === String(unitId)) uArch.push(row);
      else uKeep.push(row);
    });
    
    if (uArch.length > 0) {
      let uArchSheet = archiveSs.getSheetByName(DB_SCHEMA.UnitMaster.name);
      if(!uArchSheet) { uArchSheet = archiveSs.insertSheet(DB_SCHEMA.UnitMaster.name); uArchSheet.appendRow(uData[0]); }
      uArchSheet.getRange(uArchSheet.getLastRow()+1, 1, uArch.length, uArch[0].length).setValues(uArch);
      
      uSheet.clearContents();
      uSheet.appendRow(uData[0]);
      if(uKeep.length > 0) uSheet.getRange(2, 1, uKeep.length, uKeep[0].length).setValues(uKeep);
    }
    
    const delSheet = archiveSs.getSheetByName('シート1');
    if(delSheet) archiveSs.deleteSheet(delSheet);

    return createSuccessResponse({
      message: `アーカイブ完了: ${movedCount}件のデータを移動しました。\nファイル名: ${archiveName}`,
      url: archiveSs.getUrl()
    });
  } catch(e) { return createErrorResponse(e); } finally { lock.releaseLock(); }
}

// --- みらいパスポート連携 & AIプロンプト ---

function getCustomAiPrompt() {
  try {
    const prompt = PROPERTIES.getProperty('CUSTOM_AI_PROMPT');
    return createSuccessResponse({ prompt: prompt });
  } catch (e) { return createErrorResponse(e); }
}

function saveCustomAiPrompt(text) {
  try {
    MiraiAuth.requireTeacher();
    PROPERTIES.setProperty('CUSTOM_AI_PROMPT', text);
    return createSuccessResponse();
  } catch (e) { return createErrorResponse(e); }
}

// ==========================================
//  5b. ワークシート/回答/AI（統合: 旧みらいパスポート機能を内部化）
// ==========================================

/* ---------- AI 設定（Gemini APIキーは ScriptProperties に school-level で保持） ---------- */

// サーバー内部だけで APIキーを読む（クライアントには絶対に返さない）。
function getGeminiApiKey_() {
  return PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY') || '';
}

// 設定済みかどうかと先生名だけを返す（生のキーは返さない）。要ログイン。
function getAiConfig() {
  MiraiAuth.requireUser();
  return {
    hasApiKey: !!getGeminiApiKey_(),
    teacherName: PropertiesService.getScriptProperties().getProperty('TEACHER_NAME') || ''
  };
}

// Gemini APIキー・先生名を保存（先生専用）。キーは入力があったときだけ上書き（空欄＝変更なし）。
function saveAiConfig(apiKey, teacherName) {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) return createErrorResponse(new Error('BUSY'));
  try {
    MiraiAuth.requireTeacher();
    var props = PropertiesService.getScriptProperties();
    var toSet = {};
    if (teacherName !== undefined && teacherName !== null) toSet['TEACHER_NAME'] = String(teacherName);
    if (apiKey !== undefined && apiKey !== null && String(apiKey).trim() !== '') {
      toSet['GEMINI_API_KEY'] = String(apiKey).trim();
    }
    if (Object.keys(toSet).length) props.setProperties(toSet);
    return { success: true };
  } catch (e) {
    return createErrorResponse(e);
  } finally {
    lock.releaseLock();
  }
}

/* ---------- ワークシート（Worksheets シート） ---------- */

// ワークシートを保存（新規 or taskId 一致で上書き）。先生専用。
function saveWorksheetToDB(data) {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) return createErrorResponse(new Error('BUSY'));
  try {
    MiraiAuth.requireTeacher();
    var sheet = getSpreadsheet().getSheetByName(DB_SCHEMA.Worksheets.name);
    var now = new Date();
    var taskId = String(data.taskId || Utilities.getUuid());
    var unitName = data.unitName || '無題';
    var stepTitle = data.stepTitle || '無題';
    var htmlContent = data.htmlContent || '';
    var jsonSource = JSON.stringify(data.jsonSource || {});
    var canvasJson = data.canvasJson ? JSON.stringify(data.canvasJson) : '';
    var rubricHtml = data.rubricHtml || '';
    var isShared = data.isShared || false;

    var found = sheet.getRange('A:A').createTextFinder(taskId).matchEntireCell(true).findNext();
    if (found) {
      var row = found.getRow();
      sheet.getRange(row, WS_COL_UNIT_NAME, 1, 8).setValues([[
        unitName, stepTitle, htmlContent, now, jsonSource, canvasJson, rubricHtml, isShared
      ]]);
    } else {
      sheet.appendRow([taskId, unitName, stepTitle, htmlContent, now, jsonSource, canvasJson, rubricHtml, isShared]);
    }
    return true;
  } catch (e) {
    return createErrorResponse(e);
  } finally {
    lock.releaseLock();
  }
}

// 指定 taskId のワークシートを読む。要ログイン。
function loadWorksheetFromDB(taskId) {
  MiraiAuth.requireUser();
  var sheet = getSpreadsheet().getSheetByName(DB_SCHEMA.Worksheets.name);
  var found = sheet.getRange('A:A').createTextFinder(String(taskId)).matchEntireCell(true).findNext();
  if (!found) return null;
  var row = found.getRow();
  var values = sheet.getRange(row, 1, 1, WS_TOTAL_COLS).getValues()[0];
  return {
    taskId: values[0], unitName: values[1], stepTitle: values[2], htmlContent: values[3],
    jsonSource: safeJSONParse(values[5]), canvasJson: safeJSONParse(values[6]),
    rubricHtml: values[7], isShared: values[8]
  };
}

// 複数 taskId のワークシートをまとめて取得（一括生成・印刷用）。要ログイン。
function getWorksheetsByIds(taskIds) {
  MiraiAuth.requireUser();
  var sheet = getSpreadsheet().getSheetByName(DB_SCHEMA.Worksheets.name);
  var data = sheet.getDataRange().getValues();
  return data.slice(1)
    .filter(function (row) { return taskIds.includes(String(row[0])); })
    .map(function (row) {
      return {
        taskId: row[0], unitName: row[1], stepTitle: row[2], htmlContent: row[3],
        canvasJson: safeJSONParse(row[6]), jsonSource: safeJSONParse(row[5])
      };
    });
}

// 保存済みワークシートの履歴（最新30件）。要ログイン。
function getHistory() {
  MiraiAuth.requireUser();
  var sheet = getSpreadsheet().getSheetByName(DB_SCHEMA.Worksheets.name);
  if (sheet.getLastRow() < 2) return [];
  var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 5).getValues();
  return data
    .map(function (r) { return { id: r[0], title: r[2] || '無題', timestamp: new Date(r[4]).getTime() }; })
    .filter(function (item) { return item.id; })
    .sort(function (a, b) { return b.timestamp - a.timestamp; })
    .slice(0, 30);
}

// 児童向け: 配信済みワークシート一覧（軽量版・HTML本文なし）。要ログイン。
function getStudentWorksheetList() {
  MiraiAuth.requireUser();
  var sheet = getSpreadsheet().getSheetByName(DB_SCHEMA.Worksheets.name);
  if (sheet.getLastRow() < 2) return [];
  var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 3).getValues();
  return data
    .filter(function (row) { return row[0]; })
    .map(function (row) { return { taskId: String(row[0]), unitName: String(row[1]), stepTitle: String(row[2]) }; });
}

/* ---------- 児童レスポンス（Responses シート） ---------- */

// 本人の全提出データをワークシート情報と結合して返す（旧 getPortfolioData）。
// 本人性はサーバーが決める（クライアントの studentId は受け取らない）。
function getMyWorksheetSubmissions() {
  var me = MiraiAuth.requireUser();
  var ss = getSpreadsheet();

  var wsSheet = ss.getSheetByName(DB_SCHEMA.Worksheets.name);
  var worksheetsMap = {};
  if (wsSheet.getLastRow() >= 2) {
    var wsData = wsSheet.getRange(2, 1, wsSheet.getLastRow() - 1, WS_TOTAL_COLS).getValues();
    wsData.forEach(function (r) {
      if (!r[0]) return;
      worksheetsMap[String(r[0])] = { unitName: String(r[1]), stepTitle: String(r[2]), htmlContent: String(r[3]) };
    });
  }

  var resSheet = ss.getSheetByName(DB_SCHEMA.Responses.name);
  var portfolio = [];
  if (resSheet.getLastRow() >= 2) {
    var resData = resSheet.getDataRange().getValues();
    for (var i = 1; i < resData.length; i++) {
      var row = resData[i];
      if (String(row[RS_COL_STUDENT_ID - 1]) === me) {
        var taskId = String(row[RS_COL_TASK_ID - 1]);
        var wi = worksheetsMap[taskId] || { unitName: '不明', stepTitle: '不明', htmlContent: '' };
        portfolio.push({
          responseId: row[RS_COL_RESPONSE_ID - 1],
          taskId: taskId,
          submittedAt: row[RS_COL_SUBMITTED_AT - 1] ? new Date(row[RS_COL_SUBMITTED_AT - 1]).getTime() : 0,
          canvasImage: row[RS_COL_CANVAS_IMAGE - 1],
          textContent: row[RS_COL_TEXT_CONTENT - 1],
          status: row[RS_COL_STATUS - 1],
          feedbackText: row[RS_COL_FEEDBACK_TXT - 1],
          canvasJson: row[RS_COL_CANVAS_JSON - 1],
          reflectionText: row[RS_COL_REFLECTION - 1] || '',
          unitName: wi.unitName, stepTitle: wi.stepTitle, htmlContent: wi.htmlContent
        });
      }
    }
  }
  return portfolio.sort(function (a, b) { return b.submittedAt - a.submittedAt; });
}

// 児童の振り返りを保存。本人の回答のみ更新可（IDOR 防止）。
function saveStudentReflection(data) {
  var me = MiraiAuth.requireUser();
  if (!data || !data.responseId || data.reflectionText === undefined || data.reflectionText === null) {
    return { success: false, message: 'パラメータが不足しています。' };
  }
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) return createErrorResponse(new Error('BUSY'));
  try {
    var sheet = getSpreadsheet().getSheetByName(DB_SCHEMA.Responses.name);
    var finder = sheet.getRange('A:A').createTextFinder(data.responseId).matchEntireCell(true).findNext();
    if (finder) {
      var row = finder.getRow();
      var owner = String(sheet.getRange(row, RS_COL_STUDENT_ID).getValue());
      if (owner !== me) return { success: false, message: '対象の回答が見つかりません。' };
      sheet.getRange(row, RS_COL_REFLECTION).setValue(data.reflectionText);
      return { success: true };
    }
    return { success: false, message: '対象の回答が見つかりません。' };
  } finally {
    lock.releaseLock();
  }
}

// 児童の回答を保存（taskId × 本人 で新規 or 上書き）。studentId は必ずサーバー判定。
function saveStudentResponse(data) {
  var me = MiraiAuth.requireUser();
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) return createErrorResponse(new Error('BUSY'));
  try {
    var sheet = getSpreadsheet().getSheetByName(DB_SCHEMA.Responses.name);
    var now = new Date();
    var lastRow = sheet.getLastRow();
    var existingRow = -1;
    if (lastRow >= 2) {
      var keys = sheet.getRange(2, RS_COL_TASK_ID, lastRow - 1, 2).getValues();
      for (var i = 0; i < keys.length; i++) {
        if (String(keys[i][0]) === String(data.taskId) && String(keys[i][1]) === me) {
          existingRow = i + 2;
          break;
        }
      }
    }
    var isPublicVal = (data.isPublic === undefined) ? true : data.isPublic;
    if (existingRow > 0) {
      sheet.getRange(existingRow, RS_COL_STUDENT_NAME, 1, 5).setValues([[
        data.studentName, now, data.canvasImage, data.textContent, data.status
      ]]);
      if (data.canvasJson) sheet.getRange(existingRow, RS_COL_CANVAS_JSON).setValue(data.canvasJson);
      sheet.getRange(existingRow, RS_COL_IS_PUBLIC).setValue(isPublicVal);
      if (data.answersJson !== undefined && data.answersJson !== null) {
        sheet.getRange(existingRow, RS_COL_ANSWERS_JSON).setValue(data.answersJson);
      }
    } else {
      sheet.appendRow([
        Utilities.getUuid(), data.taskId, me, data.studentName, now,
        data.canvasImage || '', data.textContent || '', data.status || 'submitted',
        '', '', '', data.canvasJson || '', isPublicVal, '[]', '', data.answersJson || ''
      ]);
    }
    return { success: true };
  } finally {
    lock.releaseLock();
  }
}

// 本人の最新回答を取得（復元用）。taskId のみ受け取り、本人性はサーバー判定。
function getMyResponse(taskId) {
  var me = MiraiAuth.requireUser();
  var sheet = getSpreadsheet().getSheetByName(DB_SCHEMA.Responses.name);
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return null;
  var keys = sheet.getRange(2, RS_COL_TASK_ID, lastRow - 1, 2).getValues();
  for (var i = keys.length - 1; i >= 0; i--) {
    if (String(keys[i][0]) === String(taskId) && String(keys[i][1]) === me) {
      var row = sheet.getRange(i + 2, 1, 1, RS_TOTAL_COLS).getValues()[0];
      return {
        responseId: row[RS_COL_RESPONSE_ID - 1],
        status: row[RS_COL_STATUS - 1],
        feedbackText: row[RS_COL_FEEDBACK_TXT - 1],
        canvasImage: row[RS_COL_CANVAS_IMAGE - 1],
        canvasJson: row[RS_COL_CANVAS_JSON - 1],
        isPublic: row[RS_COL_IS_PUBLIC - 1],
        reactions: ensureArray(safeJSONParse(row[RS_COL_REACTIONS - 1])),
        reflectionText: row[RS_COL_REFLECTION - 1] || '',
        answersJson: row[RS_COL_ANSWERS_JSON - 1] || ''
      };
    }
  }
  return null;
}

// 指定タスクの全提出データ（先生用ダッシュボード）。先生専用。
function getTaskSubmissions(taskId) {
  MiraiAuth.requireTeacher();
  var sheet = getSpreadsheet().getSheetByName(DB_SCHEMA.Responses.name);
  var values = sheet.getDataRange().getValues();
  var results = [];
  for (var i = 1; i < values.length; i++) {
    if (String(values[i][1]) === String(taskId)) {
      results.push({
        rowIndex: i + 1,
        studentId: values[i][2],
        studentName: values[i][3],
        submittedAt: values[i][4],
        canvasImage: values[i][5],
        status: values[i][7],
        feedbackText: values[i][8],
        teacherCanvasJson: values[i][RS_COL_FEEDBACK_JSON - 1],
        canvasJson: values[i][11]
      });
    }
  }
  return results;
}

// 指定行の提出データから canvasJson とワークシートHTMLを取得（添削プレビュー用）。先生専用。
function getSubmissionDetail(rowIndex) {
  MiraiAuth.requireTeacher();
  var ss = getSpreadsheet();
  var resSheet = ss.getSheetByName(DB_SCHEMA.Responses.name);
  var taskId = String(resSheet.getRange(rowIndex, RS_COL_TASK_ID).getValue());
  var canvasJson = resSheet.getRange(rowIndex, RS_COL_CANVAS_JSON).getValue();
  var wsSheet = ss.getSheetByName(DB_SCHEMA.Worksheets.name);
  var htmlContent = '';
  var found = wsSheet.getRange('A:A').createTextFinder(taskId).matchEntireCell(true).findNext();
  if (found) htmlContent = wsSheet.getRange(found.getRow(), WS_COL_HTML_CONTENT).getValue();
  return { canvasJson: safeJSONParse(canvasJson), htmlContent: htmlContent };
}

// 全提出データ + ワークシート一覧を一括取得（先生用管理画面）。先生専用。
function getDashboardData() {
  MiraiAuth.requireTeacher();
  var ss = getSpreadsheet();
  var wsSheet = ss.getSheetByName(DB_SCHEMA.Worksheets.name);
  var worksheets = [];
  if (wsSheet.getLastRow() >= 2) {
    var wsData = wsSheet.getRange(2, 1, wsSheet.getLastRow() - 1, 3).getValues();
    worksheets = wsData
      .filter(function (r) { return r[0]; })
      .map(function (r) { return { taskId: String(r[0]), unitName: String(r[1]), stepTitle: String(r[2]) }; });
  }
  var resSheet = ss.getSheetByName(DB_SCHEMA.Responses.name);
  var submissions = [];
  if (resSheet.getLastRow() >= 2) {
    var resData = resSheet.getDataRange().getValues();
    for (var i = 1; i < resData.length; i++) {
      var row = resData[i];
      if (!row[0]) continue;
      submissions.push({
        rowIndex: i + 1,
        responseId: row[0],
        taskId: String(row[1]),
        studentId: row[2],
        studentName: row[3],
        submittedAt: row[4] ? new Date(row[4]).getTime() : 0,
        canvasImage: row[5],
        textContent: row[6],
        status: row[7],
        feedbackText: row[8]
      });
    }
  }
  return { submissions: submissions, worksheets: worksheets };
}

// 児童の回答に対する添削（コメント + 赤ペン K列）を保存。先生専用。
function saveFeedback(data) {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) return createErrorResponse(new Error('BUSY'));
  try {
    MiraiAuth.requireTeacher();
    var sheet = getSpreadsheet().getSheetByName(DB_SCHEMA.Responses.name);
    if (data.rowIndex) {
      sheet.getRange(data.rowIndex, RS_COL_STATUS).setValue('graded');
      sheet.getRange(data.rowIndex, RS_COL_FEEDBACK_TXT).setValue(data.feedbackText);
      if (data.canvasJson) sheet.getRange(data.rowIndex, RS_COL_FEEDBACK_JSON).setValue(data.canvasJson);
    }
    return { success: true };
  } finally {
    lock.releaseLock();
  }
}

// 複数提出物へのコメントを一括保存。先生専用。
function batchSaveFeedback(feedbacks) {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) return createErrorResponse(new Error('BUSY'));
  try {
    MiraiAuth.requireTeacher();
    var sheet = getSpreadsheet().getSheetByName(DB_SCHEMA.Responses.name);
    for (var i = 0; i < feedbacks.length; i++) {
      var fb = feedbacks[i];
      if (fb.rowIndex && fb.feedbackText) {
        sheet.getRange(fb.rowIndex, RS_COL_STATUS).setValue('graded');
        sheet.getRange(fb.rowIndex, RS_COL_FEEDBACK_TXT).setValue(fb.feedbackText);
      }
    }
    return { success: true, count: feedbacks.length };
  } finally {
    lock.releaseLock();
  }
}

// 広場に公開されている他児童の回答を取得。要ログイン。
// 自分の作品は除外し、他児童の studentId（メール）は payload に含めない（プライバシー保護）。
function getSharedResponses(taskId) {
  var me = MiraiAuth.requireUser();
  var sheet = getSpreadsheet().getSheetByName(DB_SCHEMA.Responses.name);
  var values = sheet.getDataRange().getValues();
  var results = [];
  for (var i = 1; i < values.length; i++) {
    var row = values[i];
    var isPublic = (row[12] === '' || row[12] === true || row[12] === 'true');
    if (String(row[1]) === String(taskId) &&
        (row[7] === 'submitted' || row[7] === 'graded') &&
        isPublic) {
      if (me && String(row[2]) === me) continue; // 自分の作品はサーバー側で除外
      results.push({
        responseId: row[0],
        studentName: row[3],
        canvasImage: row[5],
        canvasJson: row[11],
        reactions: ensureArray(safeJSONParse(row[13]))
      });
    }
  }
  return results;
}

// 友達の作品にリアクションを送る。要ログイン。
function savePeerReaction(data) {
  MiraiAuth.requireUser();
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) return createErrorResponse(new Error('BUSY'));
  try {
    var sheet = getSpreadsheet().getSheetByName(DB_SCHEMA.Responses.name);
    var finder = sheet.getRange('A:A').createTextFinder(data.targetResponseId).matchEntireCell(true).findNext();
    if (finder) {
      var row = finder.getRow();
      var cell = sheet.getRange(row, RS_COL_REACTIONS);
      var current = ensureArray(safeJSONParse(cell.getValue()));
      data.reaction.timestamp = new Date().getTime();
      current.push(data.reaction);
      cell.setValue(JSON.stringify(current));
      return { success: true, reactions: current };
    }
    return { success: false };
  } finally {
    lock.releaseLock();
  }
}

/* ---------- AI 連携（Gemini API） ---------- */

// Gemini API を呼び出しテキストを返す（内部関数）。キーは x-goog-api-key ヘッダで送る。
function callGeminiAPI(prompt) {
  // トップレベル関数は google.script.run から誰でも呼べるため、AI 実行は先生に限定する
  // （Gemini 枠の悪用・任意プロンプト実行の防止）。内部の generate* からは先生文脈で呼ばれる。
  MiraiAuth.requireTeacher();
  var apiKey = getGeminiApiKey_();
  if (!apiKey) {
    throw new Error('Gemini APIキーが設定されていません。先生モードの設定を確認してください。');
  }
  var url = 'https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent';
  var payload = { contents: [{ parts: [{ text: prompt }] }] };
  var options = {
    method: 'post',
    contentType: 'application/json',
    headers: { 'x-goog-api-key': apiKey },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };
  var res = UrlFetchApp.fetch(url, options);
  var json = JSON.parse(res.getContentText());
  if (json.error) {
    throw new Error('AIエラー: ' + json.error.message);
  }
  if (!json.candidates || !json.candidates.length ||
      !json.candidates[0].content ||
      !json.candidates[0].content.parts ||
      !json.candidates[0].content.parts.length) {
    throw new Error('AIから応答が得られませんでした。');
  }
  return json.candidates[0].content.parts[0].text;
}

// 統一ワークシート生成プロンプトを構築する（純関数）。
function buildWorksheetPrompt(data) {
  var grade       = data.grade       || '';
  var unitName    = data.unitName    || '';
  var stepTitle   = data.stepTitle   || '';
  var description = data.description || '';
  var ocrContext  = data.ocrContext  || '';

  var ocrSection = '';
  if (ocrContext) {
    ocrSection = '\n【参考資料テキスト】\n' + ocrContext + '\n※この資料の内容を授業に反映させてください。\n';
  }

  return 'あなたは「教育工学」と「クリエイティブ・コーディング」に精通したフルスタックエンジニアです。\n'
    + '日本の小学校の授業で使う、高品質なワークシートのHTMLを生成してください。\n\n'
    + '【授業情報】\n'
    + '学年: ' + grade + '\n'
    + '単元名: ' + unitName + '\n'
    + '授業タイトル: ' + stepTitle + '\n'
    + '活動内容: ' + description + '\n'
    + ocrSection + '\n'
    + '【出力形式の制約（厳守）】\n'
    + '- HTMLの body 内部のみを出力すること。<!DOCTYPE>, <html>, <head>, <body> タグは不要。\n'
    + '- Markdown記法は禁止。```html ブロックで囲まないこと。\n'
    + '- 外部リソース（img src="https://...", CDN, 外部ライブラリ）は一切使用禁止。\n'
    + '- すべて標準API（HTML, CSS, インラインSVG, JavaScript）のみで完結させること。\n'
    + '- via.placeholder.com 等の外部画像サービスも禁止。\n\n'
    + '【HTMLレイアウト構造（この構造を厳守すること）】\n'
    + '<div class="ws-sheet">\n'
    + '  <div class="ws-header-fixed">\n'
    + '    <div class="ws-header-left">\n'
    + '      <span class="ws-unit-name">' + grade + ' ' + unitName + '</span>\n'
    + '      <h1 class="ws-title">' + stepTitle + '</h1>\n'
    + '    </div>\n'
    + '    <table class="ws-meta-table">\n'
    + '      <tr><td class="ws-meta-label">年</td><td class="ws-meta-input"></td>\n'
    + '          <td class="ws-meta-label">組</td><td class="ws-meta-input"></td>\n'
    + '          <td class="ws-meta-label">名前</td><td class="ws-meta-input" style="min-width:120px;"></td></tr>\n'
    + '    </table>\n'
    + '  </div>\n\n'
    + '  <div class="ws-content">\n'
    + '  </div>\n\n'
    + '  <div class="ws-footer-fixed">\n'
    + '    <div class="ws-assessment-grid">\n'
    + '      <div class="ws-reflection-box">\n'
    + '        <span class="ws-reflection-title">ふりかえり</span>\n'
    + '        <div class="ws-box ws-lines" style="height:4.5em; background-image:linear-gradient(#ccc 1px, transparent 1px); background-size:100% 1.5em;"></div>\n'
    + '      </div>\n'
    + '      <div>\n'
    + '        <table class="table table-bordered table-sm mb-0" style="font-size:0.85em; text-align:center;">\n'
    + '          <tr><td class="bg-light" style="width:40%;">わかった</td><td><button type="button" class="eval-btn" data-value="△">△</button><button type="button" class="eval-btn" data-value="◯">◯</button><button type="button" class="eval-btn" data-value="◎">◎</button></td></tr>\n'
    + '          <tr><td class="bg-light">考えた</td><td><button type="button" class="eval-btn" data-value="△">△</button><button type="button" class="eval-btn" data-value="◯">◯</button><button type="button" class="eval-btn" data-value="◎">◎</button></td></tr>\n'
    + '          <tr><td class="bg-light">進んで取り組んだ</td><td><button type="button" class="eval-btn" data-value="△">△</button><button type="button" class="eval-btn" data-value="◯">◯</button><button type="button" class="eval-btn" data-value="◎">◎</button></td></tr>\n'
    + '        </table>\n'
    + '      </div>\n'
    + '    </div>\n'
    + '  </div>\n'
    + '</div>\n\n'
    + '【ws-content 内に含めるセクション】\n'
    + '1. 今日のめあて: 背景 #e3f2fd, 左ボーダー #2196f3, border-radius:8px のボックス。活動内容から子供向けのめあてを生成。\n'
    + '2. AIコーチのヒント: 背景 #fff3e0, 左ボーダー #ff9800, border-radius:8px。つまずきやすい点や考えるコツを1-2行で。\n'
    + '3. 学習課題・問題: 問題文は通常の div/p で記述（ws-box を付けない）。\n'
    + '4. 記述欄・解答欄: 児童が書き込む欄には class="ws-box" を付ける。罫線付きは class="ws-box ws-lines" + background-image で罫線を描画。\n\n'
    + '【利用可能なCSSクラス一覧】\n'
    + 'レイアウト: ws-sheet(flex column), ws-header-fixed, ws-header-left, ws-content(flex:1), ws-footer-fixed(margin-top:auto)\n'
    + 'テキスト: ws-title(1.3rem bold), ws-unit-name(badge風), ws-date-small(0.75rem)\n'
    + 'メタ情報: ws-meta-table, ws-meta-label(背景#eee), ws-meta-input(入力欄)\n'
    + '記述欄: ws-box(リサイズ可能な入力ボックス), ws-lines(罫線付き), ws-instruction(指示文ボックス)\n'
    + '評価: ws-assessment-grid(grid 2fr 1fr), ws-reflection-box, ws-reflection-title, eval-btn(◎○△ボタン)\n'
    + '教科別: grid-paper(方眼紙40px), graph-paper(グラフ用紙20px), math-grid/math-cell/math-line(算数マス目), mode-kokugo(国語縦書き)\n'
    + 'Bootstrap 5: p-3, mb-3, bg-light, table, table-bordered, card, badge 等のユーティリティクラスも使用可\n\n'
    + '【図・グラフの描画技術】\n'
    + 'テーマに応じて最適な描画技術を選択し、視覚的に美しい図版を積極的に生成すること:\n'
    + '- インラインSVG（推奨）: グラフ（棒・折れ線・円）、座標平面、地図、図形、フローチャート、イラスト。viewBox で A4幅に収まるサイズに。\n'
    + '- CSS Art: 単純な図形（円、三角形、矢印）、実験器具のアイコン。\n'
    + '  例: <div style="width:80px;height:80px;border-radius:50%;border:2px solid #333;"></div>\n'
    + '- HTML table: 表、時間割、比較表。\n'
    + '- 再帰/フラクタル: 自然物（木、雪の結晶）の描画にはSVG + JavaScript。\n'
    + '- 数式/三角関数: 周期的な動き、波形、天体の軌道はSVG path + Math.sin/cos。\n\n'
    + 'SVGの具体例（棒グラフ）:\n'
    + '<svg viewBox="0 0 300 200" style="width:100%;max-width:400px;" xmlns="http://www.w3.org/2000/svg">\n'
    + '  <rect x="30" y="20" width="40" height="150" fill="#4CAF50" rx="4"/>\n'
    + '  <rect x="90" y="60" width="40" height="110" fill="#2196F3" rx="4"/>\n'
    + '  <rect x="150" y="100" width="40" height="70" fill="#FF9800" rx="4"/>\n'
    + '  <line x1="20" y1="170" x2="280" y2="170" stroke="#333" stroke-width="2"/>\n'
    + '  <text x="50" y="190" text-anchor="middle" font-size="12">A</text>\n'
    + '  <text x="110" y="190" text-anchor="middle" font-size="12">B</text>\n'
    + '  <text x="170" y="190" text-anchor="middle" font-size="12">C</text>\n'
    + '</svg>\n\n'
    + '【印刷への配慮】\n'
    + '- コントロールパネルには class="no-print" を付け、印刷時に非表示にする。\n'
    + '- SVGは印刷時にも正しく表示される。\n'
    + '- A4用紙（210mm×297mm）に収まるサイズを意識し、余白を適切にとる。\n'
    + '- 図版が大きすぎないよう max-width を設定する。\n\n'
    + '【美学】\n'
    + '教育用であっても視覚的な美しさ（線の滑らかさ、色の調和、余白のバランス）を意識してください。\n'
    + '小学生が親しみやすく、学習意欲が高まるデザインにしてください。\n';
}

// 統一プロンプトでワークシートHTMLを生成（クリーニング済み文字列を返す）。先生専用。
function generateSingleWorksheet(data) {
  MiraiAuth.requireTeacher();
  var prompt = buildWorksheetPrompt(data);
  var result = callGeminiAPI(prompt);
  result = result
    .replace(/```html/gi, '').replace(/```/g, '')
    .replace(/^[\s\S]*?<body[^>]*>/i, '').replace(/<\/body>[\s\S]*$/i, '')
    .replace(/<img\s+[^>]*src\s*=\s*["']https?:\/\/[^"']+["'][^>]*\/?>/gi, '')
    .trim();
  return result;
}

// AIでルーブリック（評価基準表）を生成。先生専用。
function generateRubricAI(data) {
  MiraiAuth.requireTeacher();
  var prompt = '教育評価専門家としてルーブリック作成。'
    + '単元:' + data.unitName
    + ',活動:' + data.stepTitle
    + ',内容:' + data.description
    + '。3観点3段階,HTMLテーブル形式(table table-bordered),具体的記述。HTMLのみ。';
  return callGeminiAPI(prompt);
}

// 指定タスクの全提出物にAIコメントを一括生成（たたき台）。先生専用。
function generateBatchComments(taskId) {
  MiraiAuth.requireTeacher();
  var ss = getSpreadsheet();

  var wsSheet = ss.getSheetByName(DB_SCHEMA.Worksheets.name);
  var unitName = '';
  var stepTitle = '';
  var wsFound = wsSheet.getRange('A:A').createTextFinder(String(taskId)).matchEntireCell(true).findNext();
  if (wsFound) {
    var wsRow = wsSheet.getRange(wsFound.getRow(), 1, 1, 3).getValues()[0];
    unitName = String(wsRow[1] || '');
    stepTitle = String(wsRow[2] || '');
  }

  var resSheet = ss.getSheetByName(DB_SCHEMA.Responses.name);
  var allData = resSheet.getDataRange().getValues();
  var submissions = [];
  for (var i = 1; i < allData.length; i++) {
    var row = allData[i];
    if (String(row[RS_COL_TASK_ID - 1]) !== String(taskId)) continue;
    if (row[RS_COL_STATUS - 1] === 'draft') continue;
    submissions.push({
      rowIndex: i + 1,
      studentName: row[RS_COL_STUDENT_NAME - 1] || '名無し',
      textContent: row[RS_COL_TEXT_CONTENT - 1] || '',
      reflection: row[RS_COL_REFLECTION - 1] || '',
      status: row[RS_COL_STATUS - 1] || ''
    });
  }

  if (submissions.length === 0) {
    return { comments: [], taskTitle: stepTitle };
  }

  var studentLines = submissions.map(function (s, idx) {
    var line = (idx + 1) + '. ' + s.studentName + ' | 自己評価: ' + (s.textContent || '未記入');
    if (s.reflection) line += ' | ふりかえり: ' + s.reflection;
    return line;
  }).join('\n');

  var prompt = 'あなたは経験豊富なベテラン小学校教師です。\n'
    + '以下の授業で提出された児童のワークシートに対して、一人ひとりに温かいコメントを書いてください。\n\n'
    + '【授業情報】\n'
    + '単元: ' + unitName + '\n'
    + '授業: ' + stepTitle + '\n\n'
    + '【コメントのルール】\n'
    + '- 児童の自己評価やふりかえりの内容に基づいて、個別化されたコメントを書く\n'
    + '- 学習の成果や成長を具体的に認め、児童が達成感を感じられるようにする\n'
    + '- 次の授業に向けたアドバイスや励ましを含める\n'
    + '- 「◎」の項目は大いに褒め、「△」の項目は改善のヒントを優しく伝える\n'
    + '- 小学生が読んで嬉しくなる、親しみやすい言葉遣いにする\n'
    + '- 各コメントは2〜4文（50〜120文字程度）で簡潔にまとめる\n'
    + '- 自己評価が未記入の場合は、提出したこと自体を認めて励ます\n\n'
    + '【児童の提出データ】\n'
    + studentLines + '\n\n'
    + '【出力形式（厳守）】\n'
    + '以下のJSON配列のみを出力してください。説明文やMarkdownは不要です。\n'
    + '[\n'
    + '  {"index": 1, "comment": "コメント内容"},\n'
    + '  {"index": 2, "comment": "コメント内容"},\n'
    + '  ...\n'
    + ']\n';

  var result = callGeminiAPI(prompt);
  var parsed = [];
  try {
    var jsonMatch = result.match(/\[[\s\S]*\]/);
    if (jsonMatch) parsed = JSON.parse(jsonMatch[0]);
  } catch (e) {
    Logger.log('AI コメントのJSONパースに失敗: ' + e.message);
  }

  var comments = submissions.map(function (s, idx) {
    var aiComment = '';
    for (var j = 0; j < parsed.length; j++) {
      if (parsed[j].index === idx + 1) { aiComment = parsed[j].comment || ''; break; }
    }
    return {
      rowIndex: s.rowIndex,
      studentName: s.studentName,
      textContent: s.textContent,
      status: s.status,
      aiComment: aiComment
    };
  });

  return { comments: comments, taskTitle: stepTitle };
}

/* ---------- Drive / ユーティリティ ---------- */

// このWebアプリの公開URLを返す。要ログイン。
function getWebAppUrl() {
  MiraiAuth.requireUser();
  return ScriptApp.getService().getUrl();
}

// 画像を Drive にアップロードしファイルIDを返す。要ログイン。
// data:image/(png|jpe?g|gif|webp);base64 のみ・8MBまで（任意ファイルの公開ホスト化を防ぐ）。
function uploadImageToDrive(base64Data) {
  MiraiAuth.requireUser();
  var FOLDER_NAME = 'みらいコンパス 添付ファイル';
  var m = (typeof base64Data === 'string')
    ? base64Data.match(/^data:image\/(png|jpe?g|gif|webp);base64,([A-Za-z0-9+/=\s]+)$/)
    : null;
  if (!m) {
    throw new Error('画像データの形式が不正です（png/jpeg/gif/webp のみ対応）。');
  }
  var b64 = m[2].replace(/\s+/g, '');
  var approxBytes = Math.floor(b64.length * 3 / 4);
  if (approxBytes > 8 * 1024 * 1024) {
    throw new Error('画像が大きすぎます（8MBまで）。');
  }
  var folders = DriveApp.getFoldersByName(FOLDER_NAME);
  var folder = folders.hasNext() ? folders.next() : DriveApp.createFolder(FOLDER_NAME);
  var contentType = 'image/' + m[1];
  var decoded = Utilities.base64Decode(b64);
  var blob = Utilities.newBlob(decoded, contentType, 'image.' + (m[1] === 'jpeg' ? 'jpg' : m[1]));
  var file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return file.getId();
}

// ==========================================
//  6. ヘルパー関数 (Helper Functions)
// ==========================================

function getSpreadsheet() {
  const ssId = PROPERTIES.getProperty('SS_ID');
  if (!ssId) throw new Error('Database not initialized.');
  const ss = SpreadsheetApp.openById(ssId);
  checkAndFixSheets(ss);
  return ss;
}

function fetchSheetData(ss, sheetName) {
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return [];
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow <= 1) return [];
  return sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
}

function checkAndFixSheets(ss) {
  // ほぼ全APIから呼ばれる高コスト処理のため、チェック結果を6時間キャッシュして高速化する
  // （スキーマ変更時はキャッシュ期限切れ後に自動で再チェックされる）
  const cache = CacheService.getScriptCache();
  const cacheKey = 'schema_ok_' + ss.getId();
  if (cache.get(cacheKey)) return;

  Object.keys(DB_SCHEMA).forEach(key => {
    const def = DB_SCHEMA[key];
    let sheet = ss.getSheetByName(def.name);
    if (!sheet) {
      sheet = ss.insertSheet(def.name);
      sheet.appendRow(def.headers);
    } else {
      const currentCols = sheet.getLastColumn();
      if (currentCols < def.headers.length) {
        const missingHeaders = def.headers.slice(currentCols);
        sheet.getRange(1, currentCols + 1, 1, missingHeaders.length).setValues([missingHeaders]);
      }
    }
  });

  cache.put(cacheKey, '1', 21600); // 6時間
}

function createSuccessResponse(data = {}) { return { success: true, ...data }; }
function createErrorResponse(error) { console.error(error); return { success: false, error: error.toString() }; }
function safeJsonParse(str) { try { return JSON.parse(str || '{}'); } catch(e) { return {}; } }
// パスポート由来の関数が使うパーサ。失敗時は null（safeJsonParse とは戻り値が異なる）。
function safeJSONParse(s) { try { return JSON.parse(s); } catch (e) { return null; } }
function ensureArray(val) { return Array.isArray(val) ? val : []; }
function formatDate(d) { try { return Utilities.formatDate(new Date(d), "JST", "HH:mm"); } catch(e) { return ""; } }

/**
 * 日付セルを "yyyy-MM-dd" 文字列に正規化する
 * シートは "2026-07-12" のような文字列を自動的にDate型へ変換するため、
 * Date型・文字列のどちらで格納されていても同じ形式で返す
 */
function normalizeDateStr(v) {
  if (v instanceof Date) return Utilities.formatDate(v, "JST", "yyyy-MM-dd");
  const s = String(v || "");
  if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return s;
  try {
    const d = new Date(s);
    if (!isNaN(d.getTime())) return Utilities.formatDate(d, "JST", "yyyy-MM-dd");
  } catch (e) {}
  return s;
}

/**
 * 時刻セルを "HH:mm" 文字列に正規化する（Date型・"8:45"のような文字列の両対応）
 */
function normalizeTimeStr(v) {
  if (v instanceof Date) return Utilities.formatDate(v, "JST", "HH:mm");
  const s = String(v || "");
  const m = s.match(/^(\d{1,2}):(\d{2})/);
  if (m) return ('0' + m[1]).slice(-2) + ':' + m[2];
  return s;
}
