/**
 * 🧭 みらいコンパス Ver. 1.0 - サーバーサイドプログラム
 * * このファイルは、Googleスプレッドシート（データベース）とのやり取りを担当します。
 * データの保存、読み出し、初期設定などの機能が含まれています。
 */

// ==========================================
//  1. 設定と基本情報 (Configuration)
// ==========================================

// プロパティサービス（設定値を保存する場所）の取得
const PROPERTIES = PropertiesService.getScriptProperties();

// アプリケーションの名前
const APP_TITLE = 'みらいコンパス Ver. 1.0';

/**
 * データベース（スプレッドシート）の設計図
 * 新しい機能を追加するときは、ここに列（headers）を追加します。
 */
const DB_SCHEMA = {
  // 単元と課題のマスタデータ
  UnitMaster: {
    name: 'UnitMaster',
    headers: ['unitId', 'taskId', 'type', 'title', 'description', 'estTime', 'deletedAt', 'category', 'step', 'textbook', 'tablet', 'print', 'prerequisites', 'format', 'unitInfo', 'totalHours']
  },
  // 毎時の学習ログ（誰が、いつ、どの課題を、どうしたか）
  LearningLogs: {
    name: 'LearningLogs',
    headers: ['logId', 'studentId', 'studentName', 'taskId', 'status', 'reflection', 'timestamp', 'classId']
  },
  // 現在の状況（LIVEモニタリング用）
  LiveStatus: {
    name: 'LiveStatus',
    headers: ['studentId', 'studentName', 'currentTask', 'mode', 'lastUpdate', 'currentUnitId', 'currentHour', 'classId', 'x', 'y']
  },
  // 先生からのスタンプなどのフィードバック
  Feedback: {
    name: 'Feedback',
    headers: ['feedbackId', 'studentName', 'taskId', 'stamp', 'timestamp', 'classId']
  },
  // 児童が自分で作成した課題（マイタスク）
  MyTasks: {
    name: 'MyTasks',
    headers: ['taskId', 'studentName', 'title', 'description', 'estTime', 'created_at', 'unitId', 'classId']
  },
  // 児童が立てた学習計画
  StudentPlans: {
    name: 'StudentPlans',
    headers: ['studentName', 'unitId', 'planData', 'lastUpdate', 'classId']
  },
  // 毎時の振り返り（達成度、コメント、スキル）
  DailyReflections: {
    name: 'DailyReflections',
    headers: ['studentName', 'unitId', 'hour', 'achievement', 'comment', 'teacherCheck', 'timestamp', 'classId', 'skills']
  },
  // 単元のまとめ（ポートフォリオ）
  Portfolios: {
    name: 'Portfolios',
    headers: ['studentName', 'unitId', 'summary', 'lastUpdate', 'classId']
  },
  // 授業スケジュール（先生からのお知らせ・時数指定）
  ClassSchedule: {
    name: 'ClassSchedule',
    headers: ['scheduleId', 'classId', 'date', 'startTime', 'endTime', 'unitId', 'hour', 'message', 'createdAt']
  }
};

// ==========================================
//  2. 基本機能 (Core Functions)
// ==========================================

/**
 * Webアプリにアクセスしたときに最初に実行される関数
 */
function doGet(e) {
  try {
    // index.html ファイルを読み込んでWebページを作成します
    const template = HtmlService.createTemplateFromFile('index');
    
    // URLパラメータ(?mode=teacherなど)があれば受け取ります
    template.mode = (e && e.parameter && e.parameter.mode) ? e.parameter.mode : 'auto';
    
    // スマホ対応などの設定をしてページを表示します
    return template.evaluate()
      .addMetaTag('viewport', 'width=device-width, initial-scale=1.0')
      .setTitle(APP_TITLE)
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } catch (error) {
    // エラーが発生した場合の画面
    return HtmlService.createHtmlOutput(`
      <div style="font-family:sans-serif; padding:20px; color:#d32f2f;">
        <h2>起動エラー (Startup Error)</h2>
        <p>アプリケーションの読み込みに失敗しました。</p>
        <p>エラー詳細: ${error.toString()}</p>
      </div>
    `);
  }
}

/**
 * アプリの初期状態データを取得する関数
 * クライアント側（ブラウザ）から最初に呼び出されます。
 */
function getAppInitialData() {
  try {
    const ssId = PROPERTIES.getProperty('SS_ID');
    const adminEmail = PROPERTIES.getProperty('ADMIN_EMAIL');
    const userEmail = Session.getActiveUser().getEmail();
    
    // 管理者かどうかを判定（簡易的）
    const isAdmin = (!adminEmail) || (adminEmail === userEmail);

    return createSuccessResponse({
      isInitialized: !!ssId, // データベースが作成済みか
      ssUrl: ssId ? `https://docs.google.com/spreadsheets/d/${ssId}/edit` : null,
      isAdmin: isAdmin,
      userEmail: userEmail
    });
  } catch (e) {
    return createErrorResponse(e);
  }
}

/**
 * 初回セットアップ：データベース（スプレッドシート）を作成する関数
 */
function initSystem() {
  try {
    let ssId = PROPERTIES.getProperty('SS_ID');
    let ss;
    
    if (ssId) {
      ss = SpreadsheetApp.openById(ssId);
    } else {
      // 新しいスプレッドシートを作成
      ss = SpreadsheetApp.create('みらいコンパス_データベース');
      ssId = ss.getId();
      PROPERTIES.setProperty('SS_ID', ssId);
      PROPERTIES.setProperty('ADMIN_EMAIL', Session.getActiveUser().getEmail());
    }
    
    // 必要なシートと列が存在するか確認し、なければ作成・修復します
    checkAndFixSheets(ss);
    
    // デフォルトの「シート1」があれば削除
    const defaultSheet = ss.getSheetByName('シート1');
    if (defaultSheet) ss.deleteSheet(defaultSheet);
    
    return createSuccessResponse({ message: 'システムを初期化しました。' });
  } catch (e) {
    return createErrorResponse(e);
  }
}

// ==========================================
//  3. データ読み込み (Read Data)
// ==========================================

/**
 * アプリに必要な全データを一括取得する関数
 */
function getData() {
  try {
    const ssId = PROPERTIES.getProperty('SS_ID');
    if (!ssId) return createSuccessResponse({ json: JSON.stringify({}) });
    
    const ss = SpreadsheetApp.openById(ssId);
    // 念のためシート構成をチェック
    checkAndFixSheets(ss);

    // 1. 単元マスタ (UnitMaster)
    const unitData = fetchSheetData(ss, DB_SCHEMA.UnitMaster.name).map(r => ({
      unitId: String(r[0]), taskId: String(r[1]), type: String(r[2]), title: String(r[3]),
      desc: String(r[4]), time: Number(r[5]), category: String(r[7]), step: String(r[8] || ''),
      textbook: String(r[9] || ''), tablet: String(r[10] || ''), print: String(r[11] || ''),
      prerequisites: String(r[12] || '').split(',').filter(x => x), format: String(r[13] || 'student'),
      unitInfo: safeJsonParse(r[14]), totalHours: Number(r[15] || 8)
    }));

    // 2. 児童の現在の状況 (LiveStatus)
    const liveData = fetchSheetData(ss, DB_SCHEMA.LiveStatus.name).map(r => ({
      id: String(r[0]), name: String(r[1]), task: String(r[2]), mode: String(r[3]),
      time: r[4] ? formatDate(r[4]) : '', currentUnitId: String(r[5] || ''),
      currentHour: Number(r[6] || 1), classId: String(r[7] || ''),
      x: Number(r[8]) || 0, y: Number(r[9]) || 0
    }));

    // 3. 学習ログの集計 (LearningLogs)
    const clsProgress = {};
    fetchSheetData(ss, DB_SCHEMA.LearningLogs.name).forEach(r => {
      const name = String(r[2]); const taskId = String(r[3]); const status = String(r[4]); const reflection = String(r[5] || "");
      if (!clsProgress[name]) clsProgress[name] = {};
      if (!clsProgress[name][taskId]) clsProgress[name][taskId] = { status: '', reflection: '' };
      // 最新のステータスで上書き
      if (status && status !== 'メモ') clsProgress[name][taskId].status = status;
      if (reflection) clsProgress[name][taskId].reflection = reflection;
    });

    // 4. フィードバック (Feedback)
    const clsFeedback = {};
    fetchSheetData(ss, DB_SCHEMA.Feedback.name).forEach(r => {
      const name = String(r[1]); const taskId = String(r[2]); const stamp = String(r[3]);
      if (!clsFeedback[name]) clsFeedback[name] = {};
      clsFeedback[name][taskId] = stamp;
    });

    // 5. 児童作成課題 (MyTasks)
    const allMyTasks = {};
    fetchSheetData(ss, DB_SCHEMA.MyTasks.name).forEach(r => {
      const name = String(r[1]);
      if (!allMyTasks[name]) allMyTasks[name] = [];
      allMyTasks[name].push({
        taskId: String(r[0]), title: String(r[2]), desc: String(r[3]), time: Number(r[4]),
        category: 'マイタスク', type: 'challenge', format: 'student', unitId: String(r[6] || '')
      });
    });

    // 6. 学習計画 (StudentPlans)
    const clsPlans = {};
    fetchSheetData(ss, DB_SCHEMA.StudentPlans.name).forEach(r => {
      const name = String(r[0]); const uid = String(r[1]);
      if (!clsPlans[name]) clsPlans[name] = {};
      clsPlans[name][uid] = safeJsonParse(r[2]);
    });

    // 7. 毎時の振り返り (DailyReflections)
    const clsReflections = {};
    fetchSheetData(ss, DB_SCHEMA.DailyReflections.name).forEach(r => {
      const name = String(r[0]); const uid = String(r[1]); const hour = String(r[2]);
      if (!clsReflections[name]) clsReflections[name] = {};
      if (!clsReflections[name][uid]) clsReflections[name][uid] = {};
      clsReflections[name][uid][hour] = { 
        achievement: r[3], comment: r[4], check: r[5],
        skills: safeJsonParse(r[8])
      };
    });

    // 8. ポートフォリオまとめ (Portfolios)
    const clsPortfolios = {};
    fetchSheetData(ss, DB_SCHEMA.Portfolios.name).forEach(r => {
      const name = String(r[0]); const uid = String(r[1]); const summary = String(r[2]);
      if (!clsPortfolios[name]) clsPortfolios[name] = {};
      clsPortfolios[name][uid] = summary;
    });

    // 9. 授業スケジュール (ClassSchedule) - 本日以降のみ
    const todayStr = Utilities.formatDate(new Date(), "JST", "yyyy-MM-dd");
    const schedules = fetchSheetData(ss, DB_SCHEMA.ClassSchedule.name)
      .filter(r => String(r[2]) >= todayStr)
      .map(r => ({
        id: String(r[0]), classId: String(r[1]), date: String(r[2]), 
        startTime: formatDate(r[3]), endTime: formatDate(r[4]), 
        unitId: String(r[5]), hour: String(r[6]), message: String(r[7])
      }));

    return createSuccessResponse({
      json: JSON.stringify({
        unit: unitData, live: liveData, progress: clsProgress, feedback: clsFeedback,
        myTasks: allMyTasks, plans: clsPlans, dailyReflections: clsReflections,
        portfolios: clsPortfolios, schedules: schedules
      })
    });
  } catch (e) { return createErrorResponse(e); }
}

/**
 * 児童個人の詳細データを取得する関数（ログイン時などに使用）
 */
function getStudentProgress(studentName, classId, currentUnitId) {
  try {
    const ssId = PROPERTIES.getProperty('SS_ID');
    if (!ssId) return createSuccessResponse({ json: JSON.stringify({}) });
    const ss = SpreadsheetApp.openById(ssId);

    // ログイン時に最新のクラスIDや単元IDをLiveStatusに記録
    if (classId || currentUnitId) {
      updateLiveStatusMeta(ss, studentName, classId, currentUnitId);
    }

    // 必要なデータを各シートからフィルタリングして取得
    const map = {}; // 進捗
    fetchSheetData(ss, DB_SCHEMA.LearningLogs.name).forEach(r => {
      if (r[2] === studentName) {
        const tid = String(r[3]);
        if (!map[tid]) map[tid] = { status: '', reflection: '' };
        if (r[4] && r[4] !== 'メモ') map[tid].status = r[4];
        if (r[5]) map[tid].reflection = r[5];
      }
    });

    const fbMap = {}; // フィードバック
    fetchSheetData(ss, DB_SCHEMA.Feedback.name).forEach(r => {
      if (r[1] === studentName) fbMap[String(r[2])] = String(r[3]);
    });

    const myTasks = []; // マイタスク
    fetchSheetData(ss, DB_SCHEMA.MyTasks.name).forEach(r => {
      if (r[1] === studentName) {
        myTasks.push({
          taskId: String(r[0]), title: String(r[2]), desc: String(r[3]), time: Number(r[4]),
          category: 'マイタスク', type: 'challenge', format: 'student', unitId: String(r[6] || '')
        });
      }
    });

    const plans = {}; // 計画
    fetchSheetData(ss, DB_SCHEMA.StudentPlans.name).forEach(r => {
      if (r[0] === studentName) plans[r[1]] = safeJsonParse(r[2]);
    });

    const reflections = {}; // 振り返り
    fetchSheetData(ss, DB_SCHEMA.DailyReflections.name).forEach(r => {
      if (r[0] === studentName) {
        const uid = String(r[1]);
        const hour = String(r[2]);
        if (!reflections[uid]) reflections[uid] = {};
        reflections[uid][hour] = {
          achievement: r[3], comment: r[4], check: r[5], skills: safeJsonParse(r[8])
        };
      }
    });

    let portfolioSummary = ""; // まとめ
    const pSheet = ss.getSheetByName(DB_SCHEMA.Portfolios.name);
    if(pSheet) {
      const pData = pSheet.getDataRange().getValues();
      for(let i=1; i<pData.length; i++) {
        if(pData[i][0] === studentName && String(pData[i][1]) === String(currentUnitId)) {
          portfolioSummary = pData[i][2];
          break;
        }
      }
    }

    return createSuccessResponse({
      json: JSON.stringify({ 
        progress: map, feedback: fbMap, myTasks: myTasks, 
        portfolio: { summary: portfolioSummary }, plans: plans, reflections: reflections
      })
    });
  } catch (e) { return createErrorResponse(e); }
}

/**
 * 分析用データを取得する関数（シンプル版）
 */
function getAnalysisData() {
  try {
    const ssId = PROPERTIES.getProperty('SS_ID');
    if (!ssId) return createSuccessResponse({ json: JSON.stringify({ tasks: [], counts: {} }) });
    const ss = SpreadsheetApp.openById(ssId);
    
    const tasks = fetchSheetData(ss, DB_SCHEMA.UnitMaster.name).map(r => ({ id: r[1], title: r[3], unitId: r[0] }));
    const counts = {};
    fetchSheetData(ss, DB_SCHEMA.LearningLogs.name).forEach(r => {
      if (String(r[4]) === '完了') {
        const tid = String(r[3]);
        counts[tid] = (counts[tid] || 0) + 1;
      }
    });
    return createSuccessResponse({ json: JSON.stringify({ tasks: tasks, counts: counts }) });
  } catch (e) { return createErrorResponse(e); }
}

// 内部関数：LiveStatusのメタデータ（クラスID、単元ID）を更新
function updateLiveStatusMeta(ss, name, classId, unitId) {
  const sheet = ss.getSheetByName(DB_SCHEMA.LiveStatus.name);
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][1] === name) {
      if (unitId) sheet.getRange(i + 1, 6).setValue(unitId);
      if (classId) sheet.getRange(i + 1, 8).setValue(classId);
      return;
    }
  }
}

// ==========================================
//  4. データ保存 (Write Data)
// ==========================================

/**
 * 学習状況（ステータス）を更新する関数
 */
function updateStatus(studentName, taskId, taskTitle, status, mode, reflection, classId, currentUnitId) {
  try {
    const ss = getSpreadsheet();
    const now = new Date();
    // ログを追記
    ss.getSheetByName(DB_SCHEMA.LearningLogs.name).appendRow([
      Utilities.getUuid(), studentName, studentName, taskId, status, reflection || "", now, classId || ""
    ]);

    // LiveStatus（現在の状態）を更新
    if (status !== 'メモ') {
      const liveSheet = ss.getSheetByName(DB_SCHEMA.LiveStatus.name);
      const data = liveSheet.getDataRange().getValues();
      let rIdx = -1;
      // 既存行を探す
      for (let i = 1; i < data.length; i++) {
        if (data[i][1] === studentName) { rIdx = i + 1; break; }
      }
      const displayTitle = taskTitle || taskId;
      
      if (rIdx > 0) {
        // 更新
        liveSheet.getRange(rIdx, 3, 1, 2).setValues([[displayTitle, mode]]);
        liveSheet.getRange(rIdx, 5).setValue(now);
        if (classId) liveSheet.getRange(rIdx, 8).setValue(classId);
        if (currentUnitId) liveSheet.getRange(rIdx, 6).setValue(currentUnitId);
      } else {
        // 新規作成（初期座標 x:0, y:0）
        liveSheet.appendRow([studentName, studentName, displayTitle, mode, now, currentUnitId || "", 1, classId || "", 0, 0]);
      }
    }
    return createSuccessResponse();
  } catch (e) { return createErrorResponse(e); }
}

/**
 * 座席配置（座標）を保存する関数
 */
function saveSeatCoordinates(coordinates) {
  try {
    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName(DB_SCHEMA.LiveStatus.name);
    const data = sheet.getDataRange().getValues();
    const nameToRow = new Map();
    for (let i = 1; i < data.length; i++) { nameToRow.set(data[i][1], i + 1); }
    
    coordinates.forEach(c => {
      const rIdx = nameToRow.get(c.name);
      if (rIdx) { sheet.getRange(rIdx, 9, 1, 2).setValues([[c.x, c.y]]); }
    });
    return createSuccessResponse();
  } catch(e) { return createErrorResponse(e); }
}

/**
 * 授業スケジュールを保存・配信する関数
 */
function saveClassSchedule(scheduleData) {
  try {
    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName(DB_SCHEMA.ClassSchedule.name);
    sheet.appendRow([
      Utilities.getUuid(), scheduleData.classId, scheduleData.date, 
      scheduleData.startTime, scheduleData.endTime, 
      scheduleData.unitId, scheduleData.hour, scheduleData.message, new Date()
    ]);
    return createSuccessResponse();
  } catch(e) { return createErrorResponse(e); }
}

/**
 * スタンプ（フィードバック）を送信する関数
 */
function sendFeedback(studentName, taskId, stamp, classId) {
  try {
    const ss = getSpreadsheet();
    ss.getSheetByName(DB_SCHEMA.Feedback.name).appendRow([Utilities.getUuid(), studentName, taskId, stamp, new Date(), classId || ""]);
    return createSuccessResponse();
  } catch (e) { return createErrorResponse(e); }
}

/**
 * マイタスクを追加する関数
 */
function addMyTask(studentName, title, desc, time, unitId, classId) {
  try {
    const ss = getSpreadsheet();
    const taskId = "MT" + Utilities.getUuid().substring(0, 8);
    ss.getSheetByName(DB_SCHEMA.MyTasks.name).appendRow([taskId, studentName, title, desc, time, new Date(), unitId || "", classId || ""]);
    return createSuccessResponse();
  } catch (e) { return createErrorResponse(e); }
}

/**
 * 学習計画を保存する関数
 */
function saveStudentPlan(studentName, unitId, planData, classId) {
  try {
    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName(DB_SCHEMA.StudentPlans.name);
    const data = sheet.getDataRange().getValues();
    const json = JSON.stringify(planData);
    const now = new Date();
    
    let rowIndex = -1;
    for(let i = 1; i < data.length; i++) {
      if(data[i][0] === studentName && data[i][1] === unitId) { rowIndex = i + 1; break; }
    }
    
    if(rowIndex > 0) {
      sheet.getRange(rowIndex, 3, 1, 2).setValues([[json, now]]);
      if(classId) sheet.getRange(rowIndex, 5).setValue(classId);
    } else {
      sheet.appendRow([studentName, unitId, json, now, classId || ""]);
    }
    return createSuccessResponse();
  } catch(e) { return createErrorResponse(e); }
}

/**
 * 毎時の振り返りを保存する関数
 */
function saveDailyReflection(studentName, unitId, hour, achievement, comment, classId, skills) {
  try {
    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName(DB_SCHEMA.DailyReflections.name);
    const data = sheet.getDataRange().getValues();
    const now = new Date();
    const skillsJson = JSON.stringify(skills || []);
    
    let rowIndex = -1;
    for(let i = 1; i < data.length; i++) {
      if(data[i][0] === studentName && data[i][1] === unitId && String(data[i][2]) === String(hour)) { rowIndex = i + 1; break; }
    }
    
    if(rowIndex > 0) {
      sheet.getRange(rowIndex, 4, 1, 2).setValues([[achievement, comment]]);
      sheet.getRange(rowIndex, 7).setValue(now);
      if(classId) sheet.getRange(rowIndex, 8).setValue(classId);
      sheet.getRange(rowIndex, 9).setValue(skillsJson);
    } else {
      sheet.appendRow([studentName, unitId, hour, achievement, comment, "", now, classId || "", skillsJson]);
    }
    return createSuccessResponse();
  } catch(e) { return createErrorResponse(e); }
}

/**
 * 単元のまとめ（ポートフォリオ）を保存する関数
 */
function savePortfolio(studentName, unitId, summary, classId) {
  try {
    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName(DB_SCHEMA.Portfolios.name);
    const data = sheet.getDataRange().getValues();
    const now = new Date();
    
    let rowIndex = -1;
    for(let i = 1; i < data.length; i++) {
      if(data[i][0] === studentName && String(data[i][1]) === String(unitId)) { rowIndex = i + 1; break; }
    }
    
    if(rowIndex > 0) {
      sheet.getRange(rowIndex, 3, 1, 2).setValues([[summary, now]]);
      if(classId) sheet.getRange(rowIndex, 5).setValue(classId);
    } else {
      sheet.appendRow([studentName, unitId, summary, now, classId || ""]);
    }
    return createSuccessResponse();
  } catch(e) { return createErrorResponse(e); }
}

/**
 * JSONデータから単元を一括登録する関数（AIインポート用）
 */
function importUnitJson(jsonStr) {
  try {
    const ssId = PROPERTIES.getProperty('SS_ID');
    if (!ssId) throw new Error('初期設定未完了');
    const ss = SpreadsheetApp.openById(ssId);
    checkAndFixSheets(ss);
    
    const data = JSON.parse(jsonStr);
    const uid = "U" + Utilities.formatDate(new Date(), "JST", "yyyyMMddHHmm");
    const uInfoStr = JSON.stringify(data.unitInfo || {});
    const totalHours = data.unitInfo?.totalHours || 8;
    
    const rows = data.tasks.map(t => [
      uid, t.id || '', t.type || 'must', t.title || '無題', t.description || '', t.estimatedTime || 10, '', 
      t.category || 'まなぶ', t.step || '', t.textbook || '', t.tablet || '', t.print || '',
      (t.prerequisites || []).join(','), t.format || 'student', uInfoStr, totalHours
    ]);
    
    if (rows.length > 0) {
      const sheet = ss.getSheetByName(DB_SCHEMA.UnitMaster.name);
      sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
    }
    return createSuccessResponse({ count: rows.length });
  } catch (e) { return createErrorResponse(e); }
}

/**
 * 古い単元データをアーカイブ（別ファイル退避）する関数
 */
function archiveUnitData(unitId, unitTitle) {
  try {
    const ss = getSpreadsheet();
    const archiveName = `アーカイブ_${unitTitle || unitId}_${Utilities.formatDate(new Date(), 'JST', 'yyyyMMdd')}`;
    const archiveSs = SpreadsheetApp.create(archiveName);
    
    const targets = [
      { key: 'MyTasks', colUnitId: 6 },
      { key: 'StudentPlans', colUnitId: 1 },
      { key: 'DailyReflections', colUnitId: 1 },
      { key: 'Portfolios', colUnitId: 1 }
    ];
    let movedCount = 0;
    
    // 各シートから該当データを移動
    targets.forEach(t => {
      const sheet = ss.getSheetByName(DB_SCHEMA[t.key].name);
      if(!sheet) return;
      const data = sheet.getDataRange().getValues();
      const headers = data[0];
      const rows = data.slice(1);
      const toArchive = [];
      const toKeep = [];
      
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

    // 単元マスタからも移動
    const uSheet = ss.getSheetByName(DB_SCHEMA.UnitMaster.name);
    const uData = uSheet.getDataRange().getValues();
    const uKeep = [];
    const uArch = [];
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
    
    // アーカイブファイルのデフォルトシート削除
    const delSheet = archiveSs.getSheetByName('シート1');
    if(delSheet) archiveSs.deleteSheet(delSheet);

    return createSuccessResponse({ 
      message: `アーカイブ完了: ${movedCount}件のデータを移動しました。\nファイル名: ${archiveName}`,
      url: archiveSs.getUrl()
    });

  } catch(e) { return createErrorResponse(e); }
}

// ==========================================
//  5. Helper Functions (内部処理用)
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
}

function createSuccessResponse(data = {}) { return { success: true, ...data }; }
function createErrorResponse(error) { console.error(error); return { success: false, error: error.toString() }; }
function safeJsonParse(str) { try { return JSON.parse(str || '{}'); } catch(e) { return {}; } }
function formatDate(d) { try { return Utilities.formatDate(new Date(d), "JST", "HH:mm"); } catch(e) { return ""; } }
