/**
 * みらいコンパス v12.0 - Server Side Logic (Refactored)
 * * 概要:
 * 自由進度学習支援アプリケーションのバックエンドロジック。
 * データの読み書き、初期化、AIインポート機能などを提供する。
 * * リファクタリングのポイント:
 * - 定数定義によるマジックナンバー/文字列の排除
 * - 関数の責務分離と構造化
 * - エラーハンドリングの統一
 * - JSDocによるドキュメント化
 */

// ==========================================
//  1. Configuration & Constants
// ==========================================

const PROPERTIES = PropertiesService.getScriptProperties();
const APP_TITLE = 'みらいコンパス v12.0';

/**
 * データベースのシート定義
 * シート名とヘッダー列の構成を管理
 */
const DB_SCHEMA = {
  UnitMaster: {
    name: 'UnitMaster',
    headers: ['unitId', 'taskId', 'type', 'title', 'description', 'estTime', 'deletedAt', 'category', 'step', 'textbook', 'tablet', 'print', 'prerequisites', 'format', 'unitInfo', 'totalHours']
  },
  LearningLogs: {
    name: 'LearningLogs',
    headers: ['logId', 'studentId', 'studentName', 'taskId', 'status', 'reflection', 'timestamp']
  },
  LiveStatus: {
    name: 'LiveStatus',
    headers: ['studentId', 'studentName', 'currentTask', 'mode', 'lastUpdate', 'currentUnitId', 'currentHour']
  },
  Feedback: {
    name: 'Feedback',
    headers: ['feedbackId', 'studentName', 'taskId', 'stamp', 'timestamp']
  },
  MyTasks: {
    name: 'MyTasks',
    headers: ['taskId', 'studentName', 'title', 'description', 'estTime', 'created_at', 'unitId']
  },
  StudentPlans: {
    name: 'StudentPlans',
    headers: ['studentName', 'unitId', 'planData', 'lastUpdate']
  },
  DailyReflections: {
    name: 'DailyReflections',
    headers: ['studentName', 'unitId', 'hour', 'achievement', 'comment', 'teacherCheck', 'timestamp']
  },
  Portfolios: {
    name: 'Portfolios',
    headers: ['studentName', 'summary', 'lastUpdate']
  }
};

// ==========================================
//  2. Core Functions (App Entry Points)
// ==========================================

/**
 * Webアプリへのアクセス時に実行 (GETリクエスト)
 */
function doGet(e) {
  try {
    const ssId = PROPERTIES.getProperty('SS_ID');
    const template = HtmlService.createTemplateFromFile('index');
    
    // クエリパラメータからモードを取得 (デフォルトは 'auto')
    template.mode = e.parameter.mode || 'auto';
    // 初期設定済みかどうかを判定
    template.isFirstRun = !ssId;

    return template.evaluate()
      .addMetaTag('viewport', 'width=device-width, initial-scale=1.0')
      .setTitle(APP_TITLE)
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } catch (error) {
    return HtmlService.createHtmlOutput(`<h1>起動エラー</h1><p>${error.toString()}</p>`);
  }
}

/**
 * クライアント初期化用データの取得
 * 権限チェックやスプレッドシートのURLなどを返す
 */
function getAppInitialData() {
  try {
    const ssId = PROPERTIES.getProperty('SS_ID');
    const adminEmail = PROPERTIES.getProperty('ADMIN_EMAIL');
    const userEmail = Session.getActiveUser().getEmail();

    // 管理者判定: 初回起動者またはプロパティに保存されたEmailと一致する場合
    const isAdmin = (!adminEmail) || (adminEmail === userEmail);

    return createSuccessResponse({
      isInitialized: !!ssId,
      ssUrl: ssId ? `https://docs.google.com/spreadsheets/d/${ssId}/edit` : null,
      isAdmin: isAdmin,
      userEmail: userEmail
    });
  } catch (e) {
    return createErrorResponse(e);
  }
}

/**
 * システムの初期化・修復
 * スプレッドシートの作成および全シートの存在確認・修復を行う
 */
function initSystem() {
  try {
    let ssId = PROPERTIES.getProperty('SS_ID');
    let ss;

    if (ssId) {
      ss = SpreadsheetApp.openById(ssId);
    } else {
      // 新規作成
      ss = SpreadsheetApp.create('みらいコンパス_データベース');
      ssId = ss.getId();
      PROPERTIES.setProperty('SS_ID', ssId);
      PROPERTIES.setProperty('ADMIN_EMAIL', Session.getActiveUser().getEmail());
    }

    // 全シートのスキーマチェックと修復
    checkAndFixSheets(ss);

    // デフォルトの「シート1」があれば削除
    const defaultSheet = ss.getSheetByName('シート1');
    if (defaultSheet) ss.deleteSheet(defaultSheet);
    
    return createSuccessResponse({ message: 'System initialized successfully.' });
  } catch (e) {
    return createErrorResponse(e);
  }
}

// ==========================================
//  3. Data Access (Read Operations)
// ==========================================

/**
 * アプリのメインデータを一括取得 (UnitMaster, LiveStatus, Progress, Feedback, MyTasks, Plans, Reflections)
 * クライアント側の負荷を減らすため、JSON文字列化して返す
 */
function getData() {
  try {
    const ssId = PROPERTIES.getProperty('SS_ID');
    if (!ssId) return createSuccessResponse({ json: JSON.stringify({}) });

    const ss = SpreadsheetApp.openById(ssId);
    // 念のためアクセス時にスキーマチェック（列不足エラー回避）
    checkAndFixSheets(ss);

    // 1. 単元マスタ (UnitMaster)
    const unitData = fetchSheetData(ss, DB_SCHEMA.UnitMaster.name).map(r => ({
      unitId: String(r[0]), 
      taskId: String(r[1]), 
      type: String(r[2]), 
      title: String(r[3]),
      desc: String(r[4]), 
      time: Number(r[5]), 
      category: String(r[7]),
      step: String(r[8] || ''), 
      textbook: String(r[9] || ''), 
      tablet: String(r[10] || ''), 
      print: String(r[11] || ''),
      prerequisites: String(r[12] || '').split(',').filter(x => x), 
      format: String(r[13] || 'student'),
      unitInfo: safeJsonParse(r[14]), 
      totalHours: Number(r[15] || 8)
    }));

    // 2. ライブ状況 (LiveStatus)
    const liveData = fetchSheetData(ss, DB_SCHEMA.LiveStatus.name).map(r => ({
      id: String(r[0]), 
      name: String(r[1]), 
      task: String(r[2]), 
      mode: String(r[3]),
      time: r[4] ? formatDate(r[4]) : '',
      currentUnitId: String(r[5] || ''), 
      currentHour: Number(r[6] || 1)
    }));

    // 3. 学習ログ集計 (LearningLogs -> clsProgress)
    const clsProgress = {};
    fetchSheetData(ss, DB_SCHEMA.LearningLogs.name).forEach(r => {
      const name = String(r[2]); 
      const taskId = String(r[3]);
      const status = String(r[4]);
      const reflection = String(r[5] || "");

      if (!clsProgress[name]) clsProgress[name] = {};
      if (!clsProgress[name][taskId]) clsProgress[name][taskId] = { status: '', reflection: '' };

      // ステータスが「メモ」の場合は進捗状態としては無視（上書きしない）
      if (status && status !== 'メモ') clsProgress[name][taskId].status = status;
      // 振り返りは常に最新
      if (reflection) clsProgress[name][taskId].reflection = reflection;
    });

    // 4. フィードバック (Feedback)
    const clsFeedback = {};
    fetchSheetData(ss, DB_SCHEMA.Feedback.name).forEach(r => {
      const name = String(r[1]);
      const taskId = String(r[2]);
      const stamp = String(r[3]);
      if (!clsFeedback[name]) clsFeedback[name] = {};
      clsFeedback[name][taskId] = stamp;
    });

    // 5. マイタスク (MyTasks)
    const allMyTasks = {};
    fetchSheetData(ss, DB_SCHEMA.MyTasks.name).forEach(r => {
      const name = String(r[1]);
      if (!allMyTasks[name]) allMyTasks[name] = [];
      allMyTasks[name].push({
        taskId: String(r[0]), 
        title: String(r[2]), 
        desc: String(r[3]), 
        time: Number(r[4]),
        category: 'マイタスク', 
        type: 'challenge', 
        format: 'student', 
        unitId: String(r[6] || '')
      });
    });

    // 6. 単元計画 (StudentPlans)
    const clsPlans = {};
    fetchSheetData(ss, DB_SCHEMA.StudentPlans.name).forEach(r => {
      const name = String(r[0]); 
      const uid = String(r[1]);
      if (!clsPlans[name]) clsPlans[name] = {};
      clsPlans[name][uid] = safeJsonParse(r[2]);
    });

    // 7. 毎時の振り返り (DailyReflections)
    const clsReflections = {};
    fetchSheetData(ss, DB_SCHEMA.DailyReflections.name).forEach(r => {
      const name = String(r[0]); 
      const uid = String(r[1]); 
      const hour = String(r[2]);
      if (!clsReflections[name]) clsReflections[name] = {};
      if (!clsReflections[name][uid]) clsReflections[name][uid] = {};
      clsReflections[name][uid][hour] = { achievement: r[3], comment: r[4], check: r[5] };
    });

    return createSuccessResponse({
      json: JSON.stringify({
        unit: unitData,
        live: liveData,
        progress: clsProgress,
        feedback: clsFeedback,
        myTasks: allMyTasks,
        plans: clsPlans,
        dailyReflections: clsReflections
      })
    });
  } catch (e) {
    return createErrorResponse(e);
  }
}

/**
 * 特定の児童の全データを取得 (Client Side: login)
 */
function getStudentProgress(studentName) {
  try {
    const ssId = PROPERTIES.getProperty('SS_ID');
    if (!ssId) return createSuccessResponse({ json: JSON.stringify({}) });

    const ss = SpreadsheetApp.openById(ssId);
    checkAndFixSheets(ss);

    // 各シートから該当生徒のデータのみ抽出して構成
    // (getDataロジックの個人版)
    
    // Logs
    const map = {};
    fetchSheetData(ss, DB_SCHEMA.LearningLogs.name).forEach(r => {
      if (r[1] === studentName) {
        const tid = String(r[3]);
        if (!map[tid]) map[tid] = { status: '', reflection: '' };
        if (r[4] && r[4] !== 'メモ') map[tid].status = r[4];
        if (r[5]) map[tid].reflection = r[5];
      }
    });

    // Feedback
    const fbMap = {};
    fetchSheetData(ss, DB_SCHEMA.Feedback.name).forEach(r => {
      if (r[1] === studentName) fbMap[String(r[2])] = String(r[3]);
    });

    // MyTasks
    const myTasks = [];
    fetchSheetData(ss, DB_SCHEMA.MyTasks.name).forEach(r => {
      if (r[1] === studentName) {
        myTasks.push({
          taskId: String(r[0]), title: String(r[2]), desc: String(r[3]), time: Number(r[4]),
          category: 'マイタスク', type: 'challenge', format: 'student'
        });
      }
    });

    // Portfolio Summary
    let portfolio = { summary: "" };
    const pfData = fetchSheetData(ss, DB_SCHEMA.Portfolios.name);
    // 最新のものを探す（簡易的にリスト走査で一致するもの）
    const pfRow = pfData.find(r => r[0] === studentName);
    if (pfRow) portfolio.summary = String(pfRow[1]);

    // Plans
    const plans = {};
    fetchSheetData(ss, DB_SCHEMA.StudentPlans.name).forEach(r => {
      if (r[0] === studentName) {
        plans[r[1]] = safeJsonParse(r[2]);
      }
    });

    return createSuccessResponse({
      json: JSON.stringify({
        progress: map,
        feedback: fbMap,
        myTasks: myTasks,
        portfolio: portfolio,
        plans: plans
      })
    });
  } catch (e) {
    return createErrorResponse(e);
  }
}

/**
 * 分析用データの取得 (Analysis Dashboard)
 * 簡易的に getData をラップするが、将来的に集計ロジックを分離可能
 */
function getAnalysisData() {
  try {
    const ssId = PROPERTIES.getProperty('SS_ID');
    if (!ssId) return createSuccessResponse({ json: JSON.stringify({ tasks: [], counts: {} }) });
    
    const ss = SpreadsheetApp.openById(ssId);
    
    // タスク一覧
    const tasks = fetchSheetData(ss, DB_SCHEMA.UnitMaster.name).map(r => ({ id: r[1], title: r[3] }));
    
    // 完了数集計
    const counts = {};
    fetchSheetData(ss, DB_SCHEMA.LearningLogs.name).forEach(r => {
      if (String(r[4]) === '完了') {
        const tid = String(r[3]);
        counts[tid] = (counts[tid] || 0) + 1;
      }
    });

    return createSuccessResponse({
      json: JSON.stringify({ tasks: tasks, counts: counts })
    });
  } catch (e) {
    return createErrorResponse(e);
  }
}

// ==========================================
//  4. Data Access (Write Operations)
// ==========================================

/**
 * 単元情報のインポート
 */
function importUnitJson(jsonStr) {
  try {
    const ssId = PROPERTIES.getProperty('SS_ID');
    if (!ssId) throw new Error('初期設定未完了');
    
    const ss = SpreadsheetApp.openById(ssId);
    checkAndFixSheets(ss);
    
    const data = JSON.parse(jsonStr);
    const uid = generateUnitId();
    const uInfoStr = JSON.stringify(data.unitInfo || {});
    const totalHours = data.unitInfo?.totalHours || 8;
    
    // UnitMasterへの書き込みデータ作成
    const rows = data.tasks.map(t => [
      uid, 
      t.id || '', 
      t.type || 'must', 
      t.title || '無題', 
      t.description || '', 
      t.estimatedTime || 10, 
      '', 
      t.category || 'まなぶ', 
      t.step || '', 
      t.textbook || '', 
      t.tablet || '', 
      t.print || '',
      (t.prerequisites || []).join(','), 
      t.format || 'student',
      uInfoStr, 
      totalHours
    ]);

    if (rows.length > 0) {
      const sheet = ss.getSheetByName(DB_SCHEMA.UnitMaster.name);
      sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
    }
    return createSuccessResponse({ count: rows.length });
  } catch (e) {
    return createErrorResponse(e);
  }
}

/**
 * 学習ステータスの更新
 */
function updateStatus(studentName, taskId, taskTitle, status, mode, reflection) {
  try {
    const ss = getSpreadsheet();
    const now = new Date();
    
    // ログ追記
    const logSheet = ss.getSheetByName(DB_SCHEMA.LearningLogs.name);
    logSheet.appendRow([
      Utilities.getUuid(), 
      studentName, // studentIdとしてnameを使用
      studentName, 
      taskId, 
      status, 
      reflection || "", 
      now
    ]);

    // ライブ状況更新（「メモ」の場合はステータス表示を変えない）
    if (status !== 'メモ') {
      const liveSheet = ss.getSheetByName(DB_SCHEMA.LiveStatus.name);
      const data = liveSheet.getDataRange().getValues();
      let rIdx = -1;
      
      // 生徒検索
      for (let i = 1; i < data.length; i++) {
        if (data[i][1] === studentName) { rIdx = i + 1; break; }
      }
      
      const displayTitle = taskTitle || taskId;
      if (rIdx > 0) {
        liveSheet.getRange(rIdx, 3, 1, 3).setValues([[displayTitle, mode, now]]);
      } else {
        // 新規追加
        liveSheet.appendRow([studentName, studentName, displayTitle, mode, now]);
      }
    }
    return createSuccessResponse();
  } catch (e) {
    return createErrorResponse(e);
  }
}

/**
 * スタンプ（フィードバック）の送信
 */
function sendFeedback(studentName, taskId, stamp) {
  try {
    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName(DB_SCHEMA.Feedback.name);
    sheet.appendRow([Utilities.getUuid(), studentName, taskId, stamp, new Date()]);
    return createSuccessResponse();
  } catch (e) {
    return createErrorResponse(e);
  }
}

/**
 * マイタスクの追加
 */
function addMyTask(studentName, title, desc, time, unitId) {
  try {
    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName(DB_SCHEMA.MyTasks.name);
    const taskId = "MT" + Utilities.getUuid().substring(0, 8);
    sheet.appendRow([taskId, studentName, title, desc, time, new Date(), unitId || ""]);
    return createSuccessResponse();
  } catch (e) {
    return createErrorResponse(e);
  }
}

/**
 * 単元計画の保存 (Upsert)
 */
function saveStudentPlan(studentName, unitId, planData) {
  try {
    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName(DB_SCHEMA.StudentPlans.name);
    const data = sheet.getDataRange().getValues();
    const json = JSON.stringify(planData);
    const now = new Date();
    
    let rowIndex = -1;
    // 既存レコード検索
    for(let i = 1; i < data.length; i++) {
      if(data[i][0] === studentName && data[i][1] === unitId) { 
        rowIndex = i + 1; 
        break; 
      }
    }
    
    if(rowIndex > 0) {
      sheet.getRange(rowIndex, 3, 1, 2).setValues([[json, now]]);
    } else {
      sheet.appendRow([studentName, unitId, json, now]);
    }
    return createSuccessResponse();
  } catch(e) {
    return createErrorResponse(e);
  }
}

/**
 * 毎時の振り返り保存 (Upsert)
 */
function saveDailyReflection(studentName, unitId, hour, achievement, comment) {
  try {
    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName(DB_SCHEMA.DailyReflections.name);
    const data = sheet.getDataRange().getValues();
    const now = new Date();
    
    let rowIndex = -1;
    for(let i = 1; i < data.length; i++) {
      // 氏名, 単元, 時数 が一致する行を探す
      if(data[i][0] === studentName && data[i][1] === unitId && String(data[i][2]) === String(hour)) { 
        rowIndex = i + 1; 
        break; 
      }
    }
    
    if(rowIndex > 0) {
      sheet.getRange(rowIndex, 4, 1, 4).setValues([[achievement, comment, "", now]]);
    } else {
      sheet.appendRow([studentName, unitId, hour, achievement, comment, "", now]);
    }
    return createSuccessResponse();
  } catch(e) {
    return createErrorResponse(e);
  }
}

/**
 * ポートフォリオ（単元のまとめ）保存 (Upsert)
 */
function savePortfolio(studentName, summary) {
  try {
    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName(DB_SCHEMA.Portfolios.name);
    const data = sheet.getDataRange().getValues();
    const now = new Date();
    
    let rowIndex = -1;
    for(let i = 1; i < data.length; i++) {
      if(data[i][0] === studentName) { rowIndex = i + 1; break; }
    }
    
    if(rowIndex > 0) {
      sheet.getRange(rowIndex, 2, 1, 2).setValues([[summary, now]]);
    } else {
      sheet.appendRow([studentName, summary, now]);
    }
    return createSuccessResponse();
  } catch(e) {
    return createErrorResponse(e);
  }
}

// ==========================================
//  5. Helper Functions
// ==========================================

/**
 * スプレッドシートオブジェクトを取得（IDチェック＆修復含む）
 */
function getSpreadsheet() {
  const ssId = PROPERTIES.getProperty('SS_ID');
  if (!ssId) throw new Error('Database not initialized.');
  const ss = SpreadsheetApp.openById(ssId);
  checkAndFixSheets(ss); // 書き込み時も念のためチェック
  return ss;
}

/**
 * シートデータの取得（ヘッダー行を除く）
 */
function fetchSheetData(ss, sheetName) {
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return [];
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow <= 1) return [];
  // 範囲指定して値を取得
  return sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
}

/**
 * 必須シートの存在確認と修復（列拡張含む）
 */
function checkAndFixSheets(ss) {
  Object.keys(DB_SCHEMA).forEach(key => {
    const def = DB_SCHEMA[key];
    let sheet = ss.getSheetByName(def.name);
    if (!sheet) {
      sheet = ss.insertSheet(def.name);
      sheet.appendRow(def.headers);
    } else {
      // ヘッダー列数が足りない場合は拡張
      const currentCols = sheet.getLastColumn();
      if (currentCols < def.headers.length) {
        const missingHeaders = def.headers.slice(currentCols);
        sheet.getRange(1, currentCols + 1, 1, missingHeaders.length).setValues([missingHeaders]);
      }
    }
  });
}

function generateUnitId() {
  return "U" + Utilities.formatDate(new Date(), "JST", "yyyyMMddHHmm");
}

function formatDate(dateObj) {
  try {
    return Utilities.formatDate(new Date(dateObj), "JST", "HH:mm");
  } catch(e) {
    return "";
  }
}

function safeJsonParse(str) {
  try {
    return JSON.parse(str || '{}');
  } catch(e) {
    return {};
  }
}

function createSuccessResponse(data = {}) {
  return { success: true, ...data };
}

function createErrorResponse(error) {
  console.error(error);
  return { success: false, error: error.toString() };
}
