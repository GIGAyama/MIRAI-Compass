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
    headers: ['feedbackId', 'studentName', 'taskId', 'stamp', 'timestamp', 'classId']
  },
  MyTasks: {
    name: 'MyTasks',
    headers: ['taskId', 'studentName', 'title', 'description', 'estTime', 'created_at', 'unitId', 'classId']
  },
  StudentPlans: {
    name: 'StudentPlans',
    headers: ['studentName', 'unitId', 'planData', 'lastUpdate', 'classId']
  },
  DailyReflections: {
    name: 'DailyReflections',
    headers: ['studentName', 'unitId', 'hour', 'achievement', 'comment', 'teacherCheck', 'timestamp', 'classId', 'skills']
  },
  Portfolios: {
    name: 'Portfolios',
    headers: ['studentName', 'unitId', 'summary', 'lastUpdate', 'classId', 'feedback', 'stamp']
  },
  ClassSchedule: {
    name: 'ClassSchedule',
    headers: ['scheduleId', 'classId', 'date', 'startTime', 'endTime', 'unitId', 'hour', 'message', 'createdAt']
  },
  StudentRoster: {
    name: 'StudentRoster',
    headers: ['rosterId', 'classId', 'studentNumber', 'studentName', 'isActive', 'updatedAt']
  }
};

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
      .addMetaTag('viewport', 'width=device-width, initial-scale=1.0')
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

/**
 * 外部（Passportなど）からのデータ受信
 * POSTリクエストを受け取り、LiveStatusを更新します。
 */
function doPost(e) {
  const output = ContentService.createTextOutput();
  output.setMimeType(ContentService.MimeType.JSON);

  try {
    if (!e.postData || !e.postData.contents) {
      throw new Error("No Data received");
    }

    const json = JSON.parse(e.postData.contents);
    
    // 排他制御ロックを取得（最大10秒待機）
    const lock = LockService.getScriptLock();
    if (lock.tryLock(10000)) {
      try {
        if (json.action === 'syncMode' || json.action === 'syncStatus') {
          updateLiveStatusFromPassport(json);
        }
      } finally {
        lock.releaseLock();
      }
    }
    
    output.setContent(JSON.stringify({ success: true }));

  } catch (err) {
    output.setContent(JSON.stringify({ success: false, error: err.message }));
  }
  
  return output;
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
    // 初期パスワード設定（未設定時のみ）
    if (!PROPERTIES.getProperty('TEACHER_PASS')) {
      PROPERTIES.setProperty('TEACHER_PASS', 'admin');
    }
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
    
    PROPERTIES.setProperty('TEACHER_PASS', 'admin');
    return createSuccessResponse({ message: 'システムを初期化しました。初期パスワードは admin です。' });
  } catch (e) {
    return createErrorResponse(e);
  }
}

function verifyPassword(inputPass) {
  try {
    const currentPass = String(PROPERTIES.getProperty('TEACHER_PASS') || 'admin');
    return createSuccessResponse({ authenticated: (String(inputPass) === currentPass) });
  } catch (e) {
    return createErrorResponse(e);
  }
}

function changeTeacherPassword(newPass) {
  try {
    if (!newPass) throw new Error("パスワードが空です");
    PROPERTIES.setProperty('TEACHER_PASS', String(newPass));
    return createSuccessResponse();
  } catch (e) {
    return createErrorResponse(e);
  }
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

    // [改修] 名簿データの取得：出席番号と在籍状況も取得する
    const rosterData = fetchSheetData(ss, DB_SCHEMA.StudentRoster.name).map(r => ({
      classId: String(r[1]), 
      name: String(r[3]), 
      number: r[2] !== "" ? Number(r[2]) : 999, // 出席番号（空なら末尾へ）
      isActive: (r[4] === true || r[4] === "TRUE" || r[4] === "") // 在籍状況（空ならTrue扱い）
    })).filter(r => r.name);

    // 学習ログを集計（最新ステータスのみ）
    const clsProgress = {};
    fetchSheetData(ss, DB_SCHEMA.LearningLogs.name).forEach(r => {
      const name = String(r[2]); const taskId = String(r[3]); const status = String(r[4]); const reflection = String(r[5] || "");
      if (!clsProgress[name]) clsProgress[name] = {};
      if (!clsProgress[name][taskId]) clsProgress[name][taskId] = { status: '', reflection: '' };
      if (status && status !== 'メモ') clsProgress[name][taskId].status = status;
      if (reflection) clsProgress[name][taskId].reflection = reflection;
    });

    // その他のデータを取得
    const clsFeedback = {};
    fetchSheetData(ss, DB_SCHEMA.Feedback.name).forEach(r => {
      const name = String(r[1]); const taskId = String(r[2]); const stamp = String(r[3]);
      if (!clsFeedback[name]) clsFeedback[name] = {};
      clsFeedback[name][taskId] = stamp;
    });

    const allMyTasks = {};
    fetchSheetData(ss, DB_SCHEMA.MyTasks.name).forEach(r => {
      const name = String(r[1]);
      if (!allMyTasks[name]) allMyTasks[name] = [];
      allMyTasks[name].push({
        taskId: String(r[0]), title: String(r[2]), desc: String(r[3]), time: Number(r[4]),
        category: 'マイタスク', type: 'challenge', format: 'student', unitId: String(r[6] || '')
      });
    });

    const clsPlans = {};
    fetchSheetData(ss, DB_SCHEMA.StudentPlans.name).forEach(r => {
      const name = String(r[0]); const uid = String(r[1]);
      if (!clsPlans[name]) clsPlans[name] = {};
      clsPlans[name][uid] = safeJsonParse(r[2]);
    });

    const clsReflections = {};
    fetchSheetData(ss, DB_SCHEMA.DailyReflections.name).forEach(r => {
      const name = String(r[0]); const uid = String(r[1]); const hour = String(r[2]);
      if (!clsReflections[name]) clsReflections[name] = {};
      if (!clsReflections[name][uid]) clsReflections[name][uid] = {};
      clsReflections[name][uid][hour] = { 
        achievement: r[3], comment: r[4], check: r[5], skills: safeJsonParse(r[8])
      };
    });

    const clsPortfolios = {};
    fetchSheetData(ss, DB_SCHEMA.Portfolios.name).forEach(r => {
      const name = String(r[0]); const uid = String(r[1]);
      if (!clsPortfolios[name]) clsPortfolios[name] = {};
      clsPortfolios[name][uid] = {
        summary: String(r[2]), feedback: String(r[5] || ""), stamp: String(r[6] || "")
      };
    });

    // 今日のスケジュールのみ取得
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
        unit: unitData, live: liveData, roster: rosterData, progress: clsProgress, feedback: clsFeedback,
        myTasks: allMyTasks, plans: clsPlans, dailyReflections: clsReflections,
        portfolios: clsPortfolios, schedules: schedules
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
    const ss = SpreadsheetApp.openById(ssId);
    
    const liveData = fetchSheetData(ss, DB_SCHEMA.LiveStatus.name).map(r => ({
      id: String(r[0]), name: String(r[1]), task: String(r[2]), mode: String(r[3]),
      time: r[4] ? formatDate(r[4]) : '', currentUnitId: String(r[5] || ''),
      currentHour: Number(r[6] || 1), classId: String(r[7] || ''),
      x: Number(r[8]) || 0, y: Number(r[9]) || 0
    }));

    return createSuccessResponse({ live: liveData });
  } catch (e) {
    return createErrorResponse(e);
  }
}

/**
 * 児童個人用データの取得（ログイン時）
 */
function getStudentProgress(studentName, classId, currentUnitId) {
  try {
    const ssId = PROPERTIES.getProperty('SS_ID');
    if (!ssId) return createSuccessResponse({ json: JSON.stringify({}) });
    const ss = SpreadsheetApp.openById(ssId);

    // ログイン時に現在の単元・クラス情報をLiveStatusに書き込む
    if (classId || currentUnitId) {
      updateLiveStatusMeta(ss, studentName, classId, currentUnitId);
    }

    // 必要なデータのみ抽出して返す
    const map = {}; 
    fetchSheetData(ss, DB_SCHEMA.LearningLogs.name).forEach(r => {
      if (r[2] === studentName) {
        const tid = String(r[3]);
        if (!map[tid]) map[tid] = { status: '', reflection: '' };
        if (r[4] && r[4] !== 'メモ') map[tid].status = r[4];
        if (r[5]) map[tid].reflection = r[5];
      }
    });

    const fbMap = {}; 
    fetchSheetData(ss, DB_SCHEMA.Feedback.name).forEach(r => {
      if (r[1] === studentName) fbMap[String(r[2])] = String(r[3]);
    });

    const myTasks = []; 
    fetchSheetData(ss, DB_SCHEMA.MyTasks.name).forEach(r => {
      if (r[1] === studentName) {
        myTasks.push({
          taskId: String(r[0]), title: String(r[2]), desc: String(r[3]), time: Number(r[4]),
          category: 'マイタスク', type: 'challenge', format: 'student', unitId: String(r[6] || '')
        });
      }
    });

    const plans = {}; 
    fetchSheetData(ss, DB_SCHEMA.StudentPlans.name).forEach(r => {
      if (r[0] === studentName) plans[r[1]] = safeJsonParse(r[2]);
    });

    const reflections = {}; 
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

    let portfolioData = { summary: "", feedback: "", stamp: "" };
    const pSheet = ss.getSheetByName(DB_SCHEMA.Portfolios.name);
    if(pSheet) {
      const pData = pSheet.getDataRange().getValues();
      for(let i=1; i<pData.length; i++) {
        if(pData[i][0] === studentName && String(pData[i][1]) === String(currentUnitId)) {
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
        portfolio: portfolioData, plans: plans, reflections: reflections
      })
    });
  } catch (e) { return createErrorResponse(e); }
}

/**
 * 内部関数: 児童のLiveStatus（クラス、単元）を更新
 */
function updateLiveStatusMeta(ss, name, classId, unitId) {
  const lock = LockService.getScriptLock();
  if (lock.tryLock(5000)) { // 5秒待機
    try {
      const sheet = ss.getSheetByName(DB_SCHEMA.LiveStatus.name);
      const data = sheet.getDataRange().getValues();
      for (let i = 1; i < data.length; i++) {
        if (data[i][1] === name) {
          if (unitId) sheet.getRange(i + 1, 6).setValue(unitId);
          if (classId) sheet.getRange(i + 1, 8).setValue(classId);
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
  if (!lock.tryLock(10000)) throw new Error("サーバーが混み合っています。もう一度お試しください。");

  try {
    const ss = getSpreadsheet();
    const now = new Date();
    
    // ログ履歴に追加
    ss.getSheetByName(DB_SCHEMA.LearningLogs.name).appendRow([
      Utilities.getUuid(), studentName, studentName, taskId, status, reflection || "", now, classId || ""
    ]);

    // LiveStatus（現在の状態）を更新
    if (status !== 'メモ') {
      const liveSheet = ss.getSheetByName(DB_SCHEMA.LiveStatus.name);
      const data = liveSheet.getDataRange().getValues();
      let rIdx = -1;
      for (let i = 1; i < data.length; i++) {
        if (data[i][1] === studentName) { rIdx = i + 1; break; }
      }
      const displayTitle = taskTitle || taskId;
      
      if (rIdx > 0) {
        // 既存行を更新
        liveSheet.getRange(rIdx, 3, 1, 2).setValues([[displayTitle, mode]]);
        liveSheet.getRange(rIdx, 5).setValue(now);
        if (classId) liveSheet.getRange(rIdx, 8).setValue(classId);
        if (currentUnitId) liveSheet.getRange(rIdx, 6).setValue(currentUnitId);
      } else {
        // 新規追加
        liveSheet.appendRow([studentName, studentName, displayTitle, mode, now, currentUnitId || "", 1, classId || "", 0, 0]);
      }
    }
    return createSuccessResponse();
  } catch (e) { 
    return createErrorResponse(e); 
  } finally {
    lock.releaseLock();
  }
}

/**
 * Passportからの通知でLiveStatusを更新（内部処理用）
 */
function updateLiveStatusFromPassport(data) {
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName(DB_SCHEMA.LiveStatus.name);
  if (!sheet) return;

  const dataValues = sheet.getDataRange().getValues();
  const now = new Date();
  
  const targetId = String(data.studentId);
  const targetName = data.studentName;
  
  let rowIndex = -1;
  // IDまたは名前で検索
  for (let i = 1; i < dataValues.length; i++) {
    if (String(dataValues[i][0]) === targetId) { rowIndex = i + 1; break; }
  }
  if (rowIndex === -1 && targetName) {
    for (let i = 1; i < dataValues.length; i++) {
      if (String(dataValues[i][1]) === targetName) { rowIndex = i + 1; break; }
    }
  }

  const currentTask = data.taskTitle || data.taskId || "";
  const mode = data.mode || (rowIndex > 0 ? dataValues[rowIndex-1][3] : "normal");
  
  if (rowIndex > 0) {
    sheet.getRange(rowIndex, 1).setValue(targetId);
    if (targetName) sheet.getRange(rowIndex, 2).setValue(targetName);
    if (currentTask) sheet.getRange(rowIndex, 3).setValue(currentTask);
    if (data.mode) sheet.getRange(rowIndex, 4).setValue(mode);
    sheet.getRange(rowIndex, 5).setValue(now);
  } else {
    sheet.appendRow([targetId, targetName || "Unknown", currentTask, mode, now, "", 1, "", 0, 0]);
  }
}

// --- 先生用管理機能（ロック推奨） ---

function updateUnitTask(taskId, updateData) {
  const lock = LockService.getScriptLock();
  if (lock.tryLock(5000)) {
    try {
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
  try {
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
  } catch (e) { return createErrorResponse(e); }
}

function updateUnitBasicInfo(unitId, infoData) {
  try {
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
  } catch (e) { return createErrorResponse(e); }
}

function updateUnitTotalHours(unitId, newTotalHours) {
  try {
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
  } catch (e) { return createErrorResponse(e); }
}


// [改修] 名簿保存処理をオブジェクト配列対応に変更

function saveClassRoster(classId, studentList) {
  const lock = LockService.getScriptLock();
  if (lock.tryLock(10000)) {
    try {
      const ss = getSpreadsheet();
      const sheet = ss.getSheetByName(DB_SCHEMA.StudentRoster.name);
      const data = sheet.getDataRange().getValues();
      
      // 既存の該当クラスデータを削除（逆順ループで安全に削除）
      for (let i = data.length - 1; i >= 1; i--) {
        if (String(data[i][1]) === classId) { sheet.deleteRow(i + 1); }
      }
      
      // studentList: [{name: '...', number: 1, isActive: true}, ...]
      // 従来の文字列リストにも対応（後方互換性）
      const rows = studentList.map(s => {
        if (typeof s === 'string') {
          return [Utilities.getUuid(), classId, '', s, true, new Date()];
        } else {
          return [Utilities.getUuid(), classId, s.number, s.name, s.isActive, new Date()];
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

// --- 生徒のアクション保存（課題追加、計画、振り返り） ---

function addMyTask(studentName, title, desc, time, unitId, classId) {
  try {
    const ss = getSpreadsheet();
    const taskId = "MT" + Utilities.getUuid().substring(0, 8);
    ss.getSheetByName(DB_SCHEMA.MyTasks.name).appendRow([taskId, studentName, title, desc, time, new Date(), unitId || "", classId || ""]);
    return createSuccessResponse({ taskId: taskId });
  } catch (e) { return createErrorResponse(e); }
}

function saveStudentPlan(studentName, unitId, planData, classId) {
  const lock = LockService.getScriptLock();
  if (lock.tryLock(5000)) {
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
    } catch(e) { return createErrorResponse(e); } finally { lock.releaseLock(); }
  } else { return createErrorResponse(new Error("Timeout")); }
}

function saveDailyReflection(studentName, unitId, hour, achievement, comment, classId, skills) {
  const lock = LockService.getScriptLock();
  if (lock.tryLock(5000)) {
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
    } catch(e) { return createErrorResponse(e); } finally { lock.releaseLock(); }
  } else { return createErrorResponse(new Error("Timeout")); }
}

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

// --- フィードバック・評価 ---

function sendFeedback(studentName, taskId, stamp, classId) {
  try {
    const ss = getSpreadsheet();
    ss.getSheetByName(DB_SCHEMA.Feedback.name).appendRow([Utilities.getUuid(), studentName, taskId, stamp, new Date(), classId || ""]);
    return createSuccessResponse();
  } catch (e) { return createErrorResponse(e); }
}

function savePortfolioFeedback(studentName, unitId, feedback, stamp) {
  try {
    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName(DB_SCHEMA.Portfolios.name);
    const data = sheet.getDataRange().getValues();
    let rowIndex = -1;
    for(let i = 1; i < data.length; i++) {
      if(data[i][0] === studentName && String(data[i][1]) === String(unitId)) { rowIndex = i + 1; break; }
    }
    
    if(rowIndex > 0) {
      sheet.getRange(rowIndex, 6, 1, 2).setValues([[feedback, stamp]]);
    } else {
      const now = new Date();
      sheet.appendRow([studentName, unitId, "", now, "", feedback, stamp]);
    }
    return createSuccessResponse();
  } catch(e) { return createErrorResponse(e); }
}

function saveAllPortfolios(feedbackList) {
  const lock = LockService.getScriptLock();
  if (lock.tryLock(10000)) {
    try {
      const ss = getSpreadsheet();
      const sheet = ss.getSheetByName(DB_SCHEMA.Portfolios.name);
      const data = sheet.getDataRange().getValues();
      const rowMap = new Map();
      for (let i = 1; i < data.length; i++) {
        const key = data[i][0] + "_" + data[i][1];
        rowMap.set(key, i);
      }

      const rowsToAppend = [];
      const now = new Date();

      feedbackList.forEach(item => {
        const key = item.studentName + "_" + item.unitId;
        if (rowMap.has(key)) {
          const rowIndex = rowMap.get(key);
          data[rowIndex][5] = item.feedback; 
          data[rowIndex][6] = item.stamp;    
        } else {
          rowsToAppend.push([item.studentName, item.unitId, "", now, "", item.feedback, item.stamp]);
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

// JSONインポート、アーカイブ、Passport連携など

function importUnitJson(jsonStr) {
  try {
    const ssId = PROPERTIES.getProperty('SS_ID');
    if (!ssId) throw new Error('初期設定未完了');
    const ss = SpreadsheetApp.openById(ssId);
    checkAndFixSheets(ss);
    
    const data = JSON.parse(jsonStr);
    if (!data.unitInfo) data.unitInfo = {};
    if (!data.unitInfo.title && data.unitInfo.unitName) data.unitInfo.title = data.unitInfo.unitName;
    if (data.unitInfo.grade && typeof data.unitInfo.grade !== 'string') data.unitInfo.grade = String(data.unitInfo.grade);
    if (!data.unitInfo.title) data.unitInfo.title = "無題の単元";

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

function archiveUnitData(unitId, unitTitle) {
  try {
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
  } catch(e) { return createErrorResponse(e); }
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
    PROPERTIES.setProperty('CUSTOM_AI_PROMPT', text);
    return createSuccessResponse();
  } catch (e) { return createErrorResponse(e); }
}

function getPassportDbId() { return PROPERTIES.getProperty('PASSPORT_DB_ID') || ""; }
function getPassportUrl() { return PROPERTIES.getProperty('PASSPORT_WEB_APP_URL') || ""; }

function savePassportConfig(dbId, url) {
  PROPERTIES.setProperty('PASSPORT_DB_ID', dbId);
  PROPERTIES.setProperty('PASSPORT_WEB_APP_URL', url);
  return true;
}

function sendUnitPlanToPassport_DirectDB(unitId) {
  try {
    const passportSsId = PROPERTIES.getProperty('PASSPORT_DB_ID');
    const passportUrl = PROPERTIES.getProperty('PASSPORT_WEB_APP_URL'); 
    
    if (!passportSsId || !passportUrl) throw new Error("みらいパスポートの連携設定（ID/URL）が完了していません。");

    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName(DB_SCHEMA.UnitMaster.name);
    const data = sheet.getDataRange().getValues();
    
    let unitInfo = { unitName: "", grade: "" };
    const tasks = [];
    
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(unitId)) {
        if (!unitInfo.unitName) {
           const infoJson = safeJsonParse(data[i][14]);
           unitInfo.unitName = infoJson.unitName || infoJson.title || "無題の単元";
           unitInfo.grade = infoJson.grade || "";
        }
        tasks.push({
          taskId: String(data[i][1]),
          title: String(data[i][3]),
          description: String(data[i][4])
        });
      }
    }
    
    if (tasks.length === 0) throw new Error("対象の単元データが見つかりません。");

    const importData = {
      unitName: unitInfo.unitName, grade: unitInfo.grade,
      tasks: tasks, timestamp: new Date().toISOString()
    };

    const transactionId = Utilities.getUuid();
    const passportSs = SpreadsheetApp.openById(passportSsId);
    let queueSheet = passportSs.getSheetByName('ImportQueue');
    if (!queueSheet) {
      queueSheet = passportSs.insertSheet('ImportQueue');
      queueSheet.appendRow(['transactionId', 'dataJson', 'createdAt']);
    }
    queueSheet.appendRow([transactionId, JSON.stringify(importData), new Date()]);

    const openUrl = `${passportUrl}?page=wizard&importId=${transactionId}`;
    return createSuccessResponse({
      message: "連携データを送信しました。パスポートを開きます。",
      passportUrl: openUrl,
      taskIds: ""
    });
  } catch (e) { return createErrorResponse(e); }
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
