/**
 * 🧭 みらいコンパス Ver. 1.0 - サーバーサイドプログラム
 * * このファイルは、Googleスプレッドシート（データベース）とのやり取りを担当します。
 * データの保存、読み出し、初期設定などの機能が含まれています。
 */

// ==========================================
//  1. 設定と基本情報 (Configuration)
// ==========================================

const PROPERTIES = PropertiesService.getScriptProperties();
const APP_TITLE = 'みらいコンパス';

/**
 * データベース（スプレッドシート）の設計図
 */
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
  // 新規追加: 児童名簿
  StudentRoster: {
    name: 'StudentRoster',
    headers: ['rosterId', 'classId', 'studentNumber', 'studentName', 'isActive', 'updatedAt']
  }
};

// ==========================================
//  2. 基本機能 (Core Functions)
// ==========================================

function doGet(e) {
  try {
    const template = HtmlService.createTemplateFromFile('index');
    template.mode = 'student'; 
    return template.evaluate()
      .addMetaTag('viewport', 'width=device-width, initial-scale=1.0')
      .setTitle(APP_TITLE)
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .setFaviconUrl('https://drive.google.com/uc?id=1zzJYaALAtpAVIkEG_k5oGoVQu0hIPS7G&.png');
  } catch (error) {
    return HtmlService.createHtmlOutput(`<h2>起動エラー</h2><p>${error.toString()}</p>`);
  }
}

function getAppInitialData() {
  try {
    const ssId = PROPERTIES.getProperty('SS_ID');
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
    checkAndFixSheets(ss);
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
    const inputStr = String(inputPass);
    return createSuccessResponse({ authenticated: (inputStr === currentPass) });
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

function getData() {
  try {
    const ssId = PROPERTIES.getProperty('SS_ID');
    if (!ssId) return createSuccessResponse({ json: JSON.stringify({}) });
    
    const ss = SpreadsheetApp.openById(ssId);
    checkAndFixSheets(ss);

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

    // 名簿データの取得
    const rosterData = fetchSheetData(ss, DB_SCHEMA.StudentRoster.name).map(r => ({
      classId: String(r[1]),
      name: String(r[3]),
      number: String(r[2] || "")
    })).filter(r => r.name); // 名前があるもののみ

    const clsProgress = {};
    fetchSheetData(ss, DB_SCHEMA.LearningLogs.name).forEach(r => {
      const name = String(r[2]); const taskId = String(r[3]); const status = String(r[4]); const reflection = String(r[5] || "");
      if (!clsProgress[name]) clsProgress[name] = {};
      if (!clsProgress[name][taskId]) clsProgress[name][taskId] = { status: '', reflection: '' };
      if (status && status !== 'メモ') clsProgress[name][taskId].status = status;
      if (reflection) clsProgress[name][taskId].reflection = reflection;
    });

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
        achievement: r[3], comment: r[4], check: r[5],
        skills: safeJsonParse(r[8])
      };
    });

    const clsPortfolios = {};
    fetchSheetData(ss, DB_SCHEMA.Portfolios.name).forEach(r => {
      const name = String(r[0]); const uid = String(r[1]);
      if (!clsPortfolios[name]) clsPortfolios[name] = {};
      clsPortfolios[name][uid] = {
        summary: String(r[2]),
        feedback: String(r[5] || ""),
        stamp: String(r[6] || "")
      };
    });

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

function getStudentProgress(studentName, classId, currentUnitId) {
  try {
    const ssId = PROPERTIES.getProperty('SS_ID');
    if (!ssId) return createSuccessResponse({ json: JSON.stringify({}) });
    const ss = SpreadsheetApp.openById(ssId);

    if (classId || currentUnitId) {
      updateLiveStatusMeta(ss, studentName, classId, currentUnitId);
    }

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

function updateStatus(studentName, taskId, taskTitle, status, mode, reflection, classId, currentUnitId) {
  try {
    const ss = getSpreadsheet();
    const now = new Date();
    ss.getSheetByName(DB_SCHEMA.LearningLogs.name).appendRow([
      Utilities.getUuid(), studentName, studentName, taskId, status, reflection || "", now, classId || ""
    ]);

    if (status !== 'メモ') {
      const liveSheet = ss.getSheetByName(DB_SCHEMA.LiveStatus.name);
      const data = liveSheet.getDataRange().getValues();
      let rIdx = -1;
      for (let i = 1; i < data.length; i++) {
        if (data[i][1] === studentName) { rIdx = i + 1; break; }
      }
      const displayTitle = taskTitle || taskId;
      
      if (rIdx > 0) {
        liveSheet.getRange(rIdx, 3, 1, 2).setValues([[displayTitle, mode]]);
        liveSheet.getRange(rIdx, 5).setValue(now);
        if (classId) liveSheet.getRange(rIdx, 8).setValue(classId);
        if (currentUnitId) liveSheet.getRange(rIdx, 6).setValue(currentUnitId);
      } else {
        liveSheet.appendRow([studentName, studentName, displayTitle, mode, now, currentUnitId || "", 1, classId || "", 0, 0]);
      }
    }
    return createSuccessResponse();
  } catch (e) { return createErrorResponse(e); }
}

function updateUnitTask(taskId, updateData) {
  try {
    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName(DB_SCHEMA.UnitMaster.name);
    const data = sheet.getDataRange().getValues();
    let rowIndex = -1;

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][1]) === String(taskId)) {
        rowIndex = i + 1;
        break;
      }
    }

    if (rowIndex > 0) {
      if (updateData.title !== undefined) sheet.getRange(rowIndex, 4).setValue(updateData.title);
      if (updateData.description !== undefined) sheet.getRange(rowIndex, 5).setValue(updateData.description);
      if (updateData.estTime !== undefined) sheet.getRange(rowIndex, 6).setValue(updateData.estTime);
      if (updateData.category !== undefined) sheet.getRange(rowIndex, 8).setValue(updateData.category);
      if (updateData.step !== undefined) sheet.getRange(rowIndex, 9).setValue(updateData.step);
      if (updateData.format !== undefined) sheet.getRange(rowIndex, 14).setValue(updateData.format);
      if (updateData.type !== undefined) sheet.getRange(rowIndex, 3).setValue(updateData.type);
      
      return createSuccessResponse({ message: '更新しました' });
    } else {
      throw new Error('タスクが見つかりません');
    }
  } catch (e) {
    return createErrorResponse(e);
  }
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
  } catch (e) {
    return createErrorResponse(e);
  }
}

function createNewUnit(unitInfo) {
  try {
    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName(DB_SCHEMA.UnitMaster.name);
    
    const unitId = "U" + Utilities.formatDate(new Date(), "JST", "yyyyMMddHHmmss");
    const taskId = "T" + Utilities.getUuid().substring(0, 8);
    
    const infoObj = {
      title: unitInfo.title,
      subject: unitInfo.subject,
      grade: unitInfo.grade,
      totalHours: Number(unitInfo.totalHours)
    };
    
    const newRow = [
      unitId,
      taskId,
      'must', 
      '【導入】' + unitInfo.title, 
      '単元の目標や計画を確認しよう', 
      15, 
      '', 
      '導入', 
      'Step 1', 
      '', 
      '', 
      '', 
      '', 
      'teacher', 
      JSON.stringify(infoObj), 
      infoObj.totalHours 
    ];
    
    sheet.appendRow(newRow);
    
    return createSuccessResponse({ 
      message: '新しい単元を作成しました',
      unitId: unitId
    });
  } catch (e) {
    return createErrorResponse(e);
  }
}

function addUnitTask(unitId, taskData) {
  try {
    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName(DB_SCHEMA.UnitMaster.name);
    const data = sheet.getDataRange().getValues();
    
    let refRow = null;
    for(let i=1; i<data.length; i++) {
      if(String(data[i][0]) === String(unitId)) {
        refRow = data[i];
        break;
      }
    }
    
    if(!refRow) throw new Error('単元が見つかりません');
    
    const taskId = "T" + Utilities.getUuid().substring(0, 8);
    
    const newRow = [
      unitId,
      taskId,
      taskData.type || 'must',
      taskData.title || '無題',
      taskData.description || '',
      taskData.estTime || 15,
      '', 
      taskData.category || 'まなぶ',
      taskData.step || '',
      '', 
      '', 
      '', 
      '', 
      taskData.format || 'student',
      refRow[14], 
      refRow[15] 
    ];
    
    sheet.appendRow(newRow);
    return createSuccessResponse({ taskId: taskId, message: 'タスクを追加しました' });
    
  } catch(e) {
    return createErrorResponse(e);
  }
}

function deleteUnitTask(taskId) {
  try {
    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName(DB_SCHEMA.UnitMaster.name);
    const data = sheet.getDataRange().getValues();
    let rowIndex = -1;

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][1]) === String(taskId)) {
        rowIndex = i + 1;
        break;
      }
    }

    if (rowIndex > 0) {
      sheet.deleteRow(rowIndex);
      return createSuccessResponse({ message: '削除しました' });
    } else {
      throw new Error('タスクが見つかりません');
    }
  } catch (e) {
    return createErrorResponse(e);
  }
}

/**
 * 先生用：クラス名簿を保存する関数
 * 既存のクラスデータを洗い替えます。
 */
function saveClassRoster(classId, nameList) {
  try {
    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName(DB_SCHEMA.StudentRoster.name);
    const data = sheet.getDataRange().getValues();
    
    // 既存の該当クラスデータを削除（逆順にループして削除）
    for (let i = data.length - 1; i >= 1; i--) {
      if (String(data[i][1]) === classId) {
        sheet.deleteRow(i + 1);
      }
    }
    
    // 新しいデータを追加
    const rows = nameList.map(name => [
      Utilities.getUuid(), // rosterId
      classId,
      '', // studentNumber (今回は省略)
      name,
      true, // isActive
      new Date()
    ]);
    
    if (rows.length > 0) {
      sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
    }
    
    return createSuccessResponse({ message: `${classId}の名簿を更新しました（${rows.length}名）` });
  } catch (e) {
    return createErrorResponse(e);
  }
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

function sendFeedback(studentName, taskId, stamp, classId) {
  try {
    const ss = getSpreadsheet();
    ss.getSheetByName(DB_SCHEMA.Feedback.name).appendRow([Utilities.getUuid(), studentName, taskId, stamp, new Date(), classId || ""]);
    return createSuccessResponse();
  } catch (e) { return createErrorResponse(e); }
}

function addMyTask(studentName, title, desc, time, unitId, classId) {
  try {
    const ss = getSpreadsheet();
    const taskId = "MT" + Utilities.getUuid().substring(0, 8);
    ss.getSheetByName(DB_SCHEMA.MyTasks.name).appendRow([taskId, studentName, title, desc, time, new Date(), unitId || "", classId || ""]);
    return createSuccessResponse({ taskId: taskId });
  } catch (e) { return createErrorResponse(e); }
}

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

/**
 * ポートフォリオのフィードバックを一括保存する関数
 */
function saveAllPortfolios(feedbackList) {
  try {
    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName(DB_SCHEMA.Portfolios.name);
    // 全データを取得してメモリ上で照合
    const data = sheet.getDataRange().getValues();
    const headers = data[0]; // ヘッダー行
    
    // 行のインデックスをマッピング (Key: "名前_単元ID", Value: 行番号(0始まり))
    const rowMap = new Map();
    for (let i = 1; i < data.length; i++) {
      const key = data[i][0] + "_" + data[i][1];
      rowMap.set(key, i);
    }

    const rowsToAppend = [];
    const now = new Date();

    // 入力されたリストをループ処理
    feedbackList.forEach(item => {
      const key = item.studentName + "_" + item.unitId;
      if (rowMap.has(key)) {
        // 既存行がある場合：配列の値を更新
        const rowIndex = rowMap.get(key);
        data[rowIndex][5] = item.feedback; // feedbackカラム
        data[rowIndex][6] = item.stamp;    // stampカラム
      } else {
        // 新規の場合：追加用配列に作成
        rowsToAppend.push([
          item.studentName, 
          item.unitId, 
          "", // summary (未提出でもフィードバック先行入力可とする)
          now, 
          "", // classId (必要なら補完)
          item.feedback, 
          item.stamp
        ]);
      }
    });

    // 1. 更新分をシートに一括書き戻し
    if (data.length > 1) {
      sheet.getRange(1, 1, data.length, data[0].length).setValues(data);
    }

    // 2. 新規分を追記
    if (rowsToAppend.length > 0) {
      sheet.getRange(sheet.getLastRow() + 1, 1, rowsToAppend.length, rowsToAppend[0].length).setValues(rowsToAppend);
    }

    return createSuccessResponse({ message: '一括保存しました' });
  } catch (e) {
    return createErrorResponse(e);
  }
}

function importUnitJson(jsonStr) {
  try {
    const ssId = PROPERTIES.getProperty('SS_ID');
    if (!ssId) throw new Error('初期設定未完了');
    const ss = SpreadsheetApp.openById(ssId);
    checkAndFixSheets(ss);
    
    const data = JSON.parse(jsonStr);

    // データの正規化（揺らぎ吸収）
    if (!data.unitInfo) data.unitInfo = {};
    
    // 1. 単元名の読み替え (unitName -> title)
    if (!data.unitInfo.title && data.unitInfo.unitName) {
      data.unitInfo.title = data.unitInfo.unitName;
    }
    
    // 2. 学年の型変換 (数値 -> 文字列)
    if (data.unitInfo.grade && typeof data.unitInfo.grade !== 'string') {
      data.unitInfo.grade = String(data.unitInfo.grade);
    }
    
    // 3. titleがない場合のフォールバック
    if (!data.unitInfo.title) {
      data.unitInfo.title = "無題の単元";
    }

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
      { key: 'MyTasks', colUnitId: 6 },
      { key: 'StudentPlans', colUnitId: 1 },
      { key: 'DailyReflections', colUnitId: 1 },
      { key: 'Portfolios', colUnitId: 1 }
    ];
    let movedCount = 0;
    
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
    
    const delSheet = archiveSs.getSheetByName('シート1');
    if(delSheet) archiveSs.deleteSheet(delSheet);

    return createSuccessResponse({ 
      message: `アーカイブ完了: ${movedCount}件のデータを移動しました。\nファイル名: ${archiveName}`,
      url: archiveSs.getUrl()
    });

  } catch(e) { return createErrorResponse(e); }
}

/**
 * 単元の基本情報（タイトル、めあて、概要など）を更新する関数
 */
function updateUnitBasicInfo(unitId, infoData) {
  try {
    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName(DB_SCHEMA.UnitMaster.name);
    const data = sheet.getDataRange().getValues();
    
    // ヘッダーを除外して全行をチェック
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(unitId)) {
        // 現在のunitInfoを取得してマージ
        let currentInfo = safeJsonParse(data[i][14]);
        
        // 更新項目を反映
        if (infoData.title !== undefined) currentInfo.title = infoData.title;
        if (infoData.subject !== undefined) currentInfo.subject = infoData.subject;
        if (infoData.grade !== undefined) currentInfo.grade = infoData.grade;
        if (infoData.goal !== undefined) currentInfo.goal = infoData.goal;
        if (infoData.description !== undefined) currentInfo.description = infoData.description;
        
        // unitInfoカラム(列15)を更新
        sheet.getRange(i + 1, 15).setValue(JSON.stringify(currentInfo));
      }
    }
    return createSuccessResponse({ message: '単元情報を更新しました' });
  } catch (e) {
    return createErrorResponse(e);
  }
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

function createSuccessResponse(data = {}) {
 return { success: true, ...data }; 
}

function createErrorResponse(error) { 
 console.error(error); return { success: false, error: error.toString() }; 
}

function safeJsonParse(str) {
 try { return JSON.parse(str || '{}'); } catch(e) { return {}; } 
}

function formatDate(d) { 
 try { return Utilities.formatDate(new Date(d), "JST", "HH:mm"); } catch(e) { return ""; } 
}

/**
 * [追加] 保存されたカスタムAIプロンプトを取得する
 */
function getCustomAiPrompt() {
  try {
    const prompt = PROPERTIES.getProperty('CUSTOM_AI_PROMPT');
    return createSuccessResponse({ prompt: prompt });
  } catch (e) {
    return createErrorResponse(e);
  }
}

/**
 * [追加] カスタムAIプロンプトを保存する
 */
function saveCustomAiPrompt(text) {
  try {
    PROPERTIES.setProperty('CUSTOM_AI_PROMPT', text);
    return createSuccessResponse();
  } catch (e) {
    return createErrorResponse(e);
  }
}
