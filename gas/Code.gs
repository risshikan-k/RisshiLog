// =========================================
// 基本設定
// =========================================
// 【重要】塾のアカウントに移行したらここを書き換える
const SHEET_ID = '1eYvBli5lOdl991ZSkinb3Bvhrrw3KqGeQLzfedHyITY'; 

const STATS_SHEET_NAME = '統計'; 
const MONITOR_SHEET_NAME = 'モニター';
const USER_SHEET_NAME = '名簿';
const FEEDBACK_SHEET_NAME = '意見箱';

// =========================================
// メインの受信処理（doGet）
// =========================================
function doGet(e) {
  let app = SpreadsheetApp.openById(SHEET_ID);
  let mode = e.parameter.mode;
  
  if (mode === 'latest') return getLatestLog(app);
  if (mode === 'get_all_status') return getAllStatus(app);
  if (mode === 'toggle_status') return toggleStatus(app, e.parameter.idm);
  
  // ★変更：nickname パラメーターを受け取って updateUser に渡す
  if (mode === 'update_user') return updateUser(app, e.parameter.idm, e.parameter.name, e.parameter.grade, e.parameter.yomi, e.parameter.nickname);
  
  if (mode === 'delete_user') return deleteUser(app, e.parameter.idm);
  if (mode === 'force_exit_all') return processForceExitAll(app, null); 
  if (mode === 'submit_feedback') return submitFeedback(app, e.parameter.idm, e.parameter.name, e.parameter.message);

  // doGetの分岐(mode判定)の中に以下を追加
  if (mode === 'get_history') return getUserHistory(app, e.parameter.idm);
  if (mode === 'update_history') return updateHistoryRow(app, e.parameter.row, e.parameter.entry, e.parameter.exit);

  /**
   * 指定したユーザーの直近3回分の履歴を取得する
   */
  function getUserHistory(app, idm) {
    let sheet = getYearlySheet(app);
    let data = sheet.getDataRange().getValues();
    let history = [];
    
    // 最新（下）から探索して最大3件取得
    for (let i = data.length - 1; i >= 1; i--) {
      if (String(data[i][0]).trim() === idm) {
        history.push({
          row: i + 1,
          date: formatDate(data[i][2]),
          entry: formatTime(data[i][3]),
          exit: formatTime(data[i][4])
        });
        if (history.length >= 3) break;
      }
    }
    return responseJSON(history);
  }

  /**
   * 履歴の時間を更新し、滞在時間を再計算する
   */
  function updateHistoryRow(app, row, entry, exit) {
    let sheet = getYearlySheet(app);
    let rowNum = parseInt(row);
    
    sheet.getRange(rowNum, 4).setValue(entry); // 入室
    sheet.getRange(rowNum, 5).setValue(exit);  // 退出
    
    // 滞在時間の再計算
    if (entry && exit) {
      let dateStr = formatDate(sheet.getRange(rowNum, 3).getValue());
      let t1 = new Date(dateStr + " " + entry);
      let t2 = new Date(dateStr + " " + exit);
      if (!isNaN(t1) && !isNaN(t2)) {
        let diff = Math.round((t2 - t1) / (1000 * 60) * 10) / 10;
        sheet.getRange(rowNum, 6).setValue(diff > 0 ? diff : 0);
      }
    }
    
    syncStatsData(app); // 統計を即座に同期
    return responseJSON({ success: true });
  }

  let idm = e.parameter.id;
  if (idm) {
    // ★変更：初期値に nickname を追加
    let result = { found: false, name: "", yomi: "", nickname: "", grade: "未設定", totalDays: 0, totalTime: 0, lastDate: "-", monthlyDays: 0, monthlyTime: 0 };
    
    // 1. 名簿から基本情報を取得
    let userSheet = getUserSheet(app);
    if (userSheet) {
      let users = userSheet.getDataRange().getValues();
      for (let i = 1; i < users.length; i++) {
        if (String(users[i][0]).trim() === idm) {
          result.found = true;
          result.name = String(users[i][1] || "");
          let gradeVal = String(users[i][2] || "").trim();
          result.grade = gradeVal !== "" ? gradeVal : "未設定";
          result.yomi = String(users[i][3] || "").trim();
          
          // ★追加：F列（インデックス5）からニックネームを読み取る
          result.nickname = String(users[i][5] || "").trim();
          break; 
        }
      }
    }

    // 2. 統計シートからデータを取得
    let statsSheet = getStatsSheet(app);
    if (statsSheet) {
      let stats = statsSheet.getDataRange().getValues();
      
      let nowMonth = new Date().getMonth() + 1;
      let offset = nowMonth >= 4 ? nowMonth - 4 : nowMonth + 8;
      let daysColIndex = 5 + offset * 2; 
      let timeColIndex = 6 + offset * 2; 

      for (let i = 1; i < stats.length; i++) {
        if (String(stats[i][0]).trim() === idm) {
          result.found = true;
          result.totalDays = stats[i][2] || 0;
          result.totalTime = stats[i][3] || 0;
          result.lastDate = formatDate(stats[i][4]);
          
          result.monthlyDays = stats[i][daysColIndex] || 0;
          result.monthlyTime = stats[i][timeColIndex] || 0;
          break; 
        }
      }
    }
    
    return responseJSON(result);
  }
  return responseJSON({ found: false });
}

// =========================================
// ダッシュボード用：全員の状況を取得
// =========================================
function getAllStatus(app) {
  let userSheet = getUserSheet(app);
  let monthlySheet = getYearlySheet(app);
  
  if (!userSheet || !monthlySheet) return responseJSON([]);

  let users = userSheet.getDataRange().getValues();
  let statusMap = {};
  
  for (let i = 1; i < users.length; i++) {
    let idm = String(users[i][0]).trim();
    let grade = users[i][2] ? String(users[i][2]) : "";
    let yomi = users[i][3] ? String(users[i][3]) : "";
    let url = users[i][4] ? String(users[i][4]) : "";
    
    if (idm) statusMap[idm] = { idm: idm, name: users[i][1], grade: grade, yomi: yomi, url: url, status: 'out', time: '-' };
  }

  let logs = monthlySheet.getDataRange().getValues();
  for (let i = logs.length - 1; i >= 1; i--) {
    let idm = String(logs[i][0]).trim();
    if (statusMap[idm]) {
      if (statusMap[idm].processed) continue; 

      if (!logs[i][4] || String(logs[i][4]).trim() === "") {
        statusMap[idm].status = 'in';
        statusMap[idm].time = formatTime(logs[i][3]);
      }
      statusMap[idm].processed = true; 
    }
  }

  let resultList = Object.values(statusMap).sort((a, b) => {
    if (a.status === 'in' && b.status === 'out') return -1;
    if (a.status === 'out' && b.status === 'in') return 1;
    return 0;
  });

  return responseJSON(resultList);
}

// =========================================
// ダッシュボード・マイページ用：ユーザー情報の変更
// =========================================
// ★変更：引数に newNickname を追加
function updateUser(app, targetIdm, newName, newGrade, newYomi, newNickname) {
  targetIdm = String(targetIdm).trim();
  
  // 1. 名簿シートの更新
  let userSheet = getUserSheet(app);
  if (userSheet) {
    let isExist = false;
    let users = userSheet.getDataRange().getValues();
    for (let i = 1; i < users.length; i++) {
      if (String(users[i][0]).trim() === targetIdm) {
        userSheet.getRange(i + 1, 2).setValue(newName);
        userSheet.getRange(i + 1, 3).setValue(newGrade);
        if (newYomi !== undefined) {
          userSheet.getRange(i + 1, 4).setValue(newYomi);
        }
        // ★追加：ニックネームをF列（6列目）に書き込む
        if (newNickname !== undefined) {
          userSheet.getRange(i + 1, 6).setValue(newNickname);
        }
        isExist = true;
        break; 
      }
    }
    
    if (!isExist) {
      let personalUrl = "https://okamuro-d.github.io/Risshi-Log/web/index.html?id=" + targetIdm;
      // ★変更：新規作成時は E列にURL、F列にニックネーム を保存
      userSheet.appendRow([targetIdm, newName, newGrade, newYomi || "", personalUrl, newNickname || ""]);
    }
  }

  // 2. 統計シートの更新
  let statsSheet = getStatsSheet(app);
  if (statsSheet) {
    let stats = statsSheet.getDataRange().getValues();
    for (let i = 1; i < stats.length; i++) {
      if (String(stats[i][0]).trim() === targetIdm) {
        statsSheet.getRange(i + 1, 2).setValue(newName);
        break; 
      }
    }
  }

  // 3. 意見箱シートの更新
  let feedbackSheet = app.getSheetByName(FEEDBACK_SHEET_NAME);
  if (feedbackSheet) {
    let feedbacks = feedbackSheet.getDataRange().getValues();
    for (let i = 1; i < feedbacks.length; i++) {
      if (String(feedbacks[i][1]).trim() === targetIdm) {
        feedbackSheet.getRange(i + 1, 3).setValue(newName);
      }
    }
  }

  // 4. すべてのログシートの更新
  let sheets = app.getSheets();
  for (let s = 0; s < sheets.length; s++) {
    let sheet = sheets[s];
    let sheetName = sheet.getName();
    if (/^\d{4}年度$/.test(sheetName) || /^\d{4}-\d{2}$/.test(sheetName)) {
      let logs = sheet.getDataRange().getValues();
      for (let i = 1; i < logs.length; i++) {
        if (String(logs[i][0]).trim() === targetIdm) {
          sheet.getRange(i + 1, 2).setValue(newName);
        }
      }
    }
  }

  return responseJSON({ success: true });
}

// =========================================
// 手動での入退室切り替え処理
// =========================================
function toggleStatus(app, targetIdm) {
  targetIdm = String(targetIdm).trim();
  let userSheet = getUserSheet(app);
  let monthlySheet = getYearlySheet(app);
  let monitorSheet = app.getSheetByName(MONITOR_SHEET_NAME);
  
  let users = userSheet.getDataRange().getValues();
  let userName = "未登録(新規)";
  let isExist = false;
  
  for (let i = 1; i < users.length; i++) {
    if (String(users[i][0]).trim() === targetIdm) {
      userName = users[i][1] || "未登録";
      isExist = true;
      break;
    }
  }

  if (!isExist) {
    let personalUrl = "https://okamuro-d.github.io/Risshi-Log/web/index.html?id=" + targetIdm;
    // ★変更：E列にURL、F列に空のニックネーム枠を作成
    userSheet.appendRow([targetIdm, "", "", "", personalUrl, ""]);
  }

  let logs = monthlySheet.getDataRange().getValues();
  let targetRowIndex = -1;
  let entryTime = null;

  for (let i = logs.length - 1; i >= 1; i--) {
    if (String(logs[i][0]).trim() === targetIdm) {
      if (!logs[i][4] || String(logs[i][4]).trim() === "") {
        targetRowIndex = i + 1; 
        let dateStr = logs[i][2];
        let timeStr = logs[i][3];
        entryTime = new Date(formatDate(dateStr) + " " + formatTime(timeStr));
      }
      break;
    }
  }

  let now = new Date();
  let dateStr = Utilities.formatDate(now, 'Asia/Tokyo', 'yyyy-MM-dd');
  let timeStr = Utilities.formatDate(now, 'Asia/Tokyo', 'HH:mm:ss');

  if (targetRowIndex !== -1) {
    let durationMin = 0;
    if (entryTime && !isNaN(entryTime.getTime())) {
      durationMin = (now.getTime() - entryTime.getTime()) / (1000 * 60);
    }
    durationMin = Math.round(durationMin * 10) / 10;
    monthlySheet.getRange(targetRowIndex, 5).setValue(timeStr);
    monthlySheet.getRange(targetRowIndex, 6).setValue(durationMin);
    if (monitorSheet) monitorSheet.appendRow([userName, "退出", dateStr, timeStr]);
    updateSingleUserStats(app, targetIdm, durationMin, dateStr); 
    return responseJSON({ success: true, action: 'out', name: userName });
  } else {
    monthlySheet.appendRow([targetIdm, userName, dateStr, timeStr, "", ""]);
    if (monitorSheet) monitorSheet.appendRow([userName, "入室", dateStr, timeStr]);
    updateSingleUserStats(app, targetIdm, 0, dateStr); // 入室時も統計更新を通す
    return responseJSON({ success: true, action: 'in', name: userName });
  }
}

// =========================================
// 一括退室処理
// =========================================
function processForceExitAll(app, forcedTimeStr = null) {
  let monthlySheet = getYearlySheet(app);
  let monitorSheet = app.getSheetByName(MONITOR_SHEET_NAME);
  let logs = monthlySheet.getDataRange().getValues();
  
  let now = new Date();
  let dateStr = Utilities.formatDate(now, 'Asia/Tokyo', 'yyyy-MM-dd');
  
  let timeStr = forcedTimeStr ? forcedTimeStr : Utilities.formatDate(now, 'Asia/Tokyo', 'HH:mm:ss');
  let exitTimeDate = forcedTimeStr ? new Date(dateStr + " " + timeStr) : now;
  
  let count = 0;

  for (let i = logs.length - 1; i >= 1; i--) {
    if (!logs[i][4] || String(logs[i][4]).trim() === "") {
      let targetRowIndex = i + 1;
      let userName = logs[i][1];
      let entryDateStr = logs[i][2];
      let entryTimeStr = logs[i][3];
      let entryTime = new Date(formatDate(entryDateStr) + " " + formatTime(entryTimeStr));
      
      let durationMin = 0;
      if (entryTime && !isNaN(entryTime.getTime())) {
        durationMin = (exitTimeDate.getTime() - entryTime.getTime()) / (1000 * 60);
      }
      
      if (durationMin < 0) durationMin = 0;
      durationMin = Math.round(durationMin * 10) / 10;
      
      monthlySheet.getRange(targetRowIndex, 5).setValue(timeStr);
      monthlySheet.getRange(targetRowIndex, 6).setValue(durationMin);
      
      let remark = forcedTimeStr ? "退出(自動22時)" : "退出(一括)";
      if (monitorSheet) monitorSheet.appendRow([userName, remark, dateStr, timeStr]);
      count++;
    }
  }
  syncStatsData(app); 
  return responseJSON({ success: true, count: count });
}

// =========================================
// その他の関数・ヘルパー
// =========================================
function submitFeedback(app, idm, name, message) {
  let sheet = app.getSheetByName(FEEDBACK_SHEET_NAME);
  if (!sheet) {
    sheet = app.insertSheet(FEEDBACK_SHEET_NAME);
    sheet.appendRow(['日時', 'IDm', '名前', '意見・要望']);
  }
  let now = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd HH:mm:ss');
  sheet.appendRow([now, idm || "不明", name || "匿名", message]);
  return responseJSON({ success: true });
}

function getYearlySheet(app) {
  let now = new Date();
  let year = now.getFullYear();
  let month = now.getMonth() + 1; 
  if (month <= 3) year -= 1;
  
  let sheetName = year + "年度";
  let sheet = app.getSheetByName(sheetName);
  if (!sheet) {
    sheet = app.insertSheet(sheetName);
    sheet.appendRow(['IDm', '名前', '日付', '入室時刻', '退出時刻', '滞在時間(分)']);
  }
  return sheet;
}

function getStatsSheet(app) {
  let sheet = app.getSheetByName(STATS_SHEET_NAME);
  if (!sheet) {
    sheet = app.insertSheet(STATS_SHEET_NAME);
    sheet.appendRow([
      'IDm', '名前', '累計入室日数', '累計時間(分)', '最終入室日',
      '4月日数', '4月時間', '5月日数', '5月時間', '6月日数', '6月時間',
      '7月日数', '7月時間', '8月日数', '8月時間', '9月日数', '9月時間',
      '10月日数', '10月時間', '11月日数', '11月時間', '12月日数', '12月時間',
      '1月日数', '1月時間', '2月日数', '2月時間', '3月日数', '3月時間'
    ]);
  }
  return sheet;
}

function getUserSheet(app) {
  let sheet = app.getSheetByName(USER_SHEET_NAME);
  if (!sheet) {
    sheet = app.insertSheet(USER_SHEET_NAME);
    sheet.appendRow(['IDm', '名前', '学年', 'ふりがな', '生徒用URL', 'ニックネーム']);
  }
  return sheet;
}

function getLatestLog(app) {
  let sheet = app.getSheetByName(MONITOR_SHEET_NAME);
  if (!sheet || sheet.getLastRow() < 2) return responseJSON({ empty: true });
  let rowData = sheet.getRange(sheet.getLastRow(), 1, 1, 4).getValues()[0];
  return responseJSON({ name: rowData[0], status: rowData[1], date: formatDate(rowData[2]), time: formatTime(rowData[3]) });
}

function responseJSON(data) {
  let output = ContentService.createTextOutput(JSON.stringify(data));
  output.setMimeType(ContentService.MimeType.JSON);
  return output;
}

function formatDate(date) {
  if (!date) return "-";
  try { return Utilities.formatDate(new Date(date), 'Asia/Tokyo', 'yyyy-MM-dd'); } catch (e) { return date; }
}

function formatTime(timeVal) {
  if (!timeVal) return "";
  if (timeVal instanceof Date) return Utilities.formatDate(timeVal, 'Asia/Tokyo', 'HH:mm:ss');
  return String(timeVal).split('.')[0]; 
}

function deleteUser(app, targetIdm) {
  targetIdm = String(targetIdm).trim();
  
  let userSheet = getUserSheet(app);
  if (userSheet) {
    let users = userSheet.getDataRange().getValues();
    for (let i = users.length - 1; i >= 1; i--) {
      if (String(users[i][0]).trim() === targetIdm) userSheet.deleteRow(i + 1);
    }
  }

  let statsSheet = app.getSheetByName(STATS_SHEET_NAME);
  if (statsSheet) {
    let stats = statsSheet.getDataRange().getValues();
    for (let i = stats.length - 1; i >= 1; i--) {
      if (String(stats[i][0]).trim() === targetIdm) statsSheet.deleteRow(i + 1);
    }
  }

  let feedbackSheet = app.getSheetByName(FEEDBACK_SHEET_NAME);
  if (feedbackSheet) {
    let feedbacks = feedbackSheet.getDataRange().getValues();
    for (let i = feedbacks.length - 1; i >= 1; i--) {
      if (String(feedbacks[i][1]).trim() === targetIdm) feedbackSheet.deleteRow(i + 1);
    }
  }

  let sheets = app.getSheets();
  for (let s = 0; s < sheets.length; s++) {
    let sheet = sheets[s];
    let sheetName = sheet.getName();
    if (/^\d{4}年度$/.test(sheetName) || /^\d{4}-\d{2}$/.test(sheetName)) {
      let logs = sheet.getDataRange().getValues();
      for (let i = logs.length - 1; i >= 1; i--) {
        if (String(logs[i][0]).trim() === targetIdm) sheet.deleteRow(i + 1);
      }
    }
  }
  return responseJSON({ success: true });
}

// =========================================
// 【退室時用】対象者1名の「今月」のデータのみを高速更新
// =========================================
function updateSingleUserStats(app, targetIdm, durationMin, dateStr) {
  let now = new Date();
  let currentMonth = now.getMonth() + 1;
  let currentYear = currentMonth <= 3 ? now.getFullYear() - 1 : now.getFullYear();

  let yearlySheet = app.getSheetByName(currentYear + "年度");
  let statsSheet = getStatsSheet(app);
  if (!yearlySheet || !statsSheet) return;

  let logs = yearlySheet.getDataRange().getValues();
  let monthlyDaysSet = new Set();
  let monthlyTime = 0;

  for (let i = 1; i < logs.length; i++) {
    if (String(logs[i][0]).trim() === targetIdm) {
      let logDateStr = formatDate(logs[i][2]);
      let timeVal = parseFloat(logs[i][5]) || 0;

      if (logDateStr !== "-" && timeVal > 0) {
        let logDate = new Date(logs[i][2]);
        if (!isNaN(logDate.getTime()) && (logDate.getMonth() + 1) === currentMonth) {
          monthlyDaysSet.add(logDateStr); 
          monthlyTime += timeVal; 
        }
      }
    }
  }

  let stats = statsSheet.getDataRange().getValues();
  for (let i = 1; i < stats.length; i++) {
    if (String(stats[i][0]).trim() === targetIdm) {
      let rowNum = i + 1;
      
      let currentTotalTime = parseFloat(stats[i][3]) || 0;
      let newTotalTime = currentTotalTime + durationMin;
      
      let currentTotalDays = parseInt(stats[i][2]) || 0;
      let lastDateInStats = String(stats[i][4]);
      let newTotalDays = (lastDateInStats !== dateStr && durationMin > 0) ? currentTotalDays + 1 : currentTotalDays;

      let offset = currentMonth >= 4 ? currentMonth - 4 : currentMonth + 8;
      let daysCol = 6 + offset * 2;
      let timeCol = 7 + offset * 2;

      statsSheet.getRange(rowNum, 3).setValue(newTotalDays);
      statsSheet.getRange(rowNum, 4).setValue(Math.round(newTotalTime * 10) / 10);
      statsSheet.getRange(rowNum, 5).setValue(dateStr);
      statsSheet.getRange(rowNum, daysCol).setValue(monthlyDaysSet.size);
      statsSheet.getRange(rowNum, timeCol).setValue(Math.round(monthlyTime * 10) / 10);
      break;
    }
  }
}

// =========================================
// 年度シート＆過去シートから完全同期（再計算）
// =========================================
function syncStatsData(app) {
  let statsSheet = getStatsSheet(app);
  let userSheet = getUserSheet(app);
  if (!userSheet || !statsSheet) return;

  let users = userSheet.getDataRange().getValues();
  let statsMap = {};

  for (let i = 1; i < users.length; i++) {
    let idm = String(users[i][0]).trim();
    if (idm) {
      statsMap[idm] = {
        name: users[i][1],
        totalDays: new Set(),
        totalTime: 0,
        lastDate: "-",
        monthly: {}
      };
      for (let m = 1; m <= 12; m++) statsMap[idm].monthly[m] = { days: new Set(), time: 0 };
    }
  }

  let sheets = app.getSheets();
  let now = new Date();
  let currentYear = now.getMonth() + 1 <= 3 ? now.getFullYear() - 1 : now.getFullYear();

  for (let s = 0; s < sheets.length; s++) {
    let sheetName = sheets[s].getName();
    let isLogSheet = false;
    let sheetFiscalYear = 0;

    let matchYearly = sheetName.match(/^(\d{4})年度$/);
    if (matchYearly) {
      isLogSheet = true;
      sheetFiscalYear = parseInt(matchYearly[1]);
    } else {
      let matchMonthly = sheetName.match(/^(\d{4})-(\d{2})$/);
      if (matchMonthly) {
        isLogSheet = true;
        let y = parseInt(matchMonthly[1]);
        let m = parseInt(matchMonthly[2]);
        sheetFiscalYear = m <= 3 ? y - 1 : y;
      }
    }

    if (isLogSheet) {
      let isCurrentYear = (sheetFiscalYear === currentYear);
      let logs = sheets[s].getDataRange().getValues();

      for (let i = 1; i < logs.length; i++) {
        let idm = String(logs[i][0]).trim();
        if (!statsMap[idm]) continue;

        let dateStr = formatDate(logs[i][2]);
        let timeVal = parseFloat(logs[i][5]) || 0; 

        if (dateStr !== "-" && timeVal > 0) {
          statsMap[idm].totalDays.add(dateStr);
          statsMap[idm].totalTime += timeVal;
          if (statsMap[idm].lastDate === "-" || dateStr > statsMap[idm].lastDate) {
            statsMap[idm].lastDate = dateStr;
          }

          if (isCurrentYear) {
            let logDate = new Date(logs[i][2]);
            if (!isNaN(logDate.getTime())) {
              let logMonth = logDate.getMonth() + 1;
              if (statsMap[idm].monthly[logMonth]) {
                statsMap[idm].monthly[logMonth].days.add(dateStr);
                statsMap[idm].monthly[logMonth].time += timeVal;
              }
            }
          }
        }
      }
    }
  }

  let outputRows = [];
  for (let i = 1; i < users.length; i++) {
    let idm = String(users[i][0]).trim();
    if (!idm || !statsMap[idm]) continue;
    let sm = statsMap[idm];

    let row = [
      idm,
      sm.name,
      sm.totalDays.size,
      Math.round(sm.totalTime * 10) / 10,
      sm.lastDate
    ];

    let months = [4, 5, 6, 7, 8, 9, 10, 11, 12, 1, 2, 3];
    months.forEach(m => {
      row.push(sm.monthly[m].days.size);
      row.push(Math.round(sm.monthly[m].time * 10) / 10);
    });
    outputRows.push(row);
  }

  if (outputRows.length > 0) {
    let lastRow = statsSheet.getLastRow();
    if (lastRow > 1) {
      statsSheet.getRange(2, 1, lastRow - 1, 29).clearContent();
    }
    statsSheet.getRange(2, 1, outputRows.length, outputRows[0].length).setValues(outputRows);
  }
}

// スプレッドシートを手作業で編集したときに自動で統計を同期する
function onEdit(e) {
  if (!e || !e.source) return;
  let sheetName = e.source.getActiveSheet().getName();
  if (/^\d{4}年度$/.test(sheetName)) {
    let app = SpreadsheetApp.getActiveSpreadsheet();
    syncStatsData(app);
  }
}

// =========================================
// 【自動実行タイマー用】毎晩の統計データ完全同期（全員分）
// =========================================
function dailyFullStatsSync() {
  let app = SpreadsheetApp.openById(SHEET_ID);
  syncStatsData(app); 
  console.log("毎日の統計データ完全同期が完了しました。");
}

// =========================================
// 【自動実行タイマー用】4月1日の学年自動繰り上げ ＆ 統計データの年度更新処理
// =========================================
function promoteGradesIfNeeded() {
  let now = new Date();
  
  if (now.getMonth() === 3) {
    let properties = PropertiesService.getScriptProperties();
    let lastRunYear = properties.getProperty('LAST_PROMOTE_YEAR');
    let currentYear = now.getFullYear().toString();
    
    if (lastRunYear === currentYear) {
      console.log("今年の年度更新はすでに完了しています。");
      return;
    }

    let app = SpreadsheetApp.openById(SHEET_ID);
    
    let userSheet = getUserSheet(app);
    if (userSheet) {
      let data = userSheet.getDataRange().getValues();
      if (data.length > 1) {
        let updatedGrades = [];
        updatedGrades.push([data[0][2]]); 

        for (let i = 1; i < data.length; i++) {
          let currentGrade = String(data[i][2]).trim();
          let newGrade = currentGrade;
          if (currentGrade === "1") newGrade = "2";
          else if (currentGrade === "2") newGrade = "3";
          else if (currentGrade === "3") newGrade = "卒業生";
          
          updatedGrades.push([newGrade]);
        }
        userSheet.getRange(1, 3, updatedGrades.length, 1).setValues(updatedGrades);
      }
    }

    let statsSheet = app.getSheetByName(STATS_SHEET_NAME);
    if (statsSheet) {
      let lastYear = (now.getFullYear() - 1).toString();
      let backupName = STATS_SHEET_NAME + "_" + lastYear + "年度";
      
      if (!app.getSheetByName(backupName)) {
        let backupSheet = statsSheet.copyTo(app);
        backupSheet.setName(backupName);
        console.log(backupName + " として過去の統計データを保存しました。");
      }

      let lastRow = statsSheet.getLastRow();
      if (lastRow > 1) {
        statsSheet.getRange(2, 6, lastRow - 1, 24).clearContent();
        console.log("今年度用の月間統計データをリセットしました。");
      }
    }

    properties.setProperty('LAST_PROMOTE_YEAR', currentYear);
    console.log(currentYear + "年度の自動更新処理（学年・統計）がすべて完了しました。");
    
  } else {
    console.log("4月ではないため、年度の更新はスキップされました。");
  }
}

// =========================================
// 【自動実行タイマー用】毎晩の一括退室処理
// =========================================
function autoForceExitAll() {
  let app = SpreadsheetApp.openById(SHEET_ID);
  processForceExitAll(app, "22:00:00");
  console.log("毎晩の自動一括退室（22時で記録）が完了しました。");
}