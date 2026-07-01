/** ════════════════════════════════════════════
 * 永森健康管理系統  v6.1
 * Google Apps Script Backend
 * ════════════════════════════════════════════ */

var MAX_ATTEMPTS = 5;
var LOCKOUT_MIN  = 15;

function doGet(e) {
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('永森健康管理系統')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * 處理非同步 API 請求 (解耦架構)
 */
function doPost(e) {
  var postData;
  try {
    postData = JSON.parse(e.postData.contents);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ success: false, msg: "解析請求失敗" }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  var action = postData.action;
  var data = postData.data;
  var result;

  try {
    switch (action) {
      case 'login':
        result = login(data.password);
        break;
      case 'saveData':
        result = saveData(data);
        break;
      case 'deleteRecord':
        result = deleteRecord(data.rowIndex, data.userName);
        break;
      case 'getHistory':
        result = getHistory(data.userName);
        break;
      case 'analyzeFood':
        result = analyzeFood(data.foodText, data.userName);
        break;
      case 'analyzeFoodImage':
        result = analyzeFoodImage(data.base64Data, data.mimeType, data.userName);
        break;
      case 'saveWeight':
        result = saveWeight(data);
        break;
      case 'getWeightHistory':
        result = getWeightHistory(data.userName);
        break;
      case 'updatePassword':
        result = updatePassword(data.userName, data.newPassword);
        break;
      default:
        result = { success: false, msg: "未知的操作指令: " + action };
    }
  } catch (err) {
    result = { success: false, msg: err.message };
  }

  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

/* ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
   LOGIN
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━ */
function login(password) {
  try {
    var props       = PropertiesService.getScriptProperties();
    var ss          = SpreadsheetApp.getActiveSpreadsheet();
    var configSheet = ss.getSheetByName('設定');
    if (!configSheet) return { success: false, msg: "找不到『設定』分頁" };

    var rows      = configSheet.getDataRange().getValues();
    var inputPass = String(password).trim();

    var matchedUser = null;
    var matchedTdee = 2000;
    for (var i = 1; i < rows.length; i++) {
      if (!rows[i][0]) continue;
      if (String(rows[i][0]).trim() === inputPass) {
        matchedUser = String(rows[i][1]).trim();
        matchedTdee = Number(rows[i][2]) || 2000;
        break;
      }
    }

    var lockKey = matchedUser ? 'lock_user_' + matchedUser : 'lock_global_failed_attempts';
    var lockRaw = props.getProperty(lockKey);
    var lock    = lockRaw ? JSON.parse(lockRaw) : { attempts: 0, lastAttempt: 0 };

    if (lock.attempts >= MAX_ATTEMPTS) {
      var elapsedMin = (Date.now() - lock.lastAttempt) / 60000;
      if (elapsedMin < LOCKOUT_MIN) {
        return { success: false, msg: '嘗試次數過多，請 ' + Math.ceil(LOCKOUT_MIN - elapsedMin) + ' 分鐘後再試' };
      }
      lock = { attempts: 0, lastAttempt: 0 };
    }

    if (matchedUser) {
      props.deleteProperty(lockKey);
      var savedHeight = props.getProperty('height_' + matchedUser);
      var height      = savedHeight ? parseFloat(savedHeight) : 170;
      var dashboard   = getAllDashboardData(matchedUser, matchedTdee);
      return { success: true, userName: matchedUser, tdee: matchedTdee, height: height, dashboard: dashboard };
    }

    lock.attempts++;
    lock.lastAttempt = Date.now();
    props.setProperty(lockKey, JSON.stringify(lock));
    var left = MAX_ATTEMPTS - lock.attempts;
    return {
      success: false,
      msg: left > 0 ? '密碼不正確（剩餘 ' + left + ' 次機會）' : '系統已暫時鎖定，請稍後再試'
    };
  } catch (e) {
    return { success: false, msg: '系統錯誤：' + e.message };
  }
}

/* ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
   DASHBOARD
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━ */
function getAllDashboardData(userName, tdee) {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('紀錄');
  if (!sheet) return { today: {}, trend: {}, todayLog: [], streak: 0 };

  var targetTdee = tdee || 2000;
  var now        = new Date();
  var todayStr   = _fmtDate(now);
  var tz         = Session.getScriptTimeZone();

  var labels = [], dailyTotals = {}, dateOrder = [];
  for (var d = 6; d >= 0; d--) {
    var dt = new Date(now.getTime() - d * 86400000);
    var ds = _fmtDate(dt);
    labels.push(Utilities.formatDate(dt, tz, 'M/d'));
    dailyTotals[ds] = 0;
    dateOrder.push(ds);
  }

  var raw        = sheet.getDataRange().getValues();
  var todayTotal = 0;
  var todayLog   = [];
  var allDays    = {};

  for (var i = 1; i < raw.length; i++) {
    if (String(raw[i][5]).trim() !== userName) continue;
    var rowDate = raw[i][1] instanceof Date ? _fmtDate(raw[i][1]) : String(raw[i][1]).trim();
    var cal     = Number(raw[i][3]) || 0;
    var time    = raw[i][0] instanceof Date ? _fmtTime(raw[i][0]) : '';

    allDays[rowDate] = (allDays[rowDate] || 0) + cal;

    if (rowDate === todayStr) {
      todayTotal += cal;
      todayLog.push({ rowIndex: i + 1, name: String(raw[i][2]), cal: cal, time: time, note: String(raw[i][4] || '') });
    }
    if (dailyTotals.hasOwnProperty(rowDate)) dailyTotals[rowDate] += cal;
  }

  todayLog.sort(function(a, b) { return b.time.localeCompare(a.time); });

  var streak   = 0;
  var checkDay = new Date(now);
  checkDay.setHours(0, 0, 0, 0);
  if (todayTotal === 0) checkDay = new Date(checkDay.getTime() - 86400000);
  while (allDays[_fmtDate(checkDay)] > 0) {
    streak++;
    checkDay = new Date(checkDay.getTime() - 86400000);
  }
  if (todayTotal > 0 && streak === 0) streak = 1;

  var values  = dateOrder.map(function(k) { return dailyTotals[k]; });
  var remain  = targetTdee - todayTotal;
  var percent = parseFloat(Math.min((todayTotal / targetTdee) * 100, 100).toFixed(1));

  return {
    today:    { total: todayTotal, remain: remain, percent: percent },
    trend:    { labels: labels, values: values, tdee: targetTdee },
    todayLog: todayLog,
    streak:   streak
  };
}

/* ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
   SAVE
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━ */
function saveData(data) {
  try {
    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('紀錄');
    var now   = new Date();

    var dateStr    = data.date ? data.date : _fmtDate(now);
    var recordTime = now;
    if (data.date && data.date !== _fmtDate(now)) {
      recordTime = new Date(data.date + 'T12:00:00');
    }

    sheet.appendRow([recordTime, dateStr, data.mealName, data.calories, data.note, String(data.currentUser)]);
    var newRowIndex = sheet.getLastRow();

    return {
      newEntry: {
        rowIndex: newRowIndex,
        name:     data.mealName,
        cal:      data.calories,
        time:     _fmtTime(recordTime),
        note:     data.note,
        dateStr:  dateStr
      }
    };
  } catch (e) {
    throw new Error('儲存失敗：' + e.message);
  }
}

/* ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
   DELETE
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━ */
function deleteRecord(rowIndex, userName) {
  try {
    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('紀錄');
    if (!sheet) throw new Error("找不到『紀錄』分頁");

    var ri = parseInt(rowIndex, 10);
    if (isNaN(ri) || ri < 2) throw new Error('無效的列索引');

    var rowData = sheet.getRange(ri, 1, 1, 6).getValues()[0];
    if (String(rowData[5]).trim() !== userName) throw new Error('無權限刪除此紀錄');

    sheet.deleteRow(ri);
    return { success: true };
  } catch (e) {
    throw new Error('刪除失敗：' + e.message);
  }
}

/* ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
   HISTORY
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━ */
function getHistory(userName) {
  try {
    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('紀錄');
    if (!sheet) return [];
    var raw    = sheet.getDataRange().getValues();
    var result = [];
    for (var i = 1; i < raw.length; i++) {
      if (!raw[i][0]) continue;
      if (String(raw[i][5]).trim() !== userName) continue;
      var rowDate = raw[i][1] instanceof Date ? _fmtDate(raw[i][1]) : String(raw[i][1]).trim();
      var time    = raw[i][0] instanceof Date ? _fmtTime(raw[i][0]) : '';
      result.push({ rowIndex: i + 1, dateStr: rowDate, time: time, name: String(raw[i][2] || ''), cal: Number(raw[i][3]) || 0, note: String(raw[i][4] || '') });
    }
    result.sort(function(a, b) {
      var dc = b.dateStr.localeCompare(a.dateStr);
      return dc !== 0 ? dc : b.time.localeCompare(a.time);
    });
    return result;
  } catch (e) { throw new Error('讀取歷史失敗：' + e.message); }
}

/* ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
   共用 AI 呼叫邏輯 (內部使用)
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━ */
function _callAiAnalysis(userName, messages) {
  var props = PropertiesService.getScriptProperties();
  var today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  var countKey = 'ai_count_' + userName + '_' + today;
  var count = parseInt(props.getProperty(countKey) || '0', 10);

  if (count >= 40) return { error: true, note: '今日分析已達上限，請明天再試' };
  
  var GITHUB_KEY = props.getProperty('GITHUB_API_KEY');
  if (!GITHUB_KEY) return { error: true, note: '系統尚未設定 GITHUB_API_KEY' };

  var url = 'https://models.inference.ai.azure.com/chat/completions';
  var payload = {
    model: 'gpt-4o',
    response_format: { type: 'json_object' },
    temperature: 0.1,
    messages: messages
  };

  var lastErrorMsg = '目前系統繁忙，請稍後再試。';

  for (var attempt = 1; attempt <= 3; attempt++) {
    try {
      var resp = UrlFetchApp.fetch(url, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json', 'Authorization': 'Bearer ' + GITHUB_KEY },
        muteHttpExceptions: true,
        payload: JSON.stringify(payload)
      });
      
      var code = resp.getResponseCode();
      if (code === 200) {
        var json = JSON.parse(resp.getContentText());
        var content = json.choices[0].message.content;
        
        // 過濾 Markdown 外框，防止解析失敗
        content = content.replace(/^```(json)?\n?/i, '').replace(/\n?```$/i, '').trim();
        var obj = JSON.parse(content);
        
        // 確定成功後才扣除次數
        props.setProperty(countKey, String(count + 1));
        return { error: false, data: obj };
      } 
      
      // 如果是 400 或 401 這類非暫時性錯誤，直接中斷不重試
      if (code === 400 || code === 401 || code === 403) {
        lastErrorMsg = 'AI 服務設定錯誤或權限不足 (' + code + ')';
        break; 
      }
      
      Utilities.sleep(500);
    } catch (e) { 
      lastErrorMsg = '系統錯誤：' + e.message;
      Utilities.sleep(500); 
    }
  }
  
  return { error: true, note: lastErrorMsg };
}

/* ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
   ANALYZE FOOD（文字版）
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━ */
function analyzeFood(foodText, userName) {
  foodText = String(foodText || '一顆大蘋果').trim().substring(0, 200);
  var messages = [
    {
      role: 'system',
      content: '你是一位資深的台灣營養師，熟悉台灣小吃、便當菜色與常見烹調手法（如熱炒、油炸、勾芡）。請針對食物進行精確分析，考量隱藏的油脂與糖分。你必須只回傳嚴格的 JSON，包含以下欄位：\'calories\'(數字,總熱量kcal), \'confidence\'(字串,high/medium/low), \'note\'(字串,150字內繁體中文，包含：1.估算依據/份量 2.烹調方式帶來的隱藏熱量提醒 3.具體且實用的健康建議)。絕不輸出 Markdown 或任何說明文字。'
    },
    { role: 'user', content: foodText }
  ];

  var result = _callAiAnalysis(userName, messages);
  if (result.error) return { calories: null, note: result.note };
  
  return {
    calories:   parseInt(result.data.calories)  || null,
    confidence: String(result.data.confidence   || 'medium'),
    note:       String(result.data.note         || '')
  };
}

/* ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
   ANALYZE FOOD IMAGE（圖片版）
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━ */
function analyzeFoodImage(base64Data, mimeType, userName) {
  var cleanBase64 = base64Data.replace(/^data:image\/\w+;base64,/, '');
  var mime        = mimeType || 'image/jpeg';
  
  var messages = [
    {
      role: 'system',
      content: '你是一位資深的台灣營養師，熟悉台灣在地飲食習慣。請發揮專業觀察力，仔細分析照片中的食物，考量比例尺、容器大小、油脂反光度及烹調方式（如煎、炸、勾芡）來推估真實份量。你必須只回傳嚴格的 JSON，包含以下欄位：\'foods_detected\'(字串陣列,每項格式為\'食物名稱(估計重量g)\'),\'calories\'(數字,總熱量kcal),\'confidence\'(字串,high/medium/low),\'note\'(字串,150字內繁體中文，包含：1.辨識與重量估算依據 2.烹調方式分析與隱藏熱量提醒 3.改善建議)。絕不輸出 Markdown。'
    },
    {
      role: 'user',
      content: [
        { type: 'text', text: '請辨識這張照片中的食物並估算熱量' },
        { type: 'image_url', image_url: { url: 'data:' + mime + ';base64,' + cleanBase64 } }
      ]
    }
  ];

  var result = _callAiAnalysis(userName, messages);
  if (result.error) return { calories: null, note: result.note, detectedFoods: '' };

  return {
    calories:      parseInt(result.data.calories)   || null,
    confidence:    String(result.data.confidence    || 'medium'),
    note:          String(result.data.note          || ''),
    detectedFoods: (result.data.foods_detected      || []).join('、')
  };
}

/* ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
   WEIGHT：儲存體重
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━ */
function saveWeight(data) {
  try {
    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('體重');
    if (!sheet) throw new Error("找不到『體重』分頁");

    var weight = parseFloat(data.weight);
    if (isNaN(weight) || weight < 20 || weight > 300) throw new Error('體重請填入 20–300 kg 之間的數值');

    var now      = new Date();
    var todayStr = _fmtDate(now);
    var userName = String(data.userName);

    var raw      = sheet.getDataRange().getValues();
    var foundRow = -1;
    for (var i = raw.length - 1; i >= 1; i--) {
      var rowDate = raw[i][1] instanceof Date ? _fmtDate(raw[i][1]) : String(raw[i][1]).trim();
      if (rowDate === todayStr && String(raw[i][3]).trim() === userName) {
        foundRow = i + 1;
        break;
      }
    }

    if (foundRow > -1) {
      sheet.getRange(foundRow, 1).setValue(now);
      sheet.getRange(foundRow, 3).setValue(weight);
    } else {
      sheet.appendRow([now, todayStr, weight, userName]);
    }

    var props       = PropertiesService.getScriptProperties();
    var savedTarget = null;
    var savedHeight = null;

    if (data.targetWeight !== undefined && data.targetWeight !== null && data.targetWeight !== '') {
      var t = parseFloat(data.targetWeight);
      if (!isNaN(t) && t >= 20 && t <= 300) {
        props.setProperty('target_' + userName, String(t));
        savedTarget = t;
      }
    }

    if (data.height !== undefined && data.height !== null && data.height !== '') {
      var h = parseFloat(data.height);
      if (!isNaN(h) && h >= 50 && h <= 250) {
        props.setProperty('height_' + userName, String(h));
        savedHeight = h;
      }
    }

    return { success: true, weight: weight, date: todayStr, targetWeight: savedTarget, height: savedHeight };
  } catch (e) {
    throw new Error('儲存失敗：' + e.message);
  }
}

/* ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
   WEIGHT：讀取歷史
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━ */
function getWeightHistory(userName) {
  try {
    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('體重');
    var props = PropertiesService.getScriptProperties();

    var targetVal    = props.getProperty('target_' + userName);
    var targetWeight = targetVal ? parseFloat(targetVal) : null;
    var heightVal    = props.getProperty('height_' + userName);
    var userHeight   = heightVal ? parseFloat(heightVal) : null;

    if (!sheet) return { records: [], stats: {}, targetWeight: targetWeight, userHeight: userHeight };

    var raw     = sheet.getDataRange().getValues();
    var records = [];
    for (var i = 1; i < raw.length; i++) {
      if (!raw[i][0]) continue;
      if (String(raw[i][3]).trim() !== userName) continue;
      var w = parseFloat(raw[i][2]);
      if (isNaN(w)) continue;
      records.push({
        date:   raw[i][1] instanceof Date ? _fmtDate(raw[i][1]) : String(raw[i][1]).trim(),
        weight: w
      });
    }

    records.sort(function(a, b) { return a.date.localeCompare(b.date); });

    var dayMap = {};
    records.forEach(function(r) { dayMap[r.date] = r.weight; });
    var dedupe = Object.keys(dayMap).sort().map(function(d) { return { date: d, weight: dayMap[d] }; });

    if (!dedupe.length) return { records: [], stats: {}, targetWeight: targetWeight, userHeight: userHeight };

    var latest  = dedupe[dedupe.length - 1].weight;
    var oldest  = dedupe[0].weight;
    var weights = dedupe.map(function(r) { return r.weight; });

    return {
      records: dedupe,
      stats: {
        latest:  latest,
        diff:    parseFloat((latest - oldest).toFixed(1)),
        highest: Math.max.apply(null, weights),
        lowest:  Math.min.apply(null, weights),
        days:    dedupe.length
      },
      targetWeight: targetWeight,
      userHeight:   userHeight
    };
  } catch (e) {
    throw new Error('讀取失敗：' + e.message);
  }
}

/* ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
   UPDATE PASSWORD
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━ */
function updatePassword(userName, newPassword) {
  try {
    var ss          = SpreadsheetApp.getActiveSpreadsheet();
    var configSheet = ss.getSheetByName('設定');
    if (!configSheet) return { success: false, msg: "找不到『設定』分頁" };

    var rows     = configSheet.getDataRange().getValues();
    var foundRow = -1;
    for (var i = 1; i < rows.length; i++) {
      if (String(rows[i][1]).trim() === userName) { foundRow = i + 1; break; }
    }

    if (foundRow > -1) {
      configSheet.getRange(foundRow, 1).setValue(newPassword);
      return { success: true };
    }
    return { success: false, msg: '找不到該使用者，修改失敗' };
  } catch (e) {
    return { success: false, msg: '系統錯誤：' + e.message };
  }
}

/* ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
   UTILS
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━ */
function _fmtDate(d) {
  return Utilities.formatDate(d instanceof Date ? d : new Date(d), Session.getScriptTimeZone(), 'yyyy-MM-dd');
}
function _fmtTime(d) {
  return Utilities.formatDate(d instanceof Date ? d : new Date(d), Session.getScriptTimeZone(), 'HH:mm');
}