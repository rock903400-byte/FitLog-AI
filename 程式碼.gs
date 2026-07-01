/** ════════════════════════════════════════════
 * 永森健康管理系統  v6.0 (商業優化版)
 * Google Apps Script Backend
 * ════════════════════════════════════════════ */

var MAX_ATTEMPTS   = 5;
var LOCKOUT_MIN    = 15;

function doGet(e) {
  // 只保留網頁渲染功能，移除 manifest 判斷
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('永森健康管理系統')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/** 
 * 處理來自前端的 POST 請求 (路由中心)
 */
function doPost(e) {
  try {
    var params = JSON.parse(e.postData.contents);
    var action = params.action;
    var data   = params.data;
    var result;

    switch (action) {
      case 'login':           result = login(data.password); break;
      case 'saveData':        result = saveData(data); break;
      case 'deleteRecord':    result = deleteRecord(data.rowIndex, data.userName); break;
      case 'getHistory':      result = getHistory(data.userName); break;
      case 'analyzeFood':     result = analyzeFood(data.foodText, data.userName); break;
      case 'analyzeFoodImage':result = analyzeFoodImage(data.base64Data, data.mimeType, data.userName); break;
      case 'saveWeight':      result = saveWeight(data); break;
      case 'getWeightHistory':result = getWeightHistory(data.userName); break;
      case 'updatePassword':  result = updatePassword(data.userName, data.newPassword); break;
      case 'createUser':      result = createUser(data.newUserName, data.newPassword); break;
      default: throw new Error('未知的 Action: ' + action);
    }

    return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ success: false, msg: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/* ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
   LOGIN  (深度優化淨化版：移除冗餘與修復崩潰錯誤)
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━ */
function login(password) {
  try {
    var props = PropertiesService.getScriptProperties();
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var configSheet = ss.getSheetByName('設定');
    if (!configSheet) return { success: false, msg: "找不到『設定』分頁" };

    var rows = configSheet.getDataRange().getValues();
    var inputPass = String(password).trim();

    // ① 一次性尋找使用者與 TDEE
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

    // ② 鎖定機制
    var lockKey = matchedUser ? 'lock_user_' + matchedUser : 'lock_global_failed_attempts';
    var lockRaw = props.getProperty(lockKey);
    var lock = lockRaw ? JSON.parse(lockRaw) : { attempts: 0, lastAttempt: 0 };

    if (lock.attempts >= MAX_ATTEMPTS) {
      var elapsedMin = (Date.now() - lock.lastAttempt) / 60000;
      if (elapsedMin < LOCKOUT_MIN) {
        return { success: false, msg: '嘗試次數過多，請 ' + Math.ceil(LOCKOUT_MIN - elapsedMin) + ' 分鐘後再試' };
      }
      lock = { attempts: 0, lastAttempt: 0 };
    }

    // ③ 登入成功處理 (已移除崩潰風險的 getRecentFoods)
    if (matchedUser) {
      props.deleteProperty(lockKey);
      
      var savedHeight = props.getProperty('height_' + matchedUser);
      var height = savedHeight ? parseFloat(savedHeight) : 170;

      var savedAge = props.getProperty('age_' + matchedUser);
      var age = savedAge ? parseInt(savedAge) : 25;

      var savedGender = props.getProperty('gender_' + matchedUser);
      var gender = savedGender ? savedGender : 'male';

      var savedActivity = props.getProperty('activity_' + matchedUser);
      var activity = savedActivity ? parseFloat(savedActivity) : 1.2;
      
      // 直接把上面找好的參數傳入，不要再重跑迴圈
      var dashboard = getAllDashboardData(matchedUser, matchedTdee); 
      
      return { 
        success: true, 
        userName: matchedUser, 
        tdee: matchedTdee, 
        height: height, 
        age: age,
        gender: gender,
        activity: activity,
        dashboard: dashboard 
      };
    }

    // ④ 登入失敗
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
   DASHBOARD (優化效能：減少重複查詢)
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━ */
function getAllDashboardData(userName, tdee) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('紀錄');
  if (!sheet) return { today: {}, trend: {}, todayLog: [], streak: 0 };

  // 若 login 已經傳入 tdee，就不必再去查設定表
  var targetTdee = tdee || 2000;

  var now = new Date();
  var todayStr = _fmtDate(now);
  var tz = Session.getScriptTimeZone();

  var labels = [], dailyTotals = {}, dateOrder = [];
  for (var d = 6; d >= 0; d--) {
    var dt = new Date(now.getTime() - d * 86400000);
    var ds = _fmtDate(dt);
    labels.push(Utilities.formatDate(dt, tz, 'M/d'));
    dailyTotals[ds] = 0;
    dateOrder.push(ds);
  }

  var raw = sheet.getDataRange().getValues();
  var todayTotal = 0;
  var todayLog = [];
  var allDays = {};

  for (var i = 1; i < raw.length; i++) {
    if (String(raw[i][5]).trim() !== userName) continue;
    var rowDate = raw[i][1] instanceof Date ? _fmtDate(raw[i][1]) : String(raw[i][1]).trim();
    var cal = Number(raw[i][3]) || 0;
    var time = raw[i][0] instanceof Date ? _fmtTime(raw[i][0]) : '';

    allDays[rowDate] = (allDays[rowDate] || 0) + cal;

    if (rowDate === todayStr) {
      todayTotal += cal;
      todayLog.push({ rowIndex: i + 1, name: String(raw[i][2]), cal: cal, time: time, note: String(raw[i][4] || '') });
    }
    if (dailyTotals.hasOwnProperty(rowDate)) dailyTotals[rowDate] += cal;
  }

  todayLog.sort(function(a, b) { return b.time.localeCompare(a.time); });

  // ── 連續天數 ──
  var streak = 0;
  var checkDay = new Date(now);
  checkDay.setHours(0, 0, 0, 0);
  if (todayTotal === 0) checkDay = new Date(checkDay.getTime() - 86400000);
  while (allDays[_fmtDate(checkDay)] > 0) {
    streak++;
    checkDay = new Date(checkDay.getTime() - 86400000);
  }
  if (todayTotal > 0 && streak === 0) streak = 1;

  var values = dateOrder.map(function(k) { return dailyTotals[k]; });
  var remain = targetTdee - todayTotal;
  var percent = parseFloat(Math.min((todayTotal / targetTdee) * 100, 100).toFixed(1));

  return {
    today: { total: todayTotal, remain: remain, percent: percent },
    trend: { labels: labels, values: values, tdee: targetTdee },
    todayLog: todayLog,
    streak: streak
  };
}

/* ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
   SAVE  (強化輸入驗證)
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━ */
function saveData(data) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('紀錄');
    
    var now = new Date();
    
    // 檢查是否有傳入自訂日期，若無則使用今天
    var dateStr = data.date ? data.date : _fmtDate(now);
    
    // 如果不是今天，產生一個對應日期的時間物件 (預設設為該日 12:00:00 以防時區誤差)
    var recordTime = now;
    if (data.date && data.date !== _fmtDate(now)) {
      recordTime = new Date(data.date + 'T12:00:00');
    }
    
    var mealName = data.mealName;
    var cal = data.calories;
    var note = data.note;

    sheet.appendRow([recordTime, dateStr, mealName, cal, note, String(data.currentUser)]);
    var newRowIndex = sheet.getLastRow(); 

    return {
      newEntry: {
        rowIndex: newRowIndex, 
        name: mealName,
        cal:  cal,
        time: _fmtTime(recordTime),
        note: note,
        dateStr: dateStr // 回傳日期以便前端判斷是否為今日
      }
    };
  } catch (e) {
    throw new Error('儲存失敗：' + e.message);
  }
}

/* ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
   DELETE  (含身份驗證)
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━ */
function deleteRecord(rowIndex, userName) {
  try {
    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('紀錄');
    if (!sheet) throw new Error("找不到『紀錄』分頁");

    var ri = parseInt(rowIndex, 10);
    if (isNaN(ri) || ri < 2) throw new Error('無效的列索引');

    // 安全驗證：確認此列確實屬於該使用者
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
   WEIGHT：儲存體重 (更新身體數據與 TDEE)
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━ */
function saveWeight(data) {
  try {
    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('體重');
    if (!sheet) throw new Error("找不到『體重』分頁");

    var weight = parseFloat(data.weight);
    if (isNaN(weight) || weight < 20 || weight > 300)
      throw new Error('體重請填入 20–300 kg 之間的數值');

    var now      = new Date();
    var todayStr = _fmtDate(now);
    var userName = String(data.userName);

    // ── 檢查今天是否已經輸入過，有則覆寫，無則新增 (保留歷史底層紀錄但不顯示) ──
    var raw      = sheet.getDataRange().getValues();
    var foundRow = -1;
    for (var i = raw.length - 1; i >= 1; i--) {
      var rowDate = raw[i][1] instanceof Date ? _fmtDate(raw[i][1]) : String(raw[i][1]).trim();
      var rowUser = String(raw[i][3]).trim();
      if (rowDate === todayStr && rowUser === userName) {
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

    var props = PropertiesService.getScriptProperties();

    // ── 儲存身高、年齡、性別、活動量至屬性 ──
    if (data.height)   props.setProperty('height_' + userName, String(data.height));
    if (data.age)      props.setProperty('age_' + userName, String(data.age));
    if (data.gender)   props.setProperty('gender_' + userName, String(data.gender));
    if (data.activity) props.setProperty('activity_' + userName, String(data.activity));
    props.setProperty('latest_weight_' + userName, String(weight));

    // ── 更新設定表的 TDEE ──
    if (data.tdee) {
      var configSheet = ss.getSheetByName('設定');
      if (configSheet) {
        var configRows = configSheet.getDataRange().getValues();
        for (var j = 1; j < configRows.length; j++) {
          if (String(configRows[j][1]).trim() === userName) {
            configSheet.getRange(j + 1, 3).setValue(data.tdee);
            break;
          }
        }
      }
    }

    return { 
      success: true, 
      weight: weight, 
      tdee: data.tdee
    };
  } catch (e) {
    throw new Error('儲存失敗：' + e.message);
  }
}

/* ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
   WEIGHT：讀取身體數據
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━ */
function getWeightHistory(userName) {
  try {
    var props = PropertiesService.getScriptProperties();
    
    var heightVal   = props.getProperty('height_' + userName);
    var userHeight  = heightVal ? parseFloat(heightVal) : 170;

    var ageVal      = props.getProperty('age_' + userName);
    var userAge     = ageVal ? parseInt(ageVal) : 25;

    var genderVal   = props.getProperty('gender_' + userName);
    var userGender  = genderVal ? genderVal : 'male';

    var activityVal = props.getProperty('activity_' + userName);
    var userActivity = activityVal ? parseFloat(activityVal) : 1.2;

    var weightVal   = props.getProperty('latest_weight_' + userName);
    var latestWeight = weightVal ? parseFloat(weightVal) : 0;

    // 如果 Properties 沒有最新體重，嘗試從試算表抓取
    if (latestWeight === 0) {
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      var sheet = ss.getSheetByName('體重');
      if (sheet) {
        var raw = sheet.getDataRange().getValues();
        for (var i = raw.length - 1; i >= 1; i--) {
          if (String(raw[i][3]).trim() === userName) {
            latestWeight = parseFloat(raw[i][2]);
            break;
          }
        }
      }
    }

    return {
      stats: { latest: latestWeight },
      userHeight: userHeight,
      userAge: userAge,
      userGender: userGender,
      userActivity: userActivity
    };
  } catch (e) {
    throw new Error('讀取身體數據失敗：' + e.message);
  }
}

/**
 * 更新使用者密碼
 */
function updatePassword(userName, newPassword) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var configSheet = ss.getSheetByName('設定');
    if (!configSheet) return { success: false, msg: "找不到『設定』分頁" };

    var rows = configSheet.getDataRange().getValues();
    var found = false;

    for (var i = 1; i < rows.length; i++) {
      if (String(rows[i][1]).trim() === userName) {
        configSheet.getRange(i + 1, 1).setValue(newPassword);
        found = true;
        break;
      }
    }

    if (found) {
      return { success: true, msg: '密碼已更新' };
    } else {
      return { success: false, msg: '找不到使用者' };
    }
  } catch (e) {
    return { success: false, msg: '密碼更新失敗：' + e.message };
  }
}

/* ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
   UTILS
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━ */
function _fmtDate(d) {
  var tz = Session.getScriptTimeZone();
  return Utilities.formatDate(d instanceof Date ? d : new Date(d), tz, 'yyyy-MM-dd');
}

function _fmtTime(d) {
  var tz = Session.getScriptTimeZone();
  return Utilities.formatDate(d instanceof Date ? d : new Date(d), tz, 'HH:mm');
}

/**
 * 建立新帳號
 */
function createUser(newUserName, newPassword) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var configSheet = ss.getSheetByName('設定');
    if (!configSheet) return { success: false, msg: "找不到『設定』分頁" };

    var rows = configSheet.getDataRange().getValues();
    var nameToCreate = String(newUserName).trim();
    var passToCreate = String(newPassword).trim();

    if (!nameToCreate || !passToCreate) {
      return { success: false, msg: '使用者名稱與密碼不得為空' };
    }

    // 檢查重複
    for (var i = 1; i < rows.length; i++) {
      var existingPass = String(rows[i][0]).trim();
      var existingName = String(rows[i][1]).trim();
      
      if (existingPass === passToCreate) {
        return { success: false, msg: '此密碼已被使用，請更換密碼' };
      }
      if (existingName === nameToCreate) {
        return { success: false, msg: '此使用者名稱已存在' };
      }
    }

    // 新增資料 (密碼, 名稱, 預設TDEE)
    configSheet.appendRow([passToCreate, nameToCreate, 2000]);
    return { success: true, msg: '帳號「' + nameToCreate + '」建立成功！' };

  } catch (e) {
    return { success: false, msg: '帳號建立失敗：' + e.message };
  }
}