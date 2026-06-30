// ============================================================
// 員工打卡系統 - 全新獨立版本 (Email + 密碼登入)
// 用法：開一個新Google Sheet → 擴充功能 → Apps Script
//       將呢個檔案全部內容貼入去 → 部署為Web應用程式
// ============================================================

var CONFIG = {
  recordSheetName: "打卡記錄",      // 主記錄分頁
  accountSheetName: "員工帳號",     // 員工Email+姓名分頁（你可以自己編輯呢頁嚟新增/刪除員工）
  companyIPPrefix: "59.148.189.22", // 公司WiFi Public IP前綴，改返你自己嘅
  checkTypes: ["返工", "放工", "放break", "完break"],
  oauthClientId: "564783152398-5dkc0roiqrgh89308u02fhmgdch8pk41.apps.googleusercontent.com"
};

// 部署為Web應用程式時用嚟接收 PWA 嘅請求
function doPost(e) {
  var body;
  try {
    body = JSON.parse(e.postData.contents);
  } catch (err) {
    return jsonOutput({ success: false, message: "請求格式錯誤" });
  }

  var action = body.action;
  if (action === "login") {
    return jsonOutput(handleLogin(body.idToken));
  } else if (action === "checkin") {
    return jsonOutput(handleCheckIn(body.idToken, body.checkType, body.remark, body.userIP));
  } else {
    return jsonOutput({ success: false, message: "未知操作" });
  }
}

// 驗證 Google ID Token，確保真係本人用Google account登入
function verifyGoogleToken(idToken) {
  try {
    if (!idToken) {
      return { valid: false, message: "未提供登入Token" };
    }
    var url = "https://oauth2.googleapis.com/tokeninfo?id_token=" + encodeURIComponent(idToken);
    var resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    var info = JSON.parse(resp.getContentText());

    if (info.error) {
      return { valid: false, message: "Token無效或已過期，請重新登入" };
    }
    if (info.aud !== CONFIG.oauthClientId) {
      return { valid: false, message: "Token唔屬於本App" };
    }
    return { valid: true, email: info.email };
  } catch (error) {
    return { valid: false, message: error.toString() };
  }
}

// ---------- 員工帳號驗證 ----------
function getAccountSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.accountSheetName);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.accountSheetName);
    sheet.getRange(1, 1, 1, 3).setValues([["姓名", "Email", "密碼"]]);
    sheet.getRange(1, 1, 1, 3).setFontWeight("bold");
    sheet.setFrozenRows(1);
    // 範例行，你可以刪走，跟住自己加返員工
    sheet.getRange(2, 1, 1, 3).setValues([["範例員工", "example@gmail.com", "abc123"]]);
  }
  return sheet;
}

function findAccount(email) {
  var sheet = getAccountSheet();
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][1]).trim().toLowerCase() === String(email).trim().toLowerCase()) {
      return { name: data[i][0], email: data[i][1] };
    }
  }
  return null;
}

function handleLogin(idToken) {
  var verify = verifyGoogleToken(idToken);
  if (!verify.valid) {
    return { success: false, message: verify.message };
  }
  var account = findAccount(verify.email);
  if (!account) {
    return { success: false, message: "你個Google account（" + verify.email + "）未登記，請聯絡管理員加入「員工帳號」分頁" };
  }
  return { success: true, name: account.name, email: account.email };
}

// ---------- 打卡 ----------
function handleCheckIn(idToken, checkType, remark, userIP) {
  // 1) 先驗證Google登入
  var loginResult = handleLogin(idToken);
  if (!loginResult.success) {
    return loginResult;
  }

  // 2) 驗證IP（一定要連公司WiFi）
  if (!userIP || !userIP.startsWith(CONFIG.companyIPPrefix)) {
    return {
      success: false,
      message: "必須連住公司WiFi先可以打卡\n檢測到IP: " + (userIP || "未知")
    };
  }

  // 3) 驗證打卡類型
  if (CONFIG.checkTypes.indexOf(checkType) === -1) {
    return { success: false, message: "無效打卡類型" };
  }

  var employeeName = loginResult.name;
  var employeeEmail = loginResult.email;
  var remarkText = remark ? String(remark).trim() : "";
  var currentTime = new Date();

  // 4) 寫入主記錄表
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var mainSheet = ss.getSheetByName(CONFIG.recordSheetName);
  if (!mainSheet) {
    mainSheet = ss.insertSheet(CONFIG.recordSheetName);
    mainSheet.getRange(1, 1, 1, 6).setValues([["時間", "員工姓名", "打卡類型", "IP地址", "Email", "備註"]]);
    mainSheet.getRange(1, 1, 1, 6).setFontWeight("bold");
    mainSheet.setFrozenRows(1);
  }
  var lastRow = mainSheet.getLastRow() + 1;
  mainSheet.getRange(lastRow, 1, 1, 6).setValues([[
    currentTime, employeeName, checkType, userIP, employeeEmail, remarkText
  ]]);
  mainSheet.getRange(lastRow, 1).setNumberFormat('yyyy-mm-dd hh:mm:ss');

  // 5) 寫入員工個人分頁
  createOrAppendEmployeeSheet(employeeName, currentTime, checkType, employeeEmail, userIP, remarkText);

  return {
    success: true,
    message: "打卡成功" + (remarkText ? "（備註已記錄）" : ""),
    time: Utilities.formatDate(currentTime, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss'),
    employee: employeeName,
    type: checkType
  };
}

function createOrAppendEmployeeSheet(employeeName, checkTime, checkType, email, userIP, remark) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(employeeName);
  if (!sheet) {
    sheet = ss.insertSheet(employeeName);
    sheet.getRange(1, 1, 1, 6).setValues([["日期", "時間", "打卡類型", "Email", "IP地址", "備註"]]);
    sheet.getRange(1, 1, 1, 6).setFontWeight("bold");
    sheet.setFrozenRows(1);
  }
  var lastRow = sheet.getLastRow() + 1;
  sheet.getRange(lastRow, 1, 1, 6).setValues([[
    Utilities.formatDate(checkTime, Session.getScriptTimeZone(), 'yyyy-MM-dd'),
    Utilities.formatDate(checkTime, Session.getScriptTimeZone(), 'HH:mm:ss'),
    checkType, email, userIP, remark || ""
  ]]);
}

function jsonOutput(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ---------- 測試用 ----------
function testSetup() {
  getAccountSheet();
  return "已建立「員工帳號」分頁，請去Sheet入面填返真實員工Email同密碼";
}
