/**
 * 志工居家訪視系統 — Google Apps Script 後端 API
 * 版本: 1.0
 *
 * ===== 部署指引 =====
 * 1. 在 Google Sheets 開啟 Apps Script（擴充功能 → Apps Script）
 * 2. 貼上本檔案內容至 Code.gs
 * 3. 先執行 InitSheet.gs 中的 initializeSheets() 建立資料表
 * 4. 點擊「部署」→「新增部署作業」
 * 5. 類型選「網路應用程式」
 * 6. 執行身分：我（Me）
 * 7. 存取權：所有人（Anyone，包含匿名）
 * 8. 複製部署網址，填入前端 frontend/app.js 的 CONFIG.GAS_API_URL
 * ====================
 */

// ============================================================
// GET 請求路由
// ============================================================
function doGet(e) {
  var params = e.parameter;
  var action = params.action;
  var result;

  try {
    switch (action) {
      case 'getQuestions':
        result = getQuestions();
        break;
      case 'getAllMembers':
        result = getAllMembers();
        break;
      case 'getMembersByBranch':
        result = getMembersByBranch(params.branch);
        break;
      case 'getBranches':
        result = getBranches();
        break;
      case 'verifyLogin':
        result = verifyLogin(params.name, params.phone);
        break;
      default:
        result = { success: false, error: '未知的 action: ' + action };
    }
  } catch (err) {
    result = { success: false, error: err.toString() };
  }

  return buildResponse(result);
}

// ============================================================
// POST 請求路由
// ============================================================
function doPost(e) {
  var result;

  try {
    var data = JSON.parse(e.postData.contents);
    var action = data.action;

    switch (action) {
      case 'submitForm':
        result = submitForm(data.record);
        break;
      default:
        result = { success: false, error: '未知的 action: ' + action };
    }
  } catch (err) {
    result = { success: false, error: err.toString() };
  }

  return buildResponse(result);
}

// ============================================================
// 建立 JSON 回應 (GAS 自動附帶 CORS Header)
// ============================================================
function buildResponse(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// ============================================================
// 取得訪視題庫
// ============================================================
function getQuestions() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('訪視題庫');
  if (!sheet) return { success: false, error: '找不到「訪視題庫」工作表，請先執行初始化腳本。' };

  var data = sheet.getDataRange().getValues();
  var questions = [];

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (!row[0]) continue;

    questions.push({
      id:       String(row[0]).trim(),
      category: String(row[1]).trim(),
      content:  String(row[2]).trim(),
      type:     String(row[3]).trim(),
      options:  row[4]
        ? String(row[4]).split(',').map(function(o) { return o.trim(); }).filter(function(o) { return o; })
        : [],
      required: row[5] === true || String(row[5]).toUpperCase() === 'TRUE'
    });
  }

  return { success: true, data: questions };
}

// ============================================================
// 取得所有人員（不回傳手機號碼，保護隱私）
// ============================================================
function getAllMembers() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('人員帳號管理');
  if (!sheet) return { success: false, error: '找不到「人員帳號管理」工作表。' };

  var data = sheet.getDataRange().getValues();
  var members = [];

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (!row[0]) continue;
    members.push({
      name:   String(row[0]).trim(),
      branch: String(row[1]).trim(),
      role:   String(row[3]).trim()
    });
  }

  return { success: true, data: members };
}

// ============================================================
// 依分隊取得人員
// ============================================================
function getMembersByBranch(branch) {
  var all = getAllMembers();
  if (!all.success) return all;

  var filtered = all.data.filter(function(m) {
    return m.branch === branch;
  });

  return { success: true, data: filtered };
}

// ============================================================
// 取得分隊列表
// ============================================================
function getBranches() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('分隊對照表');
  if (!sheet) return { success: false, error: '找不到「分隊對照表」工作表。' };

  var data = sheet.getDataRange().getValues();
  var branches = [];

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (!row[0]) continue;
    branches.push({
      name:         String(row[0]).trim(),
      managerEmail: String(row[1]).trim()
    });
  }

  return { success: true, data: branches };
}

// ============================================================
// 登入驗證：比對姓名 + 手機末三碼
// ============================================================
function verifyLogin(name, phoneLast3) {
  if (!name || !phoneLast3) {
    return { success: false, error: '請填寫姓名與手機末三碼' };
  }

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('人員帳號管理');
  if (!sheet) return { success: false, error: '系統錯誤：找不到人員資料' };

  var data = sheet.getDataRange().getValues();

  for (var i = 1; i < data.length; i++) {
    var row        = data[i];
    var memberName = String(row[0]).trim();
    var memberPhone = String(row[2]).trim().replace(/\D/g, ''); // 移除非數字
    var memberLast3 = memberPhone.slice(-3);

    if (memberName === name.trim() && memberLast3 === String(phoneLast3).trim()) {
      return {
        success: true,
        user: {
          name:   memberName,
          branch: String(row[1]).trim(),
          role:   String(row[3]).trim()
        }
      };
    }
  }

  return { success: false, error: '姓名或手機末三碼不符，請重新確認' };
}

// ============================================================
// 提交訪視表單
// ============================================================
function submitForm(record) {
  if (!record) return { success: false, error: '未收到表單資料' };

  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('訪視紀錄表');
  if (!sheet) return { success: false, error: '找不到「訪視紀錄表」工作表' };

  // 生成流水號（目前列數 = 已有筆數，不含 header）
  var id = sheet.getLastRow(); // header 是第1列，資料從第2列起，所以 lastRow 即為新 ID

  // 取得題庫順序，確保答案欄位對齊
  var qSheet = ss.getSheetByName('訪視題庫');
  var questionIds = [];
  if (qSheet) {
    var qData = qSheet.getDataRange().getValues();
    for (var i = 1; i < qData.length; i++) {
      if (qData[i][0]) questionIds.push(String(qData[i][0]).trim());
    }
  }

  // 組合列資料
  var rowData = [
    id,
    new Date(),               // 填報時間
    record.visitDate || '',   // 訪視日期
    record.submitter || '',
    Array.isArray(record.teamMembers)
      ? record.teamMembers.join(',')
      : (record.teamMembers || ''),
    record.branch    || '',
    record.clientName || '',
    record.clientPhone || '',
    record.clientAddress || '',
    record.gps       || ''
  ];

  // 依題目順序附加答案
  var answers = record.answers || {};
  questionIds.forEach(function(qid) {
    var ans = answers[qid];
    if (Array.isArray(ans)) {
      rowData.push(ans.join(','));
    } else {
      rowData.push(ans !== undefined ? String(ans) : '');
    }
  });

  sheet.appendRow(rowData);

  // 觸發分隊資料同步（非阻塞，失敗不影響主流程）
  try {
    splitToBranch(record, id);
  } catch (err) {
    Logger.log('【分隊同步失敗】' + err.toString());
  }

  return { success: true, id: id, message: '訪視紀錄已成功儲存！' };
}

// ============================================================
// 分隊自動拆分：將紀錄同步至分隊專屬試算表
// ============================================================
function splitToBranch(record, id) {
  var ss          = SpreadsheetApp.getActiveSpreadsheet();
  var branchSheet = ss.getSheetByName('分隊對照表');
  if (!branchSheet) return;

  var branchData = branchSheet.getDataRange().getValues();

  for (var i = 1; i < branchData.length; i++) {
    var row = branchData[i];
    if (String(row[0]).trim() !== record.branch) continue;

    var subSheetId = String(row[2]).trim();
    if (!subSheetId) {
      Logger.log('分隊「' + record.branch + '」尚未設定專屬試算表 ID，跳過同步。');
      return;
    }

    try {
      var subSS    = SpreadsheetApp.openById(subSheetId);
      var subSheet = subSS.getSheetByName('訪視紀錄');

      // 若分隊表尚無此工作表則自動建立並複製標題列
      if (!subSheet) {
        subSheet = subSS.insertSheet('訪視紀錄');
        var mainSheet   = ss.getSheetByName('訪視紀錄表');
        var headerRange = mainSheet.getRange(1, 1, 1, mainSheet.getLastColumn());
        subSheet.appendRow(headerRange.getValues()[0]);
      }

      // 找到剛寫入的那一列並複製
      var mainSheet = ss.getSheetByName('訪視紀錄表');
      var mainData  = mainSheet.getDataRange().getValues();
      for (var j = 1; j < mainData.length; j++) {
        if (String(mainData[j][0]) === String(id)) {
          subSheet.appendRow(mainData[j]);
          Logger.log('已同步紀錄 #' + id + ' → 分隊「' + record.branch + '」');
          break;
        }
      }
    } catch (err) {
      Logger.log('無法開啟分隊專屬試算表（ID: ' + subSheetId + '）：' + err.toString());
    }
    break;
  }
}
