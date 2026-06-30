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
        result = verifyLogin(params.idCard, params.phone, params.name);
        break;
      case 'getDashboardData':
        result = getDashboardData();
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
      id:         String(row[0]).trim(),
      visitType:  String(row[1]).trim(),
      dependency: String(row[2]).trim(),
      category:   String(row[3]).trim(),
      content:    String(row[4]).trim(),
      type:       String(row[5]).trim(),
      options:    row[6]
        ? String(row[6]).split(',').map(function(o) { return o.trim(); }).filter(function(o) { return o; })
        : [],
      required:   row[7] === true || String(row[7]).toUpperCase() === 'TRUE'
    });
  }

  return { success: true, data: questions };
}

// ============================================================
// 輔助函式：動態取得「人員帳號管理」各欄位的欄位索引 (0-based)
// ============================================================
function getColumnIndexes(headers) {
  var indexes = {
    name: -1,
    branch: -1,
    phone: -1,
    role: -1,
    idCard: -1
  };
  
  for (var i = 0; i < headers.length; i++) {
    var header = String(headers[i]).trim();
    if (header.indexOf('姓名') !== -1) {
      indexes.name = i;
    } else if (header.indexOf('單位') !== -1 || header.indexOf('分隊') !== -1) {
      indexes.branch = i;
    } else if (header.indexOf('手機') !== -1 || header.indexOf('電話') !== -1) {
      indexes.phone = i;
    } else if (header.indexOf('角色') !== -1 || header.indexOf('權限') !== -1) {
      indexes.role = i;
    } else if (header.indexOf('身分證') !== -1 || header.indexOf('身份證') !== -1) {
      indexes.idCard = i;
    }
  }
  return indexes;
}

// ============================================================
// 取得所有人員（不回傳手機號碼與身分證字號，保護隱私）
// ============================================================
function getAllMembers() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('人員帳號管理');
  if (!sheet) return { success: false, error: '找不到「人員帳號管理」工作表。' };

  var data = sheet.getDataRange().getValues();
  var idx = getColumnIndexes(data[0]);

  if (idx.name === -1) {
    return { success: false, error: '人員工作表缺少「姓名」欄位' };
  }

  var members = [];
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var name = String(row[idx.name]).trim();
    if (!name) continue;

    members.push({
      name:   name,
      branch: idx.branch !== -1 ? String(row[idx.branch]).trim() : '無分隊',
      role:   idx.role !== -1 ? String(row[idx.role]).trim() : '志工'
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
// 登入驗證：比對身分證字號後三碼 + 手機末三碼（若重複則需比對姓名）
// ============================================================
function verifyLogin(idCardLast3, phoneLast3, name) {
  if (!idCardLast3 || !phoneLast3) {
    return { success: false, error: '請填寫身分證末三碼與手機末三碼' };
  }

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('人員帳號管理');
  if (!sheet) return { success: false, error: '系統錯誤：找不到人員資料工作表' };

  var data = sheet.getDataRange().getValues();
  var idx = getColumnIndexes(data[0]);

  if (idx.name === -1 || idx.phone === -1 || idx.idCard === -1) {
    return { success: false, error: '系統錯誤：人員工作表缺少必要欄位（姓名、手機或身分證字號）' };
  }

  var matches = [];

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var memberName = String(row[idx.name]).trim();
    if (!memberName) continue;

    var memberPhone = String(row[idx.phone]).trim().replace(/\D/g, ''); // 移除非數字
    var memberIdCard = String(row[idx.idCard]).trim().replace(/[^a-zA-Z0-9]/g, ''); // 移除非英數

    var memberPhoneLast3 = memberPhone.slice(-3);
    var memberIdCardLast3 = memberIdCard.slice(-3);

    if (
      memberIdCardLast3.toUpperCase() === String(idCardLast3).trim().toUpperCase() && 
      memberPhoneLast3 === String(phoneLast3).trim()
    ) {
      matches.push({
        name:   memberName,
        branch: idx.branch !== -1 ? String(row[idx.branch]).trim() : '',
        role:   idx.role !== -1 ? String(row[idx.role]).trim() : '志工'
      });
    }
  }

  if (matches.length === 0) {
    return { success: false, error: '驗證失敗，查無此人員，請確認填寫的末三碼是否正確' };
  }

  // 如果有姓名參數，進行精確匹配
  if (name) {
    var cleanName = String(name).trim();
    for (var j = 0; j < matches.length; j++) {
      if (matches[j].name === cleanName) {
        return {
          success: true,
          user: matches[j]
        };
      }
    }
  }

  // 如果只有一個匹配，直接登入
  if (matches.length === 1) {
    return {
      success: true,
      user: matches[0]
    };
  }

  // 如果有多個匹配，返回需要姓名核對
  return {
    success: false,
    needNameDisambiguation: true,
    candidates: matches.map(function(m) {
      return { name: m.name, branch: m.branch };
    })
  };
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

  // 儲存受訪人手寫簽名圖檔至 Google Drive
  var signatureUrl = '';
  if (record.signature && String(record.signature).indexOf('data:') === 0) {
    try {
      var folders = DriveApp.getFoldersByName('志工訪視系統_受訪人簽名');
      var folder;
      if (folders.hasNext()) {
        folder = folders.next();
      } else {
        folder = DriveApp.createFolder('志工訪視系統_受訪人簽名');
      }

      var contentType = 'image/jpeg';
      var ext = '.jpg';
      if (record.signature.indexOf('image/png') !== -1) {
        contentType = 'image/png';
        ext = '.png';
      }

      var matches = record.signature.match(/^data:(image\/[a-z]+);base64,(.+)$/);
      if (matches && matches.length === 3) {
        contentType = matches[1];
        var base64Data = matches[2];
        var imageBlob = Utilities.newBlob(Utilities.base64Decode(base64Data), contentType, '簽名_' + id + '_' + (record.clientName || '未命名') + ext);
        var file = folder.createFile(imageBlob);
        file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        signatureUrl = file.getUrl();
      }
    } catch (e) {
      Logger.log('【簽名檔儲存失敗】' + e.toString());
      signatureUrl = '儲存失敗：' + e.toString();
    }
  }

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
    new Date(),                       // 填報時間
    record.visitDate || '',           // 訪視日期
    record.visitType || '',           // 訪視類型
    record.submitter || '',
    Array.isArray(record.teamMembers)
      ? record.teamMembers.join(',')
      : (record.teamMembers || ''),
    record.branch    || '',
    record.clientName || '',
    record.clientGender || '',        // 案家性別
    record.clientPhone || '',
    record.clientAddress || '',
    record.gps       || '',           // GPS定位座標
    record.houseAge !== undefined ? String(record.houseAge) : '', // 房屋屋齡
    record.residentialType || '',     // 住宅形式
    record.totalFloors !== undefined ? String(record.totalFloors) : '', // 總樓層
    record.residingFloor !== undefined ? String(record.residingFloor) : '', // 居住樓層
    record.buildingStructure || '',   // 建築結構
    record.familySize !== undefined ? String(record.familySize) : '', // 家庭總人數
    record.family65Plus !== undefined ? String(record.family65Plus) : '', // 65歲以上
    record.familyDisabled !== undefined ? String(record.familyDisabled) : '', // 行動不便
    record.familyUnder6 !== undefined ? String(record.familyUnder6) : '', // 6歲以下
    record.familyForeigner !== undefined ? String(record.familyForeigner) : '', // 外籍人士
    signatureUrl                      // 受訪者簽名
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

// ============================================================
// 取得儀表板分析所需的訪視紀錄統計資料
// ============================================================
function getDashboardData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('訪視紀錄表');
  if (!sheet) return { success: false, error: '找不到「訪視紀錄表」工作表。' };

  var data = sheet.getDataRange().getValues();
  if (data.length <= 1) {
    return { success: true, data: [] };
  }

  var headers = data[0];
  var colIdx = {};
  for (var c = 0; c < headers.length; c++) {
    colIdx[String(headers[c]).trim()] = c;
  }

  var records = [];
  
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var id = row[colIdx['流水號'] || 0];
    if (!id) continue;

    var answers = {};
    
    // 將 F01~F24, D01~D18 問卷回答加入 answers 物件
    for (var f = 1; f <= 24; f++) {
      var fid = 'F' + String(f).padStart(2, '0');
      if (colIdx[fid] !== undefined) {
        answers[fid] = row[colIdx[fid]];
      }
    }
    for (var d = 1; d <= 18; d++) {
      var did = 'D' + String(d).padStart(2, '0');
      if (colIdx[did] !== undefined) {
        answers[did] = row[colIdx[did]];
      }
    }

    // 格式化日期防止 json 解析出錯
    var visitDateVal = '';
    try {
      if (row[colIdx['訪視日期'] || 2]) {
        visitDateVal = Utilities.formatDate(new Date(row[colIdx['訪視日期'] || 2]), ss.getSpreadsheetTimeZone(), 'yyyy-MM-dd');
      }
    } catch (e) {
      visitDateVal = String(row[colIdx['訪視日期'] || 2]);
    }

    records.push({
      id:                id,
      timestamp:         row[colIdx['填報時間'] || 1],
      visitDate:         visitDateVal,
      visitType:         row[colIdx['訪視類型'] || 3],
      submitter:         row[colIdx['主填寫人姓名'] || 4],
      branch:            row[colIdx['所屬分隊'] || 6],
      clientName:        row[colIdx['案家姓名'] || 7],
      gps:               row[colIdx['GPS定位座標'] || 11],
      houseAge:          parseInt(row[colIdx['房屋屋齡'] || 12]) || 0,
      residentialType:   row[colIdx['住宅形式'] || 13],
      buildingStructure: row[colIdx['建築結構'] || 16],
      familySize:        parseInt(row[colIdx['家庭總人數'] || 17]) || 0,
      answers:           answers
    });
  }

  return { success: true, data: records };
}
