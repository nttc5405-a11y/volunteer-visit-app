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
        result = getDashboardData(params.idCard, params.phone, params.name);
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
  if (data.length <= 1) return { success: true, data: [] };

  // 依標題列名稱定位欄位，日後在題庫增減欄位不影響此處邏輯
  var col = mapQuestionColumns_(data[0]);
  var questions = [];

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (!row[col.id]) continue;
    if (!isQuestionEnabled_(row, col)) continue;  // 「啟用」欄填 FALSE 的題目不出現在表單

    questions.push({
      id:         String(row[col.id]).trim(),
      visitType:  String(row[col.visitType]).trim(),
      dependency: String(row[col.dependency]).trim(),
      category:   String(row[col.category]).trim(),
      content:    String(row[col.content]).trim(),
      type:       String(row[col.type]).trim(),
      options:    row[col.options]
        ? String(row[col.options]).split(',').map(function(o) { return o.trim(); }).filter(function(o) { return o; })
        : [],
      required:   row[col.required] === true || String(row[col.required]).toUpperCase() === 'TRUE'
    });
  }

  return { success: true, data: questions };
}

// ============================================================
// 題庫 = 唯一真相來源：以下輔助函式讓「新增／調整題目」
// 只需要在「訪視題庫」工作表操作，不必修改程式碼。
// ============================================================

// 「訪視紀錄表」的固定基本欄位（答案欄以外的欄位）
var BASE_RECORD_HEADERS = [
  '流水號', '填報時間', '訪視日期', '訪視類型', '主填寫人姓名', '協同志工',
  '所屬分隊', '案家姓名', '案家性別', '案家電話', '案家地址', 'GPS定位座標',
  '房屋屋齡', '住宅形式', '總樓層', '居住樓層', '建築結構',
  '家庭總人數', '家庭65歲以上人數', '家庭行動不便人數', '家庭6歲以下人數', '家庭外籍人士人數',
  '受訪者簽名'
];

// 訪視題庫標題列 → 欄位索引（找不到的欄位沿用原本的固定位置，相容舊試算表）
function mapQuestionColumns_(headerRow) {
  var col = {
    id: 0, visitType: 1, dependency: 2, category: 3,
    content: 4, type: 5, options: 6, required: 7, enabled: -1
  };
  var names = {
    '題目代碼': 'id',      '訪視類型': 'visitType', '依賴條件': 'dependency',
    '題目分類': 'category', '題目內容': 'content',   '題型':     'type',
    '選項內容': 'options',  '必填':     'required',  '啟用':     'enabled'
  };

  for (var c = 0; c < headerRow.length; c++) {
    var key = names[String(headerRow[c]).trim()];
    if (key) col[key] = c;
  }
  return col;
}

// 「啟用」欄：留空或 TRUE 視為啟用（舊試算表沒有此欄時一律視為啟用）
function isQuestionEnabled_(row, col) {
  if (col.enabled < 0) return true;
  var v = row[col.enabled];
  if (v === '' || v === null || v === undefined) return true;
  if (v === false) return false;
  return String(v).toUpperCase() !== 'FALSE';
}

// 判斷標題是否為題目答案欄（題目代碼格式：1~3 個英文字母 + 數字，例 F01、D18、F25）
function isAnswerColumn_(name) {
  return /^[A-Za-z]{1,3}\d{1,3}$/.test(String(name).trim());
}

// 取得題庫全部題目代碼（含已停用者：欄位保留才不會弄丟歷史資料）
function getQuestionCodes_(ss) {
  var qSheet = ss.getSheetByName('訪視題庫');
  if (!qSheet) return [];

  var qData = qSheet.getDataRange().getValues();
  if (qData.length <= 1) return [];

  var col   = mapQuestionColumns_(qData[0]);
  var codes = [];

  for (var i = 1; i < qData.length; i++) {
    var code = String(qData[i][col.id] || '').trim();
    if (code && codes.indexOf(code) === -1) codes.push(code);
  }
  return codes;
}

// 讓目標工作表具備 mainHeaders 的所有欄位（依「欄位名稱」比對，缺少的補在最後）。
// 用於把主表的欄位同步到各分隊專屬試算表。
// 回傳目標工作表同步後的標題列陣列。
function syncHeadersByName_(sheet, mainHeaders) {
  var lastCol = sheet.getLastColumn();
  var headers = lastCol > 0 ? sheet.getRange(1, 1, 1, lastCol).getValues()[0] : [];
  var missing = [];

  mainHeaders.forEach(function(h) {
    var name = String(h).trim();
    if (!name) return;
    for (var i = 0; i < headers.length; i++) {
      if (String(headers[i]).trim() === name) return;
    }
    if (missing.indexOf(name) === -1) missing.push(name);
  });

  if (missing.length > 0) {
    sheet.getRange(1, headers.length + 1, 1, missing.length).setValues([missing]);
    headers = headers.concat(missing);
    Logger.log('已於「' + sheet.getName() + '」補上欄位：' + missing.join('、'));
  }

  return headers;
}

// 確保工作表已含題庫所有題目代碼欄位，缺少的自動補在最後。
// 回傳同步後的標題列陣列。
function ensureAnswerColumns_(sheet, codes) {
  var lastCol = sheet.getLastColumn();
  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  var missing = [];

  codes.forEach(function(code) {
    for (var i = 0; i < headers.length; i++) {
      if (String(headers[i]).trim() === code) return;
    }
    if (missing.indexOf(code) === -1) missing.push(code);
  });

  if (missing.length > 0) {
    sheet.getRange(1, lastCol + 1, 1, missing.length).setValues([missing]);
    headers = headers.concat(missing);
    _styleHeader(sheet, headers.length);
    Logger.log('已於「' + sheet.getName() + '」補上欄位：' + missing.join('、'));
  }

  return headers;
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

  // 題庫若新增了題目，先自動補上對應欄位（漏點選單也不會掉資料）
  var headers = ensureAnswerColumns_(sheet, getQuestionCodes_(ss));

  // 基本欄位值（以「欄位名稱」為鍵）
  var baseValues = {
    '流水號':   id,
    '填報時間': new Date(),
    '訪視日期': record.visitDate  || '',
    '訪視類型': record.visitType  || '',
    '主填寫人姓名': record.submitter || '',
    '協同志工': Array.isArray(record.teamMembers)
      ? record.teamMembers.join(',')
      : (record.teamMembers || ''),
    '所屬分隊':   record.branch        || '',
    '案家姓名':   record.clientName    || '',
    '案家性別':   record.clientGender  || '',
    '案家電話':   record.clientPhone   || '',
    '案家地址':   record.clientAddress || '',
    'GPS定位座標': record.gps          || '',
    '房屋屋齡':   record.houseAge          !== undefined ? String(record.houseAge) : '',
    '住宅形式':   record.residentialType   || '',
    '總樓層':     record.totalFloors       !== undefined ? String(record.totalFloors) : '',
    '居住樓層':   record.residingFloor     !== undefined ? String(record.residingFloor) : '',
    '建築結構':   record.buildingStructure || '',
    '家庭總人數': record.familySize        !== undefined ? String(record.familySize) : '',
    '家庭65歲以上人數': record.family65Plus    !== undefined ? String(record.family65Plus) : '',
    '家庭行動不便人數': record.familyDisabled  !== undefined ? String(record.familyDisabled) : '',
    '家庭6歲以下人數':  record.familyUnder6    !== undefined ? String(record.familyUnder6) : '',
    '家庭外籍人士人數': record.familyForeigner !== undefined ? String(record.familyForeigner) : '',
    '受訪者簽名': signatureUrl
  };

  // 逐欄對照「欄位名稱」填值：答案不再依題庫列順序排列，
  // 因此題庫可自由拖曳排序、於中間插入新題目而不會錯位。
  var answers = record.answers || {};
  var rowData = headers.map(function(h) {
    var name = String(h).trim();

    if (baseValues.hasOwnProperty(name)) return baseValues[name];

    if (isAnswerColumn_(name)) {
      var ans = answers[name];
      if (Array.isArray(ans)) return ans.join(',');
      return (ans !== undefined && ans !== null) ? String(ans) : '';
    }

    return '';
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
      var subSS     = SpreadsheetApp.openById(subSheetId);
      var subSheet  = subSS.getSheetByName('訪視紀錄');
      var mainSheet = ss.getSheetByName('訪視紀錄表');
      var mainHeaders = mainSheet.getRange(1, 1, 1, mainSheet.getLastColumn()).getValues()[0];

      // 若分隊表尚無此工作表則自動建立並複製標題列
      if (!subSheet) {
        subSheet = subSS.insertSheet('訪視紀錄');
        subSheet.appendRow(mainHeaders);
      }

      // 主表新增題目欄位後，分隊表標題列一併補齊（依欄位名稱比對，
      // 分隊自行加過的欄位不會被覆蓋或誤判為已同步）
      var subHeaders = syncHeadersByName_(subSheet, mainHeaders);

      // 找到剛寫入的那一列，依「欄位名稱」對應後再寫入分隊表，
      // 分隊表的欄位順序與主表不同也不會錯位
      var mainData = mainSheet.getDataRange().getValues();
      for (var j = 1; j < mainData.length; j++) {
        if (String(mainData[j][0]) === String(id)) {
          var valueByName = {};
          for (var k = 0; k < mainHeaders.length; k++) {
            valueByName[String(mainHeaders[k]).trim()] = mainData[j][k];
          }

          subSheet.appendRow(subHeaders.map(function(h) {
            var name = String(h).trim();
            return valueByName.hasOwnProperty(name) ? valueByName[name] : '';
          }));

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
function getDashboardData(idCardLast3, phoneLast3, name) {
  // 儀表板含案家姓名、地址、電話等個資，必須先驗證身分才回傳資料。
  var auth = verifyLogin(idCardLast3, phoneLast3, name);
  if (!auth.success || !auth.user) {
    return { success: false, error: '請重新登入後再開啟儀表板。' };
  }

  var role = String(auth.user.role || '志工').trim();
  if (role !== '管理員' && role !== '分隊承辦人') {
    return { success: false, error: '權限不足：僅限管理員與分隊承辦人檢視統計資料。' };
  }

  // 分隊承辦人只能看自己分隊；管理員看全部
  var scopeBranch = (role === '分隊承辦人') ? String(auth.user.branch || '').trim() : '';

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('訪視紀錄表');
  if (!sheet) return { success: false, error: '找不到「訪視紀錄表」工作表。' };

  var data = sheet.getDataRange().getValues();
  if (data.length <= 1) {
    return { success: true, data: [], viewer: { name: auth.user.name, role: role, branch: scopeBranch } };
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

    // 掃描標題列中所有題目答案欄（F01、D18、F25…）
    // 新增題目時不需要修改這段程式
    for (var c = 0; c < headers.length; c++) {
      var h = String(headers[c]).trim();
      if (isAnswerColumn_(h)) answers[h] = row[c];
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

    var branchName = row[colIdx['所屬分隊'] || 6];

    // 分隊承辦人只取自己分隊的紀錄
    if (scopeBranch && String(branchName).trim() !== scopeBranch) continue;

    // 取數值欄位（空白視為 0）
    var num = function(header) {
      if (colIdx[header] === undefined) return 0;
      return parseInt(row[colIdx[header]], 10) || 0;
    };
    var text = function(header) {
      if (colIdx[header] === undefined) return '';
      return row[colIdx[header]];
    };

    records.push({
      id:                id,
      timestamp:         row[colIdx['填報時間'] || 1],
      visitDate:         visitDateVal,
      visitType:         row[colIdx['訪視類型'] || 3],
      submitter:         row[colIdx['主填寫人姓名'] || 4],
      branch:            branchName,
      clientName:        row[colIdx['案家姓名'] || 7],
      clientGender:      text('案家性別'),
      clientPhone:       text('案家電話'),
      clientAddress:     text('案家地址'),
      gps:               row[colIdx['GPS定位座標'] || 11],
      houseAge:          parseInt(row[colIdx['房屋屋齡'] || 12]) || 0,
      residentialType:   row[colIdx['住宅形式'] || 13],
      totalFloors:       num('總樓層'),
      residingFloor:     num('居住樓層'),
      buildingStructure: row[colIdx['建築結構'] || 16],
      familySize:        parseInt(row[colIdx['家庭總人數'] || 17]) || 0,
      family65Plus:      num('家庭65歲以上人數'),
      familyDisabled:    num('家庭行動不便人數'),
      familyUnder6:      num('家庭6歲以下人數'),
      familyForeigner:   num('家庭外籍人士人數'),
      answers:           answers
    });
  }

  return {
    success: true,
    data: records,
    viewer: { name: auth.user.name, role: role, branch: scopeBranch }
  };
}
