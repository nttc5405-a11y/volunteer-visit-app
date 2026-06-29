/**
 * 志工居家訪視系統 — Google Sheets 初始化腳本
 *
 * ╔══════════════════════════════════════════════════════════════╗
 * ║  可執行的函式列表：                                           ║
 * ║  1. initializeSheets()   → 建立主資料庫的 4 張工作表          ║
 * ║  2. createBranchSheets() → 自動建立 3 個分隊專屬試算表        ║
 * ║  3. resetAllSheets()     → 重置所有資料（謹慎使用）           ║
 * ╚══════════════════════════════════════════════════════════════╝
 *
 * ===== 使用方式 =====
 * 在 Google Apps Script 編輯器中，
 * 選擇函式「initializeSheets」，點擊執行（▶）。
 * 首次執行需授予試算表存取權限。
 * 執行完成後，4 張工作表與所有範例資料將自動建立完成。
 * ====================
 */

// ============================================================
// 主初始化函式（執行此函式）
// ============================================================
function initializeSheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  _createVisitRecordSheet(ss);
  _createMemberSheet(ss);
  _createBranchSheet(ss);
  _createQuestionSheet(ss);

  // 刪除預設的 "工作表1"（如存在）
  var defaultSheet = ss.getSheetByName('工作表1') || ss.getSheetByName('Sheet1');
  if (defaultSheet && ss.getNumSheets() > 4) {
    ss.deleteSheet(defaultSheet);
  }

  Logger.log('✅ 初始化完成！4 張工作表已建立，範例資料已填入。');

  try {
    SpreadsheetApp.getUi().alert(
      '✅ 初始化完成！\n\n' +
      '已建立以下工作表：\n' +
      '• 訪視紀錄表\n' +
      '• 人員帳號管理（7 筆範例人員）\n' +
      '• 分隊對照表（3 個分隊）\n' +
      '• 訪視題庫（5 題範例題目）\n\n' +
      '請接著部署 Code.gs 為 Web App 取得 API 網址。'
    );
  } catch (e) {
    // 若在非 UI 環境執行，跳過彈窗
  }
}

// ============================================================
// 建立「訪視紀錄表」
// ============================================================
function _createVisitRecordSheet(ss) {
  var sheet = ss.getSheetByName('訪視紀錄表');
  if (!sheet) {
    sheet = ss.insertSheet('訪視紀錄表', 0);
  } else {
    sheet.clearContents();
    sheet.clearFormats();
  }

  var headers = [
    '流水號', '訪視日期時間', '主填寫人姓名', '協同志工',
    '所屬分隊', '案家姓名', '案家電話', '案家地址', 'GPS定位座標',
    'Q01', 'Q02', 'Q03', 'Q04', 'Q05'
  ];

  sheet.appendRow(headers);
  _styleHeader(sheet, headers.length);

  // 凍結標題列
  sheet.setFrozenRows(1);

  // 設定欄寬
  sheet.setColumnWidth(1, 70);  // 流水號
  sheet.setColumnWidth(2, 160); // 日期時間
  sheet.setColumnWidth(3, 100); // 主填寫人
  sheet.setColumnWidth(4, 120); // 協同志工
  sheet.setColumnWidth(5, 90);  // 分隊
  sheet.setColumnWidth(6, 100); // 案家姓名
  sheet.setColumnWidth(7, 120); // 電話
  sheet.setColumnWidth(8, 200); // 地址
  sheet.setColumnWidth(9, 180); // GPS

  Logger.log('✓ 訪視紀錄表 建立完成');
}

// ============================================================
// 建立「人員帳號管理」（含範例人員）
// ============================================================
function _createMemberSheet(ss) {
  var sheet = ss.getSheetByName('人員帳號管理');
  if (!sheet) {
    sheet = ss.insertSheet('人員帳號管理', 1);
  } else {
    sheet.clearContents();
    sheet.clearFormats();
  }

  var headers = ['人員姓名', '所屬分隊', '手機號碼', '權限角色'];
  sheet.appendRow(headers);
  _styleHeader(sheet, headers.length);
  sheet.setFrozenRows(1);

  // 範例人員資料（3 分隊 × 各2志工 + 1管理員）
  var members = [
    // 第一分隊
    ['張小明', '第一分隊', '0912345678', '分隊承辦人'],
    ['李美玲', '第一分隊', '0923456789', '志工'],
    ['陳建宏', '第一分隊', '0934567890', '志工'],
    // 第二分隊
    ['王淑芬', '第二分隊', '0945678901', '分隊承辦人'],
    ['林志偉', '第二分隊', '0956789012', '志工'],
    ['黃雅婷', '第二分隊', '0967890123', '志工'],
    // 第三分隊
    ['吳家豪', '第三分隊', '0978901234', '分隊承辦人'],
    ['蔡佩君', '第三分隊', '0989012345', '志工'],
    ['許文哲', '第三分隊', '0990123456', '志工'],
    // 管理員
    ['系統管理員', '（管理員）', '0900000000', '管理員']
  ];

  members.forEach(function(row) { sheet.appendRow(row); });

  // 設定欄寬
  sheet.setColumnWidth(1, 100);
  sheet.setColumnWidth(2, 100);
  sheet.setColumnWidth(3, 130);
  sheet.setColumnWidth(4, 100);

  Logger.log('✓ 人員帳號管理 建立完成（' + members.length + ' 筆）');
}

// ============================================================
// 建立「分隊對照表」（含範例分隊）
// ============================================================
function _createBranchSheet(ss) {
  var sheet = ss.getSheetByName('分隊對照表');
  if (!sheet) {
    sheet = ss.insertSheet('分隊對照表', 2);
  } else {
    sheet.clearContents();
    sheet.clearFormats();
  }

  var headers = ['分隊名稱', '分隊承辦人帳號', '分隊專屬試算表ID'];
  sheet.appendRow(headers);
  _styleHeader(sheet, headers.length);
  sheet.setFrozenRows(1);

  // 範例分隊（試算表ID 待建立後填入）
  var branches = [
    ['第一分隊', 'branch1-manager@example.com', ''],
    ['第二分隊', 'branch2-manager@example.com', ''],
    ['第三分隊', 'branch3-manager@example.com', '']
  ];

  branches.forEach(function(row) { sheet.appendRow(row); });

  // 加入說明備註
  sheet.getRange('C2').setNote(
    '在此填入各分隊獨立試算表的 ID\n' +
    '（從試算表網址中的 /d/[試算表ID]/ 取得）\n' +
    '若留空，則不進行自動同步。'
  );

  sheet.setColumnWidth(1, 100);
  sheet.setColumnWidth(2, 220);
  sheet.setColumnWidth(3, 250);

  Logger.log('✓ 分隊對照表 建立完成');
}

// ============================================================
// 建立「訪視題庫」（含 PRD 範例題目）
// ============================================================
function _createQuestionSheet(ss) {
  var sheet = ss.getSheetByName('訪視題庫');
  if (!sheet) {
    sheet = ss.insertSheet('訪視題庫', 3);
  } else {
    sheet.clearContents();
    sheet.clearFormats();
  }

  var headers = ['題目代碼', '題目分類', '題目內容', '題型', '選項內容', '必填'];
  sheet.appendRow(headers);
  _styleHeader(sheet, headers.length);
  sheet.setFrozenRows(1);

  // 題型說明備註
  sheet.getRange('D1').setNote(
    '支援題型：\n是否題 / 單選題 / 複選題 / 簡答題'
  );
  sheet.getRange('E1').setNote(
    '選項內容以半形逗號「,」分隔\n簡答題此欄留空'
  );
  sheet.getRange('F1').setNote(
    '填 TRUE 表必填（前端驗證），\n填 FALSE 或空白表選填'
  );

  // PRD 定義的 5 道範例題目（第6欄：必填）
  var questions = [
    ['Q01', '宣導題',   '是否已向案家宣導住宅用火災警報器之重要性？',   '是否題', '是,否',                                    true],
    ['Q02', '宣導題',   '是否已向案家宣導用電安全與防範電氣火災？',     '是否題', '是,否',                                    true],
    ['Q03', '生理健康', '案主今日精神與氣色狀況如何？',                 '單選題', '良好,普通,稍微疲倦,極度不適',               true],
    ['Q04', '生活需求', '案主目前是否有迫切物資或醫療照護需求？',       '複選題', '餐食協助,醫療就醫,物資缺乏,心理諮商,無需求', true],
    ['Q05', '綜合紀錄', '請簡述本次訪視的具體狀況或特殊交辦事項。',     '簡答題', '',                                         false]
  ];

  questions.forEach(function(row) { sheet.appendRow(row); });

  sheet.setColumnWidth(1, 80);
  sheet.setColumnWidth(2, 90);
  sheet.setColumnWidth(3, 280);
  sheet.setColumnWidth(4, 80);
  sheet.setColumnWidth(5, 220);
  sheet.setColumnWidth(6, 60);

  Logger.log('✓ 訪視題庫 建立完成（' + questions.length + ' 題）');
}

// ============================================================
// 套用標題列樣式
// ============================================================
function _styleHeader(sheet, colCount) {
  var headerRange = sheet.getRange(1, 1, 1, colCount);
  headerRange
    .setBackground('#1E3A8A')
    .setFontColor('#FFFFFF')
    .setFontWeight('bold')
    .setFontSize(11)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
  sheet.setRowHeight(1, 36);
}

// ============================================================
// 重置所有資料（謹慎使用：會清空全部內容）
// ============================================================
function resetAllSheets() {
  var ui = SpreadsheetApp.getUi();
  var confirm = ui.alert(
    '⚠️ 警告',
    '此操作將清空所有工作表的資料，是否繼續？',
    ui.ButtonSet.YES_NO
  );

  if (confirm === ui.Button.YES) {
    initializeSheets();
  }
}

// ============================================================
// 自動建立三個分隊專屬試算表，並回填 ID 至「分隊對照表」
//
// 使用方式：
//   在 Apps Script 編輯器選擇「createBranchSheets」→ 點執行
//   執行後請查看「執行記錄」取得各分隊試算表連結
// ============================================================
function createBranchSheets() {
  var mainSS      = SpreadsheetApp.getActiveSpreadsheet();
  var branchSheet = mainSS.getSheetByName('分隊對照表');

  if (!branchSheet) {
    SpreadsheetApp.getUi().alert('❌ 找不到「分隊對照表」，請先執行 initializeSheets()。');
    return;
  }

  var branchData = branchSheet.getDataRange().getValues();
  var resultLines = [];

  // 取得主試算表標題列（用於同步至分隊表）
  var mainRecordSheet = mainSS.getSheetByName('訪視紀錄表');
  var mainHeaders = mainRecordSheet
    ? mainRecordSheet.getRange(1, 1, 1, mainRecordSheet.getLastColumn()).getValues()[0]
    : ['流水號','訪視日期時間','主填寫人姓名','協同志工','所屬分隊',
       '案家姓名','案家電話','案家地址','GPS定位座標','Q01','Q02','Q03','Q04','Q05'];

  // 逐列處理分隊
  for (var i = 1; i < branchData.length; i++) {
    var row        = branchData[i];
    var branchName = String(row[0]).trim();
    var manager    = String(row[1]).trim();
    var existingId = String(row[2]).trim();

    if (!branchName) continue;

    // 若已有試算表 ID 則跳過，不重複建立
    if (existingId && existingId.length > 10) {
      Logger.log('⏭️ 「' + branchName + '」已有試算表，略過（ID: ' + existingId + '）');
      resultLines.push('⏭️ ' + branchName + '：已存在，未重新建立');
      continue;
    }

    try {
      // 建立新試算表
      var newSS   = SpreadsheetApp.create('志工訪視_' + branchName);
      var newId   = newSS.getId();
      var newUrl  = newSS.getUrl();

      // 建立「訪視紀錄」工作表
      var recordSheet = newSS.getSheets()[0];
      recordSheet.setName('訪視紀錄');
      recordSheet.appendRow(mainHeaders);
      _styleHeader(recordSheet, mainHeaders.length);
      recordSheet.setFrozenRows(1);

      // 調整欄寬（與主表相同）
      recordSheet.setColumnWidth(1, 70);
      recordSheet.setColumnWidth(2, 160);
      recordSheet.setColumnWidth(3, 100);
      recordSheet.setColumnWidth(4, 120);
      recordSheet.setColumnWidth(5, 90);
      recordSheet.setColumnWidth(6, 100);
      recordSheet.setColumnWidth(7, 120);
      recordSheet.setColumnWidth(8, 200);
      recordSheet.setColumnWidth(9, 180);

      // 建立「分隊資訊」說明工作表
      var infoSheet = newSS.insertSheet('分隊資訊');
      infoSheet.appendRow(['分隊名稱', branchName]);
      infoSheet.appendRow(['承辦人', manager]);
      infoSheet.appendRow(['建立時間', new Date()]);
      infoSheet.appendRow(['說明', '本試算表由系統自動建立，存放「' + branchName + '」的訪視紀錄。']);
      infoSheet.appendRow(['', '請勿修改「訪視紀錄」工作表的標題列。']);
      infoSheet.setColumnWidth(1, 100);
      infoSheet.setColumnWidth(2, 300);

      // 將試算表 ID 回填至「分隊對照表」的 C 欄
      branchSheet.getRange(i + 1, 3).setValue(newId);

      Logger.log('✅ 建立完成：' + branchName);
      Logger.log('   試算表ID：' + newId);
      Logger.log('   連結：' + newUrl);

      resultLines.push('✅ ' + branchName + '\n   ID: ' + newId + '\n   連結: ' + newUrl);

    } catch (err) {
      Logger.log('❌ 建立「' + branchName + '」失敗：' + err.toString());
      resultLines.push('❌ ' + branchName + '：建立失敗（' + err.message + '）');
    }
  }

  // 完成提示
  var summary = resultLines.join('\n\n');
  Logger.log('\n===== 建立完畢 =====\n' + summary);

  try {
    SpreadsheetApp.getUi().alert(
      '✅ 分隊試算表建立完成！\n\n' +
      '各分隊試算表已建立，ID 已自動填入「分隊對照表」。\n\n' +
      '⚠️ 下一步：\n' +
      '請在各分隊試算表的「共用」設定中，\n' +
      '將各分隊承辦人的 Gmail 設為「編輯者」。\n\n' +
      '詳細連結請查看「執行記錄」（View → Logs）。'
    );
  } catch (e) {
    // 非 UI 環境
  }
}
