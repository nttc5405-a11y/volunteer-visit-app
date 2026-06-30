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

  // 先定義基本欄位
  var headers = [
    '流水號', '填報時間', '訪視日期', '訪視類型', '主填寫人姓名', '協同志工',
    '所屬分隊', '案家姓名', '案家性別', '案家電話', '案家地址', 'GPS定位座標',
    '房屋屋齡', '住宅形式', '總樓層', '居住樓層', '建築結構',
    '家庭總人數', '家庭65歲以上人數', '家庭行動不便人數', '家庭6歲以下人數', '家庭外籍人士人數',
    '受訪者簽名'
  ];
  
  // 動態把 F01~F24, D01~D18 加入標頭
  for (var f = 1; f <= 24; f++) {
    headers.push('F' + String(f).padStart(2, '0'));
  }
  for (var d = 1; d <= 18; d++) {
    headers.push('D' + String(d).padStart(2, '0'));
  }

  sheet.appendRow(headers);
  _styleHeader(sheet, headers.length);

  // 凍結標題列
  sheet.setFrozenRows(1);

  // 設定欄寬
  sheet.setColumnWidth(1, 70);   // 流水號
  sheet.setColumnWidth(2, 160);  // 填報時間
  sheet.setColumnWidth(3, 110);  // 訪視日期
  sheet.setColumnWidth(4, 90);   // 訪視類型
  sheet.setColumnWidth(5, 100);  // 主填寫人
  sheet.setColumnWidth(6, 120);  // 協同志工
  sheet.setColumnWidth(7, 90);   // 分隊
  sheet.setColumnWidth(8, 100);  // 案家姓名
  sheet.setColumnWidth(9, 70);   // 案家性別
  sheet.setColumnWidth(10, 120); // 電話
  sheet.setColumnWidth(11, 200); // 地址
  sheet.setColumnWidth(12, 180); // GPS
  sheet.setColumnWidth(13, 80);  // 房屋屋齡
  sheet.setColumnWidth(14, 90);  // 住宅形式
  sheet.setColumnWidth(15, 70);  // 總樓層
  sheet.setColumnWidth(16, 70);  // 居住樓層
  sheet.setColumnWidth(17, 90);  // 建築結構
  sheet.setColumnWidth(18, 90);  // 家庭總人數
  sheet.setColumnWidth(19, 110); // 65歲以上
  sheet.setColumnWidth(20, 110); // 行動不便
  sheet.setColumnWidth(21, 110); // 6歲以下
  sheet.setColumnWidth(22, 110); // 外籍人士
  sheet.setColumnWidth(23, 160); // 受訪者簽名

  Logger.log('✓ 訪視紀錄表 建立完成');
}

// ============================================================
// 建立「人員帳號管理」（含範例人員，欄位與 Excel 完全一致）
// ============================================================
function _createMemberSheet(ss) {
  var sheet = ss.getSheetByName('人員帳號管理');
  if (!sheet) {
    sheet = ss.insertSheet('人員帳號管理', 1);
  } else {
    sheet.clearContents();
    sheet.clearFormats();
  }

  var headers = ['所屬大隊', '所屬單位', '職稱', '姓名', '身分證字號', '手機', '權限角色'];
  sheet.appendRow(headers);
  _styleHeader(sheet, headers.length);
  sheet.setFrozenRows(1);

  // 範例人員資料（所屬大隊, 所屬單位, 職稱, 姓名, 身分證字號, 手機, 權限角色）
  var members = [
    // 臺東防宣
    ['台東大隊', '臺東防宣', '分隊長', '張小明', 'V123456789', '0912-345678', '分隊承辦人'],
    ['台東大隊', '臺東防宣', '隊員', '李美玲', 'V220031182', '0987-838406', '志工'],
    ['台東大隊', '臺東防宣', '隊員', '陳建宏', 'V220430658', '0928-787898', '志工'],
    // 關山防宣
    ['台東大隊', '關山防宣', '分隊長', '王淑芬', 'Q123456789', '0945-678901', '分隊承辦人'],
    ['台東大隊', '關山防宣', '隊員', '林志偉', 'Q220668764', '0988-231281', '志工'],
    ['台東大隊', '關山防宣', '隊員', '黃雅婷', 'N221225092', '0911-227392', '志工'],
    // 成功防宣
    ['台東大隊', '成功防宣', '分隊長', '吳家豪', 'P123456789', '0978-901234', '分隊承辦人'],
    ['台東大隊', '成功防宣', '隊員', '蔡佩君', 'V221171201', '0937-168008', '志工'],
    ['台東大隊', '成功防宣', '隊員', '許文哲', 'N224758810', '0912-151341', '志工'],
    // 管理員
    ['台東大隊', '（管理員）', '系統管理員', '系統管理員', 'A123456789', '0900-000000', '管理員']
  ];

  members.forEach(function(row) { sheet.appendRow(row); });

  // 設定欄寬
  sheet.setColumnWidth(1, 100); // 所屬大隊
  sheet.setColumnWidth(2, 100); // 所屬單位
  sheet.setColumnWidth(3, 100); // 職稱
  sheet.setColumnWidth(4, 100); // 姓名
  sheet.setColumnWidth(5, 120); // 身分證字號
  sheet.setColumnWidth(6, 120); // 手機
  sheet.setColumnWidth(7, 100); // 權限角色

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
    ['臺東防宣', 'taitung-manager@example.com', ''],
    ['關山防宣', 'guanshan-manager@example.com', ''],
    ['成功防宣', 'chenggong-manager@example.com', '']
  ];

  branches.forEach(function(row) { sheet.appendRow(row); });

  // 加入說明備註
  sheet.getRange('C2').setNote(
    '在此填入各分隊獨立試算表的 ID\n' +
    '（從試算表網址中的 /d/[試算表ID]/ 取得）\n' +
    '若留空，則不進行自動同步。'
  );

  sheet.setColumnWidth(1, 120);
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

  var headers = ['題目代碼', '訪視類型', '依賴條件', '題目分類', '題目內容', '題型', '選項內容', '必填'];
  sheet.appendRow(headers);
  _styleHeader(sheet, headers.length);
  sheet.setFrozenRows(1);

  // 題型說明備註
  sheet.getRange('F1').setNote(
    '支援題型：\n是否題 / 單選題 / 複選題 / 簡答題'
  );
  sheet.getRange('G1').setNote(
    '選項內容以半形逗號「,」分隔\n簡答題此欄留空'
  );
  sheet.getRange('H1').setNote(
    '填 TRUE 表必填（前端驗證），\n填 FALSE 或空白表選填'
  );

  // 防火/防災宣導 題目資料庫
  var questions = [
    // === 防火宣導 (F01 ~ F24) ===
    ['F01', '防火宣導', '', '瓦斯安全', '家中是否使用桶裝瓦斯？', '是否題', '是,未使用', true],
    ['F02', '防火宣導', 'F01=是', '瓦斯安全', '瓦斯桶檢驗期限是否合格無逾期？', '是否題', '是,否', false],
    ['F03', '防火宣導', 'F01=是', '瓦斯安全', '瓦斯桶外觀及皮管是否有定期檢查，無龜裂、老化、鬆脫、腐蝕或變形現象？', '是否題', '是,否', false],
    ['F04', '防火宣導', 'F01=是', '瓦斯安全', '使用瓦斯完畢後是否有隨手關緊開關之習慣？', '是否題', '是,否', false],
    ['F05', '防火宣導', 'F01=是', '瓦斯安全', '是否有與瓦斯業者簽訂液化石油氣定型化契約？', '是否題', '是,否', false],
    ['F06', '防火宣導', '', '瓦斯安全', '是否知道瓦斯熱水器裝設室內容易造成一氧化碳中毒？', '是否題', '是,否', true],
    ['F07', '防火宣導', '', '瓦斯安全', '家中是否使用瓦斯熱水器？', '是否題', '是,否', true],
    ['F08', '防火宣導', 'F07=是', '瓦斯安全', '熱水器是否貼有 CNS 及 TGAS 的標示？', '是否題', '是,否', false],
    ['F09', '防火宣導', 'F07=是', '瓦斯安全', '裝設熱水器位置(如陽台)是否通風良好？', '是否題', '是,否', false],
    ['F10', '防火宣導', 'F07=是', '瓦斯安全', '家中是否有裝置一氧化碳警報器？', '是否題', '是,否', false],
    ['F11', '防火宣導', '', '電氣安全', '家中電器插座、插頭是否插定位並保持乾淨？', '是否題', '是,否', true],
    ['F12', '防火宣導', '', '電氣安全', '是否知道電器用品電源線不可捆綁或重壓？', '是否題', '是,否', true],
    ['F13', '防火宣導', '', '電氣安全', '家中高用電量電器電源線是否使用個別插座？', '是否題', '是,否', true],
    ['F14', '防火宣導', '', '電氣安全', '神桌燈具或魚缸馬達是否定期檢查或更換？', '是否題', '是,否', true],
    ['F15', '防火宣導', '', '避難逃生', '排煙機及排煙管油垢是否定期清洗？', '是否題', '是,否', true],
    ['F16', '防火宣導', '', '避難逃生', '樓梯處及其他逃生通道是否無堆放雜物？', '是否題', '是,否', true],
    ['F17', '防火宣導', '', '避難逃生', '家中陽台裝設之鐵窗是否有預留安全出口並保持可以開啟？', '是否題', '是,否', true],
    ['F18', '防火宣導', '', '避難逃生', '家中是否有 2 個以上不同方向逃生出口？', '是否題', '是,否', true],
    ['F19', '防火宣導', '', '避難逃生', '是否有訂定家庭逃生計畫及演練？', '是否題', '是,否', true],
    ['F20', '防火宣導', '', '消防常識', '您家中是否有消防設備？', '是否題', '是,否', true],
    ['F21', '防火宣導', 'F20=是', '消防常識', '家中擁有的消防設備 (可多選)', '複選題', '滅火器,緊急照明燈,住宅用火災警報器,其他', false],
    ['F22', '防火宣導', 'F20=是', '消防常識', '家中消防設備外觀、性能是否正常？', '是否題', '是,否', false],
    ['F23', '防火宣導', '', '改善建議', '防火改善建議事項 (可多選)', '複選題', '更換合格瓦斯鋼瓶,定期檢查瓦斯桶與皮管,簽訂瓦斯定型化契約,定期檢查熱水器與配線,定期檢查插座與延長線,養成巡視並關閉火源習慣,排煙機及排煙管定期清洗,勿在逃生通道堆放雜物,鐵窗預留出口並保持開啟,訂定家庭逃生計畫,購設消防設備,定期檢查消防設備,其他建議', false],
    ['F24', '防火宣導', '', '追蹤訪視', '需要追蹤訪視？', '是否題', '是,否', true],

    // === 防災宣導 (D01 ~ D18) ===
    ['D01', '防災宣導', '', '颱風防汛', '是否知道颱風易發生於什麼季節及帶來的主要災害？', '是否題', '是,否', true],
    ['D02', '防災宣導', '', '颱風防汛', '是否知道颱風海上、陸上颱風警報發佈時機？', '是否題', '是,否', true],
    ['D03', '防災宣導', '', '土石流防災', '是否有檢視住在土石流危險區？若是，是否知道紅黃色警戒值發佈時機？', '是否題', '是,否', true],
    ['D04', '防災宣導', '', '颱風防汛', '是否知道颱風來臨前應做的防範措施為何？', '是否題', '是,否', true],
    ['D05', '防災宣導', '', '颱風防汛', '颱風來臨時是否知道不可至海堤或河床觀潮、戲水？', '是否題', '是,否', true],
    ['D06', '防災宣導', '', '緊急避難', '是否知道緊急避難包內『緊急帶出品』應備妥什麼？', '是否題', '是,否', true],
    ['D07', '防災宣導', '', '緊急避難', '是否知道會由哪些人員來執行及協助疏散動作？', '是否題', '是,否', true],
    ['D08', '防災宣導', '', '緊急避難', '是否知道該區規劃的避難收容所及避難路線？', '是否題', '是,否', true],
    ['D09', '防災宣導', '', '緊急避難', '若發現有人員受困是否知道如何報案及報案電話？', '是否題', '是,否', true],
    ['D10', '防災宣導', '', '地震防災', '是否對地震災害有所了解？', '是否題', '是,否', true],
    ['D11', '防災宣導', '', '地震防災', '是否有針對居家環境安全進行檢視？', '是否題', '是,否', true],
    ['D12', '防災宣導', '', '地震防災', '家中大型傢俱是否有固定措施以避免倒塌？', '是否題', '是,否', true],
    ['D13', '防災宣導', '', '地震防災', '是否知道地震災害時，應優先關閉電源、火源及保持房門暢通？', '是否題', '是,否', true],
    ['D14', '防災宣導', '', '地震防災', '是否知道地震來臨時，室內或室外應如何進行應變？', '是否題', '是,否', true],
    ['D15', '防災宣導', '', '地震防災', '是否知道居家四周的空曠地點及避難疏散安全路線？', '是否題', '是,否', true],
    ['D16', '防災宣導', '', '地震防災', '是否知道地震過後，居家建築檢查應注意事項？', '是否題', '是,否', true],
    ['D17', '防災宣導', '', '災害認知', '對災害了解程度', '單選題', '了解(答否3項以下),尚可(答否4項以下),不了解(答否6項以上)', true],
    ['D18', '防災宣導', '', '改善建議', '防災改善建議事項 (可多選)', '複選題', '制定家庭避難逃生計畫,家具燈具盡可能固定,養成巡視並關閉火源習慣,勿在逃生通道堆放雜物,準備緊急避難包,其他建議', false]
  ];

  questions.forEach(function(row) { sheet.appendRow(row); });

  sheet.setColumnWidth(1, 80);
  sheet.setColumnWidth(2, 90);
  sheet.setColumnWidth(3, 90);
  sheet.setColumnWidth(4, 90);
  sheet.setColumnWidth(5, 280);
  sheet.setColumnWidth(6, 80);
  sheet.setColumnWidth(7, 220);
  sheet.setColumnWidth(8, 60);

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
  var mainHeaders = [];
  if (mainRecordSheet) {
    mainHeaders = mainRecordSheet.getRange(1, 1, 1, mainRecordSheet.getLastColumn()).getValues()[0];
  } else {
    mainHeaders = [
      '流水號', '填報時間', '訪視日期', '訪視類型', '主填寫人姓名', '協同志工',
      '所屬分隊', '案家姓名', '案家性別', '案家電話', '案家地址', 'GPS定位座標',
      '房屋屋齡', '住宅形式', '總樓層', '居住樓層', '建築結構',
      '家庭總人數', '家庭65歲以上人數', '家庭行動不便人數', '家庭6歲以下人數', '家庭外籍人士人數',
      '受訪者簽名'
    ];
    for (var f = 1; f <= 24; f++) mainHeaders.push('F' + String(f).padStart(2, '0'));
    for (var d = 1; d <= 18; d++) mainHeaders.push('D' + String(d).padStart(2, '0'));
  }

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
      recordSheet.setColumnWidth(1, 70);   // 流水號
      recordSheet.setColumnWidth(2, 160);  // 填報時間
      recordSheet.setColumnWidth(3, 110);  // 訪視日期
      recordSheet.setColumnWidth(4, 90);   // 訪視類型
      recordSheet.setColumnWidth(5, 100);  // 主填寫人
      recordSheet.setColumnWidth(6, 120);  // 協同志工
      recordSheet.setColumnWidth(7, 90);   // 分隊
      recordSheet.setColumnWidth(8, 100);  // 案家姓名
      recordSheet.setColumnWidth(9, 70);   // 案家性別
      recordSheet.setColumnWidth(10, 120); // 電話
      recordSheet.setColumnWidth(11, 200); // 地址
      recordSheet.setColumnWidth(12, 180); // GPS
      recordSheet.setColumnWidth(13, 80);  // 房屋屋齡
      recordSheet.setColumnWidth(14, 90);  // 住宅形式
      recordSheet.setColumnWidth(15, 70);  // 總樓層
      recordSheet.setColumnWidth(16, 70);  // 居住樓層
      recordSheet.setColumnWidth(17, 90);  // 建築結構
      recordSheet.setColumnWidth(18, 90);  // 家庭總人數
      recordSheet.setColumnWidth(19, 110); // 65歲以上
      recordSheet.setColumnWidth(20, 110); // 行動不便
      recordSheet.setColumnWidth(21, 110); // 6歲以下
      recordSheet.setColumnWidth(22, 110); // 外籍人士
      recordSheet.setColumnWidth(23, 160); // 受訪者簽名

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
