# 志工居家訪視表單電子化系統

行動裝置優先、大按鈕點擊式的志工居家訪視電子化填報平台。

| 項目 | 技術 |
|------|------|
| 前端 | 純 HTML / CSS / Vanilla JS |
| 後端 | Google Apps Script (GAS) Web App |
| 資料庫 | Google Sheets |
| 部署 | Render Static Site (免費) |

---

## 🚀 部署步驟（依序完成）

### Step 1 — 建立 Google Sheets 主試算表

1. 前往 [Google Sheets](https://sheets.google.com) 新增一個空白試算表
2. 將試算表命名為「**志工訪視系統_主資料庫**」（或任何您喜歡的名稱）
3. 記下試算表的網址，格式為：
   ```
   https://docs.google.com/spreadsheets/d/【這段是試算表ID】/edit
   ```

---

### Step 2 — 部署 Google Apps Script (GAS) 後端

#### 2-1 建立 Apps Script 專案

1. 在試算表中點擊上方選單：**擴充功能 → Apps Script**
2. 進入 Apps Script 編輯器

#### 2-2 貼上後端程式碼

1. 將 `gas/Code.gs` 的全部內容**複製貼上**至編輯器的 `Code.gs`
2. 點擊 **「+」** 新增檔案，命名為 `InitSheet`
3. 將 `gas/InitSheet.gs` 的全部內容**複製貼上**至新建的 `InitSheet.gs`

#### 2-3 初始化試算表結構

1. 在編輯器上方下拉選單選擇函式：`initializeSheets`
2. 點擊 **▶ 執行**
3. 首次執行需授予試算表存取權限，點擊「允許」
4. 執行完成後，4 張工作表將自動建立：
   - `訪視紀錄表`
   - `人員帳號管理`（含 10 筆範例人員）
   - `分隊對照表`（3 個分隊）
   - `訪視題庫`（5 道題目）

#### 2-4 部署為 Web App

1. 點擊右上角 **「部署」→「新增部署作業」**
2. 類型選擇：**網路應用程式**
3. 設定如下：

   | 設定項 | 值 |
   |--------|-----|
   | 說明 | 志工訪視 API v1.0 |
   | 執行身分 | **我** (Me) |
   | 存取權限 | **所有人** (Anyone) |

4. 點擊 **「部署」**
5. **複製部署後的網址**（格式：`https://script.google.com/macros/s/【...】/exec`）

---

### Step 3 — 設定前端 API 網址

開啟 `frontend/app.js`，找到第 15 行：

```js
// 修改前
GAS_API_URL: 'YOUR_GAS_WEB_APP_URL_HERE',

// 修改後（填入您的 GAS Web App 網址）
GAS_API_URL: 'https://script.google.com/macros/s/AKfycbx.../exec',
```

---

### Step 4 — 上傳至 GitHub

```bash
# 在本專案資料夾中初始化 Git
git init
git add .
git commit -m "初始化：志工訪視系統"

# 建立 GitHub 新 Repo（在 github.com 操作後）
git remote add origin https://github.com/【您的帳號】/volunteer-visit-app.git
git push -u origin main
```

---

### Step 5 — 部署至 Render Static Site

1. 前往 [Render Dashboard](https://dashboard.render.com)
2. 點擊 **New → Static Site**
3. 連結您的 GitHub 儲存庫
4. 設定如下：

   | 設定項 | 值 |
   |--------|-----|
   | Name | volunteer-visit-app |
   | Root Directory | （留空） |
   | Publish directory | `frontend` |
   | Build command | （留空） |

5. 點擊 **Create Static Site**
6. 部署完成後取得網址（格式：`https://volunteer-visit-app.onrender.com`）

---

## 👥 範例帳號（初始化後可直接測試）

| 姓名 | 分隊 | 手機末三碼 | 角色 |
|------|------|-----------|------|
| 張小明 | 第一分隊 | **678** | 分隊承辦人 |
| 李美玲 | 第一分隊 | **789** | 志工 |
| 陳建宏 | 第一分隊 | **890** | 志工 |
| 王淑芬 | 第二分隊 | **901** | 分隊承辦人 |
| 林志偉 | 第二分隊 | **012** | 志工 |
| 黃雅婷 | 第二分隊 | **123** | 志工 |
| 吳家豪 | 第三分隊 | **234** | 分隊承辦人 |
| 蔡佩君 | 第三分隊 | **345** | 志工 |
| 許文哲 | 第三分隊 | **456** | 志工 |
| 系統管理員 | — | **000** | 管理員 |

> **手機末三碼**：取手機號碼最後 3 位數字，例如 `0912345678` 的末三碼是 `678`

---

## ✏️ 自訂人員與分隊

所有設定直接在 Google Sheets 修改，前端會即時連動：

### 新增人員
在 `人員帳號管理` 工作表中新增一列：
```
姓名 | 所屬分隊 | 手機號碼（10碼） | 角色
```
> 角色可填：`管理員` / `分隊承辦人` / `志工`

### 新增分隊
在 `分隊對照表` 新增一列：
```
分隊名稱 | 承辦人 Email | 分隊專屬試算表ID（可留空）
```

### 修改訪視題目
在 `訪視題庫` 工作表中增刪題目，前端下次載入即自動更新。
- 題型支援：`是否題` / `單選題` / `複選題` / `簡答題`
- 選項以**半形逗號**分隔，例如：`良好,普通,稍微疲倦`

---

## 🔒 分隊資料隔離（進階設定）

若需要各分隊只能看到自己的資料：

1. 為各分隊建立獨立的 Google Sheets
2. 將試算表 ID 填入 `分隊對照表` 的「分隊專屬試算表ID」欄位
3. 在該試算表的「共用」設定中，僅授予分隊承辦人的 Google 帳號權限

志工提交的每筆紀錄，GAS 將自動同步複製至對應分隊的試算表中。

---

## 📁 專案結構

```
志工訪視電子化/
├── frontend/
│   ├── index.html      # 登入頁
│   ├── form.html       # 訪視表單頁（動態題目）
│   ├── preview.html    # 預覽確認頁
│   ├── style.css       # 共用樣式
│   └── app.js          # 共用 JS（API / Session / GPS）
├── gas/
│   ├── Code.gs         # GAS 主後端 API
│   └── InitSheet.gs    # 試算表初始化腳本
├── render.yaml         # Render 部署設定
└── README.md           # 本文件
```

---

## ❓ 常見問題

**Q：送出表單後沒有資料？**
確認 `frontend/app.js` 的 `GAS_API_URL` 已填入正確網址，且 GAS 部署存取權限設定為「所有人」。

**Q：GAS 修改後前端仍顯示舊資料？**
GAS 有快取機制，每次更新程式碼後需重新部署（部署 → 管理部署作業 → 編輯 → 版本：新版本）。

**Q：手機定位失敗？**
iOS 需使用 Safari 或 Chrome，且網頁必須為 HTTPS（Render 已自動提供）。
