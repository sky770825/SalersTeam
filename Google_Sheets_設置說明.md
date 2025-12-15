# 📋 蔣哥行銷網頁 - Google Sheets 設置說明

## 🎯 目標
將網頁表單數據直接提交到你的Google Sheets中，無需跳轉到Google表單。

## 📊 你的Google Sheets
- **Sheet ID**: `1X8l3vEAecBEldAVoRB_iezN7szWLPgnf4ZovvqX2IIU`
- **連結**: https://docs.google.com/spreadsheets/d/1X8l3vEAecBEldAVoRB_iezN7szWLPgnf4ZovvqX2IIU/edit?usp=sharing

## 🔧 設置步驟

### 第一步：創建Google Apps Script
1. 前往 [Google Apps Script](https://script.google.com/)
2. 點擊「新增專案」
3. 將 `Google_Apps_Script_代碼.js` 文件中的代碼複製貼入
4. 點擊「儲存」按鈕

### 第二步：設置權限
1. 第一次運行時，系統會要求授權
2. 點擊「檢閱權限」
3. 選擇你的Google帳戶
4. 點擊「進階」→「前往 [專案名稱] (不安全)」
5. 點擊「允許」

### 第三步：部署為Web應用程式
1. 點擊「部署」→「新增部署作業」
2. 選擇「網頁應用程式」
3. 設置：
   - **執行身分**: 我
   - **存取權**: 任何人
4. 點擊「部署」
5. **重要**: 複製生成的Web App URL

### 第四步：更新網頁代碼
1. 打開 `蔣哥行銷網頁.html` 文件
2. 找到這一行：
   ```javascript
   const scriptUrl = 'https://script.google.com/macros/s/AKfycbwYOUR_SCRIPT_ID_HERE/exec';
   ```
3. 將 `AKfycbwYOUR_SCRIPT_ID_HERE` 替換為你從第三步複製的URL中的Script ID

## 📝 表單欄位對應

| 網頁欄位 | Google Sheets欄位 | 說明 |
|---------|------------------|------|
| 姓名 | A列 | 用戶姓名 |
| 性別 | B列 | 男性/女性 |
| 電話 | C列 | 手機號碼 |
| 居住地 | D列 | 台北/新北/桃園/新竹/其他 |
| LINE ID | E列 | LINE帳號 |
| 資訊需求 | F列 | 複選項目 |
| 參加動機 | G列 | 文字描述 |
| 同意條款 | H列 | 是/否 |
| 確認活動 | I列 | 是/否 |
| 時間戳記 | J列 | 提交時間 |
| 瀏覽器 | K列 | 用戶代理 |
| 來源 | L列 | 推薦來源 |

## 🧪 測試功能
1. 在Google Apps Script編輯器中
2. 選擇 `testFunction` 函數
3. 點擊「執行」按鈕
4. 檢查你的Google Sheets是否出現測試數據

## 🔍 故障排除

### 常見問題：
1. **權限錯誤**: 確保已正確授權Google Apps Script
2. **Sheet ID錯誤**: 確認Sheet ID正確無誤
3. **部署失敗**: 檢查部署設置是否正確
4. **數據未出現**: 檢查瀏覽器控制台是否有錯誤訊息

### 檢查方法：
1. 打開瀏覽器開發者工具 (F12)
2. 查看Console標籤是否有錯誤訊息
3. 查看Network標籤確認請求是否成功發送

## 📈 數據管理
- 數據會自動按時間順序排列
- 標題行會自動格式化（藍色背景，白色文字）
- 可選：使用 `cleanupOldData` 函數清理30天前的舊數據

## ⚡ 快速複製一套「新案場」表單串接流程（北屯 / 台南模式）

> 適用情境：每個案場一個獨立 landing page + 一份獨立 Google Sheet（例如：北屯一份、台南一份），流程都一樣，只是換 Sheet ID 而已。

### 1️⃣ 準備新的 Google Sheet
- 建一份新的試算表（或複製舊的當模板也可以）。
- 複製網址中間的那串就是 **Sheet ID**，例如：  
  - `https://docs.google.com/spreadsheets/d/【這一串就是 Sheet ID】/edit`

### 2️⃣ 建立或複製對應的 Apps Script
- 開啟這份新試算表 → 點「擴充功能 → Apps Script」。
- 如果是 **北屯那套** 模式：
  - 使用 `GoogleAppsScript.js`（主要欄位：姓名 / Line / 手機 / 地區 + 7 題問題）。
- 如果是 **台南那套** 模式：
  - 使用 `GoogleAppsScript_Tainan.js`（主要欄位：姓名 / 手機 / 聯絡方式 / 預約時段 / 複選問題 / 備註等）。
- 在 Apps Script 編輯器中：
  - 刪掉原本程式（若有），整段貼上對應的 JS 檔內容。
  - 把檔案裡 `SHEET_ID` 改成你這份新試算表的 Sheet ID（`const SHEET_ID = '你的新ID';`）。
  - 如有需要，`SHEET_NAME` 也可從 `工作表1` 改成你自己命名的分頁。

### 3️⃣ 部署成 Web App（一次完成授權 + URL）
- 在 Apps Script 中點「部署 → 新增部署作業 → 網頁應用程式」。
- 選項：
  - **執行身分**：我
  - **存取權**：任何人
- 部署後複製產生的 Web App URL（長得像 `https://script.google.com/macros/s/xxxx/exec`）。

### 4️⃣ 在對應的 HTML 場景頁塞入 API_ENDPOINT
- 打開該案場的 HTML（例如：
  - 北屯：`北屯預售屋報名表單.html`
  - 台南：`台南新都心段」發展副都核心！.html`
- 找到前端腳本裡的設定：
  - 北屯範例：
    ```js
    const SCRIPT_URL = '...'; // 或類似命名
    ```
  - 台南範例：
    ```js
    const API_ENDPOINT = ""; // 留空 = 先用前端示範
    ```
- 把這個變數改成你剛剛部署出的 Web App URL，即完成串接。

### 5️⃣ 測試一筆，確認有寫進「正確那份」 Sheet
- 在瀏覽器開該案場的 HTML 頁面，實際填一筆測試資料送出。
- 回到你對應的 Google Sheet：
  - 檢查是否多出一列資料。
  - 欄位順序、時間戳記是否正確。
- （可選）在 Apps Script 裡執行內建測試函式（例如 `test_TainanFormWrite()`）先寫一筆測試行。

> 之後如果再新增別的案場，只要「**新 Sheet → 複製一份對應的 Apps Script，改 SHEET_ID → 部署 → 把 Web App URL 塞進新的 HTML**」這幾步，就可以很快複製一整套。

## 🎉 完成！
設置完成後，用戶填寫表單時數據會直接寫入你的Google Sheets，無需任何跳轉！
