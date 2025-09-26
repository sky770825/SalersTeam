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

## 🎉 完成！
設置完成後，用戶填寫表單時數據會直接寫入你的Google Sheets，無需任何跳轉！
