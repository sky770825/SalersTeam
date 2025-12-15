/**
 * Google Apps Script - 「台南新都心段」表單寫入專用
 *
 * 使用步驟：
 * 1. 打開這份試算表：[台南新都心段表單對接表](https://docs.google.com/spreadsheets/d/1dyDlkPC6mIYzS6bisw28h3iAhCWlrHpNQpsO80fcN2Y/edit?usp=sharing)
 * 2. 在試算表中點「擴充功能」→「Apps Script」
 * 3. 刪掉原本範例程式（如果有），整份貼上本檔內容
 * 4. 儲存，然後點「部署」→「新增部署作業」→ 選「網頁應用程式」
 * 5. 執行身分選「我」，存取權限選「所有人」
 * 6. 部署後複製 Web App URL，貼回 `台南新都心段」發展副都核心！.html` 裡的 `API_ENDPOINT`
 *
 * 👉 前端送出的 JSON 結構（來自「台南新都心段」HTML）：
 * {
 *   name: string,
 *   phone: string,
 *   contact: 'phone' | 'line' | 'either',
 *   time: string,
 *   questions: string[], // 最多 2 題
 *   need: string,
 *   source: string,
 *   ts: string (ISO 時間字串)
 * }
 */

// ✅ 換成「台南新都心段」這份試算表的 ID
// 對應網址：https://docs.google.com/spreadsheets/d/1dyDlkPC6mIYzS6bisw28h3iAhCWlrHpNQpsO80fcN2Y/edit
const SHEET_ID = '1dyDlkPC6mIYzS6bisw28h3iAhCWlrHpNQpsO80fcN2Y';
const SHEET_NAME = '工作表1'; // 如有改名，這裡一起改

/**
 * 初始化工作表欄位：只會在第一次建立或沒有標題時執行
 */
function initializeSheet_Tainan(sheet) {
  Logger.log('開始初始化「台南新都心段」工作表...');

  // 如果第一格已經有「時間戳記」，代表初始化過，直接跳過
  try {
    const firstCell = sheet.getRange(1, 1).getValue();
    if (firstCell === '時間戳記') {
      Logger.log('工作表已初始化過，略過初始化。');
      return;
    }
  } catch (e) {
    Logger.log('檢查標題時發生錯誤（可能是空表），繼續初始化：' + e.toString());
  }

  // 欄位標題設計：對應前端送來的欄位
  const headers = [
    '時間戳記',              // A
    '姓名',                  // B
    '手機',                  // C
    '偏好聯絡方式',          // D
    '預計賞屋時間',          // E
    '問題1',                 // F （複選的第 1 題）
    '問題2',                 // G （複選的第 2 題）
    '問題清單（合併）',      // H （所有問題用「 / 」串起來）
    '需求備註',              // I
    '來源（source）',        // J
    '前端時間戳記 ts'        // K
  ];

  // 清空整張表，然後重新寫入標題
  sheet.clear();
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  // 標題列樣式
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange
    .setBackground('#ffd93d')
    .setFontColor('#000000')
    .setFontWeight('bold')
    .setFontSize(11)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setWrap(true);

  sheet.setRowHeight(1, 42);
  sheet.setFrozenRows(1);

  const columnWidths = [150, 120, 120, 120, 140, 220, 220, 260, 260, 160, 200];
  columnWidths.forEach((w, i) => sheet.setColumnWidth(i + 1, w));

  sheet.getRange(1, 1).setNumberFormat('yyyy/mm/dd hh:mm:ss');

  Logger.log('✅ 初始化完成：「台南新都心段」標題與格式已設定');
}

/**
 * GET：簡單回傳服務狀態
 */
function doGet(e) {
  const payload = {
    success: true,
    message: '台南新都心段表單後端已就緒，請使用 POST 提交資料。'
  };

  return HtmlService.createHtmlOutput(JSON.stringify(payload))
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * POST：處理前端送來的表單資料（支援 JSON 與表單參數兩種模式）
 */
function doPost(e) {
  try {
    Logger.log('=== 收到「台南新都心段」表單 POST ===');
    Logger.log('raw parameter: ' + JSON.stringify(e.parameter));
    if (e.postData && e.postData.contents) {
      Logger.log('raw postData: ' + e.postData.contents);
    }

    let data;

    // 1) 優先嘗試 JSON body（對應前端 fetch + JSON.stringify）
    if (e.postData && e.postData.contents) {
      try {
        data = JSON.parse(e.postData.contents);
        Logger.log('解析 JSON 成功：' + JSON.stringify(data));
      } catch (err) {
        Logger.log('JSON 解析失敗，改用表單參數模式：' + err.toString());
      }
    }

    // 2) 如果 JSON 沒成功，就用表單參數
    if (!data && e.parameter) {
      const p = e.parameter;
      data = {
        name: p.name || '',
        phone: p.phone || '',
        contact: p.contact || '',
        time: p.time || '',
        // 若是以 q1 / q2 帶進來也可以自行補上
        questions: [
          p.q1 || '',
          p.q2 || ''
        ].filter(function (x) { return x; }),
        need: p.need || '',
        source: p.source || '',
        ts: p.ts || ''
      };
      Logger.log('使用表單參數模式解析後資料：' + JSON.stringify(data));
    }

    if (!data) {
      throw new Error('沒有收到可解析的資料');
    }

    // 開啟試算表與工作表
    const spreadsheet = SpreadsheetApp.openById(SHEET_ID);
    let sheet = spreadsheet.getSheetByName(SHEET_NAME);
    if (!sheet) {
      sheet = spreadsheet.insertSheet(SHEET_NAME);
    }

    // 確保工作表已初始化
    const firstCell = sheet.getRange(1, 1).getValue();
    if (firstCell !== '時間戳記') {
      initializeSheet_Tainan(sheet);
    }

    // questions 處理：最多 2 題
    var qs = data.questions || [];
    if (Object.prototype.toString.call(qs) !== '[object Array]') {
      qs = [String(qs)];
    }
    const q1 = qs[0] || '';
    const q2 = qs[1] || '';
    const qAll = qs.length ? qs.join(' / ') : '';

    // 準備要寫入的一列
    const row = [
      new Date(),        // 時間戳記
      data.name || '',
      data.phone || '',
      data.contact || '',
      data.time || '',
      q1,
      q2,
      qAll,
      data.need || '',
      data.source || '',
      data.ts || ''
    ];

    const newRow = sheet.getLastRow() + 1;
    sheet.getRange(newRow, 1, 1, row.length).setValues([row]);
    SpreadsheetApp.flush();

    // 基本排版
    const newRange = sheet.getRange(newRow, 1, 1, row.length);
    newRange.setVerticalAlignment('middle');
    sheet.getRange(newRow, 1).setNumberFormat('yyyy/mm/dd hh:mm:ss');

    const totalRows = sheet.getLastRow();

    // 回傳一個簡單的 HTML，讓前端只要看 status / ok 即可
    const responseHtml = [
      '<!doctype html>',
      '<html><head><meta charset="UTF-8"><title>提交成功</title></head>',
      '<body>',
      '<p>提交成功，已寫入第 ' + newRow + ' 行（目前共 ' + totalRows + ' 行）。</p>',
      '</body></html>'
    ].join('');

    return HtmlService.createHtmlOutput(responseHtml)
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);

  } catch (error) {
    Logger.log('❌ 錯誤：' + error.toString());
    Logger.log(error.stack);

    const errorHtml = [
      '<!doctype html>',
      '<html><head><meta charset="UTF-8"><title>提交失敗</title></head>',
      '<body>',
      '<p>提交失敗：' + error.toString() + '</p>',
      '</body></html>'
    ].join('');

    return HtmlService.createHtmlOutput(errorHtml)
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
}

/**
 * 測試：在 Apps Script 介面中直接執行，確認能順利寫入一列假資料
 */
function test_TainanFormWrite() {
  const mockEvent = {
    postData: {
      contents: JSON.stringify({
        name: '測試_台南_' + new Date().getTime(),
        phone: '0912345678',
        contact: 'line',
        time: '平日晚上',
        questions: [
          '1｜想要了解更多台南新都心段重大發展',
          '4｜40萬不到能購買景觀公園戶嗎？'
        ],
        need: '這是測試資料，請無視。',
        source: '研森｜RWD報名頁（測試）',
        ts: new Date().toISOString()
      })
    }
  };

  const res = doPost(mockEvent);
  Logger.log('測試結果 HTML：' + res.getContent().substring(0, 300));
}


