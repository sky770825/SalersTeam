/**
 * 蔣哥行銷網頁表單提交處理器 - 修復版
 * 用於接收網頁表單數據並寫入Google Sheets
 */

// 處理 GET 請求（用於測試）
function doGet(e) {
  try {
    return ContentService
      .createTextOutput(JSON.stringify({
        success: true,
        message: 'Google Apps Script 運行正常',
        timestamp: new Date().toISOString()
      }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    return ContentService
      .createTextOutput(JSON.stringify({
        success: false,
        error: error.toString()
      }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doPost(e) {
  try {
    // 調試：檢查接收到的參數
    console.log('接收到的完整參數:', e);
    
    // 檢查參數是否存在
    if (!e) {
      console.error('錯誤：沒有接收到任何參數');
      return ContentService
        .createTextOutput(JSON.stringify({
          success: false, 
          error: '沒有接收到任何參數',
          message: '表單數據傳送失敗'
        }))
        .setMimeType(ContentService.MimeType.JSON);
    }
    
    if (!e.parameter) {
      console.error('錯誤：參數中沒有parameter屬性');
      return ContentService
        .createTextOutput(JSON.stringify({
          success: false, 
          error: '參數格式錯誤',
          message: '表單數據格式不正確'
        }))
        .setMimeType(ContentService.MimeType.JSON);
    }
    
    // 獲取表單數據
    const name = e.parameter.name || '';
    const gender = e.parameter.gender || '';
    const phone = e.parameter.phone || '';
    const location = e.parameter.location || '';
    const email = e.parameter.email || '';
    const lineId = e.parameter.lineId || '';
    const referrerType = e.parameter.referrerType || '';
    const referrerName = e.parameter.referrerName || '';
    const infoNeeds = e.parameter.infoNeeds || '';
    const motivation = e.parameter.motivation || '';
    const agreeTerms = e.parameter.agreeTerms || '';
    const agreeEvent = e.parameter.agreeEvent || '';
    const timestamp = e.parameter.timestamp || '';
    const userAgent = e.parameter.userAgent || '';
    const referrer = e.parameter.referrer || '';

    console.log('接收到的數據:', {
      name, gender, phone, location, email, lineId, 
      referrerType, referrerName, infoNeeds, motivation
    });

    // 打開Google Sheets - 明確指定工作表1
    const spreadsheet = SpreadsheetApp.openById('1rxuODXlZpQ5PZ8Gm4lq5gxLimhcpaBHtZ2Q2z_xScz0');
    const sheet = spreadsheet.getSheetByName('工作表1') || spreadsheet.getSheets()[0];
    
    // 如果是第一次運行，添加標題行
    if (sheet.getLastRow() === 0) {
      sheet.appendRow([
        '時間戳記', '姓名', '性別', '電話', '居住地', 'Email', 'LINE ID', 
        '介紹人類型', '介紹人姓名', '資訊需求', '參加動機', '同意條款', '確認活動', '瀏覽器', '來源'
      ]);
      
      // 設置標題行格式
      const headerRange = sheet.getRange(1, 1, 1, 15);
      headerRange.setBackground('#4285F4');
      headerRange.setFontColor('white');
      headerRange.setFontWeight('bold');
      headerRange.setHorizontalAlignment('center');
    }
    
    // 添加數據行
    const newRow = [
      timestamp, 
      name, 
      gender, 
      phone, 
      location, 
      email,
      lineId,
      referrerType,
      referrerName,
      infoNeeds, 
      motivation, 
      agreeTerms, 
      agreeEvent, 
      userAgent, 
      referrer
    ];
    
    sheet.appendRow(newRow);
    
    // 設置新行的格式
    const lastRow = sheet.getLastRow();
    const dataRange = sheet.getRange(lastRow, 1, 1, 15);
    dataRange.setBorder(true, true, true, true, true, true);
    
    // 自動調整列寬
    sheet.autoResizeColumns(1, 15);
    
    // 記錄成功日誌
    console.log('成功提交表單數據:', {
      name: name,
      phone: phone,
      email: email,
      timestamp: timestamp
    });
    
    // 發送確認信件
    try {
      sendConfirmationEmail(name, phone, email, lineId, referrerType);
      console.log('確認信件已發送');
    } catch (emailError) {
      console.error('發送確認信件失敗:', emailError);
    }
    
    return ContentService
      .createTextOutput(JSON.stringify({
        success: true, 
        message: '數據已成功提交到Google Sheets，確認信件已發送',
        row: lastRow
      }))
      .setMimeType(ContentService.MimeType.JSON);
      
  } catch (error) {
    // 記錄錯誤日誌
    console.error('提交表單數據時發生錯誤:', error);
    
    return ContentService
      .createTextOutput(JSON.stringify({
        success: false, 
        error: error.toString(),
        message: '提交失敗，請稍後再試'
      }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * 發送確認信件函數
 */
function sendConfirmationEmail(name, phone, email, lineId, referrerType) {
  try {
    // 信件內容
    const subject = '🎉 蔣哥房產分析說明會 - 報名確認信';
    const body = `
親愛的 ${name} 您好，

感謝您報名參加「蔣哥房產分析說明會」！

1、活動資訊：
• 時間：2025年10月5日（日） 下午2:00 準時開始
• 地點：桃園市桃園區大興西路三段90號（銷售中心）

2、現場抽獎：
大金空調清淨機（市價2萬元）

3.注意事項：
1. 請提前15分鐘到達會場
2. 請攜帶身分證件
3. 現場提供茶水點心
4. 活動全程免費，無推銷壓力

4.如有任何問題，請聯繫：
介紹人：${referrerType || '其他'}

感謝您的報名，期待與您相見！

BNI華地產房產行銷組團隊
${new Date().toLocaleDateString('zh-TW')}
    `;
    
    // 如果有Email地址，發送Email
    if (email && email.trim() !== '') {
      console.log('準備發送確認信件給:', {
        name: name,
        email: email,
        phone: phone,
        lineId: lineId,
        subject: subject
      });
      
      // 實際發送Email
      try {
        GmailApp.sendEmail(email, subject, body);
        console.log('✅ Gmail發送成功');
      } catch (gmailError) {
        console.error('❌ Gmail發送失敗:', gmailError);
        // 備用方案：使用MailApp
        try {
          MailApp.sendEmail(email, subject, body);
          console.log('✅ MailApp發送成功（備用方案）');
        } catch (mailError) {
          console.error('❌ MailApp也發送失敗:', mailError);
          throw mailError;
        }
      }
      
      // 記錄到Google Sheets的備註欄位
      const spreadsheet = SpreadsheetApp.openById('1rxuODXlZpQ5PZ8Gm4lq5gxLimhcpaBHtZ2Q2z_xScz0');
      const sheet = spreadsheet.getSheetByName('工作表1') || spreadsheet.getSheets()[0];
      const lastRow = sheet.getLastRow();
      sheet.getRange(lastRow, 16).setValue('確認信件已發送'); // 在P列添加備註
      
      console.log('確認信件已成功發送到:', email);
    } else {
      console.log('沒有Email地址，無法發送確認信件');
      
      // 記錄到Google Sheets的備註欄位
      const spreadsheet = SpreadsheetApp.openById('1rxuODXlZpQ5PZ8Gm4lq5gxLimhcpaBHtZ2Q2z_xScz0');
      const sheet = spreadsheet.getSheetByName('工作表1') || spreadsheet.getSheets()[0];
      const lastRow = sheet.getLastRow();
      sheet.getRange(lastRow, 16).setValue('無Email地址，未發送確認信'); // 在P列添加備註
    }
    
  } catch (error) {
    console.error('發送確認信件時發生錯誤:', error);
    throw error;
  }
}

/**
 * 測試函數 - 用於驗證設置是否正確
 */
function testFunction() {
  const testData = {
    parameter: {
      name: '測試用戶',
      gender: 'male',
      phone: '0912-345-678',
      location: 'taipei',
      email: 'test@example.com',
      lineId: 'test123',
      referrerType: '朋友介紹',
      referrerName: '張三',
      infoNeeds: '購屋與房地產趨勢分析, 桃園房市投資價值剖析：桃園房屋真的值得買嗎？',
      motivation: '想要了解房產投資',
      agreeTerms: '是',
      agreeEvent: '是',
      timestamp: new Date().toISOString(),
      userAgent: 'Test Browser',
      referrer: 'https://test.com'
    }
  };
  
  console.log('🧪 開始測試Google Apps Script...');
  console.log('🧪 測試數據:', testData);
  
  const result = doPost(testData);
  console.log('🧪 測試結果:', result.getContent());
  return result;
}

/**
 * 檢查工作表名稱的函數
 */
function checkSheetNames() {
  try {
    console.log('🔍 檢查Google Sheets中的工作表名稱...');
    const spreadsheet = SpreadsheetApp.openById('1rxuODXlZpQ5PZ8Gm4lq5gxLimhcpaBHtZ2Q2z_xScz0');
    const sheets = spreadsheet.getSheets();
    
    console.log('📋 找到的工作表:');
    sheets.forEach((sheet, index) => {
      console.log(`  ${index + 1}. "${sheet.getName()}"`);
    });
    
    // 檢查是否有「工作表1」
    const targetSheet = spreadsheet.getSheetByName('工作表1');
    if (targetSheet) {
      console.log('✅ 找到「工作表1」工作表');
      console.log('📊 工作表1的行數:', targetSheet.getLastRow());
      console.log('📊 工作表1的列數:', targetSheet.getLastColumn());
    } else {
      console.log('❌ 沒有找到「工作表1」工作表');
      console.log('💡 將使用第一個工作表:', sheets[0].getName());
    }
    
    return sheets.map(sheet => sheet.getName());
  } catch (error) {
    console.error('❌ 檢查工作表名稱失敗:', error);
    return '檢查失敗: ' + error.toString();
  }
}

/**
 * 新增測試數據到指定工作表
 */
function addTestDataToSheet() {
  try {
    console.log('🧪 開始新增測試數據...');
    const spreadsheet = SpreadsheetApp.openById('1rxuODXlZpQ5PZ8Gm4lq5gxLimhcpaBHtZ2Q2z_xScz0');
    
    // 嘗試不同的工作表名稱
    const possibleSheetNames = ['工作表1', '表_1', 'Sheet1', '工作表', '表1'];
    let targetSheet = null;
    let usedSheetName = '';
    
    for (const sheetName of possibleSheetNames) {
      targetSheet = spreadsheet.getSheetByName(sheetName);
      if (targetSheet) {
        usedSheetName = sheetName;
        console.log(`✅ 找到工作表: "${sheetName}"`);
        break;
      }
    }
    
    // 如果都沒找到，使用第一個工作表
    if (!targetSheet) {
      targetSheet = spreadsheet.getSheets()[0];
      usedSheetName = targetSheet.getName();
      console.log(`⚠️ 使用第一個工作表: "${usedSheetName}"`);
    }
    
    // 檢查是否需要添加標題行
    if (targetSheet.getLastRow() === 0) {
      console.log('📝 添加標題行...');
      targetSheet.appendRow([
        '時間戳記', '姓名', '性別', '電話', '居住地', 'Email', 'LINE ID', 
        '介紹人類型', '介紹人姓名', '資訊需求', '參加動機', '同意條款', '確認活動', '瀏覽器', '來源'
      ]);
      
      // 設置標題行格式
      const headerRange = targetSheet.getRange(1, 1, 1, 15);
      headerRange.setBackground('#4285F4');
      headerRange.setFontColor('white');
      headerRange.setFontWeight('bold');
      headerRange.setHorizontalAlignment('center');
    }
    
    // 添加測試數據
    const testData = [
      new Date().toISOString(),
      '測試用戶_' + new Date().getTime(),
      '男性',
      '0912-345-678',
      '台北',
      'test@example.com',
      'test123',
      '朋友介紹',
      '張三',
      '購屋與房地產趨勢分析, 桃園房市投資價值剖析',
      '想要了解房產投資',
      '是',
      '是',
      'Test Browser',
      'https://test.com'
    ];
    
    console.log('📝 添加測試數據行...');
    targetSheet.appendRow(testData);
    
    // 設置新行的格式
    const lastRow = targetSheet.getLastRow();
    const dataRange = targetSheet.getRange(lastRow, 1, 1, 15);
    dataRange.setBorder(true, true, true, true, true, true);
    
    // 自動調整列寬
    targetSheet.autoResizeColumns(1, 15);
    
    console.log(`✅ 測試數據已成功添加到工作表 "${usedSheetName}" 的第 ${lastRow} 行`);
    console.log('📊 當前工作表總行數:', targetSheet.getLastRow());
    console.log('📊 當前工作表總列數:', targetSheet.getLastColumn());
    
    return {
      success: true,
      sheetName: usedSheetName,
      rowAdded: lastRow,
      totalRows: targetSheet.getLastRow(),
      totalColumns: targetSheet.getLastColumn()
    };
    
  } catch (error) {
    console.error('❌ 新增測試數據失敗:', error);
    return {
      success: false,
      error: error.toString()
    };
  }
}

/**
 * 手動新增一筆資料到Google Sheets
 */
function addManualData() {
  try {
    console.log('📝 開始手動新增資料...');
    const spreadsheet = SpreadsheetApp.openById('1rxuODXlZpQ5PZ8Gm4lq5gxLimhcpaBHtZ2Q2z_xScz0');
    
    // 嘗試不同的工作表名稱
    const possibleSheetNames = ['工作表1', '表_1', 'Sheet1', '工作表', '表1'];
    let targetSheet = null;
    let usedSheetName = '';
    
    for (const sheetName of possibleSheetNames) {
      targetSheet = spreadsheet.getSheetByName(sheetName);
      if (targetSheet) {
        usedSheetName = sheetName;
        console.log(`✅ 找到工作表: "${sheetName}"`);
        break;
      }
    }
    
    // 如果都沒找到，使用第一個工作表
    if (!targetSheet) {
      targetSheet = spreadsheet.getSheets()[0];
      usedSheetName = targetSheet.getName();
      console.log(`⚠️ 使用第一個工作表: "${usedSheetName}"`);
    }
    
    // 檢查是否需要添加標題行
    if (targetSheet.getLastRow() === 0) {
      console.log('📝 添加標題行...');
      targetSheet.appendRow([
        '時間戳記', '姓名', '性別', '電話', '居住地', 'Email', 'LINE ID', 
        '介紹人類型', '介紹人姓名', '資訊需求', '參加動機', '同意條款', '確認活動', '瀏覽器', '來源'
      ]);
      
      // 設置標題行格式
      const headerRange = targetSheet.getRange(1, 1, 1, 15);
      headerRange.setBackground('#4285F4');
      headerRange.setFontColor('white');
      headerRange.setFontWeight('bold');
      headerRange.setHorizontalAlignment('center');
    }
    
    // 手動新增的資料
    const manualData = [
      new Date().toISOString(), // 時間戳記
      '蔣哥測試用戶', // 姓名
      '男性', // 性別
      '0987-654-321', // 電話
      '桃園', // 居住地
      'jiang.test@gmail.com', // Email
      'jiang_test_2024', // LINE ID
      '網路搜尋', // 介紹人類型
      '', // 介紹人姓名
      '購屋與房地產趨勢分析, 桃園房市投資價值剖析：桃園房屋真的值得買嗎？, 購屋需求解析：自住、投資、首購或首換，該如何規劃？', // 資訊需求
      '想要了解桃園房市投資機會', // 參加動機
      '是', // 同意條款
      '是', // 確認活動
      'Google Apps Script 手動新增', // 瀏覽器
      '手動測試' // 來源
    ];
    
    console.log('📝 添加手動資料...');
    targetSheet.appendRow(manualData);
    
    // 設置新行的格式
    const lastRow = targetSheet.getLastRow();
    const dataRange = targetSheet.getRange(lastRow, 1, 1, 15);
    dataRange.setBorder(true, true, true, true, true, true);
    
    // 自動調整列寬
    targetSheet.autoResizeColumns(1, 15);
    
    console.log(`✅ 手動資料已成功添加到工作表 "${usedSheetName}" 的第 ${lastRow} 行`);
    console.log('📊 新增的資料:');
    console.log('  姓名: 蔣哥測試用戶');
    console.log('  電話: 0987-654-321');
    console.log('  Email: jiang.test@gmail.com');
    console.log('  居住地: 桃園');
    console.log('  來源: 手動測試');
    
    return {
      success: true,
      message: `資料已成功添加到工作表 "${usedSheetName}" 的第 ${lastRow} 行`,
      sheetName: usedSheetName,
      rowAdded: lastRow,
      data: {
        name: '蔣哥測試用戶',
        phone: '0987-654-321',
        email: 'jiang.test@gmail.com',
        location: '桃園'
      }
    };
    
  } catch (error) {
    console.error('❌ 手動新增資料失敗:', error);
    return {
      success: false,
      error: error.toString()
    };
  }
}

/**
 * 詳細診斷函數 - 檢查為什麼沒有看到資料
 */
function diagnoseSheetIssue() {
  try {
    console.log('🔍 開始診斷 Google Sheets 問題...');
    const spreadsheet = SpreadsheetApp.openById('1rxuODXlZpQ5PZ8Gm4lq5gxLimhcpaBHtZ2Q2z_xScz0');
    
    console.log('📊 試算表基本資訊:');
    console.log('  試算表名稱:', spreadsheet.getName());
    console.log('  試算表 ID:', spreadsheet.getId());
    console.log('  試算表 URL:', spreadsheet.getUrl());
    
    // 檢查所有工作表
    const sheets = spreadsheet.getSheets();
    console.log('📋 所有工作表列表:');
    sheets.forEach((sheet, index) => {
      console.log(`  ${index + 1}. 工作表名稱: "${sheet.getName()}"`);
      console.log(`     行數: ${sheet.getLastRow()}`);
      console.log(`     列數: ${sheet.getLastColumn()}`);
      console.log(`     是否為活躍工作表: ${sheet === spreadsheet.getActiveSheet()}`);
      
      // 檢查前幾行的內容
      if (sheet.getLastRow() > 0) {
        console.log(`     前3行內容:`);
        for (let row = 1; row <= Math.min(3, sheet.getLastRow()); row++) {
          const rowData = sheet.getRange(row, 1, 1, Math.min(5, sheet.getLastColumn())).getValues()[0];
          console.log(`       第${row}行:`, rowData);
        }
      }
      console.log('     ---');
    });
    
    // 嘗試找到目標工作表
    const possibleSheetNames = ['工作表1', '表_1', 'Sheet1', '工作表', '表1'];
    let targetSheet = null;
    let usedSheetName = '';
    
    console.log('🎯 尋找目標工作表:');
    for (const sheetName of possibleSheetNames) {
      const foundSheet = spreadsheet.getSheetByName(sheetName);
      if (foundSheet) {
        targetSheet = foundSheet;
        usedSheetName = sheetName;
        console.log(`✅ 找到工作表: "${sheetName}"`);
        console.log(`   行數: ${foundSheet.getLastRow()}`);
        console.log(`   列數: ${foundSheet.getLastColumn()}`);
        break;
      } else {
        console.log(`❌ 未找到工作表: "${sheetName}"`);
      }
    }
    
    // 如果都沒找到，使用第一個工作表
    if (!targetSheet) {
      targetSheet = sheets[0];
      usedSheetName = targetSheet.getName();
      console.log(`⚠️ 使用第一個工作表: "${usedSheetName}"`);
    }
    
    // 檢查目標工作表的詳細內容
    console.log('📝 目標工作表詳細內容:');
    console.log(`  工作表名稱: "${usedSheetName}"`);
    console.log(`  總行數: ${targetSheet.getLastRow()}`);
    console.log(`  總列數: ${targetSheet.getLastColumn()}`);
    
    if (targetSheet.getLastRow() > 0) {
      console.log('  所有行內容:');
      for (let row = 1; row <= targetSheet.getLastRow(); row++) {
        const rowData = targetSheet.getRange(row, 1, 1, targetSheet.getLastColumn()).getValues()[0];
        console.log(`    第${row}行:`, rowData);
      }
    } else {
      console.log('  ⚠️ 工作表是空的！');
    }
    
    // 嘗試新增一筆測試資料
    console.log('🧪 嘗試新增測試資料...');
    const testData = [
      new Date().toISOString(),
      '診斷測試用戶',
      '男性',
      '0911-222-333',
      '台北',
      'diagnose@test.com',
      'diagnose_test',
      '手動測試',
      '',
      '診斷測試資訊需求',
      '診斷測試動機',
      '是',
      '是',
      '診斷測試瀏覽器',
      '診斷測試來源'
    ];
    
    targetSheet.appendRow(testData);
    const newRow = targetSheet.getLastRow();
    console.log(`✅ 測試資料已添加到第 ${newRow} 行`);
    
    // 再次檢查內容
    console.log('📝 新增後的內容:');
    const newRowData = targetSheet.getRange(newRow, 1, 1, targetSheet.getLastColumn()).getValues()[0];
    console.log(`  第${newRow}行:`, newRowData);
    
    return {
      success: true,
      spreadsheetName: spreadsheet.getName(),
      spreadsheetUrl: spreadsheet.getUrl(),
      sheets: sheets.map(sheet => ({
        name: sheet.getName(),
        rows: sheet.getLastRow(),
        columns: sheet.getLastColumn(),
        isActive: sheet === spreadsheet.getActiveSheet()
      })),
      targetSheet: {
        name: usedSheetName,
        rows: targetSheet.getLastRow(),
        columns: targetSheet.getLastColumn()
      },
      testDataAdded: true,
      testRow: newRow
    };
    
  } catch (error) {
    console.error('❌ 診斷過程中發生錯誤:', error);
    return {
      success: false,
      error: error.toString(),
      stack: error.stack
    };
  }
}

/**
 * 檢查工作表1的最新資料
 */
function checkLatestData() {
  try {
    console.log('🔍 檢查工作表1的最新資料...');
    const spreadsheet = SpreadsheetApp.openById('1rxuODXlZpQ5PZ8Gm4lq5gxLimhcpaBHtZ2Q2z_xScz0');
    const sheet = spreadsheet.getSheetByName('工作表1');
    
    if (!sheet) {
      console.log('❌ 找不到「工作表1」');
      return { success: false, error: '找不到工作表1' };
    }
    
    const lastRow = sheet.getLastRow();
    const lastColumn = sheet.getLastColumn();
    
    console.log(`📊 工作表1資訊:`);
    console.log(`  總行數: ${lastRow}`);
    console.log(`  總列數: ${lastColumn}`);
    
    if (lastRow > 0) {
      console.log('📝 所有資料行:');
      for (let row = 1; row <= lastRow; row++) {
        const rowData = sheet.getRange(row, 1, 1, lastColumn).getValues()[0];
        console.log(`  第${row}行:`, rowData.slice(0, 8)); // 只顯示前8列避免太長
      }
      
      // 顯示最後3行的完整資料
      console.log('📋 最後3行完整資料:');
      for (let row = Math.max(1, lastRow - 2); row <= lastRow; row++) {
        const rowData = sheet.getRange(row, 1, 1, lastColumn).getValues()[0];
        console.log(`  第${row}行完整資料:`, rowData);
      }
    } else {
      console.log('⚠️ 工作表是空的');
    }
    
    return {
      success: true,
      totalRows: lastRow,
      totalColumns: lastColumn,
      lastRowData: lastRow > 0 ? sheet.getRange(lastRow, 1, 1, lastColumn).getValues()[0] : null
    };
    
  } catch (error) {
    console.error('❌ 檢查資料時發生錯誤:', error);
    return { success: false, error: error.toString() };
  }
}

/**
 * 強制新增一筆資料並立即檢查
 */
function forceAddAndCheck() {
  try {
    console.log('🚀 開始強制新增資料...');
    const spreadsheet = SpreadsheetApp.openById('1rxuODXlZpQ5PZ8Gm4lq5gxLimhcpaBHtZ2Q2z_xScz0');
    const sheet = spreadsheet.getSheetByName('工作表1');
    
    if (!sheet) {
      console.log('❌ 找不到「工作表1」，嘗試使用第一個工作表');
      const firstSheet = spreadsheet.getSheets()[0];
      console.log('✅ 使用第一個工作表:', firstSheet.getName());
      return forceAddToSheet(firstSheet);
    }
    
    console.log('✅ 找到「工作表1」');
    return forceAddToSheet(sheet);
    
  } catch (error) {
    console.error('❌ 強制新增失敗:', error);
    return { success: false, error: error.toString() };
  }
}

function forceAddToSheet(targetSheet) {
  try {
    const sheetName = targetSheet.getName();
    console.log(`📝 開始向「${sheetName}」新增資料...`);
    
    // 檢查當前行數
    const beforeRows = targetSheet.getLastRow();
    console.log(`📊 新增前總行數: ${beforeRows}`);
    
    // 創建唯一的測試資料
    const timestamp = new Date().toISOString();
    const uniqueId = Date.now();
    
    const testData = [
      timestamp, // 時間戳記
      `蔣哥測試_${uniqueId}`, // 姓名
      '男性', // 性別
      `09${String(uniqueId).slice(-8)}`, // 電話
      '桃園', // 居住地
      `test_${uniqueId}@jiang.com`, // Email
      `jiang_test_${uniqueId}`, // LINE ID
      '手動測試', // 介紹人類型
      '蔣哥', // 介紹人姓名
      '購屋與房地產趨勢分析, 桃園房市投資價值剖析', // 資訊需求
      '想要了解房產投資機會', // 參加動機
      '是', // 同意條款
      '是', // 確認活動
      'Google Apps Script 強制測試', // 瀏覽器
      '強制新增測試' // 來源
    ];
    
    console.log('📝 準備新增的資料:', testData);
    
    // 強制新增資料
    targetSheet.appendRow(testData);
    
    // 立即檢查結果
    const afterRows = targetSheet.getLastRow();
    console.log(`📊 新增後總行數: ${afterRows}`);
    
    if (afterRows > beforeRows) {
      console.log('✅ 資料新增成功！');
      
      // 讀取新增的資料確認
      const newRowData = targetSheet.getRange(afterRows, 1, 1, testData.length).getValues()[0];
      console.log('📋 新增的資料內容:', newRowData);
      
      // 設置格式
      const newRowRange = targetSheet.getRange(afterRows, 1, 1, testData.length);
      newRowRange.setBorder(true, true, true, true, true, true);
      newRowRange.setBackground('#E8F5E8'); // 淺綠色背景標示新資料
      
      console.log(`🎉 成功！資料已添加到「${sheetName}」的第 ${afterRows} 行`);
      console.log(`📧 測試Email: test_${uniqueId}@jiang.com`);
      console.log(`📱 測試電話: 09${String(uniqueId).slice(-8)}`);
      
      return {
        success: true,
        message: `資料已成功添加到「${sheetName}」的第 ${afterRows} 行`,
        sheetName: sheetName,
        rowNumber: afterRows,
        data: {
          name: `蔣哥測試_${uniqueId}`,
          email: `test_${uniqueId}@jiang.com`,
          phone: `09${String(uniqueId).slice(-8)}`,
          timestamp: timestamp
        }
      };
    } else {
      console.log('❌ 資料新增失敗，行數沒有增加');
      return { success: false, error: '資料新增失敗，行數沒有增加' };
    }
    
  } catch (error) {
    console.error('❌ 新增資料時發生錯誤:', error);
    return { success: false, error: error.toString() };
  }
}

/**
 * 檢查 Google Apps Script 實際連接的 Google Sheets
 */
function checkActualConnection() {
  try {
    console.log('🔍 檢查 Google Apps Script 實際連接的 Google Sheets...');
    
    // 檢查當前使用的 Google Sheets ID
    const currentSheetId = '1rxuODXlZpQ5PZ8Gm4lq5gxLimhcpaBHtZ2Q2z_xScz0';
    console.log('📋 當前設定的 Google Sheets ID:', currentSheetId);
    
    // 嘗試打開 Google Sheets
    const spreadsheet = SpreadsheetApp.openById(currentSheetId);
    console.log('✅ 成功打開 Google Sheets');
    console.log('📊 試算表名稱:', spreadsheet.getName());
    console.log('🔗 試算表 URL:', spreadsheet.getUrl());
    console.log('🆔 試算表 ID:', spreadsheet.getId());
    
    // 檢查所有工作表
    const sheets = spreadsheet.getSheets();
    console.log('📋 所有工作表:');
    sheets.forEach((sheet, index) => {
      console.log(`  ${index + 1}. "${sheet.getName()}" - 行數: ${sheet.getLastRow()}, 列數: ${sheet.getLastColumn()}`);
    });
    
    // 檢查「工作表1」的內容
    const targetSheet = spreadsheet.getSheetByName('工作表1');
    if (targetSheet) {
      console.log('✅ 找到「工作表1」');
      console.log('📊 工作表1行數:', targetSheet.getLastRow());
      console.log('📊 工作表1列數:', targetSheet.getLastColumn());
      
      if (targetSheet.getLastRow() > 0) {
        console.log('📝 工作表1的資料:');
        for (let row = 1; row <= Math.min(5, targetSheet.getLastRow()); row++) {
          const rowData = targetSheet.getRange(row, 1, 1, Math.min(5, targetSheet.getLastColumn())).getValues()[0];
          console.log(`  第${row}行:`, rowData);
        }
      }
    } else {
      console.log('❌ 找不到「工作表1」');
    }
    
    // 檢查是否有其他可能的工作表
    console.log('🔍 檢查是否有其他包含資料的工作表...');
    sheets.forEach((sheet, index) => {
      if (sheet.getLastRow() > 0) {
        console.log(`📊 工作表 "${sheet.getName()}" 有 ${sheet.getLastRow()} 行資料`);
        // 顯示前幾行
        for (let row = 1; row <= Math.min(3, sheet.getLastRow()); row++) {
          const rowData = sheet.getRange(row, 1, 1, Math.min(5, sheet.getLastColumn())).getValues()[0];
          console.log(`  第${row}行:`, rowData);
        }
      }
    });
    
    return {
      success: true,
      spreadsheetName: spreadsheet.getName(),
      spreadsheetUrl: spreadsheet.getUrl(),
      spreadsheetId: spreadsheet.getId(),
      sheets: sheets.map(sheet => ({
        name: sheet.getName(),
        rows: sheet.getLastRow(),
        columns: sheet.getLastColumn()
      }))
    };
    
  } catch (error) {
    console.error('❌ 檢查連接時發生錯誤:', error);
    return {
      success: false,
      error: error.toString()
    };
  }
}

/**
 * 檢查所有可能的 Google Sheets 連接
 */
function checkAllPossibleSheets() {
  try {
    console.log('🔍 檢查所有可能的 Google Sheets 連接...');
    
    // 可能的 Google Sheets ID 列表
    const possibleSheetIds = [
      '1rxuODXlZpQ5PZ8Gm4lq5gxLimhcpaBHtZ2Q2z_xScz0', // 您指定的
      '1X8l3vEAecBEldAVoRB_iezN7szWLPgnf4ZovvqX2IIU', // 之前的
    ];
    
    const results = [];
    
    for (const sheetId of possibleSheetIds) {
      try {
        console.log(`\n🔍 檢查 Google Sheets ID: ${sheetId}`);
        const spreadsheet = SpreadsheetApp.openById(sheetId);
        
        const result = {
          id: sheetId,
          name: spreadsheet.getName(),
          url: spreadsheet.getUrl(),
          sheets: []
        };
        
        const sheets = spreadsheet.getSheets();
        sheets.forEach(sheet => {
          result.sheets.push({
            name: sheet.getName(),
            rows: sheet.getLastRow(),
            columns: sheet.getLastColumn(),
            hasData: sheet.getLastRow() > 0
          });
        });
        
        results.push(result);
        console.log(`✅ 成功連接: ${spreadsheet.getName()}`);
        console.log(`📊 工作表數量: ${sheets.length}`);
        
      } catch (error) {
        console.log(`❌ 無法連接 ID: ${sheetId} - ${error.message}`);
        results.push({
          id: sheetId,
          error: error.message
        });
      }
    }
    
    return {
      success: true,
      results: results
    };
    
  } catch (error) {
    console.error('❌ 檢查所有 Sheets 時發生錯誤:', error);
    return {
      success: false,
      error: error.toString()
    };
  }
}

/**
 * 簡單測試函數 - 直接測試Sheets連接
 */
function testSheetsConnection() {
  try {
    console.log('🧪 測試Google Sheets連接...');
    const spreadsheet = SpreadsheetApp.openById('1rxuODXlZpQ5PZ8Gm4lq5gxLimhcpaBHtZ2Q2z_xScz0');
    const sheet = spreadsheet.getSheetByName('工作表1') || spreadsheet.getSheets()[0];
    const lastRow = sheet.getLastRow();
    console.log('🧪 Sheets連接成功，最後一行:', lastRow);
    
    // 添加測試行
    sheet.appendRow([
      new Date().toISOString(),
      '測試用戶',
      'male',
      '0912-345-678',
      'taipei',
      'test@example.com',
      'test123',
      '朋友介紹',
      '張三',
      '測試資訊需求',
      '測試動機',
      '是',
      '是',
      'Test Browser',
      'https://test.com'
    ]);
    
    console.log('🧪 測試行添加成功');
    return '測試成功';
  } catch (error) {
    console.error('🧪 測試失敗:', error);
    return '測試失敗: ' + error.toString();
  }
}
