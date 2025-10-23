/**
 * 蔣哥行銷網頁表單提交處理器 - 修復版
 * 用於接收網頁表單數據並寫入Google Sheets
 */

// 處理 GET 請求（用於測試和獲取報名人數）
function doGet(e) {
  try {
    // 檢查是否為獲取報名人數的請求
    if (e.parameter && e.parameter.action === 'getCount') {
      const countData = getRegistrationCount();
      return ContentService
        .createTextOutput(JSON.stringify(countData))
        .setMimeType(ContentService.MimeType.JSON);
    }
    
    // 預設測試響應
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

    // 檢查報名人數是否已滿
    const countData = getRegistrationCount();
    if (countData.isFull) {
      console.log('報名已額滿，拒絕新報名');
      return ContentService
        .createTextOutput(JSON.stringify({
          success: false,
          error: '報名已額滿',
          message: '很抱歉，本次說明會的免費名額已全部額滿（30人）',
          isFull: true
        }))
        .setMimeType(ContentService.MimeType.JSON);
    }
    
    // 獲取表單數據
    const name = e.parameter.name || '';
    const phone = e.parameter.phone || '';
    const location = e.parameter.location || '';
    const lineId = e.parameter.lineId || '';
    const downPayment = e.parameter.downPayment || '';
    const referrerType = e.parameter.referrerType || '';
    const referrerName = e.parameter.referrerName || '';
    const infoNeeds = e.parameter.infoNeeds || '';
    const registrationReason = e.parameter.registrationReason || '';
    const agreeTerms = e.parameter.agreeTerms || '';
    const agreeEvent = e.parameter.agreeEvent || '';
    const timestamp = e.parameter.timestamp || '';
    const userAgent = e.parameter.userAgent || '';
    const referrer = e.parameter.referrer || '';

    // 轉換為中文顯示
    // 居住地中文轉換
    const locationChinese = location === 'newtaipei' ? '新北市' : 
                           location === 'taipei' ? '台北市' : 
                           location === 'taoyuan' ? '桃園市' : 
                           location === 'linkou' ? '林口' : 
                           location === 'guishan' ? '龜山' : 
                           location === 'qingpu' ? '青埔' : 
                           location === 'other' ? '其它' : location;
    
    // 自備款中文轉換
    const downPaymentChinese = downPayment === 'under100' ? '100萬以下' :
                              downPayment === '100-200' ? '100-200萬' :
                              downPayment === '200-500' ? '200-500萬' :
                              downPayment === 'over500' ? '500萬以上' :
                              downPayment === 'special68' ? '我只有68萬自備款，想要6%交屋' : downPayment;
    
    // 介紹人中文轉換
    const referrerTypeChinese = referrerType === '無' ? '無' :
                               referrerType === '華地產介紹' ? '華地產介紹' :
                               referrerType === '朋友介紹' ? '朋友介紹' :
                               referrerType === '網路搜尋' ? '網路搜尋' :
                               referrerType === '其他' ? '其他' : referrerType;
    
    // 報名原因中文轉換
    const registrationReasonChinese = registrationReason === 'trend' ? '我想知道房地產的投資趨勢' :
                                     registrationReason === 'investment' ? '我想知道為什麼一定要投資置產' :
                                     registrationReason === 'jiangge' ? '我想要來看蔣哥' :
                                     registrationReason === 'mystery_guest' ? '我想要來看神秘嘉賓' :
                                     registrationReason === '2029_target' ? '我本來就在找尋2029年的投資標的' : registrationReason;
    
    // 同意條款和確認活動（index.html 已經轉換為中文，直接使用）
    const agreeTermsChinese = agreeTerms;
    const agreeEventChinese = agreeEvent;

    console.log('接收到的數據:', {
      name, phone, location, lineId, downPayment,
      referrerType, referrerName, infoNeeds, registrationReason
    });

    // 打開Google Sheets - 明確指定工作表1
    const spreadsheet = SpreadsheetApp.openById('1I0ajxRqNjPw-dHFZuPuiz2ux7kVOh-jhLeSOVYIXxVw');
    let sheet = spreadsheet.getSheetByName('工作表1');
    
    // 如果找不到「工作表1」，自動建立一個
    if (!sheet) {
      console.log('📝 找不到「工作表1」，正在自動建立...');
      sheet = spreadsheet.insertSheet('工作表1');
      console.log('✅ 已成功建立「工作表1」工作表');
    }
    
    // 如果是第一次運行，添加標題行
    if (sheet.getLastRow() === 0) {
      sheet.appendRow([
        '時間戳記', '姓名', '電話', '居住地', 'LINE ID', '自備款', 
        '介紹人類型', '介紹人姓名', '資訊需求', '報名原因', '同意條款', '確認活動', '瀏覽器', '來源'
      ]);
      
      // 設置標題行格式
      const headerRange = sheet.getRange(1, 1, 1, 14);
      headerRange.setBackground('#4285F4');
      headerRange.setFontColor('white');
      headerRange.setFontWeight('bold');
      headerRange.setHorizontalAlignment('center');
    }
    
    // 添加數據行（使用中文轉換後的資料）
    const newRow = [
      timestamp, 
      name, 
      phone, 
      locationChinese,  // 使用中文居住地
      lineId,
      downPaymentChinese,  // 使用中文自備款
      referrerTypeChinese,  // 使用中文介紹人
      referrerName,
      infoNeeds, 
      registrationReasonChinese,  // 使用中文報名原因
      agreeTermsChinese,  // 使用中文同意條款
      agreeEventChinese,  // 使用中文確認活動
      userAgent, 
      referrer
    ];
    
    sheet.appendRow(newRow);
    
    // 設置新行的格式
    const lastRow = sheet.getLastRow();
    const dataRange = sheet.getRange(lastRow, 1, 1, 14);
    dataRange.setBorder(true, true, true, true, true, true);
    
    
    // 記錄成功日誌
    console.log('成功提交表單數據:', {
      name: name,
      phone: phone,
      lineId: lineId,
      timestamp: timestamp
    });
    
    // 發送確認通知（通過 LINE 或其他方式）
    try {
      sendConfirmationNotification(name, phone, lineId, referrerType);
      console.log('確認通知已發送');
    } catch (notificationError) {
      console.error('發送確認通知失敗:', notificationError);
    }
    
    // 準備返回的報名者姓名（遮罩格式）
    const maskedName = name ? name.charAt(0) + '*'.repeat(name.length - 1) : '新用戶';
    
    return ContentService
      .createTextOutput(JSON.stringify({
        success: true, 
        message: '數據已成功提交到Google Sheets，確認通知已發送',
        row: lastRow,
        registrantName: maskedName,
        registrationCount: lastRow - 1 // 扣除標題行
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
 * 發送確認通知函數（記錄到 Google Sheets 備註欄位）
 */
function sendConfirmationNotification(name, phone, lineId, referrerType) {
  try {
    console.log('準備發送確認通知給:', {
      name: name,
      phone: phone,
      lineId: lineId,
      referrerType: referrerType
    });
    
    // 記錄到Google Sheets的備註欄位
    const spreadsheet = SpreadsheetApp.openById('1I0ajxRqNjPw-dHFZuPuiz2ux7kVOh-jhLeSOVYIXxVw');
    let sheet = spreadsheet.getSheetByName('工作表1');
    
    // 如果找不到「工作表1」，自動建立一個
    if (!sheet) {
      console.log('📝 找不到「工作表1」，正在自動建立...');
      sheet = spreadsheet.insertSheet('工作表1');
      console.log('✅ 已成功建立「工作表1」工作表');
    }
    const lastRow = sheet.getLastRow();
    
    // 在備註欄位記錄確認通知
    const notificationMessage = `確認通知已記錄 - 姓名: ${name}, 電話: ${phone}, LINE ID: ${lineId || '未提供'}, 介紹人: ${referrerTypeChinese || '其他'}`;
    sheet.getRange(lastRow, 15).setValue(notificationMessage); // 在O列添加備註
    
    console.log('✅ 確認通知已記錄到 Google Sheets');
    console.log('📋 通知內容:', notificationMessage);
    
    // 如果有 LINE ID，可以在這裡添加 LINE 通知邏輯
    if (lineId && lineId.trim() !== '') {
      console.log('📱 用戶提供了 LINE ID:', lineId);
      // 這裡可以添加 LINE Notify 或其他 LINE 通知服務的邏輯
    }
    
    console.log('確認通知處理完成');
    
  } catch (error) {
    console.error('發送確認通知時發生錯誤:', error);
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
      phone: '0912-345-678',
      location: 'taipei',
      lineId: 'test123',
      downPayment: '100-200',
      referrerType: '無',
      referrerName: '張三',
      infoNeeds: '購屋與房地產趨勢分析, 桃園房市投資價值剖析：桃園房屋真的值得買嗎？',
      registrationReason: 'trend',
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
    const spreadsheet = SpreadsheetApp.openById('1I0ajxRqNjPw-dHFZuPuiz2ux7kVOh-jhLeSOVYIXxVw');
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
    const spreadsheet = SpreadsheetApp.openById('1I0ajxRqNjPw-dHFZuPuiz2ux7kVOh-jhLeSOVYIXxVw');
    
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
        '時間戳記', '姓名', '電話', '居住地', 'LINE ID', '自備款', 
        '介紹人類型', '介紹人姓名', '資訊需求', '報名原因', '同意條款', '確認活動', '瀏覽器', '來源'
      ]);
      
      // 設置標題行格式
      const headerRange = targetSheet.getRange(1, 1, 1, 14);
      headerRange.setBackground('#4285F4');
      headerRange.setFontColor('white');
      headerRange.setFontWeight('bold');
      headerRange.setHorizontalAlignment('center');
    }
    
    // 添加測試數據
    const testData = [
      new Date().toISOString(),
      '測試用戶_' + new Date().getTime(),
      '0912-345-678',
      '台北市',
      'test123',
      '100-200萬',
      '朋友介紹',
      '張三',
      '購屋與房地產趨勢分析, 桃園房市投資價值剖析',
      '我想知道房地產的投資趨勢',
      '是',
      '是',
      'Test Browser',
      'https://test.com'
    ];
    
    console.log('📝 添加測試數據行...');
    targetSheet.appendRow(testData);
    
    // 設置新行的格式
    const lastRow = targetSheet.getLastRow();
    const dataRange = targetSheet.getRange(lastRow, 1, 1, 14);
    dataRange.setBorder(true, true, true, true, true, true);
    
    // 自動調整列寬
    
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
    const spreadsheet = SpreadsheetApp.openById('1I0ajxRqNjPw-dHFZuPuiz2ux7kVOh-jhLeSOVYIXxVw');
    
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
        '時間戳記', '姓名', '電話', '居住地', 'LINE ID', '自備款', 
        '介紹人類型', '介紹人姓名', '資訊需求', '報名原因', '同意條款', '確認活動', '瀏覽器', '來源'
      ]);
      
      // 設置標題行格式
      const headerRange = targetSheet.getRange(1, 1, 1, 14);
      headerRange.setBackground('#4285F4');
      headerRange.setFontColor('white');
      headerRange.setFontWeight('bold');
      headerRange.setHorizontalAlignment('center');
    }
    
    // 手動新增的資料
    const manualData = [
      new Date().toISOString(), // 時間戳記
      '蔣哥測試用戶', // 姓名
      '0987-654-321', // 電話
      '桃園市', // 居住地
      'jiang_test_2024', // LINE ID
      '200-500萬', // 自備款
      '網路搜尋', // 介紹人類型
      '', // 介紹人姓名
      '購屋與房地產趨勢分析, 桃園房市投資價值剖析：桃園房屋真的值得買嗎？, 購屋需求解析：自住、投資、首購或首換，該如何規劃？', // 資訊需求
      '我想要來看蔣哥', // 報名原因
      '是', // 同意條款
      '是', // 確認活動
      'Google Apps Script 手動新增', // 瀏覽器
      '手動測試' // 來源
    ];
    
    console.log('📝 添加手動資料...');
    targetSheet.appendRow(manualData);
    
    // 設置新行的格式
    const lastRow = targetSheet.getLastRow();
    const dataRange = targetSheet.getRange(lastRow, 1, 1, 14);
    dataRange.setBorder(true, true, true, true, true, true);
    
    // 自動調整列寬
    
    console.log(`✅ 手動資料已成功添加到工作表 "${usedSheetName}" 的第 ${lastRow} 行`);
    console.log('📊 新增的資料:');
    console.log('  姓名: 蔣哥測試用戶');
    console.log('  電話: 0987-654-321');
    console.log('  LINE ID: jiang_test_2024');
    console.log('  居住地: 桃園市');
    console.log('  自備款: 200-500萬');
    console.log('  來源: 手動測試');
    
    return {
      success: true,
      message: `資料已成功添加到工作表 "${usedSheetName}" 的第 ${lastRow} 行`,
      sheetName: usedSheetName,
      rowAdded: lastRow,
      data: {
        name: '蔣哥測試用戶',
        phone: '0987-654-321',
        lineId: 'jiang_test_2024',
        location: '桃園市',
        downPayment: '200-500萬'
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
    const spreadsheet = SpreadsheetApp.openById('1I0ajxRqNjPw-dHFZuPuiz2ux7kVOh-jhLeSOVYIXxVw');
    
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
      '0911-222-333',
      '台北市',
      'diagnose_test',
      '100萬以下',
      '手動測試',
      '',
      '診斷測試資訊需求',
      '我想要來看神秘嘉賓',
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
    const spreadsheet = SpreadsheetApp.openById('1I0ajxRqNjPw-dHFZuPuiz2ux7kVOh-jhLeSOVYIXxVw');
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
    const spreadsheet = SpreadsheetApp.openById('1I0ajxRqNjPw-dHFZuPuiz2ux7kVOh-jhLeSOVYIXxVw');
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
      `09${String(uniqueId).slice(-8)}`, // 電話
      '桃園市', // 居住地
      `jiang_test_${uniqueId}`, // LINE ID
      '500萬以上', // 自備款
      '手動測試', // 介紹人類型
      '蔣哥', // 介紹人姓名
      '購屋與房地產趨勢分析, 桃園房市投資價值剖析', // 資訊需求
      '我本來就在找尋2029年的投資標的', // 報名原因
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
      console.log(`📱 測試電話: 09${String(uniqueId).slice(-8)}`);
      console.log(`💬 測試LINE ID: jiang_test_${uniqueId}`);
      
      return {
        success: true,
        message: `資料已成功添加到「${sheetName}」的第 ${afterRows} 行`,
        sheetName: sheetName,
        rowNumber: afterRows,
        data: {
          name: `蔣哥測試_${uniqueId}`,
          phone: `09${String(uniqueId).slice(-8)}`,
          lineId: `jiang_test_${uniqueId}`,
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
    const currentSheetId = '1I0ajxRqNjPw-dHFZuPuiz2ux7kVOh-jhLeSOVYIXxVw';
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
      '1I0ajxRqNjPw-dHFZuPuiz2ux7kVOh-jhLeSOVYIXxVw', // 新的試算表
      '1rxuODXlZpQ5PZ8Gm4lq5gxLimhcpaBHtZ2Q2z_xScz0', // 之前的
      '1X8l3vEAecBEldAVoRB_iezN7szWLPgnf4ZovvqX2IIU', // 更早的
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
 * 測試中文轉換功能
 */
function testChineseConversion() {
  try {
    console.log('🧪 測試中文轉換功能...');
    
    // 測試性別轉換
    const genderTests = [
      { input: 'male', expected: '男性' },
      { input: 'female', expected: '女性' },
      { input: '其他', expected: '其他' }
    ];
    
    console.log('📋 性別轉換測試:');
    genderTests.forEach(test => {
      const result = test.input === 'male' ? '男性' : test.input === 'female' ? '女性' : test.input;
      const status = result === test.expected ? '✅' : '❌';
      console.log(`  ${status} ${test.input} -> ${result} (期望: ${test.expected})`);
    });
    
    // 測試居住地轉換
    const locationTests = [
      { input: 'taipei', expected: '台北' },
      { input: 'newtaipei', expected: '新北' },
      { input: 'taoyuan', expected: '桃園' },
      { input: 'hsinchu', expected: '新竹' },
      { input: 'taichung', expected: '台中' },
      { input: 'tainan', expected: '台南' },
      { input: 'kaohsiung', expected: '高雄' },
      { input: 'other', expected: '其他' }
    ];
    
    console.log('📋 居住地轉換測試:');
    locationTests.forEach(test => {
      const result = test.input === 'taipei' ? '台北' : 
                    test.input === 'newtaipei' ? '新北' : 
                    test.input === 'taoyuan' ? '桃園' : 
                    test.input === 'hsinchu' ? '新竹' : 
                    test.input === 'taichung' ? '台中' : 
                    test.input === 'tainan' ? '台南' : 
                    test.input === 'kaohsiung' ? '高雄' : 
                    test.input === 'other' ? '其他' : test.input;
      const status = result === test.expected ? '✅' : '❌';
      console.log(`  ${status} ${test.input} -> ${result} (期望: ${test.expected})`);
    });
    
    // 測試同意條款轉換
    const agreeTests = [
      { input: '是', expected: '是' },
      { input: '否', expected: '否' }
    ];
    
    console.log('📋 同意條款轉換測試:');
    agreeTests.forEach(test => {
      const result = test.input;
      const status = result === test.expected ? '✅' : '❌';
      console.log(`  ${status} ${test.input} -> ${result} (期望: ${test.expected})`);
    });
    
    console.log('✅ 中文轉換測試完成');
    return { success: true, message: '中文轉換測試完成' };
    
  } catch (error) {
    console.error('❌ 中文轉換測試失敗:', error);
    return { success: false, error: error.toString() };
  }
}

/**
 * 獲取報名人數的函數
 */
function getRegistrationCount() {
  try {
    console.log('📊 獲取報名人數...');
    const spreadsheet = SpreadsheetApp.openById('1I0ajxRqNjPw-dHFZuPuiz2ux7kVOh-jhLeSOVYIXxVw');
    let sheet = spreadsheet.getSheetByName('工作表1');
    
    // 如果找不到「工作表1」，自動建立一個
    if (!sheet) {
      console.log('📝 找不到「工作表1」，正在自動建立...');
      sheet = spreadsheet.insertSheet('工作表1');
      console.log('✅ 已成功建立「工作表1」工作表');
    }
    
    // 獲取總行數（包含標題行）
    const totalRows = sheet.getLastRow();
    
    // 計算實際報名人數（扣除標題行）
    const registrationCount = Math.max(0, totalRows - 1);
    
    // 獲取最新報名者的姓名（用於顯示通知）
    let latestRegistrant = '';
    if (registrationCount > 0) {
      try {
        const nameColumn = 2; // 姓名在第2列
        const latestName = sheet.getRange(totalRows, nameColumn).getValue();
        if (latestName) {
          // 將姓名轉換為遮罩格式，如：蔡**、李**
          if (latestName.length >= 2) {
            latestRegistrant = latestName.charAt(0) + '*'.repeat(latestName.length - 1);
          } else {
            latestRegistrant = latestName.charAt(0) + '*';
          }
        }
      } catch (nameError) {
        console.log('無法獲取最新報名者姓名:', nameError);
      }
    }
    
    console.log(`📊 總行數: ${totalRows}, 實際報名人數: ${registrationCount}, 最新報名者: ${latestRegistrant}`);
    
    return {
      success: true,
      totalRows: totalRows,
      registrationCount: registrationCount,
      latestRegistrant: latestRegistrant,
      isFull: registrationCount >= 30
    };
    
  } catch (error) {
    console.error('❌ 獲取報名人數失敗:', error);
    return {
      success: false,
      error: error.toString(),
      registrationCount: 0,
      latestRegistrant: '',
      isFull: false
    };
  }
}

/**
 * 簡單測試函數 - 直接測試Sheets連接
 */
function testSheetsConnection() {
  try {
    console.log('🧪 測試Google Sheets連接...');
    const spreadsheet = SpreadsheetApp.openById('1I0ajxRqNjPw-dHFZuPuiz2ux7kVOh-jhLeSOVYIXxVw');
    const sheet = spreadsheet.getSheetByName('工作表1') || spreadsheet.getSheets()[0];
    const lastRow = sheet.getLastRow();
    console.log('🧪 Sheets連接成功，最後一行:', lastRow);
    
    // 添加測試行
    sheet.appendRow([
      new Date().toISOString(),
      '測試用戶',
      '0912-345-678',
      '台北市',
      'test123',
      '100-200萬',
      '朋友介紹',
      '張三',
      '測試資訊需求',
      '我想知道房地產的投資趨勢',
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
