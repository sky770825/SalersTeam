/**
 * 蔣哥行銷網頁表單提交處理器
 * 用於接收網頁表單數據並寫入Google Sheets
 */

/**
 * 處理 GET 請求 - 用於測試和健康檢查
 */
function doGet(e) {
  try {
    console.log('收到 GET 請求:', e);
    
    // 返回簡單的狀態訊息
    return ContentService
      .createTextOutput(JSON.stringify({
        success: true,
        message: 'Google Apps Script 運行正常',
        timestamp: new Date().toISOString(),
        method: 'GET'
      }))
      .setMimeType(ContentService.MimeType.JSON);
      
  } catch (error) {
    console.error('GET 請求處理錯誤:', error);
    return ContentService
      .createTextOutput(JSON.stringify({
        success: false,
        error: error.toString(),
        message: 'GET 請求處理失敗'
      }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doPost(e) {
  try {
    // 調試：檢查接收到的參數
    console.log('接收到的完整參數:', e);
    console.log('參數類型:', typeof e);
    
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
      console.log('實際接收到的參數結構:', JSON.stringify(e, null, 2));
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
    
    // 處理介紹人資訊
    let referrer = '';
    if (referrerType) {
      if (referrerType === '華地產介紹' || referrerType === '朋友介紹') {
        referrer = referrerName ? `${referrerType}${referrerName}` : referrerType;
      } else {
        referrer = referrerType;
      }
    }
    
    // 詳細調試：記錄所有接收到的參數
    console.log('📋 接收到的所有參數:');
    console.log('  name:', name);
    console.log('  gender:', gender);
    console.log('  phone:', phone);
    console.log('  location:', location);
    console.log('  email:', email);
    console.log('  lineId:', lineId);
    console.log('  referrerType:', referrerType);
    console.log('  referrerName:', referrerName);
    console.log('  referrer (合併後):', referrer);
    console.log('  infoNeeds:', infoNeeds);
    
    // 記錄詳細的資訊需求數據
    console.log('接收到的資訊需求:', infoNeeds);
    console.log('📧 接收到的Email:', email);
    const motivation = e.parameter.motivation || '';
    const agreeTerms = e.parameter.agreeTerms || '';
    const agreeEvent = e.parameter.agreeEvent || '';
    const timestamp = e.parameter.timestamp || '';
    const userAgent = e.parameter.userAgent || '';

    // 打開Google Sheets
    const sheet = SpreadsheetApp.openById('1X8l3vEAecBEldAVoRB_iezN7szWLPgnf4ZovvqX2IIU').getActiveSheet();
    
    // 如果是第一次運行，添加標題行
    if (sheet.getLastRow() === 0) {
      sheet.appendRow([
        '時間戳記', '姓名', '性別', '電話', '居住地', 'Email', 'LINE ID', '介紹人',
        '資訊需求', '參加動機', '同意條款', '確認活動', '瀏覽器', '來源'
      ]);
      
      // 設置標題行格式
      const headerRange = sheet.getRange(1, 1, 1, 14);
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
      referrer, // 介紹人欄位
      infoNeeds, 
      motivation, 
      agreeTerms, 
      agreeEvent, 
      userAgent, 
      referrer // 來源欄位（重複使用介紹人資訊）
    ];
    
    sheet.appendRow(newRow);
    
    // 設置新行的格式
    const lastRow = sheet.getLastRow();
    const dataRange = sheet.getRange(lastRow, 1, 1, 14);
    dataRange.setBorder(true, true, true, true, true, true);
    
    // 自動調整列寬
    sheet.autoResizeColumns(1, 14);
    
    // 記錄成功日誌
    console.log('成功提交表單數據:', {
      name: name,
      phone: phone,
      timestamp: timestamp
    });
    
    // 發送確認信件
    try {
      sendConfirmationEmail(name, phone, email, lineId, referrer);
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
function sendConfirmationEmail(name, phone, email, lineId, referrer) {
  try {
    // 檢查是否有Email地址（這裡需要從表單中獲取Email）
    // 如果沒有Email，可以發送到LINE或使用其他通知方式
    
    // 活動資訊設定（可輕鬆修改）
    const eventInfo = {
      title: '蔣哥房產分析說明會',
      date: '2025年10月5日（日）',
      time: '下午2:00 準時開始',
      location: '桃園市桃園區大興西路三段90號（銷售中心）',
      prize: '大金空調清淨機（市價2萬元）',
      contactPhone: '03-123-4567', // 可以設定固定聯絡電話
      teamName: 'BNI華地產房產行銷組團隊',
      referrer: '蔣哥房產分析團隊' // 介紹人/推薦人
    };
    
    // Email 發送設定
    const emailConfig = {
      senderName: '蔣哥房產分析團隊', // 發送者名稱
      replyTo: '', // 回覆地址（可選，留空則使用預設）
      // 注意：實際發送帳號是 Google Apps Script 專案的擁有者帳號
    };
    
    // 信件內容
    const subject = `報名確認信 - ${eventInfo.title}`;
    const body = `
親愛的 ${name} 您好，

感謝您報名參加「${eventInfo.title}」！

1、 活動資訊：
• 時間：${eventInfo.date} ${eventInfo.time}
• 地點：${eventInfo.location}

2、現場抽獎：
${eventInfo.prize}

3.注意事項：
1. 請提前15分鐘到達會場
2. 請攜帶身分證件
3. 現場提供茶水點心
4. 活動全程免費，無推銷壓力

4.如有任何問題，請聯繫：
介紹人：${referrer || '無'}

感謝您的報名，期待與您相見！

${eventInfo.teamName}
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
      const sheet = SpreadsheetApp.openById('1X8l3vEAecBEldAVoRB_iezN7szWLPgnf4ZovvqX2IIU').getActiveSheet();
      const lastRow = sheet.getLastRow();
      sheet.getRange(lastRow, 14).setValue('確認信件已發送'); // 在N列添加備註
      
      console.log('確認信件已成功發送到:', email);
    } else {
      console.log('沒有Email地址，無法發送確認信件');
      
      // 記錄到Google Sheets的備註欄位
      const sheet = SpreadsheetApp.openById('1X8l3vEAecBEldAVoRB_iezN7szWLPgnf4ZovvqX2IIU').getActiveSheet();
      const lastRow = sheet.getLastRow();
      sheet.getRange(lastRow, 14).setValue('無Email地址，未發送確認信'); // 在N列添加備註
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
      email: 'test@example.com', // 添加email測試
      lineId: 'test123',
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
 * 測試參數接收函數
 */
function testParameterReceiving() {
  console.log('🧪 測試參數接收...');
  
  // 模擬從網頁表單接收的數據
  const mockEvent = {
    parameter: {
      name: '張三',
      gender: 'male',
      phone: '0912-345-678',
      location: 'taipei',
      email: 'zhang@example.com',
      lineId: 'zhang123',
      infoNeeds: '購屋與房地產趨勢分析, 桃園房市投資價值剖析：桃園房屋真的值得買嗎？',
      motivation: '想要了解房產投資',
      agreeTerms: '是',
      agreeEvent: '是',
      timestamp: new Date().toISOString(),
      userAgent: 'Mozilla/5.0...',
      referrer: 'https://example.com'
    }
  };
  
  console.log('🧪 模擬接收到的參數:');
  Object.keys(mockEvent.parameter).forEach(key => {
    console.log(`  ${key}: ${mockEvent.parameter[key]}`);
  });
  
  return '參數接收測試完成';
}

/**
 * 簡單測試函數 - 直接測試Sheets連接
 */
function testSheetsConnection() {
  try {
    console.log('🧪 測試Google Sheets連接...');
    const sheet = SpreadsheetApp.openById('1X8l3vEAecBEldAVoRB_iezN7szWLPgnf4ZovvqX2IIU').getActiveSheet();
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

/**
 * 測試確認信件發送功能
 */
function testEmailSending() {
  try {
    console.log('🧪 測試確認信件發送功能...');
    
    // 測試數據
    const testName = '測試用戶';
    const testPhone = '0912-345-678';
    const testEmail = 'test@example.com'; // 請替換為您的測試email
    const testLineId = 'test123';
    
    console.log('🧪 測試參數:', {
      name: testName,
      phone: testPhone,
      email: testEmail,
      lineId: testLineId
    });
    
    // 調用發送確認信件函數
    sendConfirmationEmail(testName, testPhone, testEmail, testLineId);
    
    console.log('🧪 確認信件發送測試完成');
    return '確認信件發送測試成功';
    
  } catch (error) {
    console.error('🧪 確認信件發送測試失敗:', error);
    return '確認信件發送測試失敗: ' + error.toString();
  }
}

/**
 * 清理舊數據函數 - 可選，用於清理超過30天的數據
 */
function cleanupOldData() {
  try {
    const sheet = SpreadsheetApp.openById('1X8l3vEAecBEldAVoRB_iezN7szWLPgnf4ZovvqX2IIU').getActiveSheet();
    const data = sheet.getDataRange().getValues();
    const cutoffDate = new Date();
    cutoffDate.setDate(cutoffDate.getDate() - 30); // 30天前
    
    let rowsToDelete = [];
    
    for (let i = data.length - 1; i > 0; i--) { // 跳過標題行
      const timestamp = new Date(data[i][0]);
      if (timestamp < cutoffDate) {
        rowsToDelete.push(i + 1); // +1 因為Sheet行號從1開始
      }
    }
    
    // 從後往前刪除，避免行號變化
    for (let row of rowsToDelete.reverse()) {
      sheet.deleteRow(row);
    }
    
    console.log(`清理了 ${rowsToDelete.length} 行舊數據`);
    
  } catch (error) {
    console.error('清理數據時發生錯誤:', error);
  }
}
