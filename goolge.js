// 處理表單資料的共用函數
function processFormData(data) {
  console.log('處理表單資料:', data);
  try {
    const SHEET_ID = '11ZfpYUcnXYVmGWGTTP3xdWhOcmHy0snBSMK1omR9-OM';
    const sheet = SpreadsheetApp.openById(SHEET_ID).getActiveSheet();

    // 黑名單檢查
    const blacklistSheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('黑名單');
    if (blacklistSheet) {
      const blacklistData = blacklistSheet.getDataRange().getValues();
      for (let i = 1; i < blacklistData.length; i++) {
        // 根據新表格結構，聯絡ID在D欄(索引3)
        if (blacklistData[i][3] === data.contactId) {
          return ContentService.createTextOutput('黑名單').setMimeType(ContentService.MimeType.JSON);
        }
      }
    }

    // 時段重複檢查 - 檢查預約星期(J欄)和預約時段(K欄)
    const existingData = sheet.getDataRange().getValues();
    for (let i = 1; i < existingData.length; i++) {
      // J欄(索引9)是預約星期，K欄(索引10)是預約時段
      if (existingData[i][9] === data.weekday && existingData[i][10] === data.timeSlot) {
        return ContentService.createTextOutput('此時段已被預約').setMimeType(ContentService.MimeType.JSON);
      }
    }

    // 生成訂單編號
    const orderDate = new Date();
    const dateStr = orderDate.toLocaleDateString('zh-TW', { timeZone: 'Asia/Taipei' }).replace(/\//g, '');
    const rowNumber = sheet.getLastRow(); // 獲取最後一行的行號
    const orderNumber = `ORD${dateStr}${String(rowNumber).padStart(3, '0')}`;
    
    const timestamp = new Date().toLocaleString('zh-TW', { timeZone: 'Asia/Taipei' });
    
    // 根據新表格結構準備資料
    const rowData = [
      orderNumber,                          // A: 訂單編號
      timestamp,                            // B: 提交時間
      data.contactMethod || '',             // C: 聯絡方式
      data.contactId || '',                 // D: 聯絡ID
      parseInt(data.legendCount) || 0,      // E: 傳說對決數量
      parseInt(data.voiceCount) || 0,       // F: 語音通話數量
      parseInt(data.chatCount) || 0,        // G: 打字聊天數量
      parseInt(data.partyCount) || 0,       // H: 全民Party唱歌數量
      parseInt(data.consultCount) || 0,     // I: 情感諮詢數量
      data.weekday || '',                   // J: 預約星期
      data.timeSlot || '',                  // K: 預約時段
      data.fullDateTime || '',              // L: 完整預約時間
      data.selectedItems || '',             // M: 所選項目
      parseInt(data.subtotal) || 0,         // N: 服務金額小計
      parseInt(data.total) || 0,            // O: 最終應付金額
      '待付款',                             // P: 付款狀態
      '待服務',                             // Q: 服務狀態
      data.remark || ''                     // R: 備註
    ];

    sheet.appendRow(rowData);
    console.log('資料已寫入工作表');

    return ContentService.createTextOutput('success').setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    console.error('錯誤詳情:', err);
    return ContentService.createTextOutput('error: ' + err.message).setMimeType(ContentService.MimeType.JSON);
  }
}

// 處理 POST 請求
function doPost(e) {
  console.log('收到 POST 請求');
  const data = JSON.parse(e.postData.contents);
  return processFormData(data);
}

// 用於 GET 測試和接收參數
function doGet(e) {
  console.log('收到 GET 請求');
  
  // 如果有參數，處理表單資料
  if (e.parameter && Object.keys(e.parameter).length > 0) {
    console.log('GET 參數:', e.parameter);
    return processFormData(e.parameter);
  }
  
  return ContentService.createTextOutput('Google Apps Script 運作正常').setMimeType(ContentService.MimeType.JSON);
}

// 處理 OPTIONS 請求 (CORS 預檢)
function doOptions(e) {
  return HtmlService.createHtmlOutput()
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

