// 處理表單資料的共用函數
function processFormData(data) {
  console.log('處理表單資料:', data);
  try {
    // 後端資料驗證
    if (!data.contactMethod || !data.contactId || !data.weekday || !data.timeSlot) {
      console.log('資料驗證失敗：缺少必填欄位');
      return ContentService.createTextOutput('資料驗證失敗：缺少必填欄位').setMimeType(ContentService.MimeType.JSON);
    }
    
    // 檢查是否至少有一個服務項目
    const totalServices = (parseInt(data.legendCount) || 0) + 
                         (parseInt(data.voiceCount) || 0) + 
                         (parseInt(data.chatCount) || 0) + 
                         (parseInt(data.partyCount) || 0) + 
                         (parseInt(data.consultCount) || 0);
    
    if (totalServices === 0) {
      console.log('資料驗證失敗：未選擇任何服務項目');
      return ContentService.createTextOutput('資料驗證失敗：請至少選擇一個服務項目').setMimeType(ContentService.MimeType.JSON);
    }
    
    const SHEET_ID = '17iFNvOb-Gl5B1nDMC9Jf0ls1s8t_a8UK9n2kB23fEo4';
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('訂單');

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
      let existingTimeSlot = existingData[i][10];
      
      // 如果現有時段是日期物件，轉換為時間字串
      if (existingTimeSlot instanceof Date) {
        const taiwanTime = new Date(existingTimeSlot.getTime() + (8 * 60 * 60 * 1000));
        const hours = taiwanTime.getUTCHours().toString().padStart(2, '0');
        const minutes = taiwanTime.getUTCMinutes().toString().padStart(2, '0');
        existingTimeSlot = `${hours}:${minutes}`;
      }
      
      if (existingData[i][9] === data.weekday && existingTimeSlot === data.timeSlot) {
        return ContentService.createTextOutput('此時段已被預約').setMimeType(ContentService.MimeType.JSON);
      }
    }

    // 生成訂單編號
    const orderDate = new Date();
    const dateStr = orderDate.toLocaleDateString('zh-TW', { timeZone: 'Asia/Taipei' }).replace(/\//g, '');
    const rowNumber = sheet.getLastRow(); // 獲取最後一行的行號
    const orderNumber = `ORD${dateStr}${String(rowNumber).padStart(3, '0')}`;
    
    const timestamp = new Date().toLocaleString('zh-TW', { timeZone: 'Asia/Taipei' });
    
    // 根據新表格結構準備資料 - 確保正確轉換資料型別
    const rowData = [
      orderNumber,                               // A: 訂單編號
      timestamp,                                 // B: 提交時間
      String(data.contactMethod || ''),          // C: 聯絡方式
      String(data.contactId || ''),              // D: 聯絡ID
      parseInt(String(data.legendCount)) || 0,   // E: 傳說對決數量
      parseInt(String(data.voiceCount)) || 0,    // F: 語音通話數量
      parseInt(String(data.chatCount)) || 0,     // G: 打字聊天數量
      parseInt(String(data.partyCount)) || 0,    // H: 全民Party唱歌數量
      parseInt(String(data.consultCount)) || 0,  // I: 情感諮詢數量
      String(data.weekday || ''),                // J: 預約星期
      String(data.timeSlot || ''),               // K: 預約時段
      String(data.fullDateTime || ''),           // L: 完整預約時間
      String(data.selectedItems || ''),          // M: 所選項目
      parseInt(String(data.subtotal)) || 0,      // N: 服務金額小計
      parseInt(String(data.total)) || 0,         // O: 最終應付金額
      '待付款',                                  // P: 付款狀態
      '待服務',                                  // Q: 服務狀態
      String(data.remark || '')                  // R: 備註
    ];

    console.log('準備寫入的資料:', rowData);
    sheet.appendRow(rowData);
    console.log('資料已寫入工作表，行數:', sheet.getLastRow());
    
    // 發送 Email 通知
    try {
      console.log('開始發送 Email 通知...');
      sendEmailNotification(orderNumber, rowData, data);
      console.log('Email 通知發送成功！');
    } catch (emailError) {
      console.error('Email 發送失敗:', emailError);
      console.error('錯誤詳情:', emailError.toString());
      // 即使 email 失敗，仍然回傳成功（資料已寫入）
    }

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

// 取得已預約時段的函數
function getBookedTimeSlots(weekday) {
  console.log('取得已預約時段:', weekday);
  try {
    // 星期數字轉文字對照表
    const weekdayMap = {
      '1': '星期一',
      '2': '星期二', 
      '3': '星期三',
      '4': '星期四',
      '5': '星期五',
      '6': '星期六',
      '7': '星期日'
    };
    
    const weekdayText = weekdayMap[weekday] || weekday;
    console.log('星期數字:', weekday, '-> 星期文字:', weekdayText);
    
    const SHEET_ID = '17iFNvOb-Gl5B1nDMC9Jf0ls1s8t_a8UK9n2kB23fEo4';
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('訂單');
    const data = sheet.getDataRange().getValues();
    const bookedSlots = [];
    
    // 從第二行開始檢查（跳過標題行）
    for (let i = 1; i < data.length; i++) {
      // J欄(索引9)是預約星期，K欄(索引10)是預約時段，P欄(索引15)是付款狀態
      if (data[i][9] === weekdayText && data[i][15] === '已付款') {
        let timeSlot = data[i][10];
        
        // 如果是日期物件，轉換為時間字串（考慮台灣時區）
        if (timeSlot instanceof Date) {
          // 使用台灣時區
          const taiwanTime = new Date(timeSlot.getTime() + (8 * 60 * 60 * 1000)); // UTC+8
          const hours = taiwanTime.getUTCHours().toString().padStart(2, '0');
          const minutes = taiwanTime.getUTCMinutes().toString().padStart(2, '0');
          timeSlot = `${hours}:${minutes}`;
        } else if (typeof timeSlot === 'string' && timeSlot.includes('T')) {
          // 如果是 ISO 字串，解析為時間（考慮台灣時區）
          const date = new Date(timeSlot);
          const taiwanTime = new Date(date.getTime() + (8 * 60 * 60 * 1000)); // UTC+8
          const hours = taiwanTime.getUTCHours().toString().padStart(2, '0');
          const minutes = taiwanTime.getUTCMinutes().toString().padStart(2, '0');
          timeSlot = `${hours}:${minutes}`;
        }
        
        bookedSlots.push(timeSlot);
        console.log('找到已付款預約:', data[i][9], data[i][10], '-> 轉換後:', timeSlot, data[i][15]);
      }
    }
    
    console.log('找到已付款的預約時段:', bookedSlots);
    return ContentService
      .createTextOutput(JSON.stringify({ bookedSlots }))
      .setMimeType(ContentService.MimeType.JSON);
      
  } catch (err) {
    console.error('取得已預約時段失敗:', err);
    return ContentService
      .createTextOutput(JSON.stringify({ error: err.message, bookedSlots: [] }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// 用於 GET 測試和接收參數
function doGet(e) {
  console.log('收到 GET 請求');
  
  // 處理取得已預約時段的請求
  if (e.parameter && e.parameter.action === 'getBookedSlots') {
    return getBookedTimeSlots(e.parameter.weekday);
  }
  
  // 如果有參數且不是 API 呼叫，處理表單資料
  if (e.parameter && Object.keys(e.parameter).length > 0 && !e.parameter.action) {
    console.log('GET 參數 (表單提交):', e.parameter);
    const result = processFormData(e.parameter);
    
    // 支援 JSONP
    if (e.parameter.callback) {
      const callback = e.parameter.callback;
      const responseText = result.getContent();
      return ContentService.createTextOutput(`${callback}("${responseText}")`)
        .setMimeType(ContentService.MimeType.JAVASCRIPT);
    }
    
    return result;
  }
  
  return ContentService.createTextOutput('Google Apps Script 運作正常').setMimeType(ContentService.MimeType.JSON);
}

// 處理 OPTIONS 請求 (CORS 預檢)
function doOptions(e) {
  return HtmlService.createHtmlOutput()
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// Email 通知函數
function sendEmailNotification(orderNumber, rowData, originalData) {
  // 設定 email 收件人
  const recipientEmail = 'cocolin0731@gmail.com';
  console.log('準備發送 Email 到:', recipientEmail);
  
  // Email 主旨
  const subject = `新訂單通知 - ${orderNumber}`;
  console.log('Email 主旨:', subject);
  
  // 建立 HTML 格式的 email 內容
  const htmlBody = `
    <div style="font-family: Arial, sans-serif; max-width: 600px; margin: 0 auto;">
      <h2 style="color: #333; border-bottom: 2px solid #ddd; padding-bottom: 10px;">新訂單通知</h2>
      
      <div style="background-color: #f9f9f9; padding: 20px; margin: 20px 0; border-radius: 8px;">
        <h3 style="color: #555; margin-top: 0;">訂單資訊</h3>
        <p><strong>訂單編號：</strong> ${orderNumber}</p>
        <p><strong>提交時間：</strong> ${rowData[1]}</p>
      </div>
      
      <div style="background-color: #f9f9f9; padding: 20px; margin: 20px 0; border-radius: 8px;">
        <h3 style="color: #555; margin-top: 0;">客戶資訊</h3>
        <p><strong>聯絡方式：</strong> ${rowData[2]}</p>
        <p><strong>聯絡ID：</strong> ${rowData[3]}</p>
      </div>
      
      <div style="background-color: #f9f9f9; padding: 20px; margin: 20px 0; border-radius: 8px;">
        <h3 style="color: #555; margin-top: 0;">預約資訊</h3>
        <p><strong>預約星期：</strong> ${rowData[9]}</p>
        <p><strong>預約時段：</strong> ${rowData[10]}</p>
        <p><strong>完整預約時間：</strong> ${rowData[11]}</p>
      </div>
      
      <div style="background-color: #f9f9f9; padding: 20px; margin: 20px 0; border-radius: 8px;">
        <h3 style="color: #555; margin-top: 0;">服務項目</h3>
        <p><strong>所選項目：</strong> ${rowData[12]}</p>
        <ul style="margin: 10px 0; padding-left: 20px;">
          ${rowData[4] > 0 ? `<li>傳說對決：${rowData[4]} 小時</li>` : ''}
          ${rowData[5] > 0 ? `<li>語音通話：${rowData[5]} 小時</li>` : ''}
          ${rowData[6] > 0 ? `<li>打字聊天：${rowData[6]} 小時</li>` : ''}
          ${rowData[7] > 0 ? `<li>全民Party唱歌：${rowData[7]} 小時</li>` : ''}
          ${rowData[8] > 0 ? `<li>情感諮詢：${rowData[8]} 小時</li>` : ''}
        </ul>
      </div>
      
      <div style="background-color: #f9f9f9; padding: 20px; margin: 20px 0; border-radius: 8px;">
        <h3 style="color: #555; margin-top: 0;">金額資訊</h3>
        <p><strong>服務金額小計：</strong> NT$ ${rowData[13]}</p>
        <p><strong>最終應付金額：</strong> <span style="color: #e74c3c; font-size: 1.2em;">NT$ ${rowData[14]}</span></p>
      </div>
      
      ${rowData[17] ? `
      <div style="background-color: #fff3cd; padding: 20px; margin: 20px 0; border-radius: 8px; border: 1px solid #ffeaa7;">
        <h3 style="color: #856404; margin-top: 0;">備註</h3>
        <p style="margin: 0;">${rowData[17]}</p>
      </div>
      ` : ''}
      
      <div style="margin-top: 30px; padding-top: 20px; border-top: 2px solid #ddd; text-align: center; color: #666;">
        <p>此為系統自動發送的通知郵件</p>
        <p><a href="https://docs.google.com/spreadsheets/d/${SpreadsheetApp.openById('17iFNvOb-Gl5B1nDMC9Jf0ls1s8t_a8UK9n2kB23fEo4').getId()}" 
              style="color: #3498db; text-decoration: none;">查看 Google Sheet</a></p>
      </div>
    </div>
  `;
  
  // 純文字版本（作為備用）
  const textBody = `
新訂單通知 - ${orderNumber}

訂單資訊：
- 訂單編號：${orderNumber}
- 提交時間：${rowData[1]}

客戶資訊：
- 聯絡方式：${rowData[2]}
- 聯絡ID：${rowData[3]}

預約資訊：
- 預約星期：${rowData[9]}
- 預約時段：${rowData[10]}
- 完整預約時間：${rowData[11]}

服務項目：${rowData[12]}
${rowData[4] > 0 ? `- 傳說對決：${rowData[4]} 小時` : ''}
${rowData[5] > 0 ? `- 語音通話：${rowData[5]} 小時` : ''}
${rowData[6] > 0 ? `- 打字聊天：${rowData[6]} 小時` : ''}
${rowData[7] > 0 ? `- 全民Party唱歌：${rowData[7]} 小時` : ''}
${rowData[8] > 0 ? `- 情感諮詢：${rowData[8]} 小時` : ''}

金額資訊：
- 服務金額小計：NT$ ${rowData[13]}
- 最終應付金額：NT$ ${rowData[14]}

${rowData[17] ? `備註：${rowData[17]}` : ''}
  `;
  
  // 發送 email
  console.log('正在發送 Email...');
  MailApp.sendEmail({
    to: recipientEmail,
    subject: subject,
    body: textBody,
    htmlBody: htmlBody
  });
  console.log('Email 已成功發送到:', recipientEmail);
}

