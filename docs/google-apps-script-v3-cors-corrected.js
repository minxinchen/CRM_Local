// ===================================================
// CRM Google Apps Script Web App API - CORS Corrected Version
// ===================================================
// 來源：Google Apps Script 官方文檔 + CORS 解決方案
// https://developers.google.com/apps-script/guides/web
// https://stackoverflow.com/questions/53433938/how-do-i-allow-a-cors-requests-in-my-google-script
// https://stackoverflow.com/questions/59686777/set-http-headers-in-google-apps-script
// 實施日期：2024-12-03
//
// 🔧 CORS 修復：
// 1. doOptions() 處理 preflight requests
// 2. 使用 .setHeaders() (複數) 設置 CORS headers
// ===================================================

// Sheet 名稱配置
const SHEET_NAMES = {
  tasks: 'tasks',
  customers: 'customers',
  cases: 'cases',
  emails: 'emails'
};

/**
 * OPTIONS 請求處理器 - CORS Preflight
 * 來源：Stack Overflow - How do I allow CORS requests in my Google script
 * https://stackoverflow.com/questions/53433938/how-do-i-allow-a-cors-requests-in-my-google-script
 *
 * 瀏覽器在發送 POST 請求前會先發送 OPTIONS 請求（preflight）
 * Google Apps Script 必須正確回應 OPTIONS 請求才能通過 CORS 檢查
 */
function doOptions(e) {
  return ContentService.createTextOutput('')
    .setMimeType(ContentService.MimeType.JSON)
    .setHeaders({
      'Access-Control-Allow-Origin': '*',
      'Access-Control-Allow-Methods': 'GET, POST, OPTIONS',
      'Access-Control-Allow-Headers': 'Content-Type',
      'Access-Control-Max-Age': '3600'
    });
}

/**
 * GET 請求處理器 - API 狀態檢查
 * 來源：Apps Script Web App 基礎範例
 * https://developers.google.com/apps-script/guides/web#request_parameters
 */
function doGet(e) {
  const response = {
    status: 'CRM API Active',
    version: '1.2-CORS-Corrected',
    timestamp: new Date().toISOString(),
    cors: 'enabled'
  };

  return ContentService.createTextOutput(JSON.stringify(response))
    .setMimeType(ContentService.MimeType.JSON)
    .setHeaders({
      'Access-Control-Allow-Origin': '*',
      'Access-Control-Allow-Methods': 'GET, POST, OPTIONS',
      'Access-Control-Allow-Headers': 'Content-Type'
    });
}

/**
 * POST 請求處理器 - CRUD 操作路由
 * 來源：Building CRUD Web App with Google Apps Script
 * https://blogmines.com/posts/build-crud-web-app-google-apps-script
 */
function doPost(e) {
  try {
    const params = JSON.parse(e.postData.contents);
    const { action, sheet, data, id } = params;

    let result;
    switch(action) {
      case 'read':
        result = readData(sheet);
        break;
      case 'create':
        result = createData(sheet, data);
        break;
      case 'update':
        result = updateData(sheet, id, data);
        break;
      case 'delete':
        result = deleteData(sheet, id);
        break;
      case 'backup':
        result = createExcelBackup(sheet);
        break;
      default:
        throw new Error('Invalid action: ' + action);
    }

    return ContentService.createTextOutput(JSON.stringify({
      success: true,
      data: result,
      timestamp: new Date().toISOString()
    }))
    .setMimeType(ContentService.MimeType.JSON)
    .setHeaders({
      'Access-Control-Allow-Origin': '*',
      'Access-Control-Allow-Methods': 'GET, POST, OPTIONS',
      'Access-Control-Allow-Headers': 'Content-Type'
    });

  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({
      success: false,
      error: error.toString(),
      timestamp: new Date().toISOString()
    }))
    .setMimeType(ContentService.MimeType.JSON)
    .setHeaders({
      'Access-Control-Allow-Origin': '*',
      'Access-Control-Allow-Methods': 'GET, POST, OPTIONS',
      'Access-Control-Allow-Headers': 'Content-Type'
    });
  }
}

/**
 * 讀取數據
 * 來源：SpreadsheetApp API 文檔
 * https://developers.google.com/apps-script/reference/spreadsheet/spreadsheet-app
 */
function readData(sheetName) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) {
    throw new Error('Sheet not found: ' + sheetName);
  }

  const range = sheet.getDataRange();
  const values = range.getValues();

  if (values.length === 0) {
    return [];
  }

  // 第一行為標題
  const headers = values[0];
  const data = [];

  // 轉換為物件陣列
  for (let i = 1; i < values.length; i++) {
    const row = {};
    for (let j = 0; j < headers.length; j++) {
      row[headers[j]] = values[i][j];
    }
    data.push(row);
  }

  return data;
}

/**
 * 新增數據
 * 來源：Google Apps Script CRUD 最佳實踐
 * https://developers.google.com/apps-script/guides/support/best-practices
 */
function createData(sheetName, data) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) {
    throw new Error('Sheet not found: ' + sheetName);
  }

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

  // 生成新 ID（如果沒有提供）
  if (!data.ID) {
    const lastRow = sheet.getLastRow();
    data.ID = lastRow > 1 ? parseInt(sheet.getRange(lastRow, 1).getValue()) + 1 : 1;
  }

  // 添加建立日期（如果有該欄位）
  if (headers.includes('建立日期') && !data['建立日期']) {
    data['建立日期'] = new Date().toISOString().split('T')[0];
  }

  // 準備數據行
  const row = headers.map(header => data[header] || '');

  // 添加到最後一行
  sheet.appendRow(row);

  return {
    success: true,
    id: data.ID,
    message: 'Data created successfully'
  };
}

/**
 * 更新數據
 * 來源：SpreadsheetApp Range 操作
 * https://developers.google.com/apps-script/reference/spreadsheet/range
 */
function updateData(sheetName, id, data) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) {
    throw new Error('Sheet not found: ' + sheetName);
  }

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const dataRange = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn());
  const values = dataRange.getValues();

  // 尋找目標行
  let targetRowIndex = -1;
  for (let i = 0; i < values.length; i++) {
    if (values[i][0].toString() === id.toString()) {
      targetRowIndex = i;
      break;
    }
  }

  if (targetRowIndex === -1) {
    throw new Error('Record not found: ID ' + id);
  }

  // 更新數據
  const updatedRow = headers.map(header => {
    return data.hasOwnProperty(header) ? data[header] : values[targetRowIndex][headers.indexOf(header)];
  });

  // 寫入更新
  sheet.getRange(targetRowIndex + 2, 1, 1, headers.length).setValues([updatedRow]);

  return {
    success: true,
    id: id,
    message: 'Data updated successfully'
  };
}

/**
 * 刪除數據
 * 來源：Google Apps Script 數據操作範例
 * https://developers.google.com/apps-script/reference/spreadsheet/sheet#deleterowrowposition
 */
function deleteData(sheetName, id) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) {
    throw new Error('Sheet not found: ' + sheetName);
  }

  const dataRange = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1);
  const ids = dataRange.getValues();

  // 尋找目標行
  for (let i = 0; i < ids.length; i++) {
    if (ids[i][0].toString() === id.toString()) {
      sheet.deleteRow(i + 2); // +2 因為標題行和索引從 1 開始
      return {
        success: true,
        id: id,
        message: 'Data deleted successfully'
      };
    }
  }

  throw new Error('Record not found: ID ' + id);
}

/**
 * 創建 Excel 備份
 * 來源：DriveApp 檔案匯出功能
 * https://developers.google.com/apps-script/reference/drive/drive-app
 */
function createExcelBackup(sheetName) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = sheetName ? spreadsheet.getSheetByName(sheetName) : null;

  // 創建備份資料夾（如果不存在）
  const folders = DriveApp.getFoldersByName('CRM_Backups');
  let backupFolder;
  if (folders.hasNext()) {
    backupFolder = folders.next();
  } else {
    backupFolder = DriveApp.createFolder('CRM_Backups');
  }

  // 生成檔案名稱
  const timestamp = new Date().toISOString().replace(/[:.]/g, '-').split('T')[0];
  const fileName = sheetName
    ? `${sheetName}_backup_${timestamp}.xlsx`
    : `CRM_full_backup_${timestamp}.xlsx`;

  // 匯出為 Excel
  const blob = spreadsheet.getAs('application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  blob.setName(fileName);

  const file = backupFolder.createFile(blob);

  return {
    success: true,
    fileName: fileName,
    fileUrl: file.getUrl(),
    message: 'Backup created successfully'
  };
}

/**
 * 自動備份觸發器（每日執行）
 * 來源：Apps Script Time-driven Triggers
 * https://developers.google.com/apps-script/guides/triggers/installable#time-driven_triggers
 *
 * 設定方式：
 * 1. 在 Apps Script 編輯器點擊「觸發條件」（時鐘圖示）
 * 2. 新增觸發條件
 * 3. 選擇函式：dailyBackup
 * 4. 事件來源：時間驅動
 * 5. 時間型觸發條件類型：日計時器
 * 6. 選擇時間：凌晨 2-3 點
 */
function dailyBackup() {
  try {
    // 備份所有工作表
    Object.values(SHEET_NAMES).forEach(sheetName => {
      createExcelBackup(sheetName);
    });

    // 清理 30 天前的備份
    cleanOldBackups(30);

    Logger.log('Daily backup completed successfully');
  } catch (error) {
    Logger.log('Daily backup failed: ' + error.toString());
  }
}

/**
 * 清理舊備份
 * 來源：DriveApp 檔案管理
 * https://developers.google.com/apps-script/reference/drive/folder#getfiles
 */
function cleanOldBackups(daysToKeep) {
  const folders = DriveApp.getFoldersByName('CRM_Backups');
  if (!folders.hasNext()) return;

  const backupFolder = folders.next();
  const files = backupFolder.getFiles();
  const cutoffDate = new Date();
  cutoffDate.setDate(cutoffDate.getDate() - daysToKeep);

  while (files.hasNext()) {
    const file = files.next();
    if (file.getDateCreated() < cutoffDate) {
      file.setTrashed(true);
      Logger.log('Deleted old backup: ' + file.getName());
    }
  }
}
