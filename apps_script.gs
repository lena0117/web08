// Apps Script 版本：支援
// 1. 讀取課程資料（doGet）
// 2. 接收報名資料並寫入 Google 試算表（doPost）

const SHEET_CONFIG = {
  courses: {
    name: '課程資料',
    headers: ['課程編號', '課程名稱', '分類', '課程簡介', '價格', '課程圖片網址', '上架狀態']
  },
  registrations: {
    name: '報名資料',
    headers: ['時間', '姓名', '信箱', '課程名稱', '課程編號', '價格', '總金額']
  }
};

function createJsonOutput(payload) {
  const output = ContentService.createTextOutput(JSON.stringify(payload));
  output.setMimeType(ContentService.MimeType.JSON);
  output.setHeader('Access-Control-Allow-Origin', '*');
  output.setHeader('Access-Control-Allow-Methods', 'GET, POST, OPTIONS');
  output.setHeader('Access-Control-Allow-Headers', 'Content-Type');
  return output;
}

function doOptions(e) {
  const output = ContentService.createTextOutput('');
  output.setMimeType(ContentService.MimeType.TEXT);
  output.setHeader('Access-Control-Allow-Origin', '*');
  output.setHeader('Access-Control-Allow-Methods', 'GET, POST, OPTIONS');
  output.setHeader('Access-Control-Allow-Headers', 'Content-Type');
  return output;
}

function getSpreadsheet() {
  const spreadsheetId = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  if (!spreadsheetId) {
    throw new Error('請先在 Apps Script 的「腳本屬性」中新增 SPREADSHEET_ID，並填入 Google 試算表 ID。');
  }
  return SpreadsheetApp.openById(spreadsheetId);
}

function getOrCreateSheet(spreadsheet, sheetConfig) {
  let sheet = spreadsheet.getSheetByName(sheetConfig.name);
  if (!sheet) {
    sheet = spreadsheet.insertSheet(sheetConfig.name);
  }

  if (sheet.getLastRow() === 0) {
    sheet.appendRow(sheetConfig.headers);
  }

  return sheet;
}

function doGet(e) {
  try {
    const spreadsheet = getSpreadsheet();
    const sheet = getOrCreateSheet(spreadsheet, SHEET_CONFIG.courses);

    const lastRow = sheet.getLastRow();
    const lastColumn = sheet.getLastColumn();
    const values = lastRow > 1
      ? sheet.getRange(2, 1, lastRow - 1, lastColumn).getValues()
      : [];

    const courses = values
      .filter(row => row.some(cell => cell !== ''))
      .map(row => ({
        id: row[0] || '',
        name: row[1] || '',
        category: row[2] || '',
        description: row[3] || '',
        price: Number(row[4]) || 0,
        imageUrl: row[5] || '',
        status: row[6] || ''
      }));

    return createJsonOutput({ value: courses, Count: courses.length });
  } catch (err) {
    return createJsonOutput({ error: err.message });
  }
}

function doPost(e) {
  try {
    const spreadsheet = getSpreadsheet();
    const sheet = getOrCreateSheet(spreadsheet, SHEET_CONFIG.registrations);

    const rawData = e.postData && e.postData.getDataAsString()
      ? e.postData.getDataAsString()
      : '{}';
    const payload = JSON.parse(rawData);

    const customerName = payload.customerName || '';
    const customerEmail = payload.customerEmail || '';
    const courses = Array.isArray(payload.courses) ? payload.courses : [];
    const total = Number(payload.total || 0);
    const timestamp = new Date().toLocaleString('zh-TW', {
      timeZone: 'Asia/Taipei'
    });

    courses.forEach(course => {
      sheet.appendRow([
        timestamp,
        customerName,
        customerEmail,
        course.name || '',
        course.id || '',
        course.price || '',
        total
      ]);
    });

    return createJsonOutput({ success: true, message: '報名資料已成功寫入試算表' });
  } catch (err) {
    return createJsonOutput({ success: false, error: err.message });
  }
}
