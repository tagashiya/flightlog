// ============================================================
// FlightLog - Google Apps Script
// ============================================================

const SHEET_NAME = '飛行記録';
const HEADERS = [
  'ID','飛行年月日','離陸時刻','操縦者氏名','機体','登録記号',
  '飛行概要','離陸場所','着陸場所','飛行時間(分)','安全影響事項','詳細','安全影響なし','保存日時'
];

function doGet(e) {
  const action = e.parameter.action || '';
  let result;
  try {
    switch (action) {
      case 'setup':  result = setupSpreadsheet(e.parameter.folderName || ''); break;
      case 'getAll': result = getAllLogs(); break;
      default:       result = { success: true, message: 'FlightLog API is running' };
    }
  } catch(err) {
    result = { success: false, message: err.message };
  }
  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  let result;
  try {
    const params = JSON.parse(e.postData.contents);
    switch (params.action) {
      case 'setup':       result = setupSpreadsheet(params.folderName || ''); break;
      case 'reset':       result = resetSpreadsheet(params.folderName || ''); break;
      case 'saveProfile': result = saveProfile(params.profile);               break;
      case 'getProfile':  result = getProfile();                              break;
      case 'save':        result = saveLog(params.data);                      break;
      case 'getAll':      result = getAllLogs();                               break;
      case 'delete':      result = deleteLog(params.id);                      break;
      default:       result = { success: false, message: '不明なアクションです' };
    }
  } catch(err) {
    result = { success: false, message: err.message };
  }
  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

// ============================================================
// プロフィール（操縦者情報）の保存・取得
// ============================================================
function saveProfile(profile) {
  try {
    const props = PropertiesService.getUserProperties();
    props.setProperty('PILOT', profile.pilot || '');
    props.setProperty('AIRCRAFT', profile.aircraft || '');
    props.setProperty('REGISTRATION', profile.registration || '');
    props.setProperty('FOLDER', profile.folder || '');
    return { success: true, message: 'プロフィールを保存しました' };
  } catch(err) {
    return { success: false, message: 'プロフィール保存エラー: ' + err.message };
  }
}

function getProfile() {
  try {
    const props = PropertiesService.getUserProperties();
    const spreadsheetId = props.getProperty('SPREADSHEET_ID');
    let spreadsheetUrl = '';
    if (spreadsheetId) {
      try { spreadsheetUrl = SpreadsheetApp.openById(spreadsheetId).getUrl(); } catch(e) {}
    }
    return {
      success: true,
      profile: {
        pilot:        props.getProperty('PILOT') || '',
        aircraft:     props.getProperty('AIRCRAFT') || '',
        registration: props.getProperty('REGISTRATION') || '',
        folder:       props.getProperty('FOLDER') || '',
        spreadsheetUrl: spreadsheetUrl
      }
    };
  } catch(err) {
    return { success: false, message: 'プロフィール取得エラー: ' + err.message };
  }
}

// ============================================================
// パス形式（例：002_tagashiya/飛行記録）でフォルダを取得する
// 存在しない階層は自動で作成する
// ============================================================
function getFolderByPath(folderPath) {
  const parts = folderPath.split('/').map(p => p.trim()).filter(p => p !== '');
  let current = DriveApp.getRootFolder();
  for (const part of parts) {
    const folders = current.getFoldersByName(part);
    if (folders.hasNext()) {
      current = folders.next();
    } else {
      current = current.createFolder(part);
    }
  }
  return current;
}

// ============================================================
// スプレッドシートIDをリセットして新規作成
// ============================================================
function resetSpreadsheet(folderPath) {
  PropertiesService.getUserProperties().deleteProperty('SPREADSHEET_ID');
  return setupSpreadsheet(folderPath);
}

// ============================================================
// スプレッドシートセットアップ
// ============================================================
function setupSpreadsheet(folderPath) {
  try {
    const existing = PropertiesService.getUserProperties().getProperty('SPREADSHEET_ID');
    if (existing) {
      try {
        const ss = SpreadsheetApp.openById(existing);
        return { success: true, message: 'すでに設定済みです', spreadsheetId: existing, spreadsheetUrl: ss.getUrl() };
      } catch(e) {
        PropertiesService.getUserProperties().deleteProperty('SPREADSHEET_ID');
      }
    }

    const ss = SpreadsheetApp.create('FlightLog - 飛行記録');
    const spreadsheetId = ss.getId();
    const sheet = ss.getActiveSheet();
    sheet.setName(SHEET_NAME);
    sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);

    const headerRange = sheet.getRange(1, 1, 1, HEADERS.length);
    headerRange.setBackground('#1a2235');
    headerRange.setFontColor('#06b6d4');
    headerRange.setFontWeight('bold');

    const widths = [180,110,80,100,130,130,200,200,200,100,200,200,110,160];
    widths.forEach((w, i) => sheet.setColumnWidth(i + 1, w));
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, HEADERS.length).createFilter();

    // パス形式でフォルダを指定できる
    if (folderPath && folderPath.trim() !== '') {
      const file = DriveApp.getFileById(spreadsheetId);
      const targetFolder = getFolderByPath(folderPath.trim());
      targetFolder.addFile(file);
      DriveApp.getRootFolder().removeFile(file);
    }

    PropertiesService.getUserProperties().setProperty('SPREADSHEET_ID', spreadsheetId);
    return { success: true, message: 'セットアップ完了', spreadsheetId: spreadsheetId, spreadsheetUrl: ss.getUrl() };

  } catch(err) {
    return { success: false, message: 'セットアップエラー: ' + err.message };
  }
}

// ============================================================
// 飛行記録保存
// ============================================================
function saveLog(data) {
  try {
    const ss = getSpreadsheet();
    if (!ss) return { success: false, message: '初期設定が完了していません' };
    const sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) return { success: false, message: 'シートが見つかりません' };

    const lastRow = sheet.getLastRow() + 1;
    const values = [
      String(data.id), String(data.flightDate), String(data.takeoffTime), String(data.pilotName),
      String(data.aircraft), String(data.registration), String(data.flightOverview),
      String(data.takeoffLocation), String(data.landingLocation), Number(data.flightMinutes),
      String((data.safetyItems || []).join(' / ')), String(data.safetyDetail || ''),
      data.noIssue ? '○' : '', String(data.savedAt)
    ];
    const range = sheet.getRange(lastRow, 1, 1, values.length);
    range.setNumberFormat('@');
    range.setValues([values]);
    // 飛行時間のみ数値形式に戻す
    sheet.getRange(lastRow, 10).setNumberFormat('0');

    return { success: true, message: '保存しました', id: data.id };
  } catch(err) {
    return { success: false, message: '保存エラー: ' + err.message };
  }
}

// ============================================================
// 全記録取得
// ============================================================
function getAllLogs() {
  try {
    const ss = getSpreadsheet();
    if (!ss) return { success: false, message: '初期設定が完了していません' };
    const sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) return { success: false, message: 'シートが見つかりません' };

    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return { success: true, logs: [] };

    const data = sheet.getRange(2, 1, lastRow - 1, HEADERS.length).getValues();
    const logs = data
      .filter(row => row[0] !== '')
      .map(row => ({
        id: row[0], flightDate: row[1], takeoffTime: row[2],
        pilotName: row[3], aircraft: row[4], registration: row[5],
        flightOverview: row[6], takeoffLocation: row[7], landingLocation: row[8],
        flightMinutes: row[9],
        safetyItems: row[10] ? row[10].split(' / ').filter(s => s) : [],
        safetyDetail: row[11], noIssue: row[12] === '○', savedAt: row[13]
      }));

    logs.sort((a, b) => new Date(b.savedAt) - new Date(a.savedAt));
    return { success: true, logs: logs };
  } catch(err) {
    return { success: false, message: '取得エラー: ' + err.message };
  }
}

// ============================================================
// 記録削除
// ============================================================
function deleteLog(id) {
  try {
    const ss = getSpreadsheet();
    if (!ss) return { success: false, message: '初期設定が完了していません' };
    const sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) return { success: false, message: 'シートが見つかりません' };

    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return { success: false, message: '記録が見つかりません' };

    const data = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    for (let i = 0; i < data.length; i++) {
      if (String(data[i][0]) === String(id)) {
        sheet.deleteRow(i + 2);
        return { success: true, message: '削除しました' };
      }
    }
    return { success: false, message: '該当する記録が見つかりません' };
  } catch(err) {
    return { success: false, message: '削除エラー: ' + err.message };
  }
}

// ============================================================
// スプレッドシート取得
// ============================================================
function getSpreadsheet() {
  const id = PropertiesService.getUserProperties().getProperty('SPREADSHEET_ID');
  if (!id) return null;
  try { return SpreadsheetApp.openById(id); } catch(e) { return null; }
}
