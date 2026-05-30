// ============================================================
// FlightLog - Google Apps Script
// DJI Mini 5 Pro 飛行記録アプリ
// ============================================================

const SHEET_NAME = '飛行記録';
const HEADERS = [
  'ID',
  '飛行年月日',
  '離陸時刻',
  '操縦者氏名',
  '機体',
  '登録記号',
  '飛行概要',
  '離陸場所',
  '着陸場所',
  '飛行時間(分)',
  '安全影響事項',
  '詳細',
  '安全影響なし',
  '保存日時'
];

// ============================================================
// メインエントリーポイント（HTMLからのリクエストを受け取る）
// ============================================================
function doPost(e) {
  try {
    const params = JSON.parse(e.postData.contents);
    const action = params.action;

    let result;
    switch (action) {
      case 'setup':   result = setupSpreadsheet(params.folderName); break;
      case 'save':    result = saveLog(params.data);                break;
      case 'getAll':  result = getAllLogs();                        break;
      case 'delete':  result = deleteLog(params.id);               break;
      default:        result = { success: false, message: '不明なアクションです' };
    }

    return ContentService
      .createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ success: false, message: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ============================================================
// 初期設定：スプレッドシートを作成する
// folderName: 保存先フォルダ名（未指定の場合はマイドライブ直下）
// ============================================================
function setupSpreadsheet(folderName) {
  try {
    // すでに設定済みかチェック
    const existing = PropertiesService.getUserProperties().getProperty('SPREADSHEET_ID');
    if (existing) {
      try {
        const ss = SpreadsheetApp.openById(existing);
        return { success: true, message: 'すでに設定済みです', spreadsheetId: existing, spreadsheetUrl: ss.getUrl() };
      } catch(e) {
        // 見つからない場合は再作成
        PropertiesService.getUserProperties().deleteProperty('SPREADSHEET_ID');
      }
    }

    // スプレッドシート作成
    const ss = SpreadsheetApp.create('FlightLog - 飛行記録');
    const spreadsheetId = ss.getId();
    const sheet = ss.getActiveSheet();
    sheet.setName(SHEET_NAME);

    // ヘッダー行を設定
    sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);

    // ヘッダー行のスタイル設定
    const headerRange = sheet.getRange(1, 1, 1, HEADERS.length);
    headerRange.setBackground('#1a2235');
    headerRange.setFontColor('#06b6d4');
    headerRange.setFontWeight('bold');
    headerRange.setFontSize(11);

    // 列幅の設定
    sheet.setColumnWidth(1, 180);  // ID
    sheet.setColumnWidth(2, 110);  // 飛行年月日
    sheet.setColumnWidth(3, 80);   // 離陸時刻
    sheet.setColumnWidth(4, 100);  // 操縦者氏名
    sheet.setColumnWidth(5, 130);  // 機体
    sheet.setColumnWidth(6, 130);  // 登録記号
    sheet.setColumnWidth(7, 200);  // 飛行概要
    sheet.setColumnWidth(8, 200);  // 離陸場所
    sheet.setColumnWidth(9, 200);  // 着陸場所
    sheet.setColumnWidth(10, 100); // 飛行時間
    sheet.setColumnWidth(11, 200); // 安全影響事項
    sheet.setColumnWidth(12, 200); // 詳細
    sheet.setColumnWidth(13, 110); // 安全影響なし
    sheet.setColumnWidth(14, 160); // 保存日時

    // 行の高さ
    sheet.setDefaultRowHeight(24);

    // フィルター設定
    sheet.getRange(1, 1, 1, HEADERS.length).createFilter();

    // ウィンドウ枠の固定（ヘッダー行）
    sheet.setFrozenRows(1);

    // 保存先フォルダに移動
    if (folderName && folderName.trim() !== '') {
      const file = DriveApp.getFileById(spreadsheetId);
      const folders = DriveApp.getFoldersByName(folderName.trim());
      if (folders.hasNext()) {
        const folder = folders.next();
        folder.addFile(file);
        DriveApp.getRootFolder().removeFile(file);
      } else {
        // フォルダが存在しない場合は新規作成
        const newFolder = DriveApp.createFolder(folderName.trim());
        newFolder.addFile(file);
        DriveApp.getRootFolder().removeFile(file);
      }
    }

    // スプレッドシートIDを保存
    PropertiesService.getUserProperties().setProperty('SPREADSHEET_ID', spreadsheetId);

    return {
      success: true,
      message: 'セットアップ完了',
      spreadsheetId: spreadsheetId,
      spreadsheetUrl: ss.getUrl()
    };

  } catch (err) {
    return { success: false, message: 'セットアップエラー: ' + err.message };
  }
}

// ============================================================
// 飛行記録を保存する
// ============================================================
function saveLog(data) {
  try {
    const ss = getSpreadsheet();
    if (!ss) return { success: false, message: '初期設定が完了していません' };

    const sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) return { success: false, message: 'シートが見つかりません' };

    const row = [
      data.id,
      data.flightDate,
      data.takeoffTime,
      data.pilotName,
      data.aircraft,
      data.registration,
      data.flightOverview,
      data.takeoffLocation,
      data.landingLocation,
      data.flightMinutes,
      (data.safetyItems || []).join(' / '),
      data.safetyDetail || '',
      data.noIssue ? '○' : '',
      data.savedAt
    ];

    sheet.appendRow(row);

    // 最新行のスタイル設定（交互背景色）
    const lastRow = sheet.getLastRow();
    if (lastRow % 2 === 0) {
      sheet.getRange(lastRow, 1, 1, HEADERS.length).setBackground('#1a2235');
    } else {
      sheet.getRange(lastRow, 1, 1, HEADERS.length).setBackground('#111827');
    }
    sheet.getRange(lastRow, 1, 1, HEADERS.length).setFontColor('#e2e8f0');

    // 安全影響ありの行は色を変える
    if (data.safetyItems && data.safetyItems.length > 0) {
      sheet.getRange(lastRow, 11).setFontColor('#ef4444');
    }

    return { success: true, message: '保存しました', id: data.id };

  } catch (err) {
    return { success: false, message: '保存エラー: ' + err.message };
  }
}

// ============================================================
// 全飛行記録を取得する
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
        id:             row[0],
        flightDate:     row[1],
        takeoffTime:    row[2],
        pilotName:      row[3],
        aircraft:       row[4],
        registration:   row[5],
        flightOverview: row[6],
        takeoffLocation:row[7],
        landingLocation:row[8],
        flightMinutes:  row[9],
        safetyItems:    row[10] ? row[10].split(' / ').filter(s => s) : [],
        safetyDetail:   row[11],
        noIssue:        row[12] === '○',
        savedAt:        row[13]
      }));

    // 新しい順に並べ替え
    logs.sort((a, b) => new Date(b.savedAt) - new Date(a.savedAt));

    return { success: true, logs: logs };

  } catch (err) {
    return { success: false, message: '取得エラー: ' + err.message };
  }
}

// ============================================================
// 飛行記録を削除する（IDで特定）
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
    let targetRow = -1;
    for (let i = 0; i < data.length; i++) {
      if (String(data[i][0]) === String(id)) {
        targetRow = i + 2;
        break;
      }
    }

    if (targetRow === -1) return { success: false, message: '該当する記録が見つかりません' };

    sheet.deleteRow(targetRow);
    return { success: true, message: '削除しました' };

  } catch (err) {
    return { success: false, message: '削除エラー: ' + err.message };
  }
}

// ============================================================
// ユーティリティ：スプレッドシートを取得する
// ============================================================
function getSpreadsheet() {
  const id = PropertiesService.getUserProperties().getProperty('SPREADSHEET_ID');
  if (!id) return null;
  try {
    return SpreadsheetApp.openById(id);
  } catch (e) {
    return null;
  }
}
