// ============================================================
// Code.gs - メインルーター ver2
// ============================================================

const SPREADSHEET_ID_KEY = 'SPREADSHEET_ID';

/**
 * WebApp エントリーポイント (GET)
 */
function doGet(e) {
  let page = e.parameter.page || 'login';

  // SPA統合: history は index に統合されたためリダイレクト
  if (page === 'history') page = 'index';

  const template = HtmlService.createTemplateFromFile(page);
  const output = template.evaluate()
    .setTitle('経費精算アプリ')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  return output;
}

/**
 * WebApp エントリーポイント (POST) - API ルーター
 */
function doPost(e) {
  try {
    const params = JSON.parse(e.postData.contents);
    const action = params.action;

    switch (action) {
      case 'login':          return jsonOut(login(params));
      case 'logout':         return jsonOut(logout(params));
      case 'getGroups':      return jsonOut(getGroups());
      case 'getUsers':       return jsonOut(getUsers(params));
      case 'validateSession':return jsonOut(validateSession(params));
      case 'processOCR':     return jsonOut(processOCR(params));
      case 'saveRecord':     return jsonOut(saveRecord(params));
      case 'getRecords':     return jsonOut(getRecords(params));
      case 'updateRecord':   return jsonOut(updateRecord(params));
      case 'deleteRecord':   return jsonOut(deleteRecord(params));
      case 'saveImage':      return jsonOut(saveReceiptImage(params));
      case 'getCategories':  return jsonOut(getCategories());
      case 'generateReport': return jsonOut(generateReport(params));
      default:
        return jsonOut({ success: false, error: '不明なアクション: ' + action });
    }
  } catch (err) {
    console.error('doPost error:', err);
    return jsonOut({ success: false, error: err.message });
  }
}

/**
 * JSON レスポンス生成
 */
function jsonOut(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * HTML インクルード用
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * スプレッドシート取得
 */
function getSpreadsheet() {
  const id = PropertiesService.getScriptProperties().getProperty(SPREADSHEET_ID_KEY);
  if (!id) throw new Error('SPREADSHEET_ID が設定されていません。initialSetup() を実行してください。');
  return SpreadsheetApp.openById(id);
}

/**
 * settings シートから値取得
 */
function getSettings(key) {
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName('settings');
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === key) return data[i][1];
  }
  return null;
}

/**
 * 初期セットアップ（初回のみ手動実行）
 */
function initialSetup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const props = PropertiesService.getScriptProperties();
  props.setProperty(SPREADSHEET_ID_KEY, ss.getId());
  console.log('✅ SPREADSHEET_ID 設定: ' + ss.getId());

  const rootFolder = DriveApp.createFolder('expense_app_root');
  const receiptFolder = rootFolder.createFolder('receipts');
  const exportsFolder = rootFolder.createFolder('exports');
  console.log('✅ Drive フォルダ作成 | root: ' + rootFolder.getId());

  const settingsSheet = ss.getSheetByName('settings');
  const data = settingsSheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === 'root_folder_id') {
      settingsSheet.getRange(i + 1, 2).setValue(rootFolder.getId());
      break;
    }
  }
  console.log('✅ Setup完了 | FolderID: ' + rootFolder.getId());
  console.log('▶ 次のステップ: settings シートの vision_api_key にAPIキーを入力してください');
}

/**
 * パスワードハッシュ生成ユーティリティ（ユーザー登録時に使用）
 */
function generatePasswordHash(password) {
  const rawHash = password || 'your_password_here';
  const bytes = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    rawHash,
    Utilities.Charset.UTF_8
  );
  const hash = bytes.map(b => ('0' + (b & 0xFF).toString(16)).slice(-2)).join('');
  console.log('Password: ' + rawHash + ' → Hash: ' + hash);
  return hash;
}