/**
 * マインドエンジニアリング・コーチング管理システム
 * ユーティリティ関数モジュール
 * 
 * 様々なモジュールで共通して使用される便利な関数を提供します。
 */

/**
 * 一意のIDを生成する
 * @param {string} prefix - IDの接頭辞（例: 'CL'=クライアント, 'SS'=セッション, 'PY'=支払い）
 * @return {string} 生成されたID
 */
function generateUniqueId(prefix) {
  const timestamp = new Date().getTime().toString().slice(-9);
  const random = Math.floor(Math.random() * 10000).toString().padStart(4, '0');
  return `${prefix}${timestamp}${random}`;
}

/**
 * 日付を'YYYY-MM-DD'形式にフォーマットする
 * @param {Date} date - フォーマットする日付
 * @return {string} フォーマットされた日付文字列
 */
function formatDate(date) {
  const d = new Date(date);
  const year = d.getFullYear();
  const month = (d.getMonth() + 1).toString().padStart(2, '0');
  const day = d.getDate().toString().padStart(2, '0');
  return `${year}-${month}-${day}`;
}

/**
 * 日時を'YYYY-MM-DD HH:MM'形式にフォーマットする
 * @param {Date} dateTime - フォーマットする日時
 * @return {string} フォーマットされた日時文字列
 */
function formatDateTime(dateTime) {
  const d = new Date(dateTime);
  const formattedDate = formatDate(d);
  const hours = d.getHours().toString().padStart(2, '0');
  const minutes = d.getMinutes().toString().padStart(2, '0');
  return `${formattedDate} ${hours}:${minutes}`;
}

/**
 * 指定したシートの最終行番号を取得する
 * @param {Sheet} sheet - Google Sheetsのシートオブジェクト
 * @param {number} columnToCheck - チェックする列番号（デフォルト: 1）
 * @return {number} 最終行番号
 */
function getLastRow(sheet, columnToCheck = 1) {
  const values = sheet.getRange(1, columnToCheck, sheet.getLastRow(), 1).getValues();
  for (let i = values.length - 1; i >= 0; i--) {
    if (values[i][0] !== '') {
      return i + 1;
    }
  }
  return 1; // ヘッダー行のみの場合
}

/**
 * システム設定を取得する
 * @param {string} key - 設定キー
 * @param {*} defaultValue - キーが見つからない場合のデフォルト値
 * @return {*} 設定値
 */
function getSetting(key, defaultValue) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const settingsSheet = ss.getSheetByName('設定');
  
  if (!settingsSheet) {
    console.error('設定シートが見つかりません');
    return defaultValue;
  }
  
  const dataRange = settingsSheet.getDataRange();
  const values = dataRange.getValues();
  
  for (let i = 1; i < values.length; i++) { // ヘッダー行をスキップ
    if (values[i][0] === key) {
      return values[i][1];
    }
  }
  
  return defaultValue;
}

/**
 * システム設定を更新する
 * @param {string} key - 設定キー
 * @param {*} value - 新しい設定値
 * @return {boolean} 更新が成功したかどうか
 */
function updateSetting(key, value) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const settingsSheet = ss.getSheetByName('設定');
  
  if (!settingsSheet) {
    console.error('設定シートが見つかりません');
    return false;
  }
  
  const dataRange = settingsSheet.getDataRange();
  const values = dataRange.getValues();
  
  for (let i = 1; i < values.length; i++) { // ヘッダー行をスキップ
    if (values[i][0] === key) {
      settingsSheet.getRange(i + 1, 2).setValue(value);
      return true;
    }
  }
  
  // キーが存在しない場合、新しい行を追加
  const lastRow = getLastRow(settingsSheet);
  settingsSheet.getRange(lastRow + 1, 1).setValue(key);
  settingsSheet.getRange(lastRow + 1, 2).setValue(value);
  return true;
}

/**
 * 二つの日付の間の月数を計算する
 * @param {Date} startDate - 開始日
 * @param {Date} endDate - 終了日
 * @return {number} 月数（小数点以下も含む）
 */
function monthsBetween(startDate, endDate) {
  const start = new Date(startDate);
  const end = new Date(endDate);
  const yearDiff = end.getFullYear() - start.getFullYear();
  const monthDiff = end.getMonth() - start.getMonth();
  const dayDiff = end.getDate() - start.getDate();
  
  let result = yearDiff * 12 + monthDiff;
  if (dayDiff < 0) {
    result -= 1;
  } else if (dayDiff > 0) {
    const daysInMonth = new Date(end.getFullYear(), end.getMonth() + 1, 0).getDate();
    result += dayDiff / daysInMonth;
  }
  
  return result;
}

/**
 * スプレッドシートのヘッダー行からカラムのインデックスを取得する
 * @param {Sheet} sheet - Google Sheetsのシートオブジェクト
 * @param {string} columnName - カラム名
 * @return {number} カラムのインデックス（0から始まる）。見つからない場合は-1
 */
function getColumnIndex(sheet, columnName) {
  const headerRange = sheet.getRange(1, 1, 1, sheet.getLastColumn());
  const headerValues = headerRange.getValues()[0];
  
  for (let i = 0; i < headerValues.length; i++) {
    if (headerValues[i] === columnName) {
      return i;
    }
  }
  
  return -1; // カラムが見つからない場合
}

/**
 * カラーコードを取得する（デフォルトはコーポレートカラー）
 * @return {string} カラーコード
 */
function getColorCode() {
  return getSetting('CORPORATE_COLOR', '#c50502');
}

/**
 * 指定したテキストをクリーンアップする（スペース削除、改行削除など）
 * @param {string} text - クリーンアップするテキスト
 * @param {boolean} trimSpaces - スペースをトリムするかどうか（デフォルト: true）
 * @param {boolean} removeNewlines - 改行を削除するかどうか（デフォルト: false）
 * @return {string} クリーンアップされたテキスト
 */
function cleanupText(text, trimSpaces = true, removeNewlines = false) {
  if (!text) return '';
  
  let result = String(text);
  
  if (trimSpaces) {
    result = result.trim();
  }
  
  if (removeNewlines) {
    result = result.replace(/[\r\n]+/g, ' ');
  }
  
  return result;
}

/**
 * メールアドレスが有効かどうかを検証する
 * @param {string} email - 検証するメールアドレス
 * @return {boolean} 有効なメールアドレスかどうか
 */
function isValidEmail(email) {
  const emailRegex = /^[a-zA-Z0-9.!#$%&'*+/=?^_`{|}~-]+@[a-zA-Z0-9-]+(?:\.[a-zA-Z0-9-]+)*$/;
  return emailRegex.test(email);
}

/**
 * 電話番号が有効かどうかを検証する（日本の電話番号形式）
 * @param {string} phoneNumber - 検証する電話番号
 * @return {boolean} 有効な電話番号かどうか
 */
function isValidPhoneNumber(phoneNumber) {
  const cleanPhone = String(phoneNumber).replace(/[- ]/g, '');
  const phoneRegex = /^(0[0-9]{9,10})$/;
  return phoneRegex.test(cleanPhone);
}

/**
 * スプレッドシートの行を取得し、オブジェクトに変換する
 * @param {Sheet} sheet - Google Sheetsのシートオブジェクト
 * @param {number} rowIndex - 行番号（1から始まる）
 * @return {Object} 行データのオブジェクト。キーはヘッダー名
 */
function getRowAsObject(sheet, rowIndex) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const rowData = sheet.getRange(rowIndex, 1, 1, sheet.getLastColumn()).getValues()[0];
  
  const result = {};
  for (let i = 0; i < headers.length; i++) {
    if (headers[i]) { // 空のヘッダーはスキップ
      result[headers[i]] = rowData[i];
    }
  }
  
  return result;
}

/**
 * CSVファイルをGoogleドライブから読み込み、スプレッドシートの特定のシートにインポートする
 * @param {string} fileId - GoogleドライブファイルID
 * @param {string} sheetName - インポート先のシート名
 * @param {boolean} hasHeader - CSVファイルにヘッダー行があるかどうか（デフォルト: true）
 * @return {boolean} インポートが成功したかどうか
 */
function importCsvToSheet(fileId, sheetName, hasHeader = true) {
  try {
    const file = DriveApp.getFileById(fileId);
    const content = file.getBlob().getDataAsString();
    const csvData = Utilities.parseCsv(content);
    
    if (csvData.length === 0) {
      console.error('CSVファイルにデータがありません');
      return false;
    }
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(sheetName);
    
    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
    }
    
    // 既存のデータをクリア
    sheet.clear();
    
    // データをシートに書き込む
    sheet.getRange(1, 1, csvData.length, csvData[0].length).setValues(csvData);
    
    if (hasHeader) {
      // ヘッダー行のスタイルを設定
      const headerRange = sheet.getRange(1, 1, 1, csvData[0].length);
      headerRange.setBackground(getColorCode());
      headerRange.setFontColor("white");
      headerRange.setFontWeight("bold");
      
      // フリーズとフィルタを設定
      sheet.setFrozenRows(1);
      headerRange.createFilter();
    }
    
    return true;
  } catch (error) {
    console.error('CSVインポートエラー:', error);
    return false;
  }
}

// モジュールをエクスポート
const Utilities = {
  generateUniqueId,
  formatDate,
  formatDateTime,
  getLastRow,
  getSetting,
  updateSetting,
  monthsBetween,
  getColumnIndex,
  getColorCode,
  cleanupText,
  isValidEmail,
  isValidPhoneNumber,
  getRowAsObject,
  importCsvToSheet
};
