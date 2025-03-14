/**
 * マインドエンジニアリング・コーチング管理システム
 * スプレッドシート初期化スクリプト
 * 
 * このスクリプトは、システムで使用するスプレッドシートの初期構造を設定します。
 * 必要なシートとカラムを自動的に作成します。
 */

/**
 * スプレッドシートの初期設定を行う関数
 */
function setupSpreadsheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 既存のシートを取得
  const existingSheets = {};
  ss.getSheets().forEach(sheet => {
    existingSheets[sheet.getName()] = sheet;
  });
  
  // コーポレートカラー
  const corporateColor = "#c50502";
  
  // 必要なシートを作成または取得
  const clientSheet = existingSheets['クライアントinfo'] || ss.insertSheet('クライアントinfo');
  const sessionSheet = existingSheets['セッション管理'] || ss.insertSheet('セッション管理');
  const paymentSheet = existingSheets['支払い管理'] || ss.insertSheet('支払い管理');
  const emailLogSheet = existingSheets['メールログ'] || ss.insertSheet('メールログ');
  const settingsSheet = existingSheets['設定'] || ss.insertSheet('設定');
  
  // 各シートのヘッダー設定
  setupClientSheet(clientSheet, corporateColor);
  setupSessionSheet(sessionSheet, corporateColor);
  setupPaymentSheet(paymentSheet, corporateColor);
  setupEmailLogSheet(emailLogSheet, corporateColor);
  setupSettingsSheet(settingsSheet, corporateColor);
  
  // 基本設定値の登録
  registerDefaultSettings(settingsSheet);
  
  // 完了メッセージ
  SpreadsheetApp.getUi().alert('マインドエンジニアリング・コーチング管理システムの初期設定が完了しました。');
}

/**
 * クライアント情報シートを設定
 */
function setupClientSheet(sheet, color) {
  // ヘッダー行を設定
  const headers = [
    'クライアントID', 
    'タイムスタンプ', 
    'メールアドレス', 
    'お名前', 
    'お名前　（カナ）', 
    '性別', 
    '生年月日', 
    '電話番号　（ハイフンなし）', 
    'ご住所', 
    '希望セッション形式', 
    'ステータス',
    '備考欄'
  ];
  
  setSheetHeaders(sheet, headers, color);
  
  // 列の幅を調整
  sheet.setColumnWidth(1, 150); // クライアントID
  sheet.setColumnWidth(2, 180); // タイムスタンプ
  sheet.setColumnWidth(3, 220); // メールアドレス
  sheet.setColumnWidth(4, 150); // お名前
  sheet.setColumnWidth(5, 150); // お名前（カナ）
  sheet.setColumnWidth(9, 300); // ご住所
  sheet.setColumnWidth(12, 300); // 備考欄
  
  // データ検証（ドロップダウンリスト）
  // 性別列
  const genderRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['男性', '女性', 'その他', '回答しない'], true)
    .build();
  sheet.getRange('F2:F1000').setDataValidation(genderRule);
  
  // セッション形式列
  const sessionTypeRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['対面', 'オンライン', '未定'], true)
    .build();
  sheet.getRange('J2:J1000').setDataValidation(sessionTypeRule);
  
  // ステータス列
  const statusRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['問い合わせ', 'トライアル前', 'トライアル済', '契約中', '完了', '中断'], true)
    .build();
  sheet.getRange('K2:K1000').setDataValidation(statusRule);
}

/**
 * セッション管理シートを設定
 */
function setupSessionSheet(sheet, color) {
  // ヘッダー行を設定
  const headers = [
    'セッションID', 
    'クライアントID', 
    'セッション種別', 
    '予定日時', 
    'Google Meet URL', 
    'ステータス', 
    '実施日時', 
    '記録',
    '備考'
  ];
  
  setSheetHeaders(sheet, headers, color);
  
  // 列の幅を調整
  sheet.setColumnWidth(1, 150); // セッションID
  sheet.setColumnWidth(2, 150); // クライアントID
  sheet.setColumnWidth(4, 180); // 予定日時
  sheet.setColumnWidth(5, 250); // Google Meet URL
  sheet.setColumnWidth(7, 180); // 実施日時
  sheet.setColumnWidth(8, 350); // 記録
  sheet.setColumnWidth(9, 200); // 備考
  
  // データ検証（ドロップダウンリスト）
  // セッション種別列
  const sessionTypeRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['トライアル', '継続（2回目）', '継続（3回目）', '継続（4回目）', '継続（5回目）', '継続（6回目）', 'フォローアップ'], true)
    .build();
  sheet.getRange('C2:C1000').setDataValidation(sessionTypeRule);
  
  // ステータス列
  const statusRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['予定', '実施済', 'キャンセル', '延期'], true)
    .build();
  sheet.getRange('F2:F1000').setDataValidation(statusRule);
}

/**
 * 支払い管理シートを設定
 */
function setupPaymentSheet(sheet, color) {
  // ヘッダー行を設定
  const headers = [
    '支払いID', 
    'クライアントID', 
    '登録日', 
    '項目', 
    '金額', 
    '状態', 
    '入金日',
    '領収書発行状態',
    '備考'
  ];
  
  setSheetHeaders(sheet, headers, color);
  
  // 列の幅を調整
  sheet.setColumnWidth(1, 150); // 支払いID
  sheet.setColumnWidth(2, 150); // クライアントID
  sheet.setColumnWidth(3, 120); // 登録日
  sheet.setColumnWidth(4, 200); // 項目
  sheet.setColumnWidth(5, 120); // 金額
  sheet.setColumnWidth(9, 300); // 備考
  
  // 金額列の書式設定
  sheet.getRange('E2:E1000').setNumberFormat('¥#,##0');
  
  // データ検証（ドロップダウンリスト）
  // 項目列
  const paymentItemRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['トライアルセッション', '継続セッション（全5回）'], true)
    .build();
  sheet.getRange('D2:D1000').setDataValidation(paymentItemRule);
  
  // 状態列
  const statusRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['未入金', '入金済', 'キャンセル'], true)
    .build();
  sheet.getRange('F2:F1000').setDataValidation(statusRule);
  
  // 領収書発行状態列
  const receiptStatusRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['未発行', '発行済', '不要'], true)
    .build();
  sheet.getRange('H2:H1000').setDataValidation(receiptStatusRule);
}

/**
 * メールログシートを設定
 */
function setupEmailLogSheet(sheet, color) {
  // ヘッダー行を設定
  const headers = [
    'メールID', 
    '送信日時', 
    'クライアントID', 
    '送信先', 
    '件名', 
    '種類', 
    'ステータス'
  ];
  
  setSheetHeaders(sheet, headers, color);
  
  // 列の幅を調整
  sheet.setColumnWidth(1, 150); // メールID
  sheet.setColumnWidth(2, 180); // 送信日時
  sheet.setColumnWidth(3, 150); // クライアントID
  sheet.setColumnWidth(4, 220); // 送信先
  sheet.setColumnWidth(5, 300); // 件名
  
  // データ検証（ドロップダウンリスト）
  // 種類列
  const mailTypeRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['トライアル案内', 'セッション確認', 'リマインダー', '支払い案内', '領収書送付', '次回日程調整', 'その他'], true)
    .build();
  sheet.getRange('F2:F1000').setDataValidation(mailTypeRule);
  
  // ステータス列
  const statusRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['送信済', '送信失敗', '下書き'], true)
    .build();
  sheet.getRange('G2:G1000').setDataValidation(statusRule);
}

/**
 * 設定シートを設定
 */
function setupSettingsSheet(sheet, color) {
  // ヘッダー行を設定
  const headers = [
    '設定キー', 
    '設定値', 
    '説明'
  ];
  
  setSheetHeaders(sheet, headers, color);
  
  // 列の幅を調整
  sheet.setColumnWidth(1, 200); // 設定キー
  sheet.setColumnWidth(2, 300); // 設定値
  sheet.setColumnWidth(3, 400); // 説明
}

/**
 * 基本設定値を登録
 */
function registerDefaultSettings(sheet) {
  const settings = [
    ['ADMIN_EMAIL', 'mindengineeringcoaching@gmail.com', '管理者メールアドレス'],
    ['SERVICE_NAME', 'マインドエンジニアリング・コーチング', 'サービス名'],
    ['MAIL_SENDER_NAME', 'マインドエンジニアリング・コーチング 森山雄太', '送信者名'],
    ['CORPORATE_COLOR', '#c50502', 'コーポレートカラー'],
    ['TRIAL_FEE', '6000', 'トライアル料金'],
    ['CONTINUATION_FEE', '214000', '継続セッション料金'],
    ['SESSION_DURATION', '30', 'セッション時間（分）'],
    ['BUSINESS_ADDRESS', '〒790-0012 愛媛県松山市湊町2-5-2リコオビル401', '事業所住所'],
    ['BUSINESS_PHONE', '090-5710-7627', '事業所電話番号'],
    ['BANK_INFO', '店名: 六一八\n店番: 618\n普通預金\n口座番号: 1396031\n口座名義: モリヤマユウタ', '振込先情報']
  ];
  
  // 現在の行数を取得（ヘッダー行を除く）
  const currentRows = sheet.getLastRow();
  const startRow = currentRows > 1 ? currentRows + 1 : 2;
  
  // 設定値を入力
  for (let i = 0; i < settings.length; i++) {
    sheet.getRange(startRow + i, 1).setValue(settings[i][0]);
    sheet.getRange(startRow + i, 2).setValue(settings[i][1]);
    sheet.getRange(startRow + i, 3).setValue(settings[i][2]);
  }
}

/**
 * シートのヘッダー行を設定するヘルパー関数
 */
function setSheetHeaders(sheet, headers, color) {
  // ヘッダー行を設定
  for (let i = 0; i < headers.length; i++) {
    sheet.getRange(1, i + 1).setValue(headers[i]);
  }
  
  // ヘッダー行のスタイル設定
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setBackground(color);
  headerRange.setFontColor("white");
  headerRange.setFontWeight("bold");
  headerRange.setHorizontalAlignment("center");
  headerRange.setVerticalAlignment("middle");
  
  // 罫線を設定
  headerRange.setBorder(true, true, true, true, true, true);
  
  // フリーズ
  sheet.setFrozenRows(1);
  
  // フィルターを設定
  sheet.getRange(1, 1, 1, headers.length).createFilter();
}

/**
 * メニューに初期設定機能を追加
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('MEC管理システム')
    .addItem('データベース初期設定', 'setupSpreadsheet')
    .addSeparator()
    .addItem('バージョン情報', 'showVersionInfo')
    .addToUi();
}

/**
 * バージョン情報を表示
 */
function showVersionInfo() {
  const ui = SpreadsheetApp.getUi();
  ui.alert(
    'マインドエンジニアリング・コーチング管理システム',
    'バージョン: 1.0.0\n開発: 森山雄太\n\n© 2025 マインドエンジニアリング・コーチング',
    ui.ButtonSet.OK
  );
}