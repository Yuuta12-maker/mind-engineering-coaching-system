/**
 * マインドエンジニアリング・コーチング業務管理システム
 * 設定ファイル
 * 
 * システム全体で使用する設定値を一箇所にまとめています。
 * 何か変更が必要な場合は、このファイルを修正するだけで全体に反映されます。
 */

// スプレッドシートの情報
const CONFIG = {
  // このシステムが使用するスプレッドシートのID
  // URLの「https://docs.google.com/spreadsheets/d/【ここの部分】/edit」がスプレッドシートIDです
  SPREADSHEET_ID: 'あなたのスプレッドシートIDを入力してください',
  
  // 各シート名の設定
  SHEET_NAMES: {
    CLIENT: 'クライアントinfo',
    SESSION: 'セッション管理',
    PAYMENT: '支払い管理',
    EMAIL_LOG: 'メールログ',
    SETTINGS: '設定'
  },
  
  // 料金設定
  FEES: {
    TRIAL: 6000,         // トライアルセッション料金（税込）
    CONTINUATION: 214000 // 継続セッション料金（税込）
  },
  
  // デザイン設定
  DESIGN: {
    CORPORATE_COLOR: '#c50502', // コーポレートカラー
    FONT_FAMILY: 'Noto Sans JP, sans-serif' // 使用フォント
  },
  
  // サービス名と管理者情報
  SERVICE_NAME: 'マインドエンジニアリング・コーチング',
  ADMIN_NAME: '森山雄太',
  ADMIN_EMAIL: 'あなたのメールアドレスを入力してください',
  
  // セッション設定
  SESSION: {
    DURATION_MINUTES: 30, // セッション時間（分）
    REMINDER_DAYS_BEFORE: 1 // リマインダーを送る日数（セッション前）
  },
  
  // メールテンプレート
  MAIL_TEMPLATES: {
    // トライアル予約確認メール
    TRIAL_CONFIRMATION: {
      SUBJECT: '【{{SERVICE_NAME}}】トライアルセッションのご予約確認',
      BODY: `{{CLIENT_NAME}} 様

{{SERVICE_NAME}}の森山雄太です。
トライアルセッションのご予約ありがとうございます。

以下の日程でセッションを実施いたします。
日時：{{SESSION_DATE}}
形式：オンライン（GoogleMeet）
URL：{{MEET_URL}}

当日はこちらのURLからアクセスしてください。
ご不明点がございましたら、お気軽にご連絡ください。

森山雄太
{{SERVICE_NAME}}
`
    },
    
    // 支払い確認メール
    PAYMENT_CONFIRMATION: {
      SUBJECT: '【{{SERVICE_NAME}}】ご入金確認のお知らせ',
      BODY: `{{CLIENT_NAME}} 様

{{SERVICE_NAME}}の森山雄太です。
ご入金を確認いたしました。ありがとうございます。

続いてのセッションは以下の日程です。
日時：{{NEXT_SESSION_DATE}}
形式：オンライン（GoogleMeet）
URL：{{MEET_URL}}

今後ともよろしくお願いいたします。

森山雄太
{{SERVICE_NAME}}
`
    }
    // 他のメールテンプレートも必要に応じて追加できます
  }
};

/**
 * 設定値を取得する関数
 * 他のファイルからこの関数を呼び出して設定値を取得できます
 */
function getConfig() {
  return CONFIG;
}

/**
 * スプレッドシートオブジェクトを取得する関数
 */
function getSpreadsheet() {
  try {
    // 設定されたIDがあれば、そのスプレッドシートを開く
    if (CONFIG.SPREADSHEET_ID && CONFIG.SPREADSHEET_ID !== 'あなたのスプレッドシートIDを入力してください') {
      return SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    }
    // IDが設定されていない場合は、アクティブなスプレッドシートを使用
    else {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      if (ss) {
        logMessage("スプレッドシートIDが設定されていないため、現在アクティブなスプレッドシートを使用します");
      } else {
        throw new Error("アクティブなスプレッドシートが見つかりません。スプレッドシートを開いた状態で実行してください。");
      }
      return ss;
    }
  } catch (error) {
    logMessage(`スプレッドシートを取得できませんでした: ${error.message}`, 'ERROR');
    throw new Error(`スプレッドシートを取得できませんでした: ${error.message}`);
  }
}

/**
 * 指定したシートを取得する関数
 */
function getSheet(sheetName) {
  const ss = getSpreadsheet();
  return ss.getSheetByName(sheetName);
}