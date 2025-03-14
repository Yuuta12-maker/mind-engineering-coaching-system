/**
 * マインドエンジニアリング・コーチング管理システム
 * メール自動化モジュール
 * 
 * 各種メールの自動送信機能を提供します。
 */

// ユーティリティ関数をインポート
// 実際のGASでは直接利用できるため、インポート文は必要ないですが
// コードの依存関係を明確にするために記述しています
// const Utilities = require('../utils/utilities.js');

/**
 * メールを送信し、ログに記録する
 * @param {string} to - 送信先メールアドレス
 * @param {string} subject - 件名
 * @param {string} body - 本文（HTMLまたはプレーンテキスト）
 * @param {Object} options - オプション（添付ファイルなど）
 * @param {string} clientId - クライアントID（ログ用）
 * @param {string} type - メールの種類（リマインダー、確認など）
 * @return {boolean} 送信が成功したかどうか
 */
function sendEmail(to, subject, body, options = {}, clientId = '', type = 'その他') {
  try {
    // 送信者名を取得
    const senderName = Utilities.getSetting('MAIL_SENDER_NAME', 'マインドエンジニアリング・コーチング 森山雄太');
    
    // メールオプションを準備
    const mailOptions = {
      name: senderName,
      htmlBody: options.isHtml ? body : null,
      ...options
    };
    
    // メールを送信
    GmailApp.sendEmail(to, subject, options.isHtml ? '' : body, mailOptions);
    
    // メールログに記録
    logEmailSent(to, subject, clientId, type);
    
    return true;
  } catch (error) {
    console.error('メール送信エラー:', error.message);
    
    // エラーをログに記録
    logEmailError(to, subject, clientId, type, error.message);
    
    return false;
  }
}

/**
 * メール送信をログに記録
 * @param {string} to - 送信先メールアドレス
 * @param {string} subject - 件名
 * @param {string} clientId - クライアントID
 * @param {string} type - メールの種類
 */
function logEmailSent(to, subject, clientId, type) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const emailLogSheet = ss.getSheetByName('メールログ');
  
  if (!emailLogSheet) {
    console.error('メールログシートが見つかりません');
    return;
  }
  
  // メールIDを生成
  const emailId = Utilities.generateUniqueId('EM');
  
  // 現在の日時
  const timestamp = new Date();
  
  // 行データを準備
  const rowData = [
    emailId,
    timestamp,
    clientId,
    to,
    subject,
    type,
    '送信済'
  ];
  
  // 最終行の次の行に追加
  const lastRow = Utilities.getLastRow(emailLogSheet);
  emailLogSheet.getRange(lastRow + 1, 1, 1, rowData.length).setValues([rowData]);
}

/**
 * メール送信エラーをログに記録
 * @param {string} to - 送信先メールアドレス
 * @param {string} subject - 件名
 * @param {string} clientId - クライアントID
 * @param {string} type - メールの種類
 * @param {string} errorMessage - エラーメッセージ
 */
function logEmailError(to, subject, clientId, type, errorMessage) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const emailLogSheet = ss.getSheetByName('メールログ');
  
  if (!emailLogSheet) {
    console.error('メールログシートが見つかりません');
    return;
  }
  
  // メールIDを生成
  const emailId = Utilities.generateUniqueId('EM');
  
  // 現在の日時
  const timestamp = new Date();
  
  // 行データを準備
  const rowData = [
    emailId,
    timestamp,
    clientId,
    to,
    subject + ' [エラー: ' + errorMessage + ']',
    type,
    '送信失敗'
  ];
  
  // 最終行の次の行に追加
  const lastRow = Utilities.getLastRow(emailLogSheet);
  emailLogSheet.getRange(lastRow + 1, 1, 1, rowData.length).setValues([rowData]);
}

/**
 * クライアントのメール履歴を取得
 * @param {string} clientId - クライアントID
 * @return {Array<Object>} メール履歴の配列
 */
function getClientEmailHistory(clientId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const emailLogSheet = ss.getSheetByName('メールログ');
  
  if (!emailLogSheet) {
    throw new Error('メールログシートが見つかりません');
  }
  
  // データの範囲を取得
  const dataRange = emailLogSheet.getDataRange();
  const values = dataRange.getValues();
  
  // ヘッダー行を取得
  const headers = values[0];
  
  // クライアントID列のインデックスを取得
  const clientIdIndex = headers.indexOf('クライアントID');
  
  if (clientIdIndex === -1) {
    throw new Error('クライアントIDカラムが見つかりません');
  }
  
  // メール履歴の配列を作成
  const emailHistory = [];
  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    
    // クライアントIDが一致するメールログを検索
    if (row[clientIdIndex] === clientId) {
      const emailLog = {};
      
      // オブジェクトを構築
      for (let j = 0; j < headers.length; j++) {
        if (headers[j]) { // 空のヘッダーはスキップ
          emailLog[headers[j]] = row[j];
        }
      }
      
      emailHistory.push(emailLog);
    }
  }
  
  // 送信日時で降順にソート
  return emailHistory.sort((a, b) => b.送信日時 - a.送信日時);
}

/**
 * メールテンプレートを取得
 * @param {string} templateName - テンプレート名
 * @return {string|null} テンプレート内容。見つからない場合はnull
 */
function getEmailTemplate(templateName) {
  // テンプレートの保存場所（実際の実装では設定から取得するなどの方法がある）
  // ここでは簡易的に基本的なテンプレートをハードコーディングする
  const templates = {
    // トライアルセッション案内メール
    'trialConfirmation': `
<p>{{クライアント名}} 様</p>

<p>マインドエンジニアリング・コーチングの森山雄太です。<br>
トライアルセッションのお申し込みありがとうございます。</p>

<p>以下の日時でトライアルセッションを承りましたのでご確認ください。</p>

<p>【セッション詳細】<br>
日時：{{セッション日時}}<br>
形式：{{セッション形式}}<br>
{{セッションURL}}</p>

<p>トライアルセッション料金（6,000円）のお支払いについては、<br>
以下の口座にお振込みいただけますようお願いいたします。</p>

<p>【お振込先】<br>
{{振込先情報}}</p>

<p>ご質問などございましたら、お気軽にご連絡ください。<br>
セッションでお会いできることを楽しみにしております。</p>

<p>--<br>
森山雄太<br>
マインドエンジニアリング・コーチング<br>
メール：mindengineeringcoaching@gmail.com<br>
電話：090-5710-7627<br>
住所：〒790-0012 愛媛県松山市湊町2-5-2リコオビル401</p>
`,

    // セッションリマインダーメール
    'sessionReminder': `
<p>{{クライアント名}} 様</p>

<p>マインドエンジニアリング・コーチングの森山雄太です。</p>

<p>明日の{{セッション日時}}に予定されているセッションのリマインダーをお送りいたします。</p>

<p>【セッション詳細】<br>
日時：{{セッション日時}}<br>
形式：{{セッション形式}}<br>
{{セッションURL}}</p>

<p>ご質問やご不明な点がございましたら、お気軽にご連絡ください。<br>
明日のセッションでお会いできることを楽しみにしております。</p>

<p>--<br>
森山雄太<br>
マインドエンジニアリング・コーチング<br>
メール：mindengineeringcoaching@gmail.com<br>
電話：090-5710-7627</p>
`,

    // 入金確認メール
    'paymentConfirmation': `
<p>{{クライアント名}} 様</p>

<p>マインドエンジニアリング・コーチングの森山雄太です。</p>

<p>{{項目}}のご入金（{{金額}}円）を確認いたしました。<br>
ありがとうございます。</p>

<p>領収書は追ってお送りいたします。</p>

<p>引き続きよろしくお願いいたします。</p>

<p>--<br>
森山雄太<br>
マインドエンジニアリング・コーチング<br>
メール：mindengineeringcoaching@gmail.com<br>
電話：090-5710-7627</p>
`,

    // 継続契約案内メール
    'continuationContract': `
<p>{{クライアント名}} 様</p>

<p>マインドエンジニアリング・コーチングの森山雄太です。</p>

<p>先日のトライアルセッションにお越しいただき、ありがとうございました。<br>
継続セッションについてのご案内をさせていただきます。</p>

<p>【継続セッション概要】<br>
・全5回のパーソナルセッション（月1回、約30分）<br>
・期間：6ヶ月間<br>
・料金：214,000円（税込）</p>

<p>継続をご希望の場合は、1週間以内に下記口座にお振込みいただき、<br>
メールにてご連絡いただけますようお願いいたします。</p>

<p>【お振込先】<br>
{{振込先情報}}</p>

<p>継続セッションでは、トライアルで設定したゴールに向けて、<br>
より深く効果的なコーチングを提供させていただきます。</p>

<p>ご質問やご不明な点がございましたら、お気軽にご連絡ください。<br>
お返事をお待ちしております。</p>

<p>--<br>
森山雄太<br>
マインドエンジニアリング・コーチング<br>
メール：mindengineeringcoaching@gmail.com<br>
電話：090-5710-7627<br>
住所：〒790-0012 愛媛県松山市湊町2-5-2リコオビル401</p>
`,

    // 次回セッション日程調整メール
    'nextSessionScheduling': `
<p>{{クライアント名}} 様</p>

<p>マインドエンジニアリング・コーチングの森山雄太です。</p>

<p>先日のセッションにご参加いただき、ありがとうございました。<br>
次回のセッション日程を調整させていただきたいと思います。</p>

<p>以下の日時で候補をいくつかご提案いたします。<br>
ご都合の良い日時をお知らせください。</p>

<p>【候補日時】<br>
{{候補日時1}}<br>
{{候補日時2}}<br>
{{候補日時3}}</p>

<p>上記以外の日時をご希望の場合も、お気軽にお知らせください。<br>
ご返信をお待ちしております。</p>

<p>--<br>
森山雄太<br>
マインドエンジニアリング・コーチング<br>
メール：mindengineeringcoaching@gmail.com<br>
電話：090-5710-7627</p>
`
  };
  
  return templates[templateName] || null;
}

/**
 * テンプレートを変数で置換する
 * @param {string} template - テンプレート
 * @param {Object} variables - 置換する変数のオブジェクト
 * @return {string} 置換後のテンプレート
 */
function replaceTemplateVariables(template, variables) {
  let result = template;
  
  // 各変数を置換
  for (const [key, value] of Object.entries(variables)) {
    const regex = new RegExp('{{' + key + '}}', 'g');
    result = result.replace(regex, value);
  }
  
  return result;
}

/**
 * トライアルセッション確認メールを送信
 * @param {Object} client - クライアント情報
 * @param {Object} session - セッション情報
 * @return {boolean} 送信が成功したかどうか
 */
function sendTrialConfirmation(client, session) {
  // テンプレートを取得
  const template = getEmailTemplate('trialConfirmation');
  if (!template) {
    console.error('テンプレートが見つかりません');
    return false;
  }
  
  // 変数を準備
  const variables = {
    'クライアント名': client.お名前,
    'セッション日時': Utilities.formatDateTime(session.予定日時),
    'セッション形式': client.希望セッション形式
  };
  
  // セッション形式に応じた情報を追加
  if (client.希望セッション形式 === 'オンライン') {
    variables['セッションURL'] = 'Google Meet URL: ' + session['Google Meet URL'];
  } else {
    variables['セッションURL'] = '場所: 〒790-0012 愛媛県松山市湊町2-5-2リコオビル401';
  }
  
  // 振込先情報を追加
  variables['振込先情報'] = Utilities.getSetting('BANK_INFO', '');
  
  // テンプレートを置換
  const emailBody = replaceTemplateVariables(template, variables);
  
  // メールを送信
  return sendEmail(
    client.メールアドレス,
    'マインドエンジニアリング・コーチング トライアルセッションのご案内',
    emailBody,
    { isHtml: true },
    client.クライアントID,
    'トライアル案内'
  );
}

/**
 * セッションリマインダーメールを送信
 * @param {Object} client - クライアント情報
 * @param {Object} session - セッション情報
 * @return {boolean} 送信が成功したかどうか
 */
function sendSessionReminder(client, session) {
  // テンプレートを取得
  const template = getEmailTemplate('sessionReminder');
  if (!template) {
    console.error('テンプレートが見つかりません');
    return false;
  }
  
  // 変数を準備
  const variables = {
    'クライアント名': client.お名前,
    'セッション日時': Utilities.formatDateTime(session.予定日時),
    'セッション形式': client.希望セッション形式
  };
  
  // セッション形式に応じた情報を追加
  if (client.希望セッション形式 === 'オンライン') {
    variables['セッションURL'] = 'Google Meet URL: ' + session['Google Meet URL'];
  } else {
    variables['セッションURL'] = '場所: 〒790-0012 愛媛県松山市湊町2-5-2リコオビル401';
  }
  
  // テンプレートを置換
  const emailBody = replaceTemplateVariables(template, variables);
  
  // メールを送信
  return sendEmail(
    client.メールアドレス,
    'マインドエンジニアリング・コーチング セッションリマインダー',
    emailBody,
    { isHtml: true },
    client.クライアントID,
    'リマインダー'
  );
}

/**
 * 入金確認メールを送信
 * @param {Object} client - クライアント情報
 * @param {Object} payment - 支払い情報
 * @return {boolean} 送信が成功したかどうか
 */
function sendPaymentConfirmation(client, payment) {
  // テンプレートを取得
  const template = getEmailTemplate('paymentConfirmation');
  if (!template) {
    console.error('テンプレートが見つかりません');
    return false;
  }
  
  // 変数を準備
  const variables = {
    'クライアント名': client.お名前,
    '項目': payment.項目,
    '金額': payment.金額.toLocaleString()
  };
  
  // テンプレートを置換
  const emailBody = replaceTemplateVariables(template, variables);
  
  // メールを送信
  return sendEmail(
    client.メールアドレス,
    'マインドエンジニアリング・コーチング ご入金確認のお知らせ',
    emailBody,
    { isHtml: true },
    client.クライアントID,
    '支払い案内'
  );
}

/**
 * 継続契約案内メールを送信
 * @param {Object} client - クライアント情報
 * @return {boolean} 送信が成功したかどうか
 */
function sendContinuationContract(client) {
  // テンプレートを取得
  const template = getEmailTemplate('continuationContract');
  if (!template) {
    console.error('テンプレートが見つかりません');
    return false;
  }
  
  // 変数を準備
  const variables = {
    'クライアント名': client.お名前,
    '振込先情報': Utilities.getSetting('BANK_INFO', '')
  };
  
  // テンプレートを置換
  const emailBody = replaceTemplateVariables(template, variables);
  
  // メールを送信
  return sendEmail(
    client.メールアドレス,
    'マインドエンジニアリング・コーチング 継続セッションのご案内',
    emailBody,
    { isHtml: true },
    client.クライアントID,
    '継続案内'
  );
}

/**
 * 次回セッション日程調整メールを送信
 * @param {Object} client - クライアント情報
 * @param {Array<Date>} candidateDates - 候補日時の配列
 * @return {boolean} 送信が成功したかどうか
 */
function sendNextSessionScheduling(client, candidateDates) {
  // テンプレートを取得
  const template = getEmailTemplate('nextSessionScheduling');
  if (!template) {
    console.error('テンプレートが見つかりません');
    return false;
  }
  
  // 変数を準備
  const variables = {
    'クライアント名': client.お名前
  };
  
  // 候補日時を追加
  for (let i = 0; i < Math.min(candidateDates.length, 3); i++) {
    variables[`候補日時${i+1}`] = Utilities.formatDateTime(candidateDates[i]);
  }
  
  // 3つ未満の候補日時の場合、残りをクリア
  for (let i = candidateDates.length + 1; i <= 3; i++) {
    variables[`候補日時${i}`] = '';
  }
  
  // テンプレートを置換
  const emailBody = replaceTemplateVariables(template, variables);
  
  // メールを送信
  return sendEmail(
    client.メールアドレス,
    'マインドエンジニアリング・コーチング 次回セッションの日程調整について',
    emailBody,
    { isHtml: true },
    client.クライアントID,
    '次回日程調整'
  );
}

/**
 * 領収書送付メールを送信
 * @param {Object} client - クライアント情報
 * @param {Object} payment - 支払い情報
 * @param {Object} receiptInfo - 領収書情報（fileId, url）
 * @return {boolean} 送信が成功したかどうか
 */
function sendReceiptEmail(client, payment, receiptInfo) {
  try {
    // 領収書ファイルを取得
    const receiptFile = DriveApp.getFileById(receiptInfo.fileId);
    
    // メール本文
    const emailBody = `
<p>${client.お名前} 様</p>

<p>マインドエンジニアリング・コーチングの森山雄太です。</p>

<p>${payment.項目}の領収書を添付いたします。<br>
ご確認ください。</p>

<p>引き続きよろしくお願いいたします。</p>

<p>--<br>
森山雄太<br>
マインドエンジニアリング・コーチング<br>
メール：mindengineeringcoaching@gmail.com<br>
電話：090-5710-7627</p>
`;
    
    // メールを送信
    return sendEmail(
      client.メールアドレス,
      'マインドエンジニアリング・コーチング 領収書の送付',
      emailBody,
      {
        isHtml: true,
        attachments: [receiptFile.getAs(MimeType.PDF)]
      },
      client.クライアントID,
      '領収書送付'
    );
  } catch (error) {
    console.error('領収書メール送信エラー:', error.message);
    return false;
  }
}

/**
 * 明日のセッションのリマインダーメールを送信
 * @return {number} 送信されたリマインダーの数
 */
function sendTomorrowSessionReminders() {
  // SessionManagerとClientManagerが必要
  
  // 明日の日付を計算
  const tomorrow = new Date();
  tomorrow.setDate(tomorrow.getDate() + 1);
  
  // 明日のセッションを取得
  const tomorrowSessions = SessionManager.getSessionsByDate(tomorrow, true);
  
  let sentCount = 0;
  
  // 各セッションについて
  for (const session of tomorrowSessions) {
    // クライアント情報を取得
    const client = ClientManager.findClientById(session.クライアントID);
    if (!client) continue;
    
    // リマインダーメールを送信
    try {
      if (sendSessionReminder(client, session)) {
        sentCount++;
      }
    } catch (error) {
      console.error(`リマインダー送信エラー (${client.お名前}): ${error.message}`);
    }
  }
  
  return sentCount;
}

// モジュールをエクスポート
const EmailManager = {
  sendEmail,
  getClientEmailHistory,
  getEmailTemplate,
  replaceTemplateVariables,
  sendTrialConfirmation,
  sendSessionReminder,
  sendPaymentConfirmation,
  sendContinuationContract,
  sendNextSessionScheduling,
  sendReceiptEmail,
  sendTomorrowSessionReminders
};