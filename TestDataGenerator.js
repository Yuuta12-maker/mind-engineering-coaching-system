/**
 * マインドエンジニアリング・コーチング管理システム
 * テストデータ生成スクリプト
 * 
 * このスクリプトは、システムの各シートにテストデータを自動的に追加します。
 * 開発環境でのテストやデモ用に利用できます。
 */

/**
 * テストデータの生成を実行する関数
 * GASエディタで実行すると、すべてのシートにテストデータが追加されます
 */
function generateTestData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // シートの存在確認と取得
  const clientSheet = ss.getSheetByName('クライアントinfo');
  const sessionSheet = ss.getSheetByName('セッション管理');
  const paymentSheet = ss.getSheetByName('支払い管理');
  const emailLogSheet = ss.getSheetByName('メールログ');
  
  if (!clientSheet || !sessionSheet || !paymentSheet || !emailLogSheet) {
    throw new Error('必要なシートが見つかりません。先に SetupDatabase.js の setupSpreadsheet() 関数を実行してください。');
  }
  
  // 既存データの有無を確認
  const clientRows = clientSheet.getLastRow();
  const hasExistingData = clientRows > 1;
  
  if (hasExistingData) {
    const ui = SpreadsheetApp.getUi();
    const response = ui.alert(
      'テストデータの生成',
      '既存のデータが見つかりました。テストデータを追加しますか？\n(既存データは保持されます)',
      ui.ButtonSet.YES_NO
    );
    
    if (response !== ui.Button.YES) {
      return;
    }
  }
  
  // トライアル料金と継続料金の設定を取得
  const trialFee = Number(Utilities.getSetting('TRIAL_FEE', '6000'));
  const continuationFee = Number(Utilities.getSetting('CONTINUATION_FEE', '214000'));
  
  // テストデータの生成
  const clients = generateClientData(20);  // 20件のクライアントデータ
  const sessions = generateSessionData(clients);
  const payments = generatePaymentData(clients, sessions, trialFee, continuationFee);
  const emailLogs = generateEmailLogData(clients, sessions);
  
  // データの挿入
  insertClientData(clientSheet, clients);
  insertSessionData(sessionSheet, sessions);
  insertPaymentData(paymentSheet, payments);
  insertEmailLogData(emailLogSheet, emailLogs);
  
  // 完了メッセージ
  const ui = SpreadsheetApp.getUi();
  ui.alert(
    'テストデータ生成完了',
    `テストデータの生成が完了しました。\n\n` +
    `クライアント: ${clients.length}件\n` +
    `セッション: ${sessions.length}件\n` +
    `支払い: ${payments.length}件\n` +
    `メールログ: ${emailLogs.length}件\n`,
    ui.ButtonSet.OK
  );
}

/**
 * クライアントデータを生成する関数
 * @param {number} count - 生成するデータ数
 * @return {Array} 生成されたクライアントデータの配列
 */
function generateClientData(count) {
  const clients = [];
  const now = new Date();
  
  // サンプル名前リスト
  const lastNames = ['山田', '鈴木', '佐藤', '田中', '伊藤', '渡辺', '高橋', '斎藤', '中村', '小林'];
  const firstNames = ['太郎', '花子', '一郎', '裕子', '健太', '美咲', '哲也', '恵子', '雄大', '愛'];
  const lastNamesKana = ['ヤマダ', 'スズキ', 'サトウ', 'タナカ', 'イトウ', 'ワタナベ', 'タカハシ', 'サイトウ', 'ナカムラ', 'コバヤシ'];
  const firstNamesKana = ['タロウ', 'ハナコ', 'イチロウ', 'ユウコ', 'ケンタ', 'ミサキ', 'テツヤ', 'ケイコ', 'ユウダイ', 'アイ'];
  
  // ステータスの分布
  const statusOptions = [
    { status: '問い合わせ', weight: 10 },
    { status: 'トライアル前', weight: 20 },
    { status: 'トライアル済', weight: 15 },
    { status: '契約中', weight: 40 },
    { status: '完了', weight: 10 },
    { status: '中断', weight: 5 }
  ];
  
  // セッション形式の分布
  const sessionTypeOptions = [
    { type: 'オンライン', weight: 70 },
    { type: '対面', weight: 25 },
    { type: '未定', weight: 5 }
  ];
  
  for (let i = 0; i < count; i++) {
    // ランダムな名前生成
    const lastNameIndex = Math.floor(Math.random() * lastNames.length);
    const firstNameIndex = Math.floor(Math.random() * firstNames.length);
    const lastName = lastNames[lastNameIndex];
    const firstName = firstNames[firstNameIndex];
    const lastNameKana = lastNamesKana[lastNameIndex];
    const firstNameKana = firstNamesKana[firstNameIndex];
    
    // ランダムな日付生成（過去3ヶ月以内）
    const registrationDate = new Date(now);
    registrationDate.setDate(now.getDate() - Math.floor(Math.random() * 90));
    
    // ランダムな生年月日生成（20-70歳）
    const birthDate = new Date();
    const age = 20 + Math.floor(Math.random() * 50);
    birthDate.setFullYear(now.getFullYear() - age);
    birthDate.setMonth(Math.floor(Math.random() * 12));
    birthDate.setDate(1 + Math.floor(Math.random() * 28));
    
    // ランダムな電話番号生成
    const phoneNumber = `0${7 + Math.floor(Math.random() * 3)}0${Math.floor(Math.random() * 10000000 + 1000000)}`;
    
    // ランダムなメールアドレス生成
    const email = `${lastName.toLowerCase()}${firstName.toLowerCase()}${Math.floor(Math.random() * 1000)}@example.com`;
    
    // ステータスを重み付けでランダム選択
    const status = weightedRandom(statusOptions);
    
    // セッション形式を重み付けでランダム選択
    const sessionType = weightedRandom(sessionTypeOptions);
    
    // クライアントIDを生成
    const clientId = `CL${registrationDate.getTime().toString().slice(-9)}${Math.floor(Math.random() * 10000).toString().padStart(4, '0')}`;
    
    // クライアントオブジェクトを作成
    const client = {
      'クライアントID': clientId,
      'タイムスタンプ': registrationDate,
      'メールアドレス': email,
      'お名前': `${lastName} ${firstName}`,
      'お名前　（カナ）': `${lastNameKana} ${firstNameKana}`,
      '性別': Math.random() > 0.5 ? '男性' : '女性',
      '生年月日': birthDate,
      '電話番号　（ハイフンなし）': phoneNumber,
      'ご住所': `〒${100 + Math.floor(Math.random() * 899)}-${Math.floor(Math.random() * 10000).toString().padStart(4, '0')} 東京都千代田区丸の内${1 + Math.floor(Math.random() * 3)}-${1 + Math.floor(Math.random() * 10)}-${1 + Math.floor(Math.random() * 20)}`,
      '希望セッション形式': sessionType,
      'ステータス': status,
      '備考欄': getRandomRemark()
    };
    
    clients.push(client);
  }
  
  return clients;
}

/**
 * セッションデータを生成する関数
 * @param {Array} clients - クライアントデータの配列
 * @return {Array} 生成されたセッションデータの配列
 */
function generateSessionData(clients) {
  const sessions = [];
  const now = new Date();
  const sessionDuration = Number(Utilities.getSetting('SESSION_DURATION', '30'));
  
  // セッション種別の順序
  const sessionTypes = ['トライアル', '継続（2回目）', '継続（3回目）', '継続（4回目）', '継続（5回目）', '継続（6回目）', 'フォローアップ'];
  
  // クライアントごとにセッションを生成
  clients.forEach(client => {
    // クライアントの状態によってセッション数を決定
    let sessionCount = 0;
    let pastSessionCount = 0;
    let futureSessionCount = 0;
    
    switch (client['ステータス']) {
      case '問い合わせ':
        // セッションなし
        break;
      case 'トライアル前':
        futureSessionCount = 1;
        break;
      case 'トライアル済':
        pastSessionCount = 1;
        futureSessionCount = Math.random() > 0.7 ? 1 : 0; // 30%は次のセッションも予定
        break;
      case '契約中':
        // 契約中は1-5のセッションが完了している
        pastSessionCount = 1 + Math.floor(Math.random() * 5);
        futureSessionCount = 6 - pastSessionCount > 0 ? 1 : 0; // 残りのセッションのうち1つを予定
        break;
      case '完了':
        pastSessionCount = 6; // すべてのセッションが完了
        futureSessionCount = Math.random() > 0.8 ? 1 : 0; // 20%はフォローアップセッションも実施
        break;
      case '中断':
        // 途中で中断したケース
        pastSessionCount = 1 + Math.floor(Math.random() * 3);
        break;
    }
    
    sessionCount = pastSessionCount + futureSessionCount;
    
    // セッションがない場合はスキップ
    if (sessionCount === 0) {
      return;
    }
    
    // 最初のセッション日を設定（クライアント登録日の1-14日後）
    const firstSessionDate = new Date(client['タイムスタンプ']);
    firstSessionDate.setDate(firstSessionDate.getDate() + 1 + Math.floor(Math.random() * 14));
    
    // 時間は9:00-17:00の間で30分刻み
    const hours = 9 + Math.floor(Math.random() * 8);
    const minutes = Math.random() > 0.5 ? 0 : 30;
    firstSessionDate.setHours(hours, minutes, 0, 0);
    
    // セッションを生成
    for (let i = 0; i < sessionCount; i++) {
      const sessionDate = new Date(firstSessionDate);
      
      // 各セッションの間隔は2週間〜4週間
      if (i > 0) {
        sessionDate.setDate(sessionDate.getDate() + 14 + Math.floor(Math.random() * 14));
      }
      
      // セッションタイプを決定
      let sessionType = i < sessionTypes.length ? sessionTypes[i] : 'フォローアップ';
      
      // セッションステータスを決定
      let sessionStatus;
      if (sessionDate < now && i < pastSessionCount) {
        sessionStatus = '実施済';
      } else if (sessionDate > now && i >= pastSessionCount) {
        sessionStatus = '予定';
      } else {
        // 当日または過去のセッションで未実施の場合（異常ケース）
        const rand = Math.random();
        if (rand < 0.7) {
          sessionStatus = '実施済';
        } else if (rand < 0.9) {
          sessionStatus = '延期';
        } else {
          sessionStatus = 'キャンセル';
        }
      }
      
      // Google Meet URLを生成（オンラインセッションの場合）
      let meetUrl = '';
      if (client['希望セッション形式'] === 'オンライン') {
        meetUrl = `https://meet.google.com/${randomString(3)}-${randomString(4)}-${randomString(3)}`;
      }
      
      // 実施日時を決定
      let implementationDate = null;
      if (sessionStatus === '実施済') {
        implementationDate = new Date(sessionDate);
      }
      
      // セッションIDを生成
      const sessionId = `SS${sessionDate.getTime().toString().slice(-9)}${Math.floor(Math.random() * 10000).toString().padStart(4, '0')}`;
      
      // セッション記録を生成
      let sessionRecord = '';
      if (sessionStatus === '実施済') {
        sessionRecord = getRandomSessionRecord(sessionType);
      }
      
      // セッションオブジェクトを作成
      const session = {
        'セッションID': sessionId,
        'クライアントID': client['クライアントID'],
        'セッション種別': sessionType,
        '予定日時': sessionDate,
        'Google Meet URL': meetUrl,
        'ステータス': sessionStatus,
        '実施日時': implementationDate,
        '記録': sessionRecord,
        '備考': sessionStatus === '延期' ? '都合により延期' : 
                sessionStatus === 'キャンセル' ? '体調不良によりキャンセル' : ''
      };
      
      sessions.push(session);
    }
  });
  
  return sessions;
}

/**
 * 支払いデータを生成する関数
 * @param {Array} clients - クライアントデータの配列
 * @param {Array} sessions - セッションデータの配列
 * @param {number} trialFee - トライアル料金
 * @param {number} continuationFee - 継続料金
 * @return {Array} 生成された支払いデータの配列
 */
function generatePaymentData(clients, sessions, trialFee, continuationFee) {
  const payments = [];
  const now = new Date();
  
  // セッションをクライアントごとにグループ化
  const sessionsByClient = {};
  sessions.forEach(session => {
    if (!sessionsByClient[session['クライアントID']]) {
      sessionsByClient[session['クライアントID']] = [];
    }
    sessionsByClient[session['クライアントID']].push(session);
  });
  
  // クライアントごとに支払いデータを生成
  clients.forEach(client => {
    const clientSessions = sessionsByClient[client['クライアントID']] || [];
    
    // クライアントのセッションを日付順にソート
    clientSessions.sort((a, b) => a['予定日時'] - b['予定日時']);
    
    // トライアルセッションがあれば支払いを生成
    const trialSession = clientSessions.find(s => s['セッション種別'] === 'トライアル');
    if (trialSession) {
      // 支払い登録日はトライアルセッションの1-3日前
      const registrationDate = new Date(trialSession['予定日時']);
      registrationDate.setDate(registrationDate.getDate() - (1 + Math.floor(Math.random() * 3)));
      
      // 入金日は登録日の0-2日後
      const paymentDate = new Date(registrationDate);
      paymentDate.setDate(paymentDate.getDate() + Math.floor(Math.random() * 3));
      
      // 支払い状態を決定
      let paymentStatus;
      if (paymentDate <= now) {
        paymentStatus = '入金済';
      } else {
        paymentStatus = '未入金';
      }
      
      // 領収書発行状態を決定
      let receiptStatus;
      if (paymentStatus === '入金済') {
        receiptStatus = Math.random() > 0.3 ? '発行済' : '未発行';
      } else {
        receiptStatus = '未発行';
      }
      
      // 支払いIDを生成
      const paymentId = `PY${registrationDate.getTime().toString().slice(-9)}${Math.floor(Math.random() * 10000).toString().padStart(4, '0')}`;
      
      // 支払いオブジェクトを作成
      const payment = {
        '支払いID': paymentId,
        'クライアントID': client['クライアントID'],
        '登録日': registrationDate,
        '項目': 'トライアルセッション',
        '金額': trialFee,
        '状態': paymentStatus,
        '入金日': paymentStatus === '入金済' ? paymentDate : null,
        '領収書発行状態': receiptStatus,
        '備考': ''
      };
      
      payments.push(payment);
    }
    
    // 継続セッションがあれば支払いを生成
    const continuationSession = clientSessions.find(s => s['セッション種別'] === '継続（2回目）');
    if (continuationSession) {
      // 支払い登録日はトライアルセッションの実施日から1-5日後
      const trialImplementationDate = trialSession && trialSession['実施日時'] ? 
        new Date(trialSession['実施日時']) : 
        new Date(client['タイムスタンプ']);
      
      const registrationDate = new Date(trialImplementationDate);
      registrationDate.setDate(registrationDate.getDate() + 1 + Math.floor(Math.random() * 5));
      
      // 入金日は登録日の0-3日後
      const paymentDate = new Date(registrationDate);
      paymentDate.setDate(paymentDate.getDate() + Math.floor(Math.random() * 4));
      
      // 支払い状態を決定
      let paymentStatus;
      if (client['ステータス'] === 'トライアル済' && Math.random() > 0.7) {
        paymentStatus = '未入金'; // 30%は支払い前
      } else if (paymentDate <= now) {
        paymentStatus = '入金済';
      } else {
        paymentStatus = '未入金';
      }
      
      // 領収書発行状態を決定
      let receiptStatus;
      if (paymentStatus === '入金済') {
        receiptStatus = Math.random() > 0.3 ? '発行済' : '未発行';
      } else {
        receiptStatus = '未発行';
      }
      
      // 支払いIDを生成
      const paymentId = `PY${registrationDate.getTime().toString().slice(-9)}${Math.floor(Math.random() * 10000).toString().padStart(4, '0')}`;
      
      // 支払いオブジェクトを作成
      const payment = {
        '支払いID': paymentId,
        'クライアントID': client['クライアントID'],
        '登録日': registrationDate,
        '項目': '継続セッション（全5回）',
        '金額': continuationFee,
        '状態': paymentStatus,
        '入金日': paymentStatus === '入金済' ? paymentDate : null,
        '領収書発行状態': receiptStatus,
        '備考': ''
      };
      
      payments.push(payment);
    }
  });
  
  return payments;
}

/**
 * メールログデータを生成する関数
 * @param {Array} clients - クライアントデータの配列
 * @param {Array} sessions - セッションデータの配列
 * @return {Array} 生成されたメールログデータの配列
 */
function generateEmailLogData(clients, sessions) {
  const emailLogs = [];
  const now = new Date();
  
  // メールの種類
  const emailTypes = ['トライアル案内', 'セッション確認', 'リマインダー', '支払い案内', '領収書送付', '次回日程調整', 'その他'];
  
  // セッションをクライアントごとにグループ化
  const sessionsByClient = {};
  sessions.forEach(session => {
    if (!sessionsByClient[session['クライアントID']]) {
      sessionsByClient[session['クライアントID']] = [];
    }
    sessionsByClient[session['クライアントID']].push(session);
  });
  
  // クライアントごとにメールログを生成
  clients.forEach(client => {
    const clientSessions = sessionsByClient[client['クライアントID']] || [];
    
    // クライアントのセッションを日付順にソート
    clientSessions.sort((a, b) => a['予定日時'] - b['予定日時']);
    
    // 問い合わせ時のメール
    const inquiryDate = new Date(client['タイムスタンプ']);
    addEmailLog(emailLogs, client, inquiryDate, 'その他', 'お問い合わせありがとうございます');
    
    // トライアルセッションのメール
    const trialSession = clientSessions.find(s => s['セッション種別'] === 'トライアル');
    if (trialSession) {
      // トライアル案内メール
      const trialInvitationDate = new Date(inquiryDate);
      trialInvitationDate.setDate(inquiryDate.getDate() + 1);
      addEmailLog(emailLogs, client, trialInvitationDate, 'トライアル案内', 'トライアルセッションのご案内');
      
      // トライアルリマインダーメール
      if (trialSession['予定日時'] <= now) {
        const reminderDate = new Date(trialSession['予定日時']);
        reminderDate.setDate(reminderDate.getDate() - 1);
        addEmailLog(emailLogs, client, reminderDate, 'リマインダー', 'トライアルセッションリマインダー');
      }
      
      // トライアル実施後のメール
      if (trialSession['ステータス'] === '実施済' && trialSession['実施日時']) {
        const followupDate = new Date(trialSession['実施日時']);
        followupDate.setDate(followupDate.getDate() + 1);
        addEmailLog(emailLogs, client, followupDate, '次回日程調整', 'トライアルセッション後のご案内');
      }
    }
    
    // 継続セッションのメール
    clientSessions.forEach((session, index) => {
      if (index === 0) return; // トライアルセッションは既に処理済み
      
      // セッション確認メール
      if (session['予定日時'] > inquiryDate) {
        const confirmationDate = new Date(session['予定日時']);
        confirmationDate.setDate(confirmationDate.getDate() - 3);
        if (confirmationDate <= now) {
          addEmailLog(emailLogs, client, confirmationDate, 'セッション確認', `${session['セッション種別']}のご案内`);
        }
      }
      
      // リマインダーメール
      if (session['予定日時'] <= now) {
        const reminderDate = new Date(session['予定日時']);
        reminderDate.setDate(reminderDate.getDate() - 1);
        addEmailLog(emailLogs, client, reminderDate, 'リマインダー', `${session['セッション種別']}リマインダー`);
      }
      
      // セッション後のフォローアップメール
      if (session['ステータス'] === '実施済' && session['実施日時']) {
        const followupDate = new Date(session['実施日時']);
        followupDate.setDate(followupDate.getDate() + 1);
        if (index < clientSessions.length - 1) {
          addEmailLog(emailLogs, client, followupDate, '次回日程調整', `${session['セッション種別']}後のご案内`);
        } else {
          addEmailLog(emailLogs, client, followupDate, 'その他', 'プログラム完了のご案内');
        }
      }
    });
    
    // 支払い関連のメール
    if (client['ステータス'] !== '問い合わせ') {
      // 支払い案内メール
      const paymentInvitationDate = new Date(inquiryDate);
      paymentInvitationDate.setDate(inquiryDate.getDate() + 2);
      addEmailLog(emailLogs, client, paymentInvitationDate, '支払い案内', 'お支払いのご案内');
      
      // 領収書送付メール
      if (['トライアル済', '契約中', '完了'].includes(client['ステータス'])) {
        const receiptDate = new Date(inquiryDate);
        receiptDate.setDate(inquiryDate.getDate() + 5);
        if (receiptDate <= now) {
          addEmailLog(emailLogs, client, receiptDate, '領収書送付', '領収書送付のお知らせ');
        }
      }
    }
  });
  
  return emailLogs;
}

/**
 * メールログを追加するヘルパー関数
 */
function addEmailLog(emailLogs, client, date, type, subject) {
  // 現在日時より未来の場合はスキップ
  if (date > new Date()) {
    return;
  }
  
  // メールIDを生成
  const emailId = `EM${date.getTime().toString().slice(-9)}${Math.floor(Math.random() * 10000).toString().padStart(4, '0')}`;
  
  // メールログオブジェクトを作成
  const emailLog = {
    'メールID': emailId,
    '送信日時': date,
    'クライアントID': client['クライアントID'],
    '送信先': client['メールアドレス'],
    '件名': `[MEC] ${subject}`,
    '種類': type,
    'ステータス': '送信済'
  };
  
  emailLogs.push(emailLog);
}

/**
 * 重み付けされたランダム選択を行うユーティリティ関数
 */
function weightedRandom(options) {
  const totalWeight = options.reduce((sum, option) => sum + option.weight, 0);
  let random = Math.random() * totalWeight;
  
  for (const option of options) {
    random -= option.weight;
    if (random <= 0) {
      return option.status || option.type;
    }
  }
  
  return options[0].status || options[0].type;
}

/**
 * ランダムな文字列を生成するユーティリティ関数
 */
function randomString(length) {
  const chars = 'abcdefghijklmnopqrstuvwxyz';
  let result = '';
  for (let i = 0; i < length; i++) {
    result += chars.charAt(Math.floor(Math.random() * chars.length));
  }
  return result;
}

/**
 * ランダムな備考文を生成するユーティリティ関数
 */
function getRandomRemark() {
  const remarks = [
    '',
    'ホームページからのお問い合わせ',
    '知人の紹介',
    'セミナー参加者',
    'Instagramのフォロワー',
    'ブログ読者',
    'メルマガ読者',
    '仕事の関係者',
    'リピート利用',
    '電話での問い合わせ'
  ];
  
  return remarks[Math.floor(Math.random() * remarks.length)];
}

/**
 * ランダムなセッション記録を生成するユーティリティ関数
 */
function getRandomSessionRecord(sessionType) {
  const baseRecords = [
    '現状の悩みについてヒアリングを行いました。ゴール設定のための第一歩として、本人の「やりたいこと」を明確化する作業を進めました。',
    'コンフォートゾーンについて説明し、現状の外側にあるゴールを設定するワークを実施しました。クライアントは新しい視点を得て意欲的です。',
    'バランスホイールを用いて、人生の各領域における満足度を可視化しました。特に仕事と趣味の領域での変化を望んでいます。',
    'ホメオスタシスの力を活用したゴール達成のプロセスについて説明しました。クライアントはエネルギーの源泉について理解を深めていました。',
    '100%やりたいことに焦点を当て、社会貢献とつながるゴールの再設定を行いました。具体的なビジョンが形成されてきています。',
    'クライアントの内省言語を引き出すことに注力しました。セッション中に重要な気づきがありました。'
  ];
  
  let record = baseRecords[Math.floor(Math.random() * baseRecords.length)];
  
  // セッション種別に応じた追加文
  if (sessionType === 'トライアル') {
    record += ' 現状の外側にゴールを設定する重要性を理解されていました。';
  } else if (sessionType.includes('2回目')) {
    record += ' 前回からの変化について振り返りを行い、新たな視点を得られました。';
  } else if (sessionType.includes('3回目')) {
    record += ' バランスホイールの各領域での小さな成功体験が出てきています。';
  } else if (sessionType.includes('4回目')) {
    record += ' ゴールに向けたアクションが自然と生まれてきています。努力感なく進んでいる実感があるようです。';
  } else if (sessionType.includes('5回目')) {
    record += ' スコトーマが外れて、新しい可能性が見えてきたとの報告がありました。';
  } else if (sessionType.includes('6回目')) {
    record += ' プログラム全体を振り返り、変化の定着と今後の展望について話し合いました。';
  } else if (sessionType === 'フォローアップ') {
    record += ' プログラム終了後の変化の定着状況を確認しました。社会貢献の視点から新たなゴールも生まれています。';
  }
  
  return record;
}

/**
 * クライアントデータをシートに挿入する関数
 */
function insertClientData(sheet, clients) {
  if (clients.length === 0) return;
  
  // ヘッダー行を取得
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  
  // 挿入開始行を決定
  const startRow = Math.max(2, sheet.getLastRow() + 1);
  
  // クライアントごとにデータを挿入
  clients.forEach((client, index) => {
    const rowIndex = startRow + index;
    
    // 各カラムにデータを設定
    headers.forEach((header, colIndex) => {
      if (header && client[header] !== undefined) {
        sheet.getRange(rowIndex, colIndex + 1).setValue(client[header]);
      }
    });
  });
}

/**
 * セッションデータをシートに挿入する関数
 */
function insertSessionData(sheet, sessions) {
  if (sessions.length === 0) return;
  
  // ヘッダー行を取得
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  
  // 挿入開始行を決定
  const startRow = Math.max(2, sheet.getLastRow() + 1);
  
  // セッションごとにデータを挿入
  sessions.forEach((session, index) => {
    const rowIndex = startRow + index;
    
    // 各カラムにデータを設定
    headers.forEach((header, colIndex) => {
      if (header && session[header] !== undefined) {
        sheet.getRange(rowIndex, colIndex + 1).setValue(session[header]);
      }
    });
  });
}

/**
 * 支払いデータをシートに挿入する関数
 */
function insertPaymentData(sheet, payments) {
  if (payments.length === 0) return;
  
  // ヘッダー行を取得
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  
  // 挿入開始行を決定
  const startRow = Math.max(2, sheet.getLastRow() + 1);
  
  // 支払いごとにデータを挿入
  payments.forEach((payment, index) => {
    const rowIndex = startRow + index;
    
    // 各カラムにデータを設定
    headers.forEach((header, colIndex) => {
      if (header && payment[header] !== undefined) {
        sheet.getRange(rowIndex, colIndex + 1).setValue(payment[header]);
      }
    });
  });
}

/**
 * メールログデータをシートに挿入する関数
 */
function insertEmailLogData(sheet, emailLogs) {
  if (emailLogs.length === 0) return;
  
  // ヘッダー行を取得
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  
  // 挿入開始行を決定
  const startRow = Math.max(2, sheet.getLastRow() + 1);
  
  // メールログごとにデータを挿入
  emailLogs.forEach((emailLog, index) => {
    const rowIndex = startRow + index;
    
    // 各カラムにデータを設定
    headers.forEach((header, colIndex) => {
      if (header && emailLog[header] !== undefined) {
        sheet.getRange(rowIndex, colIndex + 1).setValue(emailLog[header]);
      }
    });
  });
}

/**
 * メニューに追加する関数
 * この関数は既存のonOpen関数を拡張するために、別の名前で定義しています
 */
function addTestDataGeneratorToMenu() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('MEC管理システム')
    .addItem('データベース初期設定', 'setupSpreadsheet')
    .addItem('テストデータ生成', 'generateTestData')
    .addSeparator()
    .addItem('バージョン情報', 'showVersionInfo')
    .addToUi();
}

/**
 * スプレッドシートが開かれたときに実行される関数（既存の関数を上書き）
 */
function onOpen() {
  addTestDataGeneratorToMenu();
}