/**
 * マインドエンジニアリング・コーチング管理システム
 * セッション管理モジュール
 * 
 * セッションの予約、更新、削除、Google Calendar連携などの機能を提供します。
 */

// ユーティリティ関数とクライアント管理モジュールをインポート
// 実際のGASでは直接利用できるため、インポート文は必要ないですが
// コードの依存関係を明確にするために記述しています
// const Utilities = require('../utils/utilities.js');
// const ClientManager = require('../client/clientManager.js');

/**
 * セッションを作成
 * @param {Object} sessionData - セッション情報のオブジェクト
 * @return {Object} 作成されたセッション情報（IDを含む）
 */
function createSession(sessionData) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sessionSheet = ss.getSheetByName('セッション管理');
  
  if (!sessionSheet) {
    throw new Error('セッション管理シートが見つかりません');
  }
  
  // 必須フィールドをチェック
  if (!sessionData.クライアントID || !sessionData.予定日時) {
    throw new Error('クライアントIDと予定日時は必須項目です');
  }
  
  // クライアントの存在確認
  const client = ClientManager.findClientById(sessionData.クライアントID);
  if (!client) {
    throw new Error('指定されたクライアントIDは存在しません');
  }
  
  // セッションIDを生成
  const sessionId = Utilities.generateUniqueId('SS');
  
  // Google MeetのURLを生成（オンラインセッションの場合）
  let meetUrl = '';
  if (client.希望セッション形式 === 'オンライン' || sessionData.Google_Meet_URL) {
    meetUrl = sessionData.Google_Meet_URL || createGoogleMeetUrl(sessionData.予定日時, client);
  }
  
  // 行データを準備
  const rowData = [
    sessionId,
    sessionData.クライアントID,
    sessionData.セッション種別 || 'トライアル',
    sessionData.予定日時 instanceof Date ? sessionData.予定日時 : new Date(sessionData.予定日時),
    meetUrl,
    sessionData.ステータス || '予定',
    sessionData.実施日時 || '',
    sessionData.記録 || '',
    sessionData.備考 || ''
  ];
  
  // 最終行の次の行に追加
  const lastRow = Utilities.getLastRow(sessionSheet);
  sessionSheet.getRange(lastRow + 1, 1, 1, rowData.length).setValues([rowData]);
  
  // Googleカレンダーにイベントを追加
  const calendarEventId = addToGoogleCalendar(
    sessionId,
    sessionData.クライアントID,
    rowData[3], // 予定日時
    sessionData.セッション種別,
    meetUrl,
    client
  );
  
  // 作成したセッション情報を返す
  return {
    セッションID: sessionId,
    クライアントID: sessionData.クライアントID,
    セッション種別: rowData[2],
    予定日時: rowData[3],
    'Google Meet URL': meetUrl,
    ステータス: rowData[5],
    実施日時: rowData[6],
    記録: rowData[7],
    備考: rowData[8],
    calendarEventId: calendarEventId
  };
}

/**
 * Googleカレンダーにセッションイベントを追加
 * @param {string} sessionId - セッションID
 * @param {string} clientId - クライアントID
 * @param {Date} scheduledDateTime - 予定日時
 * @param {string} sessionType - セッション種別
 * @param {string} meetUrl - Google Meet URL
 * @param {Object} clientInfo - クライアント情報
 * @return {string} 作成されたカレンダーイベントのID
 */
function addToGoogleCalendar(sessionId, clientId, scheduledDateTime, sessionType, meetUrl, clientInfo) {
  // セッション時間（分）を取得
  const sessionDuration = parseInt(Utilities.getSetting('SESSION_DURATION', '30'));
  
  // 終了時間を計算
  const endTime = new Date(scheduledDateTime.getTime() + sessionDuration * 60 * 1000);
  
  // タイトルを作成
  const title = `【${sessionType}】${clientInfo.お名前} 様`;
  
  // 説明文を作成
  let description = `セッションID: ${sessionId}\n`;
  description += `クライアントID: ${clientId}\n`;
  description += `クライアント名: ${clientInfo.お名前}\n`;
  description += `メールアドレス: ${clientInfo.メールアドレス}\n`;
  description += `電話番号: ${clientInfo['電話番号　（ハイフンなし）'] || '未登録'}\n`;
  description += `セッション形式: ${clientInfo.希望セッション形式}\n`;
  
  if (meetUrl) {
    description += `\nGoogle Meet URL: ${meetUrl}`;
  }
  
  // カレンダーを取得
  const calendar = CalendarApp.getDefaultCalendar();
  
  // イベントを作成
  const event = calendar.createEvent(
    title,
    scheduledDateTime,
    endTime,
    {
      description: description,
      location: clientInfo.希望セッション形式 === '対面' ? Utilities.getSetting('BUSINESS_ADDRESS', '') : 'オンライン',
      guests: clientInfo.メールアドレス,
      sendInvites: false // 招待状は自動メール機能で別途送信
    }
  );
  
  // Google Meetを使用する場合
  if (meetUrl && clientInfo.希望セッション形式 === 'オンライン') {
    event.setVideoCallLink(meetUrl);
  }
  
  return event.getId();
}

/**
 * Google Meet URLを生成
 * @param {Date} scheduledDateTime - 予定日時
 * @param {Object} clientInfo - クライアント情報
 * @return {string} 生成されたGoogle Meet URL
 */
function createGoogleMeetUrl(scheduledDateTime, clientInfo) {
  // 実際の環境ではGoogle Meet APIを使用してURLを生成するが、
  // ここではシミュレーションのみ行う
  
  // Google MeetのURLは通常、サービスアカウントを使用するか、
  // Google Calendar APIを使ってイベントを作成する際に自動的に生成される
  
  // 一意の会議IDを生成
  const meetingId = Utilities.generateUniqueId('meet').substring(4, 16).toLowerCase();
  
  // Google Meet URLを返す
  return `https://meet.google.com/${meetingId}`;
}

/**
 * すべてのセッション情報を取得
 * @param {boolean} activeOnly - アクティブなセッションのみを取得するか（デフォルト: false）
 * @return {Array<Object>} セッション情報の配列
 */
function getAllSessions(activeOnly = false) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sessionSheet = ss.getSheetByName('セッション管理');
  
  if (!sessionSheet) {
    throw new Error('セッション管理シートが見つかりません');
  }
  
  // データの範囲を取得
  const dataRange = sessionSheet.getDataRange();
  const values = dataRange.getValues();
  
  // ヘッダー行を取得
  const headers = values[0];
  
  // ステータス列のインデックスを取得
  const statusIndex = headers.indexOf('ステータス');
  
  // セッション情報の配列を作成
  const sessions = [];
  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    
    // 行が空でないことを確認
    if (row[0]) {
      const session = {};
      
      // オブジェクトを構築
      for (let j = 0; j < headers.length; j++) {
        if (headers[j]) { // 空のヘッダーはスキップ
          session[headers[j]] = row[j];
        }
      }
      
      // アクティブなセッションのみを取得する場合
      if (activeOnly) {
        // 「予定」または「延期」のステータスのセッションのみを追加
        if (statusIndex >= 0 && (row[statusIndex] === '予定' || row[statusIndex] === '延期')) {
          sessions.push(session);
        }
      } else {
        sessions.push(session);
      }
    }
  }
  
  return sessions;
}

/**
 * クライアントIDによるセッション情報の検索
 * @param {string} clientId - 検索するクライアントID
 * @param {boolean} activeOnly - アクティブなセッションのみを取得するか（デフォルト: false）
 * @return {Array<Object>} セッション情報の配列
 */
function findSessionsByClientId(clientId, activeOnly = false) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sessionSheet = ss.getSheetByName('セッション管理');
  
  if (!sessionSheet) {
    throw new Error('セッション管理シートが見つかりません');
  }
  
  // データの範囲を取得
  const dataRange = sessionSheet.getDataRange();
  const values = dataRange.getValues();
  
  // ヘッダー行を取得
  const headers = values[0];
  
  // クライアントID列とステータス列のインデックスを取得
  const clientIdIndex = headers.indexOf('クライアントID');
  const statusIndex = headers.indexOf('ステータス');
  
  if (clientIdIndex === -1) {
    throw new Error('クライアントIDカラムが見つかりません');
  }
  
  // セッション情報の配列を作成
  const sessions = [];
  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    
    // クライアントIDが一致するセッションを検索
    if (row[clientIdIndex] === clientId) {
      const session = {};
      
      // オブジェクトを構築
      for (let j = 0; j < headers.length; j++) {
        if (headers[j]) { // 空のヘッダーはスキップ
          session[headers[j]] = row[j];
        }
      }
      
      // アクティブなセッションのみを取得する場合
      if (activeOnly) {
        // 「予定」または「延期」のステータスのセッションのみを追加
        if (statusIndex >= 0 && (row[statusIndex] === '予定' || row[statusIndex] === '延期')) {
          sessions.push(session);
        }
      } else {
        sessions.push(session);
      }
    }
  }
  
  return sessions;
}

/**
 * セッションIDによるセッション情報の検索
 * @param {string} sessionId - 検索するセッションID
 * @return {Object|null} セッション情報。見つからない場合はnull
 */
function findSessionById(sessionId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sessionSheet = ss.getSheetByName('セッション管理');
  
  if (!sessionSheet) {
    throw new Error('セッション管理シートが見つかりません');
  }
  
  // データの範囲を取得
  const dataRange = sessionSheet.getDataRange();
  const values = dataRange.getValues();
  
  // ヘッダー行を取得
  const headers = values[0];
  
  // セッションID列のインデックスを取得
  const sessionIdIndex = headers.indexOf('セッションID');
  
  if (sessionIdIndex === -1) {
    throw new Error('セッションIDカラムが見つかりません');
  }
  
  // セッションIDによる検索
  for (let i = 1; i < values.length; i++) {
    if (values[i][sessionIdIndex] === sessionId) {
      const session = {};
      
      // オブジェクトを構築
      for (let j = 0; j < headers.length; j++) {
        if (headers[j]) { // 空のヘッダーはスキップ
          session[headers[j]] = values[i][j];
        }
      }
      
      return session;
    }
  }
  
  return null; // セッションが見つからない場合
}

/**
 * セッション情報を更新
 * @param {string} sessionId - 更新するセッションID
 * @param {Object} sessionData - 更新するセッション情報
 * @return {Object|null} 更新されたセッション情報。セッションが見つからない場合はnull
 */
function updateSession(sessionId, sessionData) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sessionSheet = ss.getSheetByName('セッション管理');
  
  if (!sessionSheet) {
    throw new Error('セッション管理シートが見つかりません');
  }
  
  // 既存のセッション情報を取得
  const existingSession = findSessionById(sessionId);
  if (!existingSession) {
    return null; // セッションが見つからない場合
  }
  
  // データの範囲を取得
  const dataRange = sessionSheet.getDataRange();
  const values = dataRange.getValues();
  
  // ヘッダー行を取得
  const headers = values[0];
  
  // セッションID列のインデックスを取得
  const sessionIdIndex = headers.indexOf('セッションID');
  
  if (sessionIdIndex === -1) {
    throw new Error('セッションIDカラムが見つかりません');
  }
  
  // カレンダー更新のために日時変更を追跡
  const isDateTimeChanged = sessionData.予定日時 && 
                          existingSession.予定日時 && 
                          new Date(sessionData.予定日時).getTime() !== new Date(existingSession.予定日時).getTime();
  
  // セッションIDによる検索
  for (let i = 1; i < values.length; i++) {
    if (values[i][sessionIdIndex] === sessionId) {
      // 既存のデータを保持
      const updatedData = { ...existingSession, ...sessionData };
      
      // セッションIDは変更しない
      updatedData['セッションID'] = sessionId;
      
      // 行データを準備
      const rowData = headers.map(header => {
        if (!header) return ''; // 空のヘッダーの場合
        return updatedData[header] !== undefined ? updatedData[header] : '';
      });
      
      // データを更新
      sessionSheet.getRange(i + 1, 1, 1, rowData.length).setValues([rowData]);
      
      // Googleカレンダーの更新が必要な場合
      if (isDateTimeChanged || sessionData.ステータス) {
        updateGoogleCalendarEvent(existingSession, updatedData);
      }
      
      return updatedData;
    }
  }
  
  return null; // セッションが見つからない場合
}

/**
 * Googleカレンダーのイベントを更新
 * @param {Object} oldSession - 更新前のセッション情報
 * @param {Object} newSession - 更新後のセッション情報
 * @return {boolean} 更新が成功したかどうか
 */
function updateGoogleCalendarEvent(oldSession, newSession) {
  try {
    // カレンダーを取得
    const calendar = CalendarApp.getDefaultCalendar();
    
    // クライアント情報を取得
    const client = ClientManager.findClientById(newSession.クライアントID || oldSession.クライアントID);
    if (!client) {
      console.error('クライアント情報が見つかりません');
      return false;
    }
    
    // 予定日時が変更された場合にカレンダーイベントを検索するクエリ
    const searchStart = new Date(oldSession.予定日時);
    searchStart.setHours(searchStart.getHours() - 1); // 1時間前
    
    const searchEnd = new Date(oldSession.予定日時);
    searchEnd.setHours(searchEnd.getHours() + 1); // 1時間後
    
    // カレンダーイベントを検索
    const events = calendar.getEvents(searchStart, searchEnd);
    
    // セッションタイトルに一致するイベントを検索
    const sessionTitle = `【${oldSession.セッション種別}】${client.お名前} 様`;
    let targetEvent = null;
    
    for (let i = 0; i < events.length; i++) {
      if (events[i].getTitle() === sessionTitle) {
        targetEvent = events[i];
        break;
      }
    }
    
    if (!targetEvent) {
      console.error('カレンダーイベントが見つかりません');
      return false;
    }
    
    // ステータスが「キャンセル」または「延期」の場合
    if (newSession.ステータス === 'キャンセル') {
      // イベントを削除
      targetEvent.deleteEvent();
      return true;
    } else if (newSession.ステータス === '延期') {
      // タイトルを更新
      targetEvent.setTitle(`【延期】${sessionTitle}`);
      return true;
    }
    
    // 予定日時が変更された場合
    if (newSession.予定日時 && new Date(newSession.予定日時).getTime() !== new Date(oldSession.予定日時).getTime()) {
      // セッション時間（分）を取得
      const sessionDuration = parseInt(Utilities.getSetting('SESSION_DURATION', '30'));
      
      // 新しい日時で更新
      const newDateTime = new Date(newSession.予定日時);
      const newEndTime = new Date(newDateTime.getTime() + sessionDuration * 60 * 1000);
      
      targetEvent.setTime(newDateTime, newEndTime);
    }
    
    // セッション種別が変更された場合
    if (newSession.セッション種別 && newSession.セッション種別 !== oldSession.セッション種別) {
      const newTitle = `【${newSession.セッション種別}】${client.お名前} 様`;
      targetEvent.setTitle(newTitle);
    }
    
    // Google Meet URLが変更された場合
    if (newSession['Google Meet URL'] && newSession['Google Meet URL'] !== oldSession['Google Meet URL']) {
      targetEvent.setVideoCallLink(newSession['Google Meet URL']);
    }
    
    return true;
  } catch (error) {
    console.error('カレンダー更新エラー:', error);
    return false;
  }
}

/**
 * セッションを実施済みに更新
 * @param {string} sessionId - 更新するセッションID
 * @param {string} notes - セッション記録
 * @return {Object|null} 更新されたセッション情報。セッションが見つからない場合はnull
 */
function completeSession(sessionId, notes) {
  const existingSession = findSessionById(sessionId);
  if (!existingSession) {
    return null; // セッションが見つからない場合
  }
  
  // 実施日時を現在にセット
  const completionTime = new Date();
  
  // セッション情報を更新
  return updateSession(sessionId, {
    ステータス: '実施済',
    実施日時: completionTime,
    記録: notes || existingSession.記録
  });
}

/**
 * 特定の日付のセッションを取得
 * @param {Date} date - 検索する日付
 * @param {boolean} activeOnly - アクティブなセッションのみを取得するか（デフォルト: true）
 * @return {Array<Object>} セッション情報の配列
 */
function getSessionsByDate(date, activeOnly = true) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sessionSheet = ss.getSheetByName('セッション管理');
  
  if (!sessionSheet) {
    throw new Error('セッション管理シートが見つかりません');
  }
  
  // データの範囲を取得
  const dataRange = sessionSheet.getDataRange();
  const values = dataRange.getValues();
  
  // ヘッダー行を取得
  const headers = values[0];
  
  // 日付列とステータス列のインデックスを取得
  const dateIndex = headers.indexOf('予定日時');
  const statusIndex = headers.indexOf('ステータス');
  
  if (dateIndex === -1) {
    throw new Error('予定日時カラムが見つかりません');
  }
  
  // 日付の範囲を設定
  const searchDate = new Date(date);
  searchDate.setHours(0, 0, 0, 0);
  const nextDay = new Date(searchDate);
  nextDay.setDate(nextDay.getDate() + 1);
  
  // セッション情報の配列を作成
  const sessions = [];
  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    
    // 日付が一致するセッションを検索
    if (row[dateIndex] instanceof Date) {
      const sessionDate = new Date(row[dateIndex]);
      if (sessionDate >= searchDate && sessionDate < nextDay) {
        const session = {};
        
        // オブジェクトを構築
        for (let j = 0; j < headers.length; j++) {
          if (headers[j]) { // 空のヘッダーはスキップ
            session[headers[j]] = row[j];
          }
        }
        
        // アクティブなセッションのみを取得する場合
        if (activeOnly) {
          // 「予定」または「延期」のステータスのセッションのみを追加
          if (statusIndex >= 0 && (row[statusIndex] === '予定' || row[statusIndex] === '延期')) {
            sessions.push(session);
          }
        } else {
          sessions.push(session);
        }
      }
    }
  }
  
  return sessions;
}

/**
 * 今日のセッションを取得
 * @param {boolean} activeOnly - アクティブなセッションのみを取得するか（デフォルト: true）
 * @return {Array<Object>} セッション情報の配列
 */
function getTodaySessions(activeOnly = true) {
  return getSessionsByDate(new Date(), activeOnly);
}

/**
 * 今週のセッションを取得
 * @param {boolean} activeOnly - アクティブなセッションのみを取得するか（デフォルト: true）
 * @return {Array<Object>} セッション情報の配列
 */
function getThisWeekSessions(activeOnly = true) {
  const today = new Date();
  const startOfWeek = new Date(today);
  startOfWeek.setDate(today.getDate() - today.getDay()); // 日曜日に設定
  startOfWeek.setHours(0, 0, 0, 0);
  
  const endOfWeek = new Date(startOfWeek);
  endOfWeek.setDate(startOfWeek.getDate() + 7); // 来週の日曜日に設定
  
  const sessions = getAllSessions(activeOnly);
  
  // 今週のセッションをフィルタリング
  return sessions.filter(session => {
    const sessionDate = new Date(session.予定日時);
    return sessionDate >= startOfWeek && sessionDate < endOfWeek;
  });
}

/**
 * 予定されたセッションをGoogle Meetリマインダーとして送信
 * @param {number} daysAhead - 何日前のセッションを対象とするか
 * @return {number} 送信されたリマインダーの数
 */
function sendSessionReminders(daysAhead = 1) {
  // 対象日を計算
  const targetDate = new Date();
  targetDate.setDate(targetDate.getDate() + daysAhead);
  
  // 対象日のセッションを取得
  const sessions = getSessionsByDate(targetDate, true);
  
  let sentCount = 0;
  
  // メール送信用モジュールがあると仮定
  // const EmailManager = require('../email/emailManager.js');
  
  for (const session of sessions) {
    // クライアント情報を取得
    const client = ClientManager.findClientById(session.クライアントID);
    if (!client) continue;
    
    // メール送信（実際の実装では EmailManager を使用）
    try {
      // リマインダーメールを送信
      // EmailManager.sendSessionReminder(session, client);
      
      console.log(`リマインダーを送信: ${client.お名前} 様 (${Utilities.formatDateTime(session.予定日時)})`);
      sentCount++;
    } catch (error) {
      console.error(`リマインダー送信エラー: ${error.message}`);
    }
  }
  
  return sentCount;
}

// モジュールをエクスポート
const SessionManager = {
  createSession,
  getAllSessions,
  findSessionsByClientId,
  findSessionById,
  updateSession,
  completeSession,
  getSessionsByDate,
  getTodaySessions,
  getThisWeekSessions,
  sendSessionReminders,
  addToGoogleCalendar,
  updateGoogleCalendarEvent
};