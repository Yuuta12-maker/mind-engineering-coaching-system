/**
 * マインドエンジニアリング・コーチング業務管理システム
 * クライアント管理モジュール
 * 
 * クライアント情報の作成、検索、更新、削除（CRUD操作）を管理します。
 */

/**
 * クライアント情報を作成
 * @param {Object} clientData - クライアント情報のオブジェクト
 * @return {Object} 作成されたクライアント情報（IDを含む）
 */
function createClient(clientData) {
  try {
    // 設定を取得
    const config = getConfig();
    
    // クライアントinfoシートを取得
    const clientSheet = getSheet(config.SHEET_NAMES.CLIENT);
    
    if (!clientSheet) {
      throw new Error('クライアントinfoシートが見つかりません');
    }
    
    // 必須フィールドをチェック
    if (!clientData.メールアドレス || !clientData.お名前) {
      throw new Error('メールアドレスと名前は必須項目です');
    }
    
    // メールアドレスの重複チェック
    if (findClientByEmail(clientData.メールアドレス)) {
      throw new Error('指定されたメールアドレスは既に登録されています');
    }
    
    // クライアントIDを生成
    const clientId = generateId('CLIENT');
    
    // 現在のタイムスタンプを取得
    const timestamp = getNowJST();
    
    // 行データを準備
    const rowData = [
      clientId,
      timestamp,
      clientData.メールアドレス,
      clientData.お名前,
      clientData['お名前　（カナ）'] || '',
      clientData.性別 || '',
      clientData.生年月日 || '',
      clientData['電話番号　（ハイフンなし）'] || '',
      clientData.ご住所 || '',
      clientData.希望セッション形式 || '未定',
      clientData.ステータス || '問い合わせ',
      clientData.備考欄 || ''
    ];
    
    // 最終行の次の行に追加
    const lastRow = clientSheet.getLastRow();
    clientSheet.getRange(lastRow + 1, 1, 1, rowData.length).setValues([rowData]);
    
    // 作成したクライアント情報を返す
    return {
      クライアントID: clientId,
      タイムスタンプ: timestamp,
      ...clientData
    };
  } catch (error) {
    // エラーログを記録
    logMessage(`クライアント作成中にエラーが発生しました: ${error.message}`, 'ERROR');
    
    // エラーを呼び出し元に伝播
    throw error;
  }
}

/**
 * すべてのクライアント情報を取得
 * @param {boolean} activeOnly - アクティブなクライアントのみを取得するか（デフォルト: false）
 * @return {Array<Object>} クライアント情報の配列
 */
function getAllClients(activeOnly = false) {
  try {
    // 設定を取得
    const config = getConfig();
    
    // クライアントinfoシートを取得
    const clientSheet = getSheet(config.SHEET_NAMES.CLIENT);
    
    if (!clientSheet) {
      throw new Error('クライアントinfoシートが見つかりません');
    }
    
    // データの範囲を取得
    const dataRange = clientSheet.getDataRange();
    const values = dataRange.getValues();
    
    // ヘッダー行を取得
    const headers = values[0];
    
    // ステータス列のインデックスを取得
    const statusIndex = headers.indexOf('ステータス');
    
    // クライアント情報の配列を作成
    const clients = [];
    for (let i = 1; i < values.length; i++) {
      const row = values[i];
      
      // 行が空でないことを確認
      if (row[0]) {
        const client = {};
        
        // オブジェクトを構築
        for (let j = 0; j < headers.length; j++) {
          if (headers[j]) { // 空のヘッダーはスキップ
            client[headers[j]] = row[j];
          }
        }
        
        // アクティブなクライアントのみを取得する場合
        if (activeOnly) {
          // 「完了」または「中断」以外のステータスのクライアントのみを追加
          if (statusIndex >= 0 && row[statusIndex] !== '完了' && row[statusIndex] !== '中断') {
            clients.push(client);
          }
        } else {
          clients.push(client);
        }
      }
    }
    
    return clients;
  } catch (error) {
    // エラーログを記録
    logMessage(`クライアント一覧取得中にエラーが発生しました: ${error.message}`, 'ERROR');
    
    // エラーを呼び出し元に伝播
    throw error;
  }
}

/**
 * クライアントIDによるクライアント情報の検索
 * @param {string} clientId - 検索するクライアントID
 * @return {Object|null} クライアント情報。見つからない場合はnull
 */
function findClientById(clientId) {
  try {
    // 設定を取得
    const config = getConfig();
    
    // クライアントinfoシートを取得
    const clientSheet = getSheet(config.SHEET_NAMES.CLIENT);
    
    if (!clientSheet) {
      throw new Error('クライアントinfoシートが見つかりません');
    }
    
    // データの範囲を取得
    const dataRange = clientSheet.getDataRange();
    const values = dataRange.getValues();
    
    // ヘッダー行を取得
    const headers = values[0];
    
    // クライアントIDの列インデックスを取得
    const idIndex = headers.indexOf('クライアントID');
    if (idIndex === -1) {
      throw new Error('クライアントIDカラムが見つかりません');
    }
    
    // クライアントIDによる検索
    for (let i = 1; i < values.length; i++) {
      if (values[i][idIndex] === clientId) {
        const client = {};
        
        // オブジェクトを構築
        for (let j = 0; j < headers.length; j++) {
          if (headers[j]) { // 空のヘッダーはスキップ
            client[headers[j]] = values[i][j];
          }
        }
        
        return client;
      }
    }
    
    return null; // クライアントが見つからない場合
  } catch (error) {
    // エラーログを記録
    logMessage(`クライアント検索中にエラーが発生しました: ${error.message}`, 'ERROR');
    
    // エラーを呼び出し元に伝播
    throw error;
  }
}

/**
 * メールアドレスによるクライアント情報の検索
 * @param {string} email - 検索するメールアドレス
 * @return {Object|null} クライアント情報。見つからない場合はnull
 */
function findClientByEmail(email) {
  try {
    // 設定を取得
    const config = getConfig();
    
    // クライアントinfoシートを取得
    const clientSheet = getSheet(config.SHEET_NAMES.CLIENT);
    
    if (!clientSheet) {
      throw new Error('クライアントinfoシートが見つかりません');
    }
    
    // データの範囲を取得
    const dataRange = clientSheet.getDataRange();
    const values = dataRange.getValues();
    
    // ヘッダー行を取得
    const headers = values[0];
    
    // メールアドレス列のインデックスを取得
    const emailIndex = headers.indexOf('メールアドレス');
    if (emailIndex === -1) {
      throw new Error('メールアドレスカラムが見つかりません');
    }
    
    // メールアドレスによる検索
    for (let i = 1; i < values.length; i++) {
      if (values[i][emailIndex] === email) {
        const client = {};
        
        // オブジェクトを構築
        for (let j = 0; j < headers.length; j++) {
          if (headers[j]) { // 空のヘッダーはスキップ
            client[headers[j]] = values[i][j];
          }
        }
        
        return client;
      }
    }
    
    return null; // クライアントが見つからない場合
  } catch (error) {
    // エラーログを記録
    logMessage(`メールアドレスでのクライアント検索中にエラーが発生しました: ${error.message}`, 'ERROR');
    
    // エラーを呼び出し元に伝播
    throw error;
  }
}

/**
 * クライアント情報を更新
 * @param {string} clientId - 更新するクライアントID
 * @param {Object} clientData - 更新するクライアント情報
 * @return {Object|null} 更新されたクライアント情報。クライアントが見つからない場合はnull
 */
function updateClient(clientId, clientData) {
  try {
    // 設定を取得
    const config = getConfig();
    
    // クライアントinfoシートを取得
    const clientSheet = getSheet(config.SHEET_NAMES.CLIENT);
    
    if (!clientSheet) {
      throw new Error('クライアントinfoシートが見つかりません');
    }
    
    // データの範囲を取得
    const dataRange = clientSheet.getDataRange();
    const values = dataRange.getValues();
    
    // ヘッダー行を取得
    const headers = values[0];
    
    // クライアントIDの列インデックスを取得
    const idIndex = headers.indexOf('クライアントID');
    if (idIndex === -1) {
      throw new Error('クライアントIDカラムが見つかりません');
    }
    
    // クライアントIDによる検索
    for (let i = 1; i < values.length; i++) {
      if (values[i][idIndex] === clientId) {
        // 既存のデータを保持
        const existingData = {};
        for (let j = 0; j < headers.length; j++) {
          if (headers[j]) {
            existingData[headers[j]] = values[i][j];
          }
        }
        
        // 更新するデータをマージ
        const updatedData = { ...existingData, ...clientData };
        
        // クライアントIDとタイムスタンプは変更しない
        updatedData['クライアントID'] = clientId;
        updatedData['タイムスタンプ'] = existingData['タイムスタンプ'];
        
        // 行データを準備
        const rowData = headers.map(header => {
          if (!header) return ''; // 空のヘッダーの場合
          return updatedData[header] !== undefined ? updatedData[header] : '';
        });
        
        // データを更新
        clientSheet.getRange(i + 1, 1, 1, rowData.length).setValues([rowData]);
        
        return updatedData;
      }
    }
    
    return null; // クライアントが見つからない場合
  } catch (error) {
    // エラーログを記録
    logMessage(`クライアント更新中にエラーが発生しました: ${error.message}`, 'ERROR');
    
    // エラーを呼び出し元に伝播
    throw error;
  }
}

/**
 * クライアント情報を削除
 * @param {string} clientId - 削除するクライアントID
 * @return {boolean} 削除が成功したかどうか
 */
function deleteClient(clientId) {
  try {
    // 設定を取得
    const config = getConfig();
    
    // クライアントinfoシートを取得
    const clientSheet = getSheet(config.SHEET_NAMES.CLIENT);
    
    if (!clientSheet) {
      throw new Error('クライアントinfoシートが見つかりません');
    }
    
    // データの範囲を取得
    const dataRange = clientSheet.getDataRange();
    const values = dataRange.getValues();
    
    // ヘッダー行を取得
    const headers = values[0];
    
    // クライアントIDの列インデックスを取得
    const idIndex = headers.indexOf('クライアントID');
    if (idIndex === -1) {
      throw new Error('クライアントIDカラムが見つかりません');
    }
    
    // クライアントIDによる検索
    for (let i = 1; i < values.length; i++) {
      if (values[i][idIndex] === clientId) {
        // 行を削除
        clientSheet.deleteRow(i + 1);
        return true;
      }
    }
    
    return false; // クライアントが見つからない場合
  } catch (error) {
    // エラーログを記録
    logMessage(`クライアント削除中にエラーが発生しました: ${error.message}`, 'ERROR');
    
    // エラーを呼び出し元に伝播
    throw error;
  }
}

/**
 * クライアントのステータスを更新
 * @param {string} clientId - 更新するクライアントID
 * @param {string} newStatus - 新しいステータス
 * @return {boolean} 更新が成功したかどうか
 */
function updateClientStatus(clientId, newStatus) {
  try {
    // 有効なステータス値をチェック
    const validStatuses = ['問い合わせ', 'トライアル前', 'トライアル済', '契約中', '完了', '中断'];
    if (!validStatuses.includes(newStatus)) {
      throw new Error('無効なステータス値です');
    }
    
    // クライアント情報を取得
    const client = findClientById(clientId);
    if (!client) {
      return false; // クライアントが見つからない場合
    }
    
    // ステータスを更新
    return updateClient(clientId, { ステータス: newStatus }) !== null;
  } catch (error) {
    // エラーログを記録
    logMessage(`クライアントステータス更新中にエラーが発生しました: ${error.message}`, 'ERROR');
    
    // エラーを呼び出し元に伝播
    throw error;
  }
}

/**
 * 指定したステータスのクライアント数を取得
 * @param {string} status - ステータス
 * @return {number} クライアント数
 */
function countClientsByStatus(status) {
  try {
    // 設定を取得
    const config = getConfig();
    
    // クライアントinfoシートを取得
    const clientSheet = getSheet(config.SHEET_NAMES.CLIENT);
    
    if (!clientSheet) {
      throw new Error('クライアントinfoシートが見つかりません');
    }
    
    // データの範囲を取得
    const dataRange = clientSheet.getDataRange();
    const values = dataRange.getValues();
    
    // ヘッダー行を取得
    const headers = values[0];
    
    // ステータス列のインデックスを取得
    const statusIndex = headers.indexOf('ステータス');
    if (statusIndex === -1) {
      throw new Error('ステータスカラムが見つかりません');
    }
    
    // ステータスごとのカウント
    let count = 0;
    for (let i = 1; i < values.length; i++) {
      if (values[i][statusIndex] === status) {
        count++;
      }
    }
    
    return count;
  } catch (error) {
    // エラーログを記録
    logMessage(`クライアント集計中にエラーが発生しました: ${error.message}`, 'ERROR');
    
    // エラーを呼び出し元に伝播
    throw error;
  }
}

/**
 * クライアントのステータス別集計を取得
 * @return {Object} ステータス別のクライアント数
 */
function getClientStatusSummary() {
  try {
    const statuses = ['問い合わせ', 'トライアル前', 'トライアル済', '契約中', '完了', '中断'];
    const summary = {};
    
    statuses.forEach(status => {
      summary[status] = countClientsByStatus(status);
    });
    
    return summary;
  } catch (error) {
    // エラーログを記録
    logMessage(`クライアント統計取得中にエラーが発生しました: ${error.message}`, 'ERROR');
    
    // エラーを呼び出し元に伝播
    throw error;
  }
}

/**
 * Google Formsからクライアント情報を取り込む
 * @param {Object} formResponse - フォーム送信データ
 * @return {Object} 作成されたクライアント情報
 */
function importClientFromForm(formResponse) {
  try {
    const clientData = {
      メールアドレス: formResponse.メールアドレス,
      お名前: formResponse.お名前,
      'お名前　（カナ）': formResponse['お名前（カナ）'] || '',
      性別: formResponse.性別 || '',
      生年月日: formResponse.生年月日 || '',
      '電話番号　（ハイフンなし）': String(formResponse.電話番号).replace(/[- ]/g, ''),
      ご住所: formResponse.ご住所 || '',
      希望セッション形式: formResponse.希望セッション形式 || '未定',
      ステータス: '問い合わせ',
      備考欄: formResponse.備考 || ''
    };
    
    return createClient(clientData);
  } catch (error) {
    // エラーログを記録
    logMessage(`フォームからのクライアント作成中にエラーが発生しました: ${error.message}`, 'ERROR');
    
    // エラーを呼び出し元に伝播
    throw error;
  }
}
