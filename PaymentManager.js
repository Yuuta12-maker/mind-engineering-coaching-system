/**
 * マインドエンジニアリング・コーチング管理システム
 * 支払い管理モジュール
 * 
 * 支払い情報の作成、検索、更新、領収書生成などの機能を提供します。
 */

// ユーティリティ関数とクライアント管理モジュールをインポート
// 実際のGASでは直接利用できるため、インポート文は必要ないですが
// コードの依存関係を明確にするために記述しています
// const Utilities = require('../utils/utilities.js');
// const ClientManager = require('../client/clientManager.js');

/**
 * 支払い情報を作成
 * @param {Object} paymentData - 支払い情報のオブジェクト
 * @return {Object} 作成された支払い情報（IDを含む）
 */
function createPayment(paymentData) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const paymentSheet = ss.getSheetByName('支払い管理');
  
  if (!paymentSheet) {
    throw new Error('支払い管理シートが見つかりません');
  }
  
  // 必須フィールドをチェック
  if (!paymentData.クライアントID || !paymentData.項目 || !paymentData.金額) {
    throw new Error('クライアントID、項目、金額は必須項目です');
  }
  
  // クライアントの存在確認
  const client = ClientManager.findClientById(paymentData.クライアントID);
  if (!client) {
    throw new Error('指定されたクライアントIDは存在しません');
  }
  
  // 支払いIDを生成
  const paymentId = Utilities.generateUniqueId('PY');
  
  // 登録日を設定
  const registrationDate = paymentData.登録日 ? new Date(paymentData.登録日) : new Date();
  
  // 金額を数値に変換
  const amount = typeof paymentData.金額 === 'string' ? 
                parseInt(paymentData.金額.replace(/[^\d]/g, '')) : 
                paymentData.金額;
  
  // 行データを準備
  const rowData = [
    paymentId,
    paymentData.クライアントID,
    registrationDate,
    paymentData.項目,
    amount,
    paymentData.状態 || '未入金',
    paymentData.入金日 || '',
    paymentData.領収書発行状態 || '未発行',
    paymentData.備考 || ''
  ];
  
  // 最終行の次の行に追加
  const lastRow = Utilities.getLastRow(paymentSheet);
  paymentSheet.getRange(lastRow + 1, 1, 1, rowData.length).setValues([rowData]);
  
  // 作成した支払い情報を返す
  return {
    支払いID: paymentId,
    クライアントID: paymentData.クライアントID,
    登録日: registrationDate,
    項目: paymentData.項目,
    金額: amount,
    状態: rowData[5],
    入金日: rowData[6],
    領収書発行状態: rowData[7],
    備考: rowData[8]
  };
}

/**
 * すべての支払い情報を取得
 * @param {boolean} unpaidOnly - 未入金の支払いのみを取得するか（デフォルト: false）
 * @return {Array<Object>} 支払い情報の配列
 */
function getAllPayments(unpaidOnly = false) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const paymentSheet = ss.getSheetByName('支払い管理');
  
  if (!paymentSheet) {
    throw new Error('支払い管理シートが見つかりません');
  }
  
  // データの範囲を取得
  const dataRange = paymentSheet.getDataRange();
  const values = dataRange.getValues();
  
  // ヘッダー行を取得
  const headers = values[0];
  
  // 状態列のインデックスを取得
  const statusIndex = headers.indexOf('状態');
  
  // 支払い情報の配列を作成
  const payments = [];
  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    
    // 行が空でないことを確認
    if (row[0]) {
      const payment = {};
      
      // オブジェクトを構築
      for (let j = 0; j < headers.length; j++) {
        if (headers[j]) { // 空のヘッダーはスキップ
          payment[headers[j]] = row[j];
        }
      }
      
      // 未入金の支払いのみを取得する場合
      if (unpaidOnly) {
        // 「未入金」の状態の支払いのみを追加
        if (statusIndex >= 0 && row[statusIndex] === '未入金') {
          payments.push(payment);
        }
      } else {
        payments.push(payment);
      }
    }
  }
  
  return payments;
}

/**
 * クライアントIDによる支払い情報の検索
 * @param {string} clientId - 検索するクライアントID
 * @return {Array<Object>} 支払い情報の配列
 */
function findPaymentsByClientId(clientId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const paymentSheet = ss.getSheetByName('支払い管理');
  
  if (!paymentSheet) {
    throw new Error('支払い管理シートが見つかりません');
  }
  
  // データの範囲を取得
  const dataRange = paymentSheet.getDataRange();
  const values = dataRange.getValues();
  
  // ヘッダー行を取得
  const headers = values[0];
  
  // クライアントID列のインデックスを取得
  const clientIdIndex = headers.indexOf('クライアントID');
  
  if (clientIdIndex === -1) {
    throw new Error('クライアントIDカラムが見つかりません');
  }
  
  // 支払い情報の配列を作成
  const payments = [];
  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    
    // クライアントIDが一致する支払いを検索
    if (row[clientIdIndex] === clientId) {
      const payment = {};
      
      // オブジェクトを構築
      for (let j = 0; j < headers.length; j++) {
        if (headers[j]) { // 空のヘッダーはスキップ
          payment[headers[j]] = row[j];
        }
      }
      
      payments.push(payment);
    }
  }
  
  return payments;
}

/**
 * 支払いIDによる支払い情報の検索
 * @param {string} paymentId - 検索する支払いID
 * @return {Object|null} 支払い情報。見つからない場合はnull
 */
function findPaymentById(paymentId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const paymentSheet = ss.getSheetByName('支払い管理');
  
  if (!paymentSheet) {
    throw new Error('支払い管理シートが見つかりません');
  }
  
  // データの範囲を取得
  const dataRange = paymentSheet.getDataRange();
  const values = dataRange.getValues();
  
  // ヘッダー行を取得
  const headers = values[0];
  
  // 支払いID列のインデックスを取得
  const paymentIdIndex = headers.indexOf('支払いID');
  
  if (paymentIdIndex === -1) {
    throw new Error('支払いIDカラムが見つかりません');
  }
  
  // 支払いIDによる検索
  for (let i = 1; i < values.length; i++) {
    if (values[i][paymentIdIndex] === paymentId) {
      const payment = {};
      
      // オブジェクトを構築
      for (let j = 0; j < headers.length; j++) {
        if (headers[j]) { // 空のヘッダーはスキップ
          payment[headers[j]] = values[i][j];
        }
      }
      
      return payment;
    }
  }
  
  return null; // 支払いが見つからない場合
}

/**
 * 支払い情報を更新
 * @param {string} paymentId - 更新する支払いID
 * @param {Object} paymentData - 更新する支払い情報
 * @return {Object|null} 更新された支払い情報。支払いが見つからない場合はnull
 */
function updatePayment(paymentId, paymentData) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const paymentSheet = ss.getSheetByName('支払い管理');
  
  if (!paymentSheet) {
    throw new Error('支払い管理シートが見つかりません');
  }
  
  // 既存の支払い情報を取得
  const existingPayment = findPaymentById(paymentId);
  if (!existingPayment) {
    return null; // 支払いが見つからない場合
  }
  
  // データの範囲を取得
  const dataRange = paymentSheet.getDataRange();
  const values = dataRange.getValues();
  
  // ヘッダー行を取得
  const headers = values[0];
  
  // 支払いID列のインデックスを取得
  const paymentIdIndex = headers.indexOf('支払いID');
  
  if (paymentIdIndex === -1) {
    throw new Error('支払いIDカラムが見つかりません');
  }
  
  // 金額を数値に変換（金額が更新される場合）
  if (paymentData.金額) {
    paymentData.金額 = typeof paymentData.金額 === 'string' ? 
                      parseInt(paymentData.金額.replace(/[^\d]/g, '')) : 
                      paymentData.金額;
  }
  
  // 支払いIDによる検索
  for (let i = 1; i < values.length; i++) {
    if (values[i][paymentIdIndex] === paymentId) {
      // 既存のデータを保持
      const updatedData = { ...existingPayment, ...paymentData };
      
      // 支払いIDは変更しない
      updatedData['支払いID'] = paymentId;
      
      // 行データを準備
      const rowData = headers.map(header => {
        if (!header) return ''; // 空のヘッダーの場合
        return updatedData[header] !== undefined ? updatedData[header] : '';
      });
      
      // データを更新
      paymentSheet.getRange(i + 1, 1, 1, rowData.length).setValues([rowData]);
      
      return updatedData;
    }
  }
  
  return null; // 支払いが見つからない場合
}

/**
 * 支払いを入金済みとして更新
 * @param {string} paymentId - 更新する支払いID
 * @param {Date} paymentDate - 入金日（デフォルト: 現在の日付）
 * @return {Object|null} 更新された支払い情報。支払いが見つからない場合はnull
 */
function markAsPaid(paymentId, paymentDate = new Date()) {
  // 既存の支払い情報を取得
  const existingPayment = findPaymentById(paymentId);
  if (!existingPayment) {
    return null; // 支払いが見つからない場合
  }
  
  // 支払い情報を更新
  return updatePayment(paymentId, {
    状態: '入金済',
    入金日: paymentDate
  });
}

/**
 * トライアルセッションの支払いを作成
 * @param {string} clientId - クライアントID
 * @return {Object} 作成された支払い情報
 */
function createTrialPayment(clientId) {
  // トライアル料金を取得
  const trialFee = parseInt(Utilities.getSetting('TRIAL_FEE', '6000'));
  
  // 支払い情報を作成
  return createPayment({
    クライアントID: clientId,
    項目: 'トライアルセッション',
    金額: trialFee,
    状態: '未入金',
    領収書発行状態: '未発行'
  });
}

/**
 * 継続セッションの支払いを作成
 * @param {string} clientId - クライアントID
 * @return {Object} 作成された支払い情報
 */
function createContinuationPayment(clientId) {
  // 継続セッション料金を取得
  const continuationFee = parseInt(Utilities.getSetting('CONTINUATION_FEE', '214000'));
  
  // 支払い情報を作成
  return createPayment({
    クライアントID: clientId,
    項目: '継続セッション（全5回）',
    金額: continuationFee,
    状態: '未入金',
    領収書発行状態: '未発行'
  });
}

/**
 * 領収書を生成
 * @param {string} paymentId - 支払いID
 * @return {Object} 領収書情報 {fileId: string, url: string}
 */
function generateReceipt(paymentId) {
  try {
    // 支払い情報を取得
    const payment = findPaymentById(paymentId);
    if (!payment) {
      throw new Error('指定された支払いIDは存在しません');
    }
    
    // 入金済みかチェック
    if (payment.状態 !== '入金済') {
      throw new Error('入金済みの支払いのみ領収書を発行できます');
    }
    
    // クライアント情報を取得
    const client = ClientManager.findClientById(payment.クライアントID);
    if (!client) {
      throw new Error('クライアント情報が見つかりません');
    }
    
    // 領収書テンプレートを取得
    // 実際の環境ではテンプレートIDを設定から取得するなどの方法がある
    const templateId = '1PGAGXZWHSMFjc0E0JUx-FNYC6rIqhxThHNBMcXy7Xqc'; // 仮のID
    
    try {
      // テンプレートファイルを開く
      const templateFile = DriveApp.getFileById(templateId);
      
      // 新しい領収書ファイル名を生成
      const receiptFileName = `領収書_${client.お名前}_${Utilities.formatDate(new Date())}.pdf`;
      
      // テンプレートをコピー
      const tempFolder = DriveApp.createFolder('temp_receipt_' + Utilities.generateUniqueId(''));
      const tempFile = templateFile.makeCopy(receiptFileName, tempFolder);
      
      // Docsとして開く
      const doc = DocumentApp.openById(tempFile.getId());
      const body = doc.getBody();
      
      // 領収書番号を生成
      const receiptNumber = `R${new Date().getFullYear()}${payment.支払いID.substring(2, 8)}`;
      
      // テンプレートの置換
      body.replaceText('{{領収書番号}}', receiptNumber);
      body.replaceText('{{日付}}', Utilities.formatDate(payment.入金日 || new Date()));
      body.replaceText('{{宛名}}', client.お名前);
      body.replaceText('{{金額}}', payment.金額.toLocaleString());
      body.replaceText('{{但し書き}}', payment.項目);
      body.replaceText('{{住所}}', Utilities.getSetting('BUSINESS_ADDRESS', ''));
      body.replaceText('{{電話番号}}', Utilities.getSetting('BUSINESS_PHONE', ''));
      body.replaceText('{{発行者名}}', 'マインドエンジニアリング・コーチング 森山雄太');
      
      // 変更を保存
      doc.saveAndClose();
      
      // PDFに変換
      const pdfBlob = tempFile.getAs('application/pdf');
      
      // クライアントのフォルダを取得または作成
      let clientFolder;
      try {
        clientFolder = DriveApp.getFoldersByName(`MEC_${payment.クライアントID}_${client.お名前}`).next();
      } catch (e) {
        // フォルダが存在しない場合、作成
        clientFolder = DriveApp.createFolder(`MEC_${payment.クライアントID}_${client.お名前}`);
      }
      
      // PDFをクライアントフォルダに保存
      const pdfFile = clientFolder.createFile(pdfBlob);
      
      // 一時フォルダを削除
      tempFolder.setTrashed(true);
      
      // 支払い情報を更新
      updatePayment(paymentId, {
        領収書発行状態: '発行済'
      });
      
      // 結果を返す
      return {
        fileId: pdfFile.getId(),
        url: pdfFile.getUrl()
      };
    } catch (error) {
      console.error('領収書生成エラー:', error);
      throw new Error('領収書の生成中にエラーが発生しました: ' + error.message);
    }
  } catch (error) {
    console.error('領収書生成エラー:', error);
    throw error;
  }
}

/**
 * 月別の売上合計を取得
 * @param {number} year - 年（デフォルト: 現在の年）
 * @param {number} month - 月（デフォルト: 現在の月）
 * @return {number} 売上合計
 */
function getMonthlySales(year = new Date().getFullYear(), month = new Date().getMonth() + 1) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const paymentSheet = ss.getSheetByName('支払い管理');
  
  if (!paymentSheet) {
    throw new Error('支払い管理シートが見つかりません');
  }
  
  // データの範囲を取得
  const dataRange = paymentSheet.getDataRange();
  const values = dataRange.getValues();
  
  // ヘッダー行を取得
  const headers = values[0];
  
  // 入金日と金額列のインデックスを取得
  const paymentDateIndex = headers.indexOf('入金日');
  const amountIndex = headers.indexOf('金額');
  const statusIndex = headers.indexOf('状態');
  
  if (paymentDateIndex === -1 || amountIndex === -1 || statusIndex === -1) {
    throw new Error('必要なカラムが見つかりません');
  }
  
  // 対象年月の範囲を設定
  const startDate = new Date(year, month - 1, 1);
  const endDate = new Date(year, month, 0);
  
  // 売上合計を計算
  let totalSales = 0;
  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    
    // 入金済みの支払いのみ集計
    if (row[statusIndex] === '入金済' && row[paymentDateIndex] instanceof Date) {
      const paymentDate = row[paymentDateIndex];
      
      // 対象年月の支払いのみ集計
      if (paymentDate >= startDate && paymentDate <= endDate) {
        totalSales += Number(row[amountIndex]);
      }
    }
  }
  
  return totalSales;
}

/**
 * 年間の月別売上データを取得
 * @param {number} year - 年（デフォルト: 現在の年）
 * @return {Array<Object>} 月別売上データの配列
 */
function getYearlySalesData(year = new Date().getFullYear()) {
  const monthlySales = [];
  
  for (let month = 1; month <= 12; month++) {
    monthlySales.push({
      month: month,
      monthName: ['1月', '2月', '3月', '4月', '5月', '6月', '7月', '8月', '9月', '10月', '11月', '12月'][month - 1],
      sales: getMonthlySales(year, month)
    });
  }
  
  return monthlySales;
}

/**
 * 支払い情報を削除
 * @param {string} paymentId - 削除する支払いID
 * @return {boolean} 削除が成功したかどうか
 */
function deletePayment(paymentId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const paymentSheet = ss.getSheetByName('支払い管理');
  
  if (!paymentSheet) {
    throw new Error('支払い管理シートが見つかりません');
  }
  
  // データの範囲を取得
  const dataRange = paymentSheet.getDataRange();
  const values = dataRange.getValues();
  
  // ヘッダー行を取得
  const headers = values[0];
  
  // 支払いID列のインデックスを取得
  const paymentIdIndex = headers.indexOf('支払いID');
  
  if (paymentIdIndex === -1) {
    throw new Error('支払いIDカラムが見つかりません');
  }
  
  // 支払いIDによる検索
  for (let i = 1; i < values.length; i++) {
    if (values[i][paymentIdIndex] === paymentId) {
      // 行を削除
      paymentSheet.deleteRow(i + 1);
      return true;
    }
  }
  
  return false; // 支払いが見つからない場合
}

/**
 * 新しい入金を確認し、必要なアクションを実行
 * @param {boolean} sendEmail - 入金確認メールを送信するかどうか（デフォルト: true）
 * @return {Array<Object>} 確認された入金情報の配列
 */
function checkNewPayments(sendEmail = true) {
  // 未入金の支払いを取得
  const unpaidPayments = getAllPayments(true);
  
  // 入金が確認された支払い情報の配列
  const confirmedPayments = [];
  
  // メール送信用モジュールがあると仮定
  // const EmailManager = require('../email/emailManager.js');
  
  // 各未入金の支払いについて
  for (const payment of unpaidPayments) {
    // ここで実際には、入金の確認処理を行う
    // 例: API経由での銀行口座の確認や手動入力された入金情報との照合など
    
    // 入金確認のシミュレーション
    // 実際の実装では、この部分は適切な確認ロジックに置き換える
    const isPaymentConfirmed = false; // デフォルトでは確認されない
    
    if (isPaymentConfirmed) {
      // 入金済みとして更新
      const updatedPayment = markAsPaid(payment.支払いID);
      
      if (updatedPayment) {
        confirmedPayments.push(updatedPayment);
        
        // クライアント情報を取得
        const client = ClientManager.findClientById(payment.クライアントID);
        
        // 入金確認メールを送信
        if (sendEmail && client) {
          try {
            // EmailManager.sendPaymentConfirmation(client, updatedPayment);
            console.log(`入金確認メールを送信: ${client.お名前} 様 (${updatedPayment.金額.toLocaleString()}円)`);
          } catch (error) {
            console.error(`メール送信エラー: ${error.message}`);
          }
        }
      }
    }
  }
  
  return confirmedPayments;
}

// モジュールをエクスポート
const PaymentManager = {
  createPayment,
  getAllPayments,
  findPaymentsByClientId,
  findPaymentById,
  updatePayment,
  markAsPaid,
  createTrialPayment,
  createContinuationPayment,
  generateReceipt,
  getMonthlySales,
  getYearlySalesData,
  deletePayment,
  checkNewPayments
};