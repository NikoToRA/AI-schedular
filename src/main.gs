/**
 * AIスケジューラ メインファイル
 * Google Apps Script で動作するAIベースのスケジュール管理システム
 */

// 設定定数
const CONFIG = {
  CALENDAR_ID: 'primary',
  SPREADSHEET_ID: '', // 使用するスプレッドシートのIDを設定
  AI_API_KEY: '', // AI APIキーを設定
  NOTIFICATION_EMAIL: '', // 通知先メールアドレス
  TIME_ZONE: 'Asia/Tokyo'
};

/**
 * メイン実行関数
 * トリガーから実行される
 */
function main() {
  try {
    Logger.log('AIスケジューラ開始: ' + new Date());

    // 設定チェック
    if (!validateConfig()) {
      throw new Error('設定が正しく設定されていません');
    }

    // カレンダーイベント取得
    const events = getCalendarEvents();
    Logger.log(`取得イベント数: ${events.length}`);

    // AI分析実行
    const analysis = analyzeScheduleWithAI(events);

    // 結果をスプレッドシートに保存
    saveAnalysisResults(analysis);

    // 必要に応じて通知送信
    if (analysis.conflicts && analysis.conflicts.length > 0) {
      sendNotification(analysis);
    }

    Logger.log('AIスケジューラ完了: ' + new Date());

  } catch (error) {
    Logger.log(`エラー: ${error.toString()}`);
    sendErrorNotification(error);
  }
}

/**
 * 設定の検証
 */
function validateConfig() {
  // 必要な設定項目のチェック
  if (!CONFIG.SPREADSHEET_ID) {
    Logger.log('SPREADSHEET_IDが設定されていません');
    return false;
  }

  if (!CONFIG.AI_API_KEY) {
    Logger.log('AI_API_KEYが設定されていません');
    return false;
  }

  return true;
}

/**
 * 手動実行用のテスト関数
 */
function testRun() {
  main();
}

/**
 * 初期設定用関数
 */
function initialize() {
  Logger.log('初期設定を開始します...');

  // スプレッドシートの作成
  const spreadsheet = createManagementSpreadsheet();
  Logger.log(`管理用スプレッドシートが作成されました: ${spreadsheet.getUrl()}`);

  // 設定シートの初期化
  initializeSettingsSheet(spreadsheet);

  // トリガーの設定
  setupTriggers();

  Logger.log('初期設定が完了しました');
}