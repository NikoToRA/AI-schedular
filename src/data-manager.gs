/**
 * データ管理モジュール
 * Googleスプレッドシートでのデータ保存・取得を行う
 */

/**
 * 管理用スプレッドシートを作成する
 * @return {Spreadsheet} 作成されたスプレッドシート
 */
function createManagementSpreadsheet() {
  try {
    const spreadsheet = SpreadsheetApp.create('AIスケジューラ管理データ');

    // 各シートを作成
    createAnalysisResultsSheet(spreadsheet);
    createSettingsSheet(spreadsheet);
    createLogSheet(spreadsheet);
    createConflictsSheet(spreadsheet);

    // デフォルトシートを削除
    const defaultSheet = spreadsheet.getSheetByName('シート1');
    if (defaultSheet) {
      spreadsheet.deleteSheet(defaultSheet);
    }

    Logger.log(`管理用スプレッドシート作成完了: ${spreadsheet.getId()}`);
    return spreadsheet;

  } catch (error) {
    Logger.log(`スプレッドシート作成エラー: ${error.toString()}`);
    throw error;
  }
}

/**
 * 分析結果シートを作成
 * @param {Spreadsheet} spreadsheet スプレッドシート
 */
function createAnalysisResultsSheet(spreadsheet) {
  const sheet = spreadsheet.insertSheet('分析結果');

  // ヘッダーを設定
  const headers = [
    '分析日時',
    'イベント数',
    '衝突数',
    '提案数',
    '時間使用率',
    'AIインサイト',
    'リスクレベル',
    '詳細'
  ];

  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length).setBackground('#4CAF50').setFontColor('white').setFontWeight('bold');

  // 列幅を調整
  sheet.setColumnWidth(1, 150); // 分析日時
  sheet.setColumnWidth(2, 80);  // イベント数
  sheet.setColumnWidth(3, 80);  // 衝突数
  sheet.setColumnWidth(4, 80);  // 提案数
  sheet.setColumnWidth(5, 100); // 時間使用率
  sheet.setColumnWidth(6, 200); // AIインサイト
  sheet.setColumnWidth(7, 100); // リスクレベル
  sheet.setColumnWidth(8, 250); // 詳細

  Logger.log('分析結果シート作成完了');
}

/**
 * 設定シートを作成
 * @param {Spreadsheet} spreadsheet スプレッドシート
 */
function createSettingsSheet(spreadsheet) {
  const sheet = spreadsheet.insertSheet('設定');

  // 設定項目とデフォルト値
  const settings = [
    ['設定項目', '値', '説明'],
    ['カレンダーID', 'primary', '使用するGoogleカレンダーのID'],
    ['AI_API_KEY', '', 'AI APIのキー（設定必須）'],
    ['通知メール', '', '通知を送信するメールアドレス'],
    ['分析対象日数', '7', '分析対象とする日数'],
    ['営業開始時間', '9', '営業開始時間（24時間制）'],
    ['営業終了時間', '18', '営業終了時間（24時間制）'],
    ['最小会議時間', '30', '最小会議時間（分）'],
    ['バッファー時間', '15', '会議間のバッファー時間（分）'],
    ['自動実行間隔', '60', 'トリガー実行間隔（分）'],
    ['通知レベル', 'medium', '通知レベル（low/medium/high）'],
    ['AI分析有効', 'true', 'AI分析機能の有効/無効']
  ];

  sheet.getRange(1, 1, settings.length, 3).setValues(settings);

  // ヘッダーのスタイル設定
  sheet.getRange(1, 1, 1, 3).setBackground('#2196F3').setFontColor('white').setFontWeight('bold');

  // 列幅調整
  sheet.setColumnWidth(1, 150); // 設定項目
  sheet.setColumnWidth(2, 200); // 値
  sheet.setColumnWidth(3, 300); // 説明

  Logger.log('設定シート作成完了');
}

/**
 * ログシートを作成
 * @param {Spreadsheet} spreadsheet スプレッドシート
 */
function createLogSheet(spreadsheet) {
  const sheet = spreadsheet.insertSheet('実行ログ');

  const headers = [
    '日時',
    'レベル',
    'カテゴリ',
    'メッセージ',
    '詳細',
    'ユーザー'
  ];

  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length).setBackground('#FF9800').setFontColor('white').setFontWeight('bold');

  // 列幅調整
  sheet.setColumnWidth(1, 150); // 日時
  sheet.setColumnWidth(2, 80);  // レベル
  sheet.setColumnWidth(3, 100); // カテゴリ
  sheet.setColumnWidth(4, 250); // メッセージ
  sheet.setColumnWidth(5, 300); // 詳細
  sheet.setColumnWidth(6, 150); // ユーザー

  Logger.log('ログシート作成完了');
}

/**
 * 衝突管理シートを作成
 * @param {Spreadsheet} spreadsheet スプレッドシート
 */
function createConflictsSheet(spreadsheet) {
  const sheet = spreadsheet.insertSheet('スケジュール衝突');

  const headers = [
    '検出日時',
    'イベント1',
    'イベント2',
    '衝突時間',
    '深刻度',
    '提案',
    '状態',
    '解決日時'
  ];

  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length).setBackground('#F44336').setFontColor('white').setFontWeight('bold');

  // 列幅調整
  sheet.setColumnWidth(1, 150); // 検出日時
  sheet.setColumnWidth(2, 200); // イベント1
  sheet.setColumnWidth(3, 200); // イベント2
  sheet.setColumnWidth(4, 150); // 衝突時間
  sheet.setColumnWidth(5, 100); // 深刻度
  sheet.setColumnWidth(6, 250); // 提案
  sheet.setColumnWidth(7, 100); // 状態
  sheet.setColumnWidth(8, 150); // 解決日時

  Logger.log('衝突管理シート作成完了');
}

/**
 * 分析結果をスプレッドシートに保存する
 * @param {Object} analysis 分析結果
 */
function saveAnalysisResults(analysis) {
  try {
    if (!CONFIG.SPREADSHEET_ID) {
      Logger.log('SPREADSHEET_IDが設定されていません');
      return;
    }

    const spreadsheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = spreadsheet.getSheetByName('分析結果');

    if (!sheet) {
      Logger.log('分析結果シートが見つかりません');
      return;
    }

    // データを準備
    const rowData = [
      analysis.timestamp,
      analysis.totalEvents,
      analysis.conflicts.length,
      analysis.suggestions.length,
      Math.round(analysis.timeUtilization * 100) + '%',
      analysis.aiInsights.join('; '),
      calculateOverallRiskLevel(analysis.riskFactors),
      JSON.stringify({
        optimizations: analysis.optimizations,
        conflicts: analysis.conflicts.map(c => c.type),
        riskCount: analysis.riskFactors.length
      })
    ];

    // 新しい行として追加
    sheet.appendRow(rowData);

    // 衝突がある場合は衝突シートにも保存
    if (analysis.conflicts.length > 0) {
      saveConflictData(analysis.conflicts, spreadsheet);
    }

    Logger.log(`分析結果保存完了: ${analysis.totalEvents}イベント, ${analysis.conflicts.length}衝突`);

  } catch (error) {
    Logger.log(`分析結果保存エラー: ${error.toString()}`);
  }
}

/**
 * 衝突データを保存する
 * @param {Array} conflicts 衝突配列
 * @param {Spreadsheet} spreadsheet スプレッドシート
 */
function saveConflictData(conflicts, spreadsheet) {
  try {
    const sheet = spreadsheet.getSheetByName('スケジュール衝突');
    if (!sheet) return;

    conflicts.forEach(conflict => {
      const rowData = [
        new Date(),
        `${conflict.event1.title} (${conflict.event1.startTime} - ${conflict.event1.endTime})`,
        `${conflict.event2.title} (${conflict.event2.startTime} - ${conflict.event2.endTime})`,
        `${new Date(conflict.event1.startTime).toLocaleString()} - ${new Date(conflict.event2.endTime).toLocaleString()}`,
        conflict.severity,
        conflict.suggestion,
        '未解決',
        ''
      ];

      sheet.appendRow(rowData);
    });

    Logger.log(`衝突データ保存完了: ${conflicts.length}件`);

  } catch (error) {
    Logger.log(`衝突データ保存エラー: ${error.toString()}`);
  }
}

/**
 * 設定を取得する
 * @param {string} key 設定キー
 * @return {string} 設定値
 */
function getSettingValue(key) {
  try {
    if (!CONFIG.SPREADSHEET_ID) {
      Logger.log('SPREADSHEET_IDが設定されていません');
      return null;
    }

    const spreadsheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = spreadsheet.getSheetByName('設定');

    if (!sheet) {
      Logger.log('設定シートが見つかりません');
      return null;
    }

    const data = sheet.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === key) {
        return data[i][1];
      }
    }

    Logger.log(`設定が見つかりません: ${key}`);
    return null;

  } catch (error) {
    Logger.log(`設定取得エラー: ${error.toString()}`);
    return null;
  }
}

/**
 * 設定を更新する
 * @param {string} key 設定キー
 * @param {string} value 設定値
 */
function updateSettingValue(key, value) {
  try {
    if (!CONFIG.SPREADSHEET_ID) {
      Logger.log('SPREADSHEET_IDが設定されていません');
      return false;
    }

    const spreadsheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = spreadsheet.getSheetByName('設定');

    if (!sheet) {
      Logger.log('設定シートが見つかりません');
      return false;
    }

    const data = sheet.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === key) {
        sheet.getRange(i + 1, 2).setValue(value);
        Logger.log(`設定更新完了: ${key} = ${value}`);
        return true;
      }
    }

    Logger.log(`更新対象設定が見つかりません: ${key}`);
    return false;

  } catch (error) {
    Logger.log(`設定更新エラー: ${error.toString()}`);
    return false;
  }
}

/**
 * 実行ログを記録する
 * @param {string} level ログレベル（INFO/WARN/ERROR）
 * @param {string} category カテゴリ
 * @param {string} message メッセージ
 * @param {string} details 詳細（オプション）
 */
function logToSpreadsheet(level, category, message, details = '') {
  try {
    if (!CONFIG.SPREADSHEET_ID) {
      return;
    }

    const spreadsheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = spreadsheet.getSheetByName('実行ログ');

    if (!sheet) {
      return;
    }

    const rowData = [
      new Date(),
      level,
      category,
      message,
      details,
      Session.getActiveUser().getEmail()
    ];

    sheet.appendRow(rowData);

    // 古いログを削除（1000行を超えた場合）
    const lastRow = sheet.getLastRow();
    if (lastRow > 1000) {
      sheet.deleteRows(2, lastRow - 1000);
    }

  } catch (error) {
    Logger.log(`ログ記録エラー: ${error.toString()}`);
  }
}

/**
 * 設定シートを初期化する
 * @param {Spreadsheet} spreadsheet スプレッドシート
 */
function initializeSettingsSheet(spreadsheet) {
  try {
    // 既に作成済みの場合は何もしない
    const settingsSheet = spreadsheet.getSheetByName('設定');
    if (settingsSheet && settingsSheet.getLastRow() > 1) {
      Logger.log('設定シートは既に初期化済みです');
      return;
    }

    // 設定シートが存在しない場合は作成
    if (!settingsSheet) {
      createSettingsSheet(spreadsheet);
    }

    Logger.log('設定シート初期化完了');

  } catch (error) {
    Logger.log(`設定シート初期化エラー: ${error.toString()}`);
  }
}

/**
 * 過去の分析データを取得する
 * @param {number} days 取得する日数
 * @return {Array} 分析データの配列
 */
function getPastAnalysisData(days = 30) {
  try {
    if (!CONFIG.SPREADSHEET_ID) {
      return [];
    }

    const spreadsheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = spreadsheet.getSheetByName('分析結果');

    if (!sheet) {
      return [];
    }

    const data = sheet.getDataRange().getValues();
    const cutoffDate = new Date(Date.now() - (days * 24 * 60 * 60 * 1000));

    const analysisData = [];

    for (let i = 1; i < data.length; i++) {
      const rowDate = new Date(data[i][0]);
      if (rowDate >= cutoffDate) {
        analysisData.push({
          date: rowDate,
          eventCount: data[i][1],
          conflictCount: data[i][2],
          suggestionCount: data[i][3],
          timeUtilization: data[i][4],
          insights: data[i][5],
          riskLevel: data[i][6],
          details: data[i][7]
        });
      }
    }

    return analysisData;

  } catch (error) {
    Logger.log(`過去データ取得エラー: ${error.toString()}`);
    return [];
  }
}

/**
 * 全体のリスクレベルを計算
 * @param {Array} riskFactors リスク要因の配列
 * @return {string} リスクレベル（low/medium/high）
 */
function calculateOverallRiskLevel(riskFactors) {
  if (riskFactors.length === 0) return 'low';

  const highRisks = riskFactors.filter(risk => risk.severity === 'high');
  const mediumRisks = riskFactors.filter(risk => risk.severity === 'medium');

  if (highRisks.length > 0) return 'high';
  if (mediumRisks.length > 1) return 'high';
  if (mediumRisks.length > 0) return 'medium';

  return 'low';
}

/**
 * データのバックアップを作成
 */
function backupData() {
  try {
    if (!CONFIG.SPREADSHEET_ID) {
      Logger.log('バックアップ対象のスプレッドシートIDが設定されていません');
      return;
    }

    const spreadsheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const backupName = `AIスケジューラバックアップ_${Utilities.formatDate(new Date(), CONFIG.TIME_ZONE, 'yyyyMMdd_HHmmss')}`;

    // スプレッドシートをコピーしてバックアップ作成
    const backup = spreadsheet.copy(backupName);

    Logger.log(`バックアップ作成完了: ${backup.getUrl()}`);
    logToSpreadsheet('INFO', 'BACKUP', 'データバックアップ作成完了', backup.getId());

    return backup.getId();

  } catch (error) {
    Logger.log(`バックアップ作成エラー: ${error.toString()}`);
    logToSpreadsheet('ERROR', 'BACKUP', 'バックアップ作成失敗', error.toString());
    throw error;
  }
}