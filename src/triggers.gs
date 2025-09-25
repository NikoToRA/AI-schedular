/**
 * トリガー管理システム
 * 定期実行やイベント駆動型トリガーの設定・管理を行う
 */

/**
 * 初期トリガーを設定する
 */
function setupTriggers() {
  try {
    Logger.log('トリガー設定を開始...');

    // 既存のトリガーを削除
    clearAllTriggers();

    // 定期実行トリガーを設定
    setupTimeDrivenTrigger();

    // カレンダー変更トリガーを設定
    setupCalendarTrigger();

    // フォーム送信トリガーを設定（設定変更用）
    setupFormTrigger();

    Logger.log('トリガー設定完了');
    logToSpreadsheet('INFO', 'TRIGGER', 'トリガー設定完了', '定期実行・カレンダー・フォーム');

  } catch (error) {
    Logger.log(`トリガー設定エラー: ${error.toString()}`);
    logToSpreadsheet('ERROR', 'TRIGGER', 'トリガー設定失敗', error.toString());
    throw error;
  }
}

/**
 * 時間駆動型トリガーを設定
 */
function setupTimeDrivenTrigger() {
  const interval = parseInt(getSettingValue('自動実行間隔') || '60', 10);

  // 毎時実行のトリガー
  ScriptApp.newTrigger('main')
    .timeBased()
    .everyHours(Math.max(1, Math.floor(interval / 60)))
    .create();

  // 毎日9:00の詳細分析トリガー
  ScriptApp.newTrigger('dailyAnalysis')
    .timeBased()
    .everyDays(1)
    .atHour(9)
    .create();

  // 週次レポートトリガー（月曜日9:00）
  ScriptApp.newTrigger('weeklyReport')
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.MONDAY)
    .atHour(9)
    .create();

  Logger.log(`時間駆動トリガー設定完了: ${interval}分間隔`);
}

/**
 * カレンダー更新トリガーを設定
 */
function setupCalendarTrigger() {
  try {
    // プライマリカレンダーの更新トリガー
    ScriptApp.newTrigger('onCalendarUpdate')
      .onEventUpdated()
      .create();

    Logger.log('カレンダー更新トリガー設定完了');

  } catch (error) {
    Logger.log(`カレンダートリガー設定エラー: ${error.toString()}`);
    // カレンダートリガーが設定できない場合でも続行
  }
}

/**
 * フォーム送信トリガーを設定
 */
function setupFormTrigger() {
  try {
    // 設定変更フォームのトリガー（将来実装）
    Logger.log('フォームトリガー設定（将来実装）');

  } catch (error) {
    Logger.log(`フォームトリガー設定エラー: ${error.toString()}`);
  }
}

/**
 * 全てのトリガーを削除
 */
function clearAllTriggers() {
  const triggers = ScriptApp.getProjectTriggers();

  triggers.forEach(trigger => {
    ScriptApp.deleteTrigger(trigger);
  });

  Logger.log(`既存トリガー削除完了: ${triggers.length}個`);
}

/**
 * カレンダー更新イベントハンドラー
 * @param {Object} e イベントオブジェクト
 */
function onCalendarUpdate(e) {
  try {
    Logger.log('カレンダー更新を検知');

    // 更新から少し待機してから分析実行
    Utilities.sleep(5000);

    // 簡易分析を実行
    const events = getCalendarEvents(1); // 1日分のみ
    const analysis = performBasicAnalysis(events);

    // 緊急の問題がある場合のみ通知
    if (analysis.conflicts.length > 0) {
      const urgentConflicts = analysis.conflicts.filter(c => c.severity === 'high');
      if (urgentConflicts.length > 0) {
        const urgentAnalysis = {
          ...analysis,
          conflicts: urgentConflicts,
          aiInsights: ['カレンダー更新により緊急の衝突が検出されました'],
          riskFactors: []
        };

        sendNotification(urgentAnalysis);
      }
    }

    logToSpreadsheet('INFO', 'CALENDAR', 'カレンダー更新対応完了', `イベント数: ${events.length}`);

  } catch (error) {
    Logger.log(`カレンダー更新処理エラー: ${error.toString()}`);
    logToSpreadsheet('ERROR', 'CALENDAR', 'カレンダー更新処理失敗', error.toString());
  }
}

/**
 * 日次詳細分析
 */
function dailyAnalysis() {
  try {
    Logger.log('日次詳細分析を開始');

    // より長期間のイベントを分析
    const events = getCalendarEvents(14); // 2週間分
    const analysis = analyzeScheduleWithAI(events);

    // 結果を保存
    saveAnalysisResults(analysis);

    // 詳細レポートを送信
    sendDailyReport(analysis);

    // データのクリーンアップ
    cleanupOldData();

    Logger.log('日次詳細分析完了');

  } catch (error) {
    Logger.log(`日次分析エラー: ${error.toString()}`);
    sendErrorNotification(error);
  }
}

/**
 * 週次レポート生成
 */
function weeklyReport() {
  try {
    Logger.log('週次レポート生成を開始');

    // 過去1週間の分析データを取得
    const weeklyData = getPastAnalysisData(7);

    // 週次統計を計算
    const weeklyStats = calculateWeeklyStats(weeklyData);

    // 週次レポートを作成・送信
    sendWeeklyReport(weeklyStats);

    // バックアップを作成
    backupData();

    Logger.log('週次レポート完了');

  } catch (error) {
    Logger.log(`週次レポートエラー: ${error.toString()}`);
    sendErrorNotification(error);
  }
}

/**
 * 日次レポートを送信
 * @param {Object} analysis 分析結果
 */
function sendDailyReport(analysis) {
  try {
    const subject = `AIスケジューラ 日次レポート - ${Utilities.formatDate(new Date(), CONFIG.TIME_ZONE, 'yyyy/MM/dd')}`;

    const body = `
AIスケジューラ 日次詳細レポート
==============================

【本日のサマリー】
- 分析イベント数: ${analysis.totalEvents}件
- 検出された衝突: ${analysis.conflicts.length}件
- AI提案数: ${analysis.suggestions.length}件
- 時間使用率: ${Math.round(analysis.timeUtilization * 100)}%

【週間トレンド】
${generateTrendAnalysis()}

【今後の注意事項】
${generateFutureAlerts()}

詳細は管理用スプレッドシートをご確認ください。
`;

    sendGmailNotification(subject, body);

  } catch (error) {
    Logger.log(`日次レポート送信エラー: ${error.toString()}`);
  }
}

/**
 * 週次統計を計算
 * @param {Array} weeklyData 週次データ
 * @return {Object} 統計情報
 */
function calculateWeeklyStats(weeklyData) {
  if (weeklyData.length === 0) {
    return {
      avgEvents: 0,
      totalConflicts: 0,
      avgUtilization: 0,
      trends: 'データが不十分です'
    };
  }

  const stats = {
    avgEvents: weeklyData.reduce((sum, data) => sum + data.eventCount, 0) / weeklyData.length,
    totalConflicts: weeklyData.reduce((sum, data) => sum + data.conflictCount, 0),
    avgUtilization: weeklyData.reduce((sum, data) => {
      const util = typeof data.timeUtilization === 'string' ?
        parseInt(data.timeUtilization.replace('%', '')) / 100 :
        data.timeUtilization;
      return sum + util;
    }, 0) / weeklyData.length,
    peakDays: identifyPeakDays(weeklyData),
    improvementAreas: identifyImprovementAreas(weeklyData)
  };

  return stats;
}

/**
 * 週次レポートを送信
 * @param {Object} stats 統計データ
 */
function sendWeeklyReport(stats) {
  try {
    const weekStart = new Date();
    weekStart.setDate(weekStart.getDate() - 7);

    const subject = `AIスケジューラ 週次レポート - ${Utilities.formatDate(weekStart, CONFIG.TIME_ZONE, 'MM/dd')}〜${Utilities.formatDate(new Date(), CONFIG.TIME_ZONE, 'MM/dd')}`;

    const body = `
AIスケジューラ 週次レポート
==========================

【先週の実績】
- 平均イベント数/日: ${Math.round(stats.avgEvents)}件
- 総衝突件数: ${stats.totalConflicts}件
- 平均時間使用率: ${Math.round(stats.avgUtilization * 100)}%

【分析結果】
- ピーク日: ${stats.peakDays || '特定できませんでした'}
- 改善提案: ${stats.improvementAreas || '効率的に管理されています'}

【来週の推奨事項】
${generateWeeklyRecommendations(stats)}

---
週次バックアップも作成されました。
継続的な改善のため、データをご活用ください。
`;

    sendGmailNotification(subject, body);

  } catch (error) {
    Logger.log(`週次レポート送信エラー: ${error.toString()}`);
  }
}

/**
 * トレンド分析を生成
 * @return {string} トレンド分析結果
 */
function generateTrendAnalysis() {
  try {
    const recentData = getPastAnalysisData(7);

    if (recentData.length < 3) {
      return '十分なデータがありません（3日以上のデータが必要）';
    }

    const conflictTrend = calculateTrend(recentData.map(d => d.conflictCount));
    const utilizationTrend = calculateTrend(recentData.map(d => {
      const util = typeof d.timeUtilization === 'string' ?
        parseInt(d.timeUtilization.replace('%', '')) / 100 :
        d.timeUtilization;
      return util;
    }));

    let trendText = '';

    if (conflictTrend > 0.1) {
      trendText += '⚠️ 衝突件数が増加傾向にあります。\n';
    } else if (conflictTrend < -0.1) {
      trendText += '✅ 衝突件数が減少傾向で改善されています。\n';
    }

    if (utilizationTrend > 0.05) {
      trendText += '📈 時間使用率が上昇傾向にあります。\n';
    } else if (utilizationTrend < -0.05) {
      trendText += '📉 時間使用率が低下傾向にあります。\n';
    }

    return trendText || '📊 安定した傾向を示しています。';

  } catch (error) {
    return `トレンド分析エラー: ${error.toString()}`;
  }
}

/**
 * 将来のアラートを生成
 * @return {string} 将来のアラート
 */
function generateFutureAlerts() {
  try {
    const futureEvents = getCalendarEvents(7);
    const futureConflicts = detectScheduleConflicts(futureEvents);

    let alerts = '';

    if (futureConflicts.length > 0) {
      alerts += `🔔 今後7日間で${futureConflicts.length}件の衝突が予想されます。\n`;

      const highPriorityFuture = futureConflicts.filter(c => c.severity === 'high');
      if (highPriorityFuture.length > 0) {
        alerts += `🚨 うち${highPriorityFuture.length}件は重要度が高く、早急な対応が必要です。\n`;
      }
    }

    // 多忙日の予測
    const busyDays = identifyBusyDays(futureEvents);
    if (busyDays.length > 0) {
      const busiestDay = busyDays.reduce((max, day) => day.eventCount > max.eventCount ? day : max);
      alerts += `📅 ${busiestDay.date}が最も多忙になる予定です（${busiestDay.eventCount}件のイベント）。\n`;
    }

    return alerts || '📅 今後1週間は比較的安定したスケジュールです。';

  } catch (error) {
    return `将来予測エラー: ${error.toString()}`;
  }
}

/**
 * 週次推奨事項を生成
 * @param {Object} stats 統計データ
 * @return {string} 推奨事項
 */
function generateWeeklyRecommendations(stats) {
  let recommendations = '';

  if (stats.avgUtilization > 0.8) {
    recommendations += '⚡ 時間使用率が高いため、バッファー時間の確保を検討してください。\n';
  }

  if (stats.totalConflicts > 5) {
    recommendations += '🔧 衝突が頻発しています。定期的なスケジュール見直しをお勧めします。\n';
  }

  if (stats.avgEvents < 3) {
    recommendations += '📈 スケジュールに余裕があります。新しいプロジェクトの開始を検討できます。\n';
  }

  return recommendations || '✅ 現在のスケジュール管理は良好です。継続してください。';
}

/**
 * 数値配列のトレンドを計算
 * @param {Array} values 数値配列
 * @return {number} トレンド（正：上昇、負：下降）
 */
function calculateTrend(values) {
  if (values.length < 2) return 0;

  const n = values.length;
  const sumX = (n * (n - 1)) / 2; // 0 + 1 + 2 + ... + (n-1)
  const sumY = values.reduce((sum, val) => sum + val, 0);
  const sumXY = values.reduce((sum, val, i) => sum + (i * val), 0);
  const sumX2 = (n * (n - 1) * (2 * n - 1)) / 6; // 0² + 1² + 2² + ... + (n-1)²

  const slope = (n * sumXY - sumX * sumY) / (n * sumX2 - sumX * sumX);
  return slope;
}

/**
 * ピーク日を特定
 * @param {Array} weeklyData 週次データ
 * @return {string} ピーク日の説明
 */
function identifyPeakDays(weeklyData) {
  if (weeklyData.length === 0) return 'データなし';

  const peakDay = weeklyData.reduce((max, data) =>
    data.eventCount > max.eventCount ? data : max
  );

  const dayName = new Date(peakDay.date).toLocaleDateString('ja-JP', { weekday: 'long' });
  return `${dayName} (${peakDay.eventCount}件)`;
}

/**
 * 改善エリアを特定
 * @param {Array} weeklyData 週次データ
 * @return {string} 改善エリアの説明
 */
function identifyImprovementAreas(weeklyData) {
  const highConflictDays = weeklyData.filter(data => data.conflictCount > 2);

  if (highConflictDays.length > 0) {
    return `衝突の多い日が${highConflictDays.length}日あります。スケジュール調整を検討してください。`;
  }

  const highUtilizationDays = weeklyData.filter(data => {
    const util = typeof data.timeUtilization === 'string' ?
      parseInt(data.timeUtilization.replace('%', '')) / 100 :
      data.timeUtilization;
    return util > 0.9;
  });

  if (highUtilizationDays.length > 2) {
    return `時間使用率の高い日が多く、休憩時間の確保が必要です。`;
  }

  return '特に改善が必要なエリアは見つかりませんでした。';
}

/**
 * 古いデータをクリーンアップ
 */
function cleanupOldData() {
  try {
    if (!CONFIG.SPREADSHEET_ID) {
      return;
    }

    const spreadsheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const analysisSheet = spreadsheet.getSheetByName('分析結果');
    const logSheet = spreadsheet.getSheetByName('実行ログ');

    const cutoffDate = new Date(Date.now() - (90 * 24 * 60 * 60 * 1000)); // 90日前

    // 分析結果の古いデータを削除
    if (analysisSheet) {
      deleteOldRows(analysisSheet, cutoffDate, 0);
    }

    // ログの古いデータを削除
    if (logSheet) {
      deleteOldRows(logSheet, cutoffDate, 0);
    }

    Logger.log('古いデータのクリーンアップ完了');

  } catch (error) {
    Logger.log(`データクリーンアップエラー: ${error.toString()}`);
  }
}

/**
 * シートの古い行を削除
 * @param {Sheet} sheet 対象シート
 * @param {Date} cutoffDate カットオフ日
 * @param {number} dateColumn 日付列のインデックス
 */
function deleteOldRows(sheet, cutoffDate, dateColumn) {
  const data = sheet.getDataRange().getValues();
  const rowsToDelete = [];

  for (let i = 1; i < data.length; i++) {
    const rowDate = new Date(data[i][dateColumn]);
    if (rowDate < cutoffDate) {
      rowsToDelete.push(i + 1); // 1-based index
    }
  }

  // 後ろから削除（インデックスが変わらないように）
  rowsToDelete.reverse().forEach(rowIndex => {
    sheet.deleteRow(rowIndex);
  });

  if (rowsToDelete.length > 0) {
    Logger.log(`${sheet.getName()}から${rowsToDelete.length}行の古いデータを削除`);
  }
}