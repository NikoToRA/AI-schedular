/**
 * ========================================
 * AIスケジューラ - 最小限版
 * ========================================
 *
 * Google Apps Script で動作するAIスケジュール管理システム（最小限版）
 *
 * 機能:
 * - Google Calendar との連携
 * - スケジュール衝突検知
 * - 基本的な分析とレポート
 * - スプレッドシートでの結果管理
 *
 * 作成日: 2025-09-25
 * バージョン: 1.0 (Minimal)
 */

// ========================================
// 設定定数
// ========================================
const CONFIG = {
  CALENDAR_ID: 'primary',
  SPREADSHEET_ID: '', // 初期化後に設定されます
  TIME_ZONE: 'Asia/Tokyo'
};

// ========================================
// メイン実行関数
// ========================================

/**
 * メイン実行関数 - トリガーから実行される
 */
function main() {
  try {
    Logger.log('AIスケジューラ開始: ' + new Date());

    // 設定チェック
    if (!CONFIG.SPREADSHEET_ID) {
      Logger.log('SPREADSHEET_IDが設定されていません。initialize()を実行してください。');
      return;
    }

    // カレンダーイベント取得
    const events = getCalendarEvents();
    Logger.log(`取得イベント数: ${events.length}`);

    // 分析実行
    const analysis = analyzeSchedule(events);

    // 結果をスプレッドシートに保存
    saveAnalysisResults(analysis);

    // ログに結果出力
    Logger.log(`分析完了: 衝突${analysis.conflicts.length}件、時間使用率${Math.round(analysis.timeUtilization * 100)}%`);

    // 衝突がある場合は詳細をログ出力
    if (analysis.conflicts.length > 0) {
      Logger.log('=== 検出された衝突 ===');
      analysis.conflicts.forEach((conflict, index) => {
        Logger.log(`${index + 1}. ${conflict.event1.title} ⇔ ${conflict.event2.title}`);
        Logger.log(`   提案: ${conflict.suggestion}`);
      });
    }

    Logger.log('AIスケジューラ完了: ' + new Date());

  } catch (error) {
    Logger.log(`エラー: ${error.toString()}`);
  }
}

/**
 * 初期設定用関数 - 最初に1回実行してください
 */
function initialize() {
  Logger.log('初期設定を開始します...');

  try {
    // スプレッドシートの作成
    const spreadsheet = createManagementSpreadsheet();
    Logger.log(`管理用スプレッドシートが作成されました: ${spreadsheet.getUrl()}`);

    // CONFIGにスプレッドシートIDを設定（手動で更新が必要）
    Logger.log(`以下のIDをCONFIG.SPREADSHEET_IDに設定してください: ${spreadsheet.getId()}`);

    // トリガーの設定
    setupTriggers();

    Logger.log('初期設定が完了しました');
    Logger.log('次のステップ: testRun()を実行して動作確認');

  } catch (error) {
    Logger.log(`初期設定エラー: ${error.toString()}`);
    throw error;
  }
}

/**
 * テスト実行用関数
 */
function testRun() {
  Logger.log('テスト実行を開始...');
  main();
}

// ========================================
// カレンダー連携
// ========================================

/**
 * カレンダーイベントを取得する
 */
function getCalendarEvents(days = 7) {
  try {
    const calendar = CalendarApp.getDefaultCalendar();
    const startTime = new Date();
    const endTime = new Date(startTime.getTime() + (days * 24 * 60 * 60 * 1000));

    const events = calendar.getEvents(startTime, endTime);

    return events.map(event => {
      return {
        id: event.getId(),
        title: event.getTitle(),
        description: event.getDescription() || '',
        startTime: event.getStartTime(),
        endTime: event.getEndTime(),
        location: event.getLocation() || '',
        isAllDay: event.isAllDayEvent()
      };
    });

  } catch (error) {
    Logger.log(`カレンダーイベント取得エラー: ${error.toString()}`);
    throw error;
  }
}

/**
 * 新しいイベントを作成する
 */
function createCalendarEvent(eventData) {
  try {
    const calendar = CalendarApp.getDefaultCalendar();

    let event;
    if (eventData.isAllDay) {
      event = calendar.createAllDayEvent(
        eventData.title,
        eventData.startTime,
        {
          description: eventData.description,
          location: eventData.location
        }
      );
    } else {
      event = calendar.createEvent(
        eventData.title,
        eventData.startTime,
        eventData.endTime,
        {
          description: eventData.description,
          location: eventData.location
        }
      );
    }

    Logger.log(`イベント作成完了: ${eventData.title}`);
    return {
      id: event.getId(),
      title: event.getTitle(),
      startTime: event.getStartTime(),
      endTime: event.getEndTime()
    };

  } catch (error) {
    Logger.log(`イベント作成エラー: ${error.toString()}`);
    throw error;
  }
}

// ========================================
// スケジュール分析
// ========================================

/**
 * スケジュールを分析する
 */
function analyzeSchedule(events) {
  try {
    Logger.log('スケジュール分析を開始...');

    const conflicts = detectScheduleConflicts(events);
    const timeUtilization = calculateTimeUtilization(events);
    const busyDays = identifyBusyDays(events);

    const result = {
      timestamp: new Date(),
      totalEvents: events.length,
      conflicts: conflicts,
      timeUtilization: timeUtilization,
      busyDays: busyDays,
      suggestions: generateSuggestions(conflicts, timeUtilization)
    };

    Logger.log(`分析完了: 衝突${result.conflicts.length}件`);
    return result;

  } catch (error) {
    Logger.log(`分析エラー: ${error.toString()}`);
    throw error;
  }
}

/**
 * スケジュールの衝突を検出する
 */
function detectScheduleConflicts(events) {
  const conflicts = [];

  for (let i = 0; i < events.length; i++) {
    for (let j = i + 1; j < events.length; j++) {
      const event1 = events[i];
      const event2 = events[j];

      if (isTimeConflict(event1, event2)) {
        conflicts.push({
          type: 'time_overlap',
          event1: {
            id: event1.id,
            title: event1.title,
            startTime: event1.startTime,
            endTime: event1.endTime
          },
          event2: {
            id: event2.id,
            title: event2.title,
            startTime: event2.startTime,
            endTime: event2.endTime
          },
          severity: calculateConflictSeverity(event1, event2),
          suggestion: generateConflictSuggestion(event1, event2)
        });
      }
    }
  }

  return conflicts;
}

/**
 * 2つのイベントが時間的に衝突するかチェック
 */
function isTimeConflict(event1, event2) {
  const start1 = new Date(event1.startTime);
  const end1 = new Date(event1.endTime);
  const start2 = new Date(event2.startTime);
  const end2 = new Date(event2.endTime);

  return (start1 < end2 && start2 < end1);
}

/**
 * 衝突の深刻度を計算
 */
function calculateConflictSeverity(event1, event2) {
  const duration1 = new Date(event1.endTime) - new Date(event1.startTime);
  const duration2 = new Date(event2.endTime) - new Date(event2.startTime);

  // 長時間のミーティングの衝突は深刻
  if (duration1 > 2 * 60 * 60 * 1000 || duration2 > 2 * 60 * 60 * 1000) {
    return 'high';
  }

  // 重要なキーワードをチェック
  const importantKeywords = ['会議', 'ミーティング', '重要', '緊急', '役員', 'プレゼン'];
  const isImportant1 = importantKeywords.some(keyword =>
    event1.title.includes(keyword) || event1.description.includes(keyword));
  const isImportant2 = importantKeywords.some(keyword =>
    event2.title.includes(keyword) || event2.description.includes(keyword));

  if (isImportant1 || isImportant2) {
    return 'high';
  }

  return 'medium';
}

/**
 * 衝突解決の提案を生成
 */
function generateConflictSuggestion(event1, event2) {
  const duration1 = new Date(event1.endTime) - new Date(event1.startTime);
  const duration2 = new Date(event2.endTime) - new Date(event2.startTime);

  if (duration1 < duration2) {
    return `「${event1.title}」を別の時間に移動することを検討してください`;
  } else {
    return `「${event2.title}」を別の時間に移動することを検討してください`;
  }
}

/**
 * 時間使用率を計算
 */
function calculateTimeUtilization(events) {
  if (events.length === 0) return 0;

  const businessHoursPerDay = 8; // 8時間
  const today = new Date();
  const weekStart = new Date(today.setDate(today.getDate() - today.getDay() + 1));
  const weekEnd = new Date(weekStart);
  weekEnd.setDate(weekEnd.getDate() + 5); // 平日のみ

  const thisWeekEvents = events.filter(event => {
    const eventDate = new Date(event.startTime);
    return eventDate >= weekStart && eventDate < weekEnd;
  });

  const totalScheduledTime = thisWeekEvents.reduce((total, event) => {
    return total + (new Date(event.endTime) - new Date(event.startTime));
  }, 0);

  const totalBusinessTime = 5 * businessHoursPerDay * 60 * 60 * 1000; // 5日 × 8時間
  return Math.min(totalScheduledTime / totalBusinessTime, 1);
}

/**
 * 多忙な日を特定
 */
function identifyBusyDays(events) {
  const dailyEvents = {};

  events.forEach(event => {
    const date = new Date(event.startTime).toDateString();
    if (!dailyEvents[date]) {
      dailyEvents[date] = [];
    }
    dailyEvents[date].push(event);
  });

  return Object.keys(dailyEvents)
    .filter(date => dailyEvents[date].length >= 4)
    .map(date => ({
      date: date,
      eventCount: dailyEvents[date].length,
      totalDuration: dailyEvents[date].reduce((total, event) =>
        total + (new Date(event.endTime) - new Date(event.startTime)), 0) / (60 * 1000) // 分
    }));
}

/**
 * 改善提案を生成
 */
function generateSuggestions(conflicts, timeUtilization) {
  const suggestions = [];

  // 衝突がある場合の提案
  if (conflicts.length > 0) {
    suggestions.push('スケジュールの衝突が検出されました。重複する会議の時間を調整することをお勧めします。');

    const highPriorityConflicts = conflicts.filter(c => c.severity === 'high');
    if (highPriorityConflicts.length > 0) {
      suggestions.push(`重要度の高い衝突が${highPriorityConflicts.length}件あります。優先的に対応してください。`);
    }
  }

  // 時間使用率に基づく提案
  if (timeUtilization > 0.8) {
    suggestions.push('スケジュールが密集しています。休憩時間を確保することを検討してください。');
  } else if (timeUtilization < 0.3) {
    suggestions.push('スケジュールに余裕があります。新しいプロジェクトの開始を検討できます。');
  }

  return suggestions;
}

// ========================================
// データ管理
// ========================================

/**
 * 管理用スプレッドシートを作成する
 */
function createManagementSpreadsheet() {
  try {
    const spreadsheet = SpreadsheetApp.create('AIスケジューラ管理データ（最小版）');

    // 分析結果シートを作成
    createAnalysisSheet(spreadsheet);

    // 衝突管理シートを作成
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
 */
function createAnalysisSheet(spreadsheet) {
  const sheet = spreadsheet.insertSheet('分析結果');

  const headers = [
    '分析日時',
    'イベント数',
    '衝突数',
    '時間使用率',
    '多忙日数',
    '提案',
    '詳細'
  ];

  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length).setBackground('#4CAF50').setFontColor('white').setFontWeight('bold');

  // 列幅を調整
  sheet.setColumnWidth(1, 150); // 分析日時
  sheet.setColumnWidth(2, 80);  // イベント数
  sheet.setColumnWidth(3, 80);  // 衝突数
  sheet.setColumnWidth(4, 100); // 時間使用率
  sheet.setColumnWidth(5, 80);  // 多忙日数
  sheet.setColumnWidth(6, 300); // 提案
  sheet.setColumnWidth(7, 200); // 詳細

  Logger.log('分析結果シート作成完了');
}

/**
 * 衝突管理シートを作成
 */
function createConflictsSheet(spreadsheet) {
  const sheet = spreadsheet.insertSheet('スケジュール衝突');

  const headers = [
    '検出日時',
    'イベント1',
    'イベント2',
    '衝突時間',
    '深刻度',
    '提案'
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

  Logger.log('衝突管理シート作成完了');
}

/**
 * 分析結果をスプレッドシートに保存する
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
      Math.round(analysis.timeUtilization * 100) + '%',
      analysis.busyDays.length,
      analysis.suggestions.join(' | '),
      `衝突: ${analysis.conflicts.length}件, 多忙日: ${analysis.busyDays.length}日`
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
 */
function saveConflictData(conflicts, spreadsheet) {
  try {
    const sheet = spreadsheet.getSheetByName('スケジュール衝突');
    if (!sheet) return;

    conflicts.forEach(conflict => {
      const rowData = [
        new Date(),
        `${conflict.event1.title} (${formatDateTime(conflict.event1.startTime)} - ${formatDateTime(conflict.event1.endTime)})`,
        `${conflict.event2.title} (${formatDateTime(conflict.event2.startTime)} - ${formatDateTime(conflict.event2.endTime)})`,
        `${formatDateTime(conflict.event1.startTime)} - ${formatDateTime(conflict.event2.endTime)}`,
        conflict.severity,
        conflict.suggestion
      ];

      sheet.appendRow(rowData);
    });

    Logger.log(`衝突データ保存完了: ${conflicts.length}件`);

  } catch (error) {
    Logger.log(`衝突データ保存エラー: ${error.toString()}`);
  }
}

// ========================================
// トリガー管理
// ========================================

/**
 * トリガーを設定する
 */
function setupTriggers() {
  try {
    Logger.log('トリガー設定を開始...');

    // 既存のトリガーを削除
    clearAllTriggers();

    // 毎時実行のトリガー
    ScriptApp.newTrigger('main')
      .timeBased()
      .everyHours(1)
      .create();

    // 毎日9:00の詳細分析トリガー
    ScriptApp.newTrigger('main')
      .timeBased()
      .everyDays(1)
      .atHour(9)
      .create();

    Logger.log('トリガー設定完了');

  } catch (error) {
    Logger.log(`トリガー設定エラー: ${error.toString()}`);
    throw error;
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

// ========================================
// ユーティリティ関数
// ========================================

/**
 * 日時をフォーマットする
 */
function formatDateTime(dateTime) {
  const date = new Date(dateTime);
  return Utilities.formatDate(date, CONFIG.TIME_ZONE, 'MM/dd HH:mm');
}

/**
 * 週間レポートを生成
 */
function generateWeeklyReport() {
  try {
    Logger.log('週間レポート生成中...');

    const events = getCalendarEvents(14); // 2週間分
    const analysis = analyzeSchedule(events);

    Logger.log('=== 週間分析レポート ===');
    Logger.log(`総イベント数: ${analysis.totalEvents}件`);
    Logger.log(`衝突件数: ${analysis.conflicts.length}件`);
    Logger.log(`時間使用率: ${Math.round(analysis.timeUtilization * 100)}%`);
    Logger.log(`多忙日数: ${analysis.busyDays.length}日`);

    if (analysis.suggestions.length > 0) {
      Logger.log('=== 改善提案 ===');
      analysis.suggestions.forEach((suggestion, index) => {
        Logger.log(`${index + 1}. ${suggestion}`);
      });
    }

    Logger.log('週間レポート生成完了');

  } catch (error) {
    Logger.log(`週間レポート生成エラー: ${error.toString()}`);
  }
}

// ========================================
// セットアップガイド
// ========================================

/**
 * セットアップガイドを表示
 */
function showSetupGuide() {
  const guide = `
=================================================
🤖 AIスケジューラ（最小版）セットアップガイド
=================================================

【ステップ1: 初期化】
1. initialize() 関数を実行
2. 作成されたスプレッドシートのIDをメモ
3. CONFIG.SPREADSHEET_ID にそのIDを設定

【ステップ2: テスト実行】
1. testRun() を実行して動作確認
2. スプレッドシートに結果が保存されることを確認

【ステップ3: 自動実行】
- 毎時と毎日9:00に自動実行されます
- generateWeeklyReport() で週次レポート生成

【主要な関数】
- main(): メイン分析実行
- testRun(): テスト実行
- initialize(): 初期化（最初に1回のみ）
- generateWeeklyReport(): 週間レポート

【機能】
✅ スケジュール衝突検知
✅ 時間使用率分析
✅ 改善提案生成
✅ スプレッドシート保存
✅ 自動実行（トリガー）

Gmail通知機能は削除されています。
結果は実行トランスクリプトとスプレッドシートで確認できます。

=================================================
`;

  Logger.log(guide);
  return guide;
}

// ========================================
// エントリーポイント
// ========================================

Logger.log(`
🤖 AIスケジューラ（最小版）が読み込まれました

📝 セットアップ手順:
1. initialize() を実行
2. showSetupGuide() で詳細ガイドを確認

🎯 主要機能:
- スケジュール衝突検知
- 基本的な分析とレポート
- スプレッドシートでの結果管理

バージョン: 1.0 (Minimal) | 作成日: 2025-09-25
`);

/**
 * ========================================
 * 🎁 AIスケジューラ 最小版 - 完成！
 * ========================================
 *
 * Gmail通知機能を削除した軽量版です。
 *
 * 含まれる機能:
 * ✅ Google Calendar 連携
 * ✅ スケジュール衝突検知
 * ✅ 時間使用率分析
 * ✅ 改善提案生成
 * ✅ スプレッドシート保存
 * ✅ 自動トリガー
 *
 * 削除された機能:
 * ❌ Gmail通知
 * ❌ 複雑なAI分析
 * ❌ Webダッシュボード
 *
 * 総コード行数: 約500行
 *
 * 使い方:
 * 1. このファイルをGoogle Apps Scriptにコピー
 * 2. initialize()を実行
 * 3. main()で分析開始
 *
 * シンプルで効果的なスケジュール管理を！ 🚀
 */