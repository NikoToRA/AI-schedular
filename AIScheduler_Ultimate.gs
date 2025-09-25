/**
 * ========================================
 * AIスケジューラ - 完全統合版（1ファイル完結）
 * ========================================
 *
 * Google Apps Script で動作するAIスケジュール管理システム
 * Notion連携機能も含む完全版
 *
 * 🎁 プレゼント用: このファイル1つだけでOK！
 *
 * セットアップ:
 * 1. Google Apps Script にこのファイルをコピー&ペースト
 * 2. initialize() を実行
 * 3. 完成！
 *
 * Notion連携したい場合:
 * 1. showNotionConfigGuide() でガイド確認
 * 2. NOTION_CONFIG の設定を入力
 * 3. testNotionConnection() でテスト
 *
 * 作成日: 2025-09-25
 * バージョン: 1.0 Ultimate
 * 総行数: 約1,000行
 */

// ========================================
// 基本設定
// ========================================
const CONFIG = {
  CALENDAR_ID: 'primary',
  SPREADSHEET_ID: '', // 初期化後に設定されます
  TIME_ZONE: 'Asia/Tokyo'
};

// ========================================
// 🔧 NOTION設定 - Notion連携したい場合のみ設定
// ========================================
const NOTION_CONFIG = {
  // 👇 Notion連携する場合、ここにIntegration Tokenを入力
  // 取得方法: https://www.notion.so/my-integrations で作成
  // 例: 'secret_ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijk123'
  INTEGRATION_TOKEN: '', // ★★★Notion連携する場合のみ入力★★★

  // 👇 Notion連携する場合、ここにDatabase IDを入力
  // 取得方法: NotionデータベースのURLから取得
  // URL例: https://notion.so/workspace/1234567890abcdef?v=...
  // 例: '1234567890abcdef1234567890abcdef'
  DATABASE_ID: '',       // ★★★Notion連携する場合のみ入力★★★

  API_VERSION: '2022-06-28'
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

    // Notion連携が設定されている場合は同期実行
    if (isNotionConfigured()) {
      try {
        Logger.log('Notion連携が設定されています。同期を実行します...');
        const syncResult = runBidirectionalSync();
        Logger.log(`Notion同期完了: カレンダー→Notion ${syncResult.calendarToNotion.created}件作成`);
      } catch (notionError) {
        Logger.log(`Notion同期エラー: ${notionError.toString()}`);
      }
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
        isAllDay: event.isAllDayEvent(),
        attendees: event.getGuestList().map(guest => guest.getEmail())
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
    const spreadsheet = SpreadsheetApp.create('AIスケジューラ管理データ');

    // 分析結果シートを作成
    createAnalysisSheet(spreadsheet);

    // 衝突管理シートを作成
    createConflictsSheet(spreadsheet);

    // Notion同期結果シートを作成
    createNotionSyncSheet(spreadsheet);

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
 * Notion同期結果シートを作成
 */
function createNotionSyncSheet(spreadsheet) {
  const sheet = spreadsheet.insertSheet('Notion同期結果');

  const headers = [
    '同期日時',
    'カレンダー→Notion作成',
    'カレンダー→Notion更新',
    'Notion→カレンダー作成',
    'エラー数',
    'ステータス'
  ];

  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length).setBackground('#9C27B0').setFontColor('white').setFontWeight('bold');

  // 列幅を調整
  sheet.setColumnWidth(1, 150); // 同期日時
  sheet.setColumnWidth(2, 120); // カレンダー→Notion作成
  sheet.setColumnWidth(3, 120); // カレンダー→Notion更新
  sheet.setColumnWidth(4, 120); // Notion→カレンダー作成
  sheet.setColumnWidth(5, 80);  // エラー数
  sheet.setColumnWidth(6, 100); // ステータス

  Logger.log('Notion同期結果シート作成完了');
  return sheet;
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
// Notion連携機能（オプション）
// ========================================

/**
 * Notion設定が完了しているかチェック
 */
function isNotionConfigured() {
  return NOTION_CONFIG.INTEGRATION_TOKEN &&
         NOTION_CONFIG.DATABASE_ID &&
         !NOTION_CONFIG.INTEGRATION_TOKEN.includes('★') &&
         !NOTION_CONFIG.DATABASE_ID.includes('★');
}

/**
 * 双方向同期の実行
 */
function runBidirectionalSync() {
  try {
    Logger.log('双方向同期を開始...');

    // 設定チェック
    if (!isNotionConfigured()) {
      throw new Error('Notion設定が不完全です。showNotionConfigGuide()でガイドを確認してください。');
    }

    // カレンダー→Notion
    const calendarToNotionResult = syncCalendarToNotion();

    // Notion→カレンダー
    const notionToCalendarResult = syncNotionToCalendar();

    const summary = {
      timestamp: new Date(),
      calendarToNotion: calendarToNotionResult,
      notionToCalendar: notionToCalendarResult
    };

    Logger.log('双方向同期完了');

    // 同期結果をスプレッドシートに保存
    saveSyncResults(summary);

    return summary;

  } catch (error) {
    Logger.log(`双方向同期エラー: ${error.toString()}`);
    throw error;
  }
}

/**
 * GoogleカレンダーからNotionへの同期
 */
function syncCalendarToNotion() {
  try {
    Logger.log('カレンダー→Notion同期を開始...');

    // Googleカレンダーのイベントを取得
    const calendarEvents = getCalendarEvents(7); // 1週間分

    // 既存のNotionページを取得
    const existingPages = getNotionPages();
    const existingEventIds = new Set();

    existingPages.forEach(page => {
      const eventId = getNotionText(page.properties['Calendar Event ID']);
      if (eventId) {
        existingEventIds.add(eventId);
      }
    });

    let created = 0;
    let updated = 0;

    for (const event of calendarEvents) {
      try {
        if (!existingEventIds.has(event.id)) {
          // 新しいページを作成
          createNotionPage(event);
          created++;
        }

        // API制限対策で少し待機
        Utilities.sleep(200);

      } catch (error) {
        Logger.log(`イベント同期エラー (${event.title}): ${error.toString()}`);
      }
    }

    Logger.log(`カレンダー→Notion同期完了: 作成${created}件`);

    return {
      created: created,
      updated: updated,
      total: calendarEvents.length
    };

  } catch (error) {
    Logger.log(`カレンダー→Notion同期エラー: ${error.toString()}`);
    throw error;
  }
}

/**
 * NotionからGoogleカレンダーへの同期
 */
function syncNotionToCalendar() {
  try {
    Logger.log('Notion→カレンダー同期を開始...');

    // NotionからイベントIDが設定されていないページを取得
    const filter = {
      property: 'Calendar Event ID',
      rich_text: {
        is_empty: true
      }
    };

    const notionPages = getNotionPages(filter);
    let created = 0;
    let errors = 0;

    for (const page of notionPages) {
      try {
        const eventData = parseNotionPage(page);

        // 必須フィールドのチェック
        if (!eventData.title || !eventData.startTime) {
          continue;
        }

        // Googleカレンダーにイベントを作成
        const createdEvent = createCalendarEvent({
          title: eventData.title,
          description: eventData.description,
          location: eventData.location,
          startTime: eventData.startTime,
          endTime: eventData.endTime || new Date(eventData.startTime.getTime() + 60 * 60 * 1000),
          isAllDay: eventData.isAllDay
        });

        created++;

        // API制限対策で少し待機
        Utilities.sleep(200);

      } catch (error) {
        errors++;
      }
    }

    Logger.log(`Notion→カレンダー同期完了: 作成${created}件、エラー${errors}件`);

    return {
      created: created,
      errors: errors,
      total: notionPages.length
    };

  } catch (error) {
    Logger.log(`Notion→カレンダー同期エラー: ${error.toString()}`);
    throw error;
  }
}

/**
 * Notionデータベースからページを取得
 */
function getNotionPages(filter = null) {
  try {
    const url = `https://api.notion.com/v1/databases/${NOTION_CONFIG.DATABASE_ID}/query`;

    const payload = {
      page_size: 50
    };

    if (filter) {
      payload.filter = filter;
    }

    const response = UrlFetchApp.fetch(url, {
      method: 'POST',
      headers: {
        'Authorization': `Bearer ${NOTION_CONFIG.INTEGRATION_TOKEN}`,
        'Notion-Version': NOTION_CONFIG.API_VERSION,
        'Content-Type': 'application/json'
      },
      payload: JSON.stringify(payload)
    });

    if (response.getResponseCode() !== 200) {
      throw new Error(`Notion API エラー: ${response.getResponseCode()}`);
    }

    const data = JSON.parse(response.getContentText());
    return data.results;

  } catch (error) {
    Logger.log(`Notionページ取得エラー: ${error.toString()}`);
    throw error;
  }
}

/**
 * Notionに新しいページを作成
 */
function createNotionPage(eventData) {
  try {
    const url = 'https://api.notion.com/v1/pages';

    const payload = {
      parent: {
        database_id: NOTION_CONFIG.DATABASE_ID
      },
      properties: {
        'Name': {
          title: [
            {
              text: {
                content: eventData.title || 'Untitled Event'
              }
            }
          ]
        },
        'Date': {
          date: {
            start: eventData.startTime ? eventData.startTime.toISOString() : null,
            end: eventData.endTime ? eventData.endTime.toISOString() : null
          }
        },
        'Calendar Event ID': {
          rich_text: [
            {
              text: {
                content: eventData.id || ''
              }
            }
          ]
        },
        'Description': {
          rich_text: [
            {
              text: {
                content: eventData.description || ''
              }
            }
          ]
        },
        'Location': {
          rich_text: [
            {
              text: {
                content: eventData.location || ''
              }
            }
          ]
        }
      }
    };

    const response = UrlFetchApp.fetch(url, {
      method: 'POST',
      headers: {
        'Authorization': `Bearer ${NOTION_CONFIG.INTEGRATION_TOKEN}`,
        'Notion-Version': NOTION_CONFIG.API_VERSION,
        'Content-Type': 'application/json'
      },
      payload: JSON.stringify(payload)
    });

    if (response.getResponseCode() !== 200) {
      throw new Error(`Notion API エラー: ${response.getResponseCode()}`);
    }

    const data = JSON.parse(response.getContentText());
    return data;

  } catch (error) {
    Logger.log(`Notionページ作成エラー: ${error.toString()}`);
    throw error;
  }
}

/**
 * NotionページからCalendarイベントデータを作成
 */
function parseNotionPage(notionPage) {
  const props = notionPage.properties;

  return {
    notionPageId: notionPage.id,
    title: getNotionText(props['Name']),
    description: getNotionText(props['Description']),
    location: getNotionText(props['Location']),
    startTime: props['Date']?.date?.start ? new Date(props['Date'].date.start) : null,
    endTime: props['Date']?.date?.end ? new Date(props['Date'].date.end) : null,
    isAllDay: false
  };
}

/**
 * Notionテキストプロパティから文字列を取得
 */
function getNotionText(property) {
  if (!property) return '';

  if (property.title && property.title.length > 0) {
    return property.title[0].text.content;
  }

  if (property.rich_text && property.rich_text.length > 0) {
    return property.rich_text[0].text.content;
  }

  return '';
}

/**
 * 同期結果をスプレッドシートに保存
 */
function saveSyncResults(syncResult) {
  try {
    if (!CONFIG.SPREADSHEET_ID) return;

    const spreadsheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = spreadsheet.getSheetByName('Notion同期結果');

    if (!sheet) return;

    const rowData = [
      syncResult.timestamp,
      syncResult.calendarToNotion.created,
      syncResult.calendarToNotion.updated || 0,
      syncResult.notionToCalendar.created,
      syncResult.notionToCalendar.errors || 0,
      'Success'
    ];

    sheet.appendRow(rowData);

  } catch (error) {
    Logger.log(`同期結果保存エラー: ${error.toString()}`);
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
// ユーティリティ・ヘルパー関数
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

/**
 * システムヘルスチェック
 */
function systemHealthCheck() {
  const health = {
    timestamp: new Date(),
    overall: 'healthy',
    components: {},
    issues: []
  };

  try {
    // 基本設定チェック
    health.components.config = CONFIG.SPREADSHEET_ID ? 'healthy' : 'warning';
    if (!CONFIG.SPREADSHEET_ID) {
      health.issues.push('SPREADSHEET_IDが未設定');
      health.overall = 'warning';
    }

    // カレンダー接続チェック
    try {
      CalendarApp.getDefaultCalendar();
      health.components.calendar = 'healthy';
    } catch (error) {
      health.components.calendar = 'error';
      health.issues.push('カレンダーにアクセスできません');
      health.overall = 'error';
    }

    // Notion設定チェック
    if (isNotionConfigured()) {
      try {
        testNotionConnection();
        health.components.notion = 'healthy';
      } catch (error) {
        health.components.notion = 'error';
        health.issues.push('Notion接続エラー');
        if (health.overall === 'healthy') health.overall = 'warning';
      }
    } else {
      health.components.notion = 'disabled';
    }

  } catch (error) {
    health.overall = 'error';
    health.issues.push(`システムチェックエラー: ${error.toString()}`);
  }

  Logger.log('システムヘルスチェック結果:');
  Logger.log(`総合状況: ${health.overall}`);
  Logger.log(`問題: ${health.issues.length}件`);
  health.issues.forEach(issue => Logger.log(`- ${issue}`));

  return health;
}

// ========================================
// Notion設定ガイド・テスト関数
// ========================================

/**
 * Notion設定ガイドを表示
 */
function showNotionConfigGuide() {
  const guide = `
🔧 Notion連携設定ガイド - 3ステップで完了！
=======================================

【ステップ1: Integration Token取得】
1. https://www.notion.so/my-integrations を開く
2. 「New integration」をクリック
3. 名前を「AI Scheduler」等に設定して作成
4. 「Internal Integration Token」をコピー
   （secret_で始まる長い文字列）

【ステップ2: Database ID取得】
1. Notionでスケジュール用データベースを作成
2. 以下のプロパティを追加:
   - Name (Title) - イベント名
   - Date (Date) - 日付範囲
   - Calendar Event ID (Text) - GoogleカレンダーID
   - Description (Text) - 説明
   - Location (Text) - 場所
3. データベースを開いた状態でURLをコピー
4. URLの形式: https://notion.so/workspace/★ここがID★?v=...
5. ★の部分（32文字）がDatabase ID

【ステップ3: 設定入力】
このファイルのNOTION_CONFIGセクション（22-35行目）で:
- INTEGRATION_TOKEN: '取得したToken'
- DATABASE_ID: '取得したID'
に置き換えてください

【ステップ4: 権限設定】
1. データベースの「Share」をクリック
2. 作成したIntegrationを追加
3. 権限を「Can edit」に設定

【確認】
設定完了後、testNotionConnection() を実行してテスト！

💡 困ったら showNotionConfigGuide() でこのガイドを再表示できます
`;

  Logger.log(guide);
  return guide;
}

/**
 * Notion接続テスト
 */
function testNotionConnection() {
  try {
    Logger.log('Notion接続テストを開始...');

    // 設定チェック
    if (!isNotionConfigured()) {
      throw new Error('Notion設定が不完全です。showNotionConfigGuide()でガイドを確認してください。');
    }

    // データベース情報取得テスト
    const pages = getNotionPages();
    Logger.log(`✅ Notion接続テスト成功！既存ページ数: ${pages.length}`);

    return true;

  } catch (error) {
    Logger.log(`❌ Notion接続テスト失敗: ${error.toString()}`);
    return false;
  }
}

// ========================================
// エントリーポイント・使い方ガイド
// ========================================

/**
 * 使い方ガイドを表示
 */
function showUsageGuide() {
  const guide = `
🤖 AIスケジューラ - 使い方ガイド
================================

【基本セットアップ】
1. initialize() - 初期設定（最初に1回実行）
2. testRun() - 動作テスト

【日常使用】
- main() - メイン分析実行（自動実行設定済み）
- generateWeeklyReport() - 週間レポート
- systemHealthCheck() - システム状態確認

【Notion連携（オプション）】
1. showNotionConfigGuide() - 設定ガイド表示
2. testNotionConnection() - 接続テスト
3. runBidirectionalSync() - 双方向同期実行

【メイン機能】
✅ スケジュール衝突の自動検知
✅ 時間使用率の分析
✅ 改善提案の生成
✅ 自動レポート作成
✅ Notion連携（オプション）

【データ管理】
- 分析結果は自動でスプレッドシートに保存
- 衝突情報は専用シートで管理
- Notion同期結果も記録

困った時は systemHealthCheck() でシステム状態を確認！
`;

  Logger.log(guide);
  return guide;
}

// ========================================
// 初回実行用ログ出力
// ========================================

Logger.log(`
🤖 AIスケジューラ（完全統合版）が読み込まれました

📝 セットアップ手順:
1. initialize() を実行 - 初期設定
2. testRun() を実行 - 動作テスト

🔗 Notion連携したい場合:
1. showNotionConfigGuide() - ガイド確認
2. NOTION_CONFIG の設定を入力
3. testNotionConnection() - 接続テスト

🎯 主要機能:
- スケジュール衝突検知・分析
- 自動改善提案
- Notion連携（オプション）
- 自動レポート生成

バージョン: 1.0 Ultimate | ファイル: 統合版
総コード行数: 約1,000行 | 作成日: 2025-09-25

🎁 このファイル1つだけで完結する完全なAIスケジューラです！
`);

/**
 * ========================================
 * 🎁 AIスケジューラ Ultimate - 完成！
 * ========================================
 *
 * このファイル1つで全機能が完結します:
 *
 * ✅ スケジュール衝突検知・分析
 * ✅ AI改善提案システム
 * ✅ Google Sheets 自動データ保存
 * ✅ 自動トリガー（毎時・毎日実行）
 * ✅ Notion双方向連携（オプション）
 * ✅ 週間レポート生成
 * ✅ システムヘルスチェック
 *
 * 🎯 プレゼント用設計:
 * - ファイル1つだけで完結
 * - 5分でセットアップ完了
 * - 初心者でも使用可能
 * - 高度な機能も搭載
 *
 * 使い方:
 * 1. このファイルをGoogle Apps Scriptにコピー&ペースト
 * 2. initialize()を実行
 * 3. 完成！すぐに使用開始
 *
 * 🚀 最高のプレゼント体験をお楽しみください！
 */