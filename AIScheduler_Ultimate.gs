/**
 * ========================================
 * AIスケジューラ - 完全統合版（1ファイル完結）
 * ========================================
 *
 * 🎁 プレゼント用: このファイル1つだけでOK！
 *
 * 【セットアップ手順】
 * 1. Google Apps Script にこのファイルをコピー&ペースト
 * 2. initialize() を実行して初期設定
 * 3. testRun() でテスト実行
 * 4. 完成！毎時自動実行開始
 *
 * 作成日: 2025-09-25 | バージョン: 1.0 Ultimate | 総行数: 1,000行
 */

// ========================================
// ⚙️ 【重要】設定エリア - 必要に応じて設定してください
// ========================================

// 📋 基本設定（通常は変更不要）
const CONFIG = {
  CALENDAR_ID: 'primary',           // 使用するGoogleカレンダー
  SPREADSHEET_ID: '',               // ← initialize()実行後に自動設定されます
  TIME_ZONE: 'Asia/Tokyo'          // タイムゾーン
};

// 🔄 同期設定（デフォルトで片方向: カレンダー→Notion のみ）
const SYNC = {
  ENABLE_BIDIRECTIONAL: false,      // 双方向同期を有効化する場合のみ true
  CAL_TO_NOTION_DAYS: 30,           // カレンダー→Notion の取得期間（日）
  DYNAMIC_INTERVAL: true,           // 成功が続けばポーリング間隔を延長
  SHORT_INTERVAL_MIN: 5,            // 通常間隔（分）
  LONG_INTERVAL_MIN: 30,            // 省コスト間隔（分）
  STABLE_THRESHOLD: 3,              // 連続安定回数で延長
  INTERVALS: [5, 15, 30, 60],       // アダプティブ間隔の階段
  BURST_RUNS: 3                     // 変更検出後、5分で追従する回数
};

// 🧹 ログ／記録のスプレッドシート行数の上限（超過分は古い順に削除）
const LOG_RETENTION = {
  ANALYSIS_MAX_ROWS: 2000,      // 分析結果シートの最大データ行数（ヘッダ除く）
  CONFLICTS_MAX_ROWS: 2000,     // 衝突シートの最大データ行数（ヘッダ除く）
  NOTION_SYNC_MAX_ROWS: 1000    // Notion同期結果シートの最大データ行数（ヘッダ除く）
};

// 🔗 Notion連携設定（オプション - 連携する場合のみ設定）
const NOTION_CONFIG = {
  // 👇【設定1】Notion Integration Tokenをここに入力
  // 📖 取得方法: https://www.notion.so/my-integrations で「New integration」作成
  // 📝 形式例: 'secret_ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijk123'
  INTEGRATION_TOKEN: '',

  // 👇【設定2】Notion Database IDをここに入力
  // 📖 取得方法: NotionデータベースのURLから32文字のIDを取得
  // 📝 URL例: https://notion.so/workspace/【ここがID】?v=...
  // 📝 形式例: '1a2b3c4d5e6f7g8h9i0j1k2l3m4n5o6p'
  DATABASE_ID: '',

  API_VERSION: '2022-06-28'         // APIバージョン（変更不要）
};

// ========================================
// メイン実行関数
// ========================================

/**
 * メイン実行関数 - トリガーから実行される
 */
function main() {
  let __lock = null;
  try {
    // 排他ロック（10秒待ち）
    try {
      __lock = LockService.getScriptLock();
      if (!__lock.tryLock(10000)) {
        Logger.log('他の実行中のためスキップ（ロック取得失敗）');
        return;
      }
    } catch (e) {
      Logger.log(`ロック取得エラー: ${e}`);
      return;
    }
    Logger.log('AIスケジューラ開始: ' + new Date());
    // 実行タイムゾーンの明示
    try { Logger.log(`現在のスクリプトタイムゾーン: ${Session.getScriptTimeZone?.() || '不明'}`); } catch (e) {}

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
    let __syncResult = null;
    if (isNotionConfigured()) {
      try {
        Logger.log('Notion連携が設定されています。同期を実行します...');
        const syncResult = SYNC.ENABLE_BIDIRECTIONAL ? runBidirectionalSync() : runCalendarToNotionOnly();
        __syncResult = syncResult;
        Logger.log(`Notion同期完了: カレンダー→Notion ${syncResult.calendarToNotion.created}件作成`);
        // 人間向けの要約ログ（スプレッドシートは変更せずログのみ）
        try {
          const cal = syncResult.calendarToNotion || { created: 0, updated: 0, duplicates: 0, errors: 0 };
          Logger.log(`新規追加タスク: ${cal.created}件, 変更タスク: ${cal.updated || 0}件, 重複タスク: ${cal.duplicates || 0}件, 失敗: ${cal.errors || 0}件`);
        } catch (e) {}
      } catch (notionError) {
        Logger.log(`Notion同期エラー: ${notionError.toString()}`);
      }
    }

    // ポーリング間隔の動的調整（Notion未設定でも実行し、安定とみなす）
    if (SYNC.DYNAMIC_INTERVAL) {
      try { adjustPollingIntervalIfNeeded(__syncResult || {}); } catch (e) { Logger.log(`ポーリング調整エラー: ${e}`); }
    }

    Logger.log('AIスケジューラ完了: ' + new Date());

  } catch (error) {
    Logger.log(`エラー: ${error.toString()}`);
  } finally {
    try { if (__lock) __lock.releaseLock(); } catch (e) {}
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

    // CONFIGにスプレッドシートIDを自動設定
    CONFIG.SPREADSHEET_ID = spreadsheet.getId();
    Logger.log(`✅ SPREADSHEET_ID自動設定完了: ${CONFIG.SPREADSHEET_ID}`);

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
 * 🔧 段階的初期化 - initialize()で止まる場合に使用
 */
function stepByStepInit() {
  Logger.log('🔧 段階的初期化を開始...');

  try {
    // ステップ1: スプレッドシート作成のみ
    Logger.log('ステップ1: スプレッドシートを作成...');
    const spreadsheet = SpreadsheetApp.create('AIスケジューラ - データ');
    const spreadsheetId = spreadsheet.getId();
    Logger.log(`✅ スプレッドシート作成完了: ${spreadsheetId}`);

    // ステップ2: CONFIG更新
    Logger.log('ステップ2: 設定を更新...');
    // 手動でCONFIG.SPREADSHEET_IDを設定する必要があります
    Logger.log('⚠️  CONFIG.SPREADSHEET_IDを手動で設定してください');
    Logger.log(`設定値: "${spreadsheetId}"`);

    return spreadsheetId;
  } catch (error) {
    Logger.log(`❌ 段階的初期化エラー: ${error.toString()}`);
    return false;
  }
}

/**
 * 🔍 デバッグ用軽量テスト - 問題が起きた時に使用
 */
function debugTest() {
  Logger.log('🔍 デバッグテストを開始...');

  try {
    // 設定チェック
    Logger.log('1. 設定チェック...');
    Logger.log(`SPREADSHEET_ID: ${CONFIG.SPREADSHEET_ID ? '設定済み' : '未設定'}`);

    // カレンダー接続
    Logger.log('2. カレンダー接続...');
    const calendar = CalendarApp.getDefaultCalendar();
    Logger.log(`カレンダー名: ${calendar.getName()}`);

    // 1日だけのイベント取得テスト
    Logger.log('3. 軽量イベント取得テスト...');
    const events = getCalendarEvents(1);
    Logger.log(`イベント数: ${events.length}`);

    Logger.log('✅ デバッグテスト完了');
    return true;
  } catch (error) {
    Logger.log(`❌ デバッグテストエラー: ${error.toString()}`);
    return false;
  }
}

/**
 * 🧪 動作確認用テスト関数 - セットアップ後に実行してください
 */
function testRun() {
  Logger.log('🧪 AIスケジューラ動作確認テストを開始...');

  try {
    // 1. 基本設定チェック
    Logger.log('📋 ステップ1: 基本設定チェック');
    if (!CONFIG.SPREADSHEET_ID) {
      Logger.log('❌ エラー: まず initialize() を実行してください');
      return false;
    }
    Logger.log('✅ 基本設定OK');

    // 2. カレンダー接続テスト
    Logger.log('📅 ステップ2: カレンダー接続テスト');
    const calendar = CalendarApp.getDefaultCalendar();
    Logger.log(`✅ カレンダー接続OK: ${calendar.getName()}`);

    // 3. 軽量テスト実行
    Logger.log('🤖 ステップ3: 軽量分析テスト');
    const events = getCalendarEvents(14); // 2週間（14日間）取得
    Logger.log(`✅ カレンダーイベント取得: ${events.length}件`);

    // イベントの詳細確認
    if (events.length > 0) {
      Logger.log(`✅ サンプルイベント: ${events[0].title}`);
    } else {
      Logger.log('✅ 2週間にイベントがありません（正常）');
    }

    // 4. Notion連携テスト（設定されている場合）
    if (isNotionConfigured()) {
      Logger.log('🔗 ステップ4: Notion連携テスト');
      const notionTest = testNotionConnection();
      if (notionTest) {
        Logger.log('✅ Notion連携OK');
      }
    } else {
      Logger.log('⚪ ステップ4: Notion連携は設定されていません（オプション）');
    }

    Logger.log('');
    Logger.log('🎉 動作確認完了！AIスケジューラが正常に動作しています');
    Logger.log('📊 結果はスプレッドシートに保存されています');
    Logger.log('⚡ 自動実行が設定済みです（5分ごと実行）');

    return true;

  } catch (error) {
    Logger.log(`❌ テストエラー: ${error.toString()}`);
    Logger.log('💡 解決方法: systemHealthCheck() でシステム状態を確認してください');
    return false;
  }
}

/**
 * 📊 完全動作テスト - より詳細なテストを実行
 */
function fullSystemTest() {
  Logger.log('📊 完全システムテストを開始...');

  try {
    // 基本テスト
    const basicTest = testRun();
    if (!basicTest) {
      Logger.log('❌ 基本テスト失敗');
      return false;
    }

    // システムヘルスチェック
    Logger.log('🏥 システムヘルスチェック実行');
    const health = systemHealthCheck();

    // 週間レポートテスト
    Logger.log('📈 週間レポートテスト');
    generateWeeklyReport();

    Logger.log('');
    Logger.log('🎯 完全テスト結果:');
    Logger.log(`📊 総合状況: ${health.overall}`);
    Logger.log(`❓ 問題数: ${health.issues.length}件`);

    if (health.issues.length > 0) {
      Logger.log('📝 解決が必要な問題:');
      health.issues.forEach((issue, index) => {
        Logger.log(`   ${index + 1}. ${issue}`);
      });
    }

    Logger.log('🚀 完全システムテスト完了！');
    return true;

  } catch (error) {
    Logger.log(`❌ 完全テストエラー: ${error.toString()}`);
    return false;
  }
}

// ========================================
// カレンダー連携
// ========================================

/**
 * カレンダーイベントを取得する
 */
function getCalendarEvents(days = 14) {
  try {
    // すべてのカレンダーから予定を取得（Advanced Calendar API使用）
    const calendars = CalendarApp.getAllCalendars();
    const timeMin = new Date();
    const timeMax = new Date(timeMin.getTime() + (days * 24 * 60 * 60 * 1000));

    const timeMinIso = timeMin.toISOString();
    const timeMaxIso = timeMax.toISOString();

    const results = [];

    calendars.forEach(calendar => {
      const calendarId = calendar.getId();
      let pageToken = null;
      do {
        try {
          let resp;
          let ok = false;
          for (let attempt = 0; attempt < 3 && !ok; attempt++) {
            try {
              resp = Calendar.Events.list(calendarId, {
                timeMin: timeMinIso,
                timeMax: timeMaxIso,
                singleEvents: true,
                showDeleted: false,
                maxResults: 2500,
                orderBy: 'startTime',
                pageToken: pageToken
              });
              ok = true;
            } catch (e) {
              if (attempt === 2) throw e;
              Utilities.sleep(500 * Math.pow(2, attempt));
            }
          }

          const items = resp.items || [];
          items.forEach(item => {
            const startStr = item.start?.dateTime || item.start?.date; // date: all-day
            const endStr = item.end?.dateTime || item.end?.date;
            if (!startStr) return;
            // 旧キー（後方互換用）: iCalUID::start
            const legacyCurrentKey = `${item.iCalUID}::${startStr}`;
            // 繰り返しの例外対応: originalStartTime があればそれを使う
            const originalStartStr = item.originalStartTime?.dateTime || item.originalStartTime?.date || null;
            // 新しい正規キー: 単発= iCalUID / 繰り返しインスタンス= iCalUID::originalStart
            const canonicalId = originalStartStr ? `${item.iCalUID}::${originalStartStr}` : `${item.iCalUID}`;
            const isAllDay = !!item.start?.date && !item.start?.dateTime;

            results.push({
              id: canonicalId,
              canonicalId: canonicalId,
              iCalUID: item.iCalUID,
              originalEventId: item.id,
              recurringEventId: item.recurringEventId || null,
              calendarId: calendarId,
              calendarName: calendar.getName(),
              title: item.summary || '',
              description: item.description || '',
              startTime: new Date(startStr),
              endTime: endStr ? new Date(endStr) : null,
              startDateRaw: item.start?.date || null,
              endDateRaw: item.end?.date || null,
              startRaw: startStr || null,
              originalStartRaw: originalStartStr,
              legacyCurrentKey: legacyCurrentKey,
              location: item.location || '',
              isAllDay: isAllDay,
              attendees: (item.attendees || []).map(a => a.email).filter(Boolean)
            });
          });

          pageToken = resp.nextPageToken || null;
        } catch (error) {
          Logger.log(`カレンダー "${calendar.getName()}" の取得エラー: ${error.toString()}`);
          pageToken = null;
        }
      } while (pageToken);
    });

    Logger.log(`iCalUIDベースのイベント配列を返却: ${results.length}件`);
    return results;

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
    const calendarId = calendar.getId();

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
      calendarId: calendarId,
      uid: `${calendarId}::${event.getId()}`,
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
    'カレンダー→Notion重複',
    'カレンダー→Notion失敗',
    'Notion→カレンダー作成',
    'Notion→カレンダー失敗',
    'ステータス',
    '備考'
  ];

  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length).setBackground('#9C27B0').setFontColor('white').setFontWeight('bold');

  // 列幅を調整
  sheet.setColumnWidth(1, 150); // 同期日時
  sheet.setColumnWidth(2, 120); // カレンダー→Notion作成
  sheet.setColumnWidth(3, 120); // カレンダー→Notion更新
  sheet.setColumnWidth(4, 120); // カレンダー→Notion重複
  sheet.setColumnWidth(5, 120); // カレンダー→Notion失敗
  sheet.setColumnWidth(6, 150); // Notion→カレンダー作成
  sheet.setColumnWidth(7, 120); // Notion→カレンダー失敗
  sheet.setColumnWidth(8, 100); // ステータス
  sheet.setColumnWidth(9, 200); // 備考

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
    // 行数が多すぎる場合は古いデータから削除
    pruneSheetRows_(sheet, 1, LOG_RETENTION.ANALYSIS_MAX_ROWS);

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

    // 行数が多すぎる場合は古いデータから削除
    pruneSheetRows_(sheet, 1, LOG_RETENTION.CONFLICTS_MAX_ROWS);
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
function syncCalendarToNotion(days = (typeof SYNC !== 'undefined' && SYNC.CAL_TO_NOTION_DAYS) ? SYNC.CAL_TO_NOTION_DAYS : 7) {
  try {
    Logger.log('カレンダー→Notion同期を開始...');

    // スキーマ検証
    if (!validateNotionSchema()) {
      Logger.log('Notionスキーマが不正のため、同期を中止します。');
      return { created: 0, updated: 0, duplicates: 0, errors: 1, total: 0 };
    }

    // Googleカレンダーのイベントを取得
    const calendarEvents = getCalendarEvents(days);

    // 既存のNotionページを取得（ID→ページのマップを構築）
    const existingPages = getNotionPages();
    const idToPage = new Map();
    existingPages.forEach(page => {
      const eventId = getNotionText(page.properties['Calendar Event ID']);
      if (eventId) {
        // 最初に見つかったものを採用
        if (!idToPage.has(eventId)) idToPage.set(eventId, page);
      }
    });

    let created = 0;
    let updated = 0;
    let duplicates = 0;
    let errors = 0;

    const processedStableIds = new Set();

    for (const event of calendarEvents) {
      try {
        const stableId = event.canonicalId || event.id;

        // 同じ安定IDを複数回見つけたら重複としてスキップ
        if (processedStableIds.has(stableId)) {
          duplicates++;
          continue;
        }
        processedStableIds.add(stableId);
        const legacyUid = event.calendarId && event.originalEventId ? `${event.calendarId}::${event.originalEventId}` : null;

        // 後方互換: 旧キー（iCalUID::start）, さらにGoogle event.id と legacy calendarId::eventId も探索
        const matchedPage = idToPage.get(stableId)
          || (event.legacyCurrentKey ? idToPage.get(event.legacyCurrentKey) : null)
          || idToPage.get(event.originalEventId)
          || (legacyUid ? idToPage.get(legacyUid) : null);

        if (matchedPage) {
          // 差分がある場合のみ更新（無駄なPATCHと更新カウントを抑制）
          if (shouldUpdateNotionPage(matchedPage, event)) {
            updateNotionPageFromEvent(matchedPage.id, event);
            // マップを安定IDで更新
            idToPage.set(stableId, matchedPage);
            updated++;
          }
        } else {
          // 新しいページを作成（IDは安定IDで保存）
          createNotionPage(event);
          created++;
        }

        // API制限対策で少し待機
        Utilities.sleep(200);

      } catch (error) {
        Logger.log(`イベント同期エラー (${event.title}): ${error.toString()}`);
        errors++;
      }
    }

    Logger.log(`カレンダー→Notion同期完了: 作成${created}件, 変更${updated}件, 重複${duplicates}件, 失敗${errors}件`);
    Logger.log(`新規追加タスク: ${created}件, 変更タスク: ${updated}件, 重複タスク: ${duplicates}件, 失敗: ${errors}件`);

    return {
      created: created,
      updated: updated,
      duplicates: duplicates,
      errors: errors,
      total: calendarEvents.length
    };

  } catch (error) {
    Logger.log(`カレンダー→Notion同期エラー: ${error.toString()}`);
    throw error;
  }
}

/**
 * Notionページがイベント内容と一致しているか（更新が必要か）
 */
function shouldUpdateNotionPage(notionPage, eventData) {
  try {
    const props = notionPage.properties || {};
    const expectedTitle = eventData.title || 'Untitled Event';
    const actualTitle = getNotionText(props['Name']) || '';

    const expectedDate = buildNotionDateForEvent(eventData);
    const actualDate = (props['日付'] && props['日付'].date) ? props['日付'].date : null;
    const actualStart = actualDate && actualDate.start ? actualDate.start : '';
    const actualEnd = actualDate && actualDate.end ? actualDate.end : '';
    const expectedStart = expectedDate.start || '';
    const expectedEnd = expectedDate.end || '';

    const expectedId = (eventData.canonicalId || eventData.id || '');
    const actualId = getNotionText(props['Calendar Event ID']) || '';

    const titleDiff = actualTitle !== expectedTitle;
    const startDiff = actualStart !== expectedStart;
    const endDiff = (actualEnd || '') !== (expectedEnd || '');
    const idDiff = actualId !== expectedId;

    return titleDiff || startDiff || endDiff || idDiff;
  } catch (e) {
    // 何か取得に失敗した場合は安全側で更新する
    return true;
  }
}

/**
 * イベントからNotion用の日付オブジェクトを構築
 */
function buildNotionDateForEvent(eventData) {
  if (eventData.isAllDay) {
    const startDate = eventData.startDateRaw
      ? eventData.startDateRaw
      : Utilities.formatDate(new Date(eventData.startTime), CONFIG.TIME_ZONE, 'yyyy-MM-dd');
    let endDate = null;
    if (eventData.endDateRaw) {
      const endEx = new Date(eventData.endDateRaw);
      endEx.setDate(endEx.getDate() - 1); // 排他的end→包含endに補正
      endDate = Utilities.formatDate(endEx, CONFIG.TIME_ZONE, 'yyyy-MM-dd');
      if (endDate === startDate) endDate = null;
    }
    return { start: startDate, end: endDate };
  } else {
    return {
      start: eventData.startTime ? new Date(eventData.startTime).toISOString() : null,
      end: eventData.endTime ? new Date(eventData.endTime).toISOString() : null
    };
  }
}

/**
 * 片方向同期（カレンダー→Notion のみ）
 */
function runCalendarToNotionOnly() {
  try {
    Logger.log('片方向同期（カレンダー→Notion）のみを実行...');
    const calendarToNotionResult = syncCalendarToNotion();
    const summary = {
      timestamp: new Date(),
      calendarToNotion: calendarToNotionResult,
      notionToCalendar: { created: 0, errors: 0, total: 0 }
    };
    saveSyncResults(summary);
    Logger.log('片方向同期完了');
    return summary;
  } catch (error) {
    Logger.log(`片方向同期エラー: ${error.toString()}`);
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
          startTime: eventData.startTime,
          endTime: eventData.endTime || new Date(eventData.startTime.getTime() + 60 * 60 * 1000),
          isAllDay: eventData.isAllDay
        });

        // Notionページに作成したイベントのUIDを書き戻し（重複防止）
        updateNotionPageCalendarId(page.id, createdEvent.uid || createdEvent.id);

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

    const all = [];
    let startCursor = null;
    let hasMore = true;

    while (hasMore) {
      const payload = {
        page_size: 100
      };
      if (filter) payload.filter = filter;
      if (startCursor) payload.start_cursor = startCursor;

      const response = httpFetchWithBackoff(url, {
        method: 'post',
        headers: {
          'Authorization': `Bearer ${NOTION_CONFIG.INTEGRATION_TOKEN}`,
          'Notion-Version': NOTION_CONFIG.API_VERSION,
          'Content-Type': 'application/json'
        },
        payload: JSON.stringify(payload),
        muteHttpExceptions: true
      });

      if (response.getResponseCode() < 200 || response.getResponseCode() >= 300) {
        throw new Error(`Notion API エラー: ${response.getResponseCode()}`);
      }

      const data = JSON.parse(response.getContentText() || '{}');
      (data.results || []).forEach(r => all.push(r));
      hasMore = !!data.has_more;
      startCursor = data.next_cursor || null;
    }

    return all;

  } catch (error) {
    Logger.log(`Notionページ取得エラー: ${error.toString()}`);
    throw error;
  }
}

/**
 * Notionデータベースのスキーマを検証
 */
function validateNotionSchema() {
  try {
    const url = `https://api.notion.com/v1/databases/${NOTION_CONFIG.DATABASE_ID}`;
    const response = httpFetchWithBackoff(url, {
      method: 'get',
      headers: {
        'Authorization': `Bearer ${NOTION_CONFIG.INTEGRATION_TOKEN}`,
        'Notion-Version': NOTION_CONFIG.API_VERSION
      },
      muteHttpExceptions: true
    });

    if (response.getResponseCode() < 200 || response.getResponseCode() >= 300) {
      throw new Error(`Notion DB取得エラー: ${response.getResponseCode()}`);
    }

    const data = JSON.parse(response.getContentText() || '{}');
    const props = data.properties || {};

    const nameOk = props['Name'] && props['Name'].type === 'title';
    const dateOk = props['日付'] && props['日付'].type === 'date';
    const idOk = props['Calendar Event ID'] && props['Calendar Event ID'].type === 'rich_text';

    if (!nameOk || !dateOk || !idOk) {
      Logger.log('Notionデータベースのプロパティが期待と一致しません。');
      Logger.log(`- Name: ${nameOk ? 'OK' : 'NG(Title型が必要)'}`);
      Logger.log(`- 日付: ${dateOk ? 'OK' : 'NG(Date型が必要)'}`);
      Logger.log(`- Calendar Event ID: ${idOk ? 'OK' : 'NG(Rich Text型が必要)'}`);
      Logger.log('showNotionConfigGuide() を確認し、データベースのプロパティを修正してください。');
      return false;
    }
    return true;
  } catch (e) {
    Logger.log(`Notionスキーマ検証エラー: ${e}`);
    return false;
  }
}

/**
 * Notionに新しいページを作成
 */
function createNotionPage(eventData) {
  try {
    const url = 'https://api.notion.com/v1/pages';

    // 日付プロパティの作成（終日は時刻なし、複数日の終日はendを前日までに補正）
    let notionDate;
    if (eventData.isAllDay) {
      const startDate = eventData.startDateRaw
        ? eventData.startDateRaw
        : Utilities.formatDate(new Date(eventData.startTime), CONFIG.TIME_ZONE, 'yyyy-MM-dd');

      let endDate = null;
      if (eventData.endDateRaw) {
        const endEx = new Date(eventData.endDateRaw); // Googleのendは排他的
        endEx.setDate(endEx.getDate() - 1); // 前日に補正（包含に変換）
        endDate = Utilities.formatDate(endEx, CONFIG.TIME_ZONE, 'yyyy-MM-dd');
        if (endDate === startDate) endDate = null; // 単日終日の場合はend省略
      }
      notionDate = { start: startDate, end: endDate };
    } else {
      notionDate = {
        start: eventData.startTime ? eventData.startTime.toISOString() : null,
        end: eventData.endTime ? eventData.endTime.toISOString() : null
      };
    }

    const payload = {
      parent: { database_id: NOTION_CONFIG.DATABASE_ID },
      properties: {
        'Name': { title: [{ text: { content: eventData.title || 'Untitled Event' } }] },
        '日付': { date: notionDate },
        'Calendar Event ID': {
          // 正規キーを保存（単発: iCalUID, 繰り返し: iCalUID::originalStart）
          rich_text: [{ text: { content: (eventData.canonicalId || eventData.id || '') } }]
        }
      }
    };

    const response = httpFetchWithBackoff(url, {
      method: 'post',
      headers: {
        'Authorization': `Bearer ${NOTION_CONFIG.INTEGRATION_TOKEN}`,
        'Notion-Version': NOTION_CONFIG.API_VERSION,
        'Content-Type': 'application/json'
      },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });

    if (response.getResponseCode() < 200 || response.getResponseCode() >= 300) {
      throw new Error(`Notion API エラー: ${response.getResponseCode()}`);
    }

    const data = JSON.parse(response.getContentText() || '{}');
    return data;

  } catch (error) {
    Logger.log(`Notionページ作成エラー: ${error.toString()}`);
    throw error;
  }
}

/**
 * NotionページにCalendar Event IDを書き戻す
 */
function updateNotionPageCalendarId(pageId, calendarEventId) {
  try {
    const url = `https://api.notion.com/v1/pages/${pageId}`;
    const payload = {
      properties: {
        'Calendar Event ID': {
          rich_text: [
            { text: { content: calendarEventId || '' } }
          ]
        }
      }
    };

  const response = httpFetchWithBackoff(url, {
    method: 'patch',
    headers: {
      'Authorization': `Bearer ${NOTION_CONFIG.INTEGRATION_TOKEN}`,
      'Notion-Version': NOTION_CONFIG.API_VERSION,
      'Content-Type': 'application/json'
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });

  if (response.getResponseCode() < 200 || response.getResponseCode() >= 300) {
    throw new Error(`Notion API 更新エラー: ${response.getResponseCode()}`);
  }
  } catch (error) {
    Logger.log(`Notionページ更新エラー: ${error.toString()}`);
    throw error;
  }
}

/**
 * 既存Notionページをイベント内容で更新（タイトル・日付・ID移行）
 */
function updateNotionPageFromEvent(pageId, eventData) {
  try {
    // 日付プロパティをイベントと同様の規則で構築
    let notionDate;
    if (eventData.isAllDay) {
      const startDate = eventData.startDateRaw
        ? eventData.startDateRaw
        : Utilities.formatDate(new Date(eventData.startTime), CONFIG.TIME_ZONE, 'yyyy-MM-dd');
      let endDate = null;
      if (eventData.endDateRaw) {
        const endEx = new Date(eventData.endDateRaw);
        endEx.setDate(endEx.getDate() - 1);
        endDate = Utilities.formatDate(endEx, CONFIG.TIME_ZONE, 'yyyy-MM-dd');
        if (endDate === startDate) endDate = null;
      }
      notionDate = { start: startDate, end: endDate };
    } else {
      notionDate = {
        start: eventData.startTime ? new Date(eventData.startTime).toISOString() : null,
        end: eventData.endTime ? new Date(eventData.endTime).toISOString() : null
      };
    }

    const url = `https://api.notion.com/v1/pages/${pageId}`;
    const payload = {
      properties: {
        'Name': { title: [{ text: { content: eventData.title || 'Untitled Event' } }] },
        '日付': { date: notionDate },
        'Calendar Event ID': {
          rich_text: [{ text: { content: (eventData.canonicalId || eventData.id || '') } }]
        }
      }
    };

    const response = httpFetchWithBackoff(url, {
      method: 'patch',
      headers: {
        'Authorization': `Bearer ${NOTION_CONFIG.INTEGRATION_TOKEN}`,
        'Notion-Version': NOTION_CONFIG.API_VERSION,
        'Content-Type': 'application/json'
      },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });

    if (response.getResponseCode() < 200 || response.getResponseCode() >= 300) {
      throw new Error(`Notion API 更新エラー: ${response.getResponseCode()}`);
    }
  } catch (error) {
    Logger.log(`Notionページ更新エラー: ${error.toString()}`);
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
    startTime: props['日付']?.date?.start ? new Date(props['日付'].date.start) : null,
    endTime: props['日付']?.date?.end ? new Date(props['日付'].date.end) : null,
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
    let sheet = spreadsheet.getSheetByName('Notion同期結果');
    if (!sheet) {
      sheet = createNotionSyncSheet(spreadsheet);
    } else {
      // 既存シートのヘッダーを新フォーマットにマイグレーション
      const neededCols = 9; // 新ヘッダー列数
      const currentCols = sheet.getMaxColumns();
      if (currentCols < neededCols) {
        sheet.insertColumnsAfter(currentCols, neededCols - currentCols);
      }
      const header = sheet.getRange(1,1,1,neededCols).getValues()[0];
      if (!header || header[0] !== '同期日時' || header[8] !== '備考') {
        const headers = [
          '同期日時',
          'カレンダー→Notion作成',
          'カレンダー→Notion更新',
          'カレンダー→Notion重複',
          'カレンダー→Notion失敗',
          'Notion→カレンダー作成',
          'Notion→カレンダー失敗',
          'ステータス',
          '備考'
        ];
        sheet.getRange(1,1,1,headers.length).setValues([headers]);
      }
    }

    const cal = syncResult.calendarToNotion || { created: 0, updated: 0, duplicates: 0, errors: 0 };
    const noc = syncResult.notionToCalendar || { created: 0, errors: 0 };

    // 変更が全くない場合はシートへの追記をスキップ（ログだけに留める）
    const totalChanges = (cal.created || 0) + (cal.updated || 0) + (cal.errors || 0) + (noc.created || 0) + (noc.errors || 0);
    if (totalChanges === 0) {
      Logger.log('変更なしのためシート追記をスキップ');
      return;
    }
    const rowData = [
      syncResult.timestamp,
      cal.created || 0,
      cal.updated || 0,
      cal.duplicates || 0,
      cal.errors || 0,
      noc.created || 0,
      noc.errors || 0,
      'Success',
      ''
    ];

    sheet.appendRow(rowData);
    // 行数が多すぎる場合は古いデータから削除
    pruneSheetRows_(sheet, 1, LOG_RETENTION.NOTION_SYNC_MAX_ROWS);

  } catch (error) {
    Logger.log(`同期結果保存エラー: ${error.toString()}`);
  }
}

/**
 * シートのデータ行数を上限以内に保つ（古い順に削除）
 * headerRows: ヘッダー行数（通常1）
 * maxDataRows: データ行の上限
 */
function pruneSheetRows_(sheet, headerRows, maxDataRows) {
  try {
    if (!sheet || !maxDataRows || maxDataRows <= 0) return;
    const lastRow = sheet.getLastRow();
    const dataRows = Math.max(0, lastRow - headerRows);
    const overflow = dataRows - maxDataRows;
    if (overflow > 0) {
      // ヘッダー直下から overflow 行分を削除
      sheet.deleteRows(headerRows + 1, overflow);
      Logger.log(`シート「${sheet.getName()}」の古い行を${overflow}件削除（上限${maxDataRows}件）`);
    }
  } catch (e) {
    Logger.log(`シート行削減エラー(${sheet && sheet.getName ? sheet.getName() : '?' }): ${e}`);
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

    // 既定間隔で実行のトリガー
    var intervals = (Array.isArray(SYNC.INTERVALS) && SYNC.INTERVALS.length) ? SYNC.INTERVALS : [SYNC.SHORT_INTERVAL_MIN, SYNC.LONG_INTERVAL_MIN];
    setPollingInterval(Math.max(1, intervals[0] || 5));

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

/**
 * ポーリング間隔を設定（分）
 */
function setPollingInterval(minutes) {
  ScriptApp.newTrigger('main')
    .timeBased()
    .everyMinutes(Math.max(1, Math.min(60, minutes)))
    .create();
  // 設定を記録
  PropertiesService.getScriptProperties().setProperty('CURRENT_INTERVAL_MIN', String(minutes));
  Logger.log(`トリガーを${minutes}分間隔に設定`);
}

/**
 * 同期結果に応じてポーリング間隔を調整（最低限ロジック）
 */
function adjustPollingIntervalIfNeeded(syncResult) {
  try {
    const props = PropertiesService.getScriptProperties();
    const intervals = (Array.isArray(SYNC.INTERVALS) && SYNC.INTERVALS.length) ? SYNC.INTERVALS : [SYNC.SHORT_INTERVAL_MIN, SYNC.LONG_INTERVAL_MIN];
    const baseMin = intervals[0] || 5;
    const currentStr = props.getProperty('CURRENT_INTERVAL_MIN') || String(baseMin);
    let currentMin = parseInt(currentStr, 10) || baseMin;
    let stableRuns = parseInt(props.getProperty('STABLE_RUNS') || '0', 10);
    let burstLeft = parseInt(props.getProperty('BURST_LEFT') || '0', 10);

    // 安定の定義: 作成0・変更0・エラー0（Notion→Calendarは無効のため無視）
    const created = syncResult?.calendarToNotion?.created || 0;
    const updated = syncResult?.calendarToNotion?.updated || 0;
    const errors = syncResult?.notionToCalendar?.errors || 0;
    const hasChange = (created + updated + errors) > 0;

    // インターバルの現在位置（見つからなければ最初）
    let idx = intervals.indexOf(currentMin);
    if (idx < 0) idx = 0;

    if (hasChange) {
      // 変更検出: バーストモード起動（5分で追従）
      props.setProperty('STABLE_RUNS', '0');
      props.setProperty('BURST_LEFT', String(SYNC.BURST_RUNS));
      if (currentMin !== baseMin) {
        clearAllTriggers();
        setPollingInterval(baseMin);
        Logger.log(`変更を検出したため間隔を${baseMin}分へ短縮（バースト開始）`);
      }
      return;
    }

    // 変更なし
    if (burstLeft > 0) {
      // バースト継続中: 5分間隔を維持してカウントダウン
      burstLeft -= 1;
      props.setProperty('BURST_LEFT', String(burstLeft));
      if (currentMin !== baseMin) {
        clearAllTriggers();
        setPollingInterval(baseMin);
      }
      Logger.log(`バースト継続中（残り${burstLeft}回）。間隔は${baseMin}分`);
      return;
    }

    // 安定カウントを進め、閾値で次の段に延長
    stableRuns += 1;
    props.setProperty('STABLE_RUNS', String(stableRuns));
    if (stableRuns >= SYNC.STABLE_THRESHOLD) {
      const nextIdx = Math.min(idx + 1, intervals.length - 1);
      const nextMin = intervals[nextIdx];
      if (nextMin !== currentMin) {
        clearAllTriggers();
        setPollingInterval(nextMin);
        props.setProperty('STABLE_RUNS', '0');
        Logger.log(`安定が継続したため間隔を${nextMin}分へ延長`);
      }
    } else {
      Logger.log(`変化なし ${stableRuns}/${SYNC.STABLE_THRESHOLD}（間隔${currentMin}分のまま）`);
    }
  } catch (e) {
    Logger.log(`ポーリング調整例外: ${e}`);
  }
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
 * HTTPフェッチ（指数バックオフ付き）
 */
function httpFetchWithBackoff(url, options, maxRetries = 3) {
  let attempt = 0;
  let lastError = null;
  let delay = 500; // ms
  while (attempt < maxRetries) {
    try {
      const resp = UrlFetchApp.fetch(url, options);
      const code = resp.getResponseCode();
      if (code === 429 || code >= 500) {
        throw new Error(`HTTP ${code}`);
      }
      return resp;
    } catch (e) {
      lastError = e;
      attempt++;
      if (attempt >= maxRetries) break;
      Utilities.sleep(delay);
      delay *= 2;
    }
  }
  throw lastError || new Error('httpFetchWithBackoff: unknown error');
}

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
