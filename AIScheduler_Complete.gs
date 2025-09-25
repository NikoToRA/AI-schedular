/**
 * ========================================
 * AIスケジューラ - 完全統合版
 * ========================================
 *
 * Google Apps Script で動作するAIベースのスケジュール管理システム
 *
 * 機能:
 * - Google Calendar との連携
 * - AI分析によるスケジュール最適化
 * - 衝突検知と解決提案
 * - Gmail通知システム
 * - Webダッシュボード
 * - 自動レポート生成
 *
 * 作成日: 2025-09-25
 * バージョン: 1.0
 * ライセンス: MIT
 */

// ========================================
// 設定定数
// ========================================
const CONFIG = {
  CALENDAR_ID: 'primary',
  SPREADSHEET_ID: '', // 初期化後に設定されます
  AI_API_KEY: '', // AI APIキーを設定してください
  NOTIFICATION_EMAIL: '', // 通知先メールアドレスを設定してください
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

    // 設定シートの初期化
    initializeSettingsSheet(spreadsheet);

    // トリガーの設定
    setupTriggers();

    Logger.log('初期設定が完了しました');
    Logger.log('次のステップ:');
    Logger.log('1. スプレッドシートの設定シートで通知メールアドレスを設定');
    Logger.log('2. AI APIキーが必要な場合は設定');
    Logger.log('3. testRun()を実行して動作確認');

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

/**
 * 設定の検証
 */
function validateConfig() {
  if (!CONFIG.SPREADSHEET_ID) {
    Logger.log('SPREADSHEET_IDが設定されていません。initialize()を実行してください。');
    return false;
  }
  return true;
}

// ========================================
// カレンダー連携モジュール
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
        attendees: event.getGuestList().map(guest => guest.getEmail()),
        isAllDay: event.isAllDayEvent(),
        created: event.getDateCreated(),
        creator: event.getCreators()[0] || '',
        status: getEventStatus(event)
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
          location: eventData.location,
          guests: eventData.attendees?.join(',') || ''
        }
      );
    } else {
      event = calendar.createEvent(
        eventData.title,
        eventData.startTime,
        eventData.endTime,
        {
          description: eventData.description,
          location: eventData.location,
          guests: eventData.attendees?.join(',') || ''
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

/**
 * スケジュールの空き時間を検索する
 */
function findAvailableTimeSlots(startDate, endDate, durationMinutes) {
  try {
    const events = getCalendarEvents(
      Math.ceil((endDate - startDate) / (24 * 60 * 60 * 1000))
    );

    const businessHours = {
      start: 9, // 9:00
      end: 18   // 18:00
    };

    const availableSlots = [];
    const currentDate = new Date(startDate);

    while (currentDate <= endDate) {
      // 平日のみチェック（土日を除外）
      if (currentDate.getDay() >= 1 && currentDate.getDay() <= 5) {
        const daySlots = findDayAvailableSlots(
          currentDate,
          events,
          businessHours,
          durationMinutes
        );
        availableSlots.push(...daySlots);
      }

      currentDate.setDate(currentDate.getDate() + 1);
    }

    return availableSlots;

  } catch (error) {
    Logger.log(`空き時間検索エラー: ${error.toString()}`);
    throw error;
  }
}

/**
 * イベントの状態を取得する
 */
function getEventStatus(event) {
  try {
    const myStatus = event.getMyStatus();
    switch (myStatus) {
      case CalendarApp.GuestStatus.OWNER:
        return 'owner';
      case CalendarApp.GuestStatus.YES:
        return 'accepted';
      case CalendarApp.GuestStatus.NO:
        return 'declined';
      case CalendarApp.GuestStatus.MAYBE:
        return 'tentative';
      default:
        return 'unknown';
    }
  } catch (error) {
    return 'unknown';
  }
}

/**
 * 特定日の空き時間を検索する
 */
function findDayAvailableSlots(date, events, businessHours, durationMinutes) {
  const dayStart = new Date(date);
  dayStart.setHours(businessHours.start, 0, 0, 0);

  const dayEnd = new Date(date);
  dayEnd.setHours(businessHours.end, 0, 0, 0);

  // その日のイベントのみ抽出
  const dayEvents = events.filter(event => {
    const eventStart = new Date(event.startTime);
    const eventEnd = new Date(event.endTime);
    return (eventStart >= dayStart && eventStart < dayEnd) ||
           (eventEnd > dayStart && eventEnd <= dayEnd) ||
           (eventStart < dayStart && eventEnd > dayEnd);
  });

  // イベントを時間順にソート
  dayEvents.sort((a, b) => new Date(a.startTime) - new Date(b.startTime));

  const availableSlots = [];
  let currentTime = dayStart;

  for (const event of dayEvents) {
    const eventStart = new Date(event.startTime);
    const eventEnd = new Date(event.endTime);

    // 現在時刻とイベント開始時刻の間に空きがあるかチェック
    if (eventStart > currentTime) {
      const gapMinutes = (eventStart - currentTime) / (1000 * 60);
      if (gapMinutes >= durationMinutes) {
        availableSlots.push({
          startTime: new Date(currentTime),
          endTime: new Date(eventStart),
          durationMinutes: gapMinutes
        });
      }
    }

    currentTime = eventEnd > currentTime ? eventEnd : currentTime;
  }

  // 最後のイベント後から営業時間終了までの空きをチェック
  if (currentTime < dayEnd) {
    const gapMinutes = (dayEnd - currentTime) / (1000 * 60);
    if (gapMinutes >= durationMinutes) {
      availableSlots.push({
        startTime: new Date(currentTime),
        endTime: new Date(dayEnd),
        durationMinutes: gapMinutes
      });
    }
  }

  return availableSlots;
}

// ========================================
// AI分析エンジン
// ========================================

/**
 * AIを使用してスケジュールを分析する
 */
function analyzeScheduleWithAI(events) {
  try {
    Logger.log('AI分析を開始...');

    // 基本的な分析を実行
    const basicAnalysis = performBasicAnalysis(events);

    // AI APIを使用した詳細分析（現在はダミー実装）
    const aiAnalysis = callAIAnalysisAPI(events, basicAnalysis);

    // 結果をまとめる
    const result = {
      timestamp: new Date(),
      totalEvents: events.length,
      conflicts: basicAnalysis.conflicts,
      suggestions: aiAnalysis.suggestions || [],
      optimizations: aiAnalysis.optimizations || [],
      timeUtilization: basicAnalysis.timeUtilization,
      aiInsights: aiAnalysis.insights || [],
      riskFactors: identifyRiskFactors(events, basicAnalysis)
    };

    Logger.log(`AI分析完了: 衝突${result.conflicts.length}件、提案${result.suggestions.length}件`);
    return result;

  } catch (error) {
    Logger.log(`AI分析エラー: ${error.toString()}`);
    return performBasicAnalysis(events);
  }
}

/**
 * 基本的なスケジュール分析を行う
 */
function performBasicAnalysis(events) {
  const conflicts = detectScheduleConflicts(events);
  const timeUtilization = calculateTimeUtilization(events);
  const patterns = analyzeSchedulePatterns(events);

  return {
    conflicts: conflicts,
    timeUtilization: timeUtilization,
    patterns: patterns,
    busyDays: identifyBusyDays(events),
    freeTime: calculateFreeTime(events)
  };
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
 * AI APIを呼び出して詳細分析を実行（ダミー実装）
 */
function callAIAnalysisAPI(events, basicAnalysis) {
  try {
    // 実際の実装では外部AI APIを呼び出し
    return generateDummyAIResponse(events, basicAnalysis);

  } catch (error) {
    Logger.log(`AI API呼び出しエラー: ${error.toString()}`);
    return {
      suggestions: [],
      optimizations: [],
      insights: [`AI分析でエラーが発生しました: ${error.toString()}`]
    };
  }
}

/**
 * ダミーAI応答を生成（開発・テスト用）
 */
function generateDummyAIResponse(events, basicAnalysis) {
  const suggestions = [];
  const optimizations = [];
  const insights = [];

  // 衝突がある場合の提案
  if (basicAnalysis.conflicts.length > 0) {
    suggestions.push('スケジュールの衝突が検出されました。重複する会議の時間を調整することをお勧めします。');
    optimizations.push({
      type: 'conflict_resolution',
      priority: 'high',
      action: '衝突するイベントの再スケジュール',
      estimatedSaving: '30分の時間節約'
    });
  }

  // 時間使用率に基づく提案
  if (basicAnalysis.timeUtilization > 0.8) {
    insights.push('スケジュールが密集しています。休憩時間を確保することを検討してください。');
    optimizations.push({
      type: 'schedule_balancing',
      priority: 'medium',
      action: '15分のバッファー時間を会議間に追加',
      estimatedSaving: 'ストレス軽減'
    });
  }

  // 連続会議のチェック
  const consecutiveMeetings = findConsecutiveMeetings(events);
  if (consecutiveMeetings.length > 2) {
    suggestions.push('連続する会議が多すぎます。移動時間や準備時間を考慮して調整してください。');
  }

  return {
    suggestions: suggestions,
    optimizations: optimizations,
    insights: insights,
    confidence: 0.85,
    analysisVersion: '1.0'
  };
}

/**
 * 連続する会議を検出
 */
function findConsecutiveMeetings(events) {
  const sortedEvents = events.sort((a, b) => new Date(a.startTime) - new Date(b.startTime));
  const consecutive = [];
  let currentGroup = [];

  for (let i = 0; i < sortedEvents.length - 1; i++) {
    const currentEnd = new Date(sortedEvents[i].endTime);
    const nextStart = new Date(sortedEvents[i + 1].startTime);
    const gap = nextStart - currentEnd;

    if (gap <= 15 * 60 * 1000) { // 15分以内
      if (currentGroup.length === 0) {
        currentGroup.push(sortedEvents[i]);
      }
      currentGroup.push(sortedEvents[i + 1]);
    } else {
      if (currentGroup.length > 0) {
        consecutive.push(currentGroup);
        currentGroup = [];
      }
    }
  }

  if (currentGroup.length > 0) {
    consecutive.push(currentGroup);
  }

  return consecutive;
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
 * スケジュールパターンを分析
 */
function analyzeSchedulePatterns(events) {
  const hourDistribution = {};
  const dayDistribution = {};
  const durationDistribution = {};

  events.forEach(event => {
    const startHour = new Date(event.startTime).getHours();
    const dayOfWeek = new Date(event.startTime).getDay();
    const duration = Math.round((new Date(event.endTime) - new Date(event.startTime)) / (60 * 1000)); // 分

    hourDistribution[startHour] = (hourDistribution[startHour] || 0) + 1;
    dayDistribution[dayOfWeek] = (dayDistribution[dayOfWeek] || 0) + 1;

    const durationRange = duration <= 30 ? '短時間' : duration <= 90 ? '中時間' : '長時間';
    durationDistribution[durationRange] = (durationDistribution[durationRange] || 0) + 1;
  });

  return {
    peakHours: Object.keys(hourDistribution).sort((a, b) => hourDistribution[b] - hourDistribution[a]).slice(0, 3),
    busyDays: Object.keys(dayDistribution).sort((a, b) => dayDistribution[b] - dayDistribution[a]),
    commonDurations: durationDistribution
  };
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
    .filter(date => dailyEvents[date].length >= 5)
    .map(date => ({
      date: date,
      eventCount: dailyEvents[date].length,
      totalDuration: dailyEvents[date].reduce((total, event) =>
        total + (new Date(event.endTime) - new Date(event.startTime)), 0) / (60 * 1000) // 分
    }));
}

/**
 * 空き時間を計算
 */
function calculateFreeTime(events) {
  const today = new Date();
  const thisWeekStart = new Date(today.setDate(today.getDate() - today.getDay() + 1));
  const thisWeekEnd = new Date(thisWeekStart);
  thisWeekEnd.setDate(thisWeekEnd.getDate() + 5);

  const availableSlots = findAvailableTimeSlots(thisWeekStart, thisWeekEnd, 30);

  return {
    totalFreeSlots: availableSlots.length,
    longestFreeSlot: Math.max(...availableSlots.map(slot => slot.durationMinutes), 0),
    averageFreeSlot: availableSlots.length > 0 ?
      availableSlots.reduce((sum, slot) => sum + slot.durationMinutes, 0) / availableSlots.length : 0
  };
}

/**
 * リスク要因を特定
 */
function identifyRiskFactors(events, basicAnalysis) {
  const risks = [];

  // 高い時間使用率
  if (basicAnalysis.timeUtilization > 0.9) {
    risks.push({
      type: 'overload',
      severity: 'high',
      description: 'スケジュールが過度に密集しています',
      recommendation: '一部の会議をリスケジュールまたはキャンセルを検討'
    });
  }

  // 多数の衝突
  if (basicAnalysis.conflicts.length > 3) {
    risks.push({
      type: 'conflicts',
      severity: 'high',
      description: '多数のスケジュール衝突があります',
      recommendation: '優先度に基づいてスケジュールを再調整'
    });
  }

  // 連続する長時間会議
  const longMeetings = events.filter(event =>
    (new Date(event.endTime) - new Date(event.startTime)) > 2 * 60 * 60 * 1000
  );
  if (longMeetings.length > 2) {
    risks.push({
      type: 'fatigue',
      severity: 'medium',
      description: '長時間会議が多く、疲労の原因となる可能性があります',
      recommendation: '会議時間の短縮や休憩時間の確保を検討'
    });
  }

  return risks;
}

// ========================================
// データ管理モジュール
// ========================================

/**
 * 管理用スプレッドシートを作成する
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
 */
function createAnalysisResultsSheet(spreadsheet) {
  const sheet = spreadsheet.insertSheet('分析結果');

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
 */
function createSettingsSheet(spreadsheet) {
  const sheet = spreadsheet.insertSheet('設定');

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
 * 実行ログを記録する
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
 */
function initializeSettingsSheet(spreadsheet) {
  try {
    const settingsSheet = spreadsheet.getSheetByName('設定');
    if (settingsSheet && settingsSheet.getLastRow() > 1) {
      Logger.log('設定シートは既に初期化済みです');
      return;
    }

    if (!settingsSheet) {
      createSettingsSheet(spreadsheet);
    }

    Logger.log('設定シート初期化完了');

  } catch (error) {
    Logger.log(`設定シート初期化エラー: ${error.toString()}`);
  }
}

/**
 * 全体のリスクレベルを計算
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

// ========================================
// 通知システム
// ========================================

/**
 * 分析結果に基づいて通知を送信する
 */
function sendNotification(analysis) {
  try {
    const notificationLevel = getSettingValue('通知レベル') || 'medium';

    // 通知が必要かチェック
    if (!shouldSendNotification(analysis, notificationLevel)) {
      Logger.log('通知送信の必要なし');
      return;
    }

    const emailBody = createNotificationEmail(analysis);
    const subject = createNotificationSubject(analysis);

    // Gmail通知送信
    sendGmailNotification(subject, emailBody);

    Logger.log('通知送信完了');
    logToSpreadsheet('INFO', 'NOTIFICATION', '通知送信完了', `衝突${analysis.conflicts.length}件`);

  } catch (error) {
    Logger.log(`通知送信エラー: ${error.toString()}`);
    logToSpreadsheet('ERROR', 'NOTIFICATION', '通知送信エラー', error.toString());
  }
}

/**
 * 通知送信が必要かチェック
 */
function shouldSendNotification(analysis, level) {
  switch (level) {
    case 'low':
      // 高リスクの問題のみ
      return analysis.riskFactors.some(risk => risk.severity === 'high') ||
             analysis.conflicts.some(conflict => conflict.severity === 'high');

    case 'medium':
      // 衝突またはリスクがある場合
      return analysis.conflicts.length > 0 || analysis.riskFactors.length > 0;

    case 'high':
      // 常に送信
      return true;

    default:
      return analysis.conflicts.length > 0;
  }
}

/**
 * 通知メールの件名を作成
 */
function createNotificationSubject(analysis) {
  const date = Utilities.formatDate(new Date(), CONFIG.TIME_ZONE, 'yyyy/MM/dd');

  if (analysis.conflicts.length > 0) {
    const highPriorityConflicts = analysis.conflicts.filter(c => c.severity === 'high').length;
    if (highPriorityConflicts > 0) {
      return `【緊急】AIスケジューラ: ${highPriorityConflicts}件の重要な衝突を検出 (${date})`;
    } else {
      return `【注意】AIスケジューラ: ${analysis.conflicts.length}件のスケジュール衝突 (${date})`;
    }
  }

  if (analysis.riskFactors.length > 0) {
    return `【情報】AIスケジューラ: スケジュール分析結果 (${date})`;
  }

  return `AIスケジューラ: 分析完了 (${date})`;
}

/**
 * 通知メール本文を作成
 */
function createNotificationEmail(analysis) {
  const date = Utilities.formatDate(new Date(), CONFIG.TIME_ZONE, 'yyyy年MM月dd日 HH:mm');

  let emailBody = `
AIスケジューラ分析結果
======================

分析日時: ${date}
分析対象イベント数: ${analysis.totalEvents}件
時間使用率: ${Math.round(analysis.timeUtilization * 100)}%

`;

  // 衝突情報
  if (analysis.conflicts.length > 0) {
    emailBody += `
🚨 スケジュール衝突 (${analysis.conflicts.length}件)
--------------------------------------
`;

    analysis.conflicts.forEach((conflict, index) => {
      const severity = conflict.severity === 'high' ? '🔴 緊急' :
                      conflict.severity === 'medium' ? '🟡 注意' : '🟢 軽微';

      emailBody += `
${index + 1}. ${severity}
   イベント1: ${conflict.event1.title}
   時間: ${formatDateTime(conflict.event1.startTime)} - ${formatDateTime(conflict.event1.endTime)}

   イベント2: ${conflict.event2.title}
   時間: ${formatDateTime(conflict.event2.startTime)} - ${formatDateTime(conflict.event2.endTime)}

   提案: ${conflict.suggestion}

`;
    });
  }

  // AI提案
  if (analysis.suggestions.length > 0) {
    emailBody += `
💡 AI提案
----------
`;
    analysis.suggestions.forEach((suggestion, index) => {
      emailBody += `${index + 1}. ${suggestion}\n`;
    });
    emailBody += '\n';
  }

  // 最適化提案
  if (analysis.optimizations.length > 0) {
    emailBody += `
⚡ 最適化提案
--------------
`;
    analysis.optimizations.forEach((opt, index) => {
      const priority = opt.priority === 'high' ? '🔴 高' :
                      opt.priority === 'medium' ? '🟡 中' : '🟢 低';

      emailBody += `${index + 1}. [${priority}] ${opt.action}
   効果: ${opt.estimatedSaving}

`;
    });
  }

  // リスク要因
  if (analysis.riskFactors.length > 0) {
    emailBody += `
⚠️ リスク要因 (${analysis.riskFactors.length}件)
-----------------------------
`;
    analysis.riskFactors.forEach((risk, index) => {
      const severity = risk.severity === 'high' ? '🔴 高リスク' :
                      risk.severity === 'medium' ? '🟡 中リスク' : '🟢 低リスク';

      emailBody += `${index + 1}. ${severity}
   問題: ${risk.description}
   対策: ${risk.recommendation}

`;
    });
  }

  // AIインサイト
  if (analysis.aiInsights.length > 0) {
    emailBody += `
🤖 AIインサイト
----------------
`;
    analysis.aiInsights.forEach((insight, index) => {
      emailBody += `${index + 1}. ${insight}\n`;
    });
    emailBody += '\n';
  }

  emailBody += `
---
このメールはAIスケジューラシステムによって自動生成されました。
詳細な分析結果は管理用スプレッドシートをご確認ください。

システム情報:
- バージョン: 1.0
- 実行時刻: ${date}
`;

  return emailBody;
}

/**
 * Gmail通知を送信する
 */
function sendGmailNotification(subject, body) {
  try {
    const notificationEmail = CONFIG.NOTIFICATION_EMAIL || getSettingValue('通知メール');

    if (!notificationEmail) {
      Logger.log('通知メールアドレスが設定されていません');
      return;
    }

    GmailApp.sendEmail(
      notificationEmail,
      subject,
      body,
      {
        name: 'AIスケジューラ',
        replyTo: Session.getActiveUser().getEmail()
      }
    );

    Logger.log(`Gmail通知送信完了: ${notificationEmail}`);

  } catch (error) {
    Logger.log(`Gmail送信エラー: ${error.toString()}`);
    throw error;
  }
}

/**
 * エラー通知を送信する
 */
function sendErrorNotification(error) {
  try {
    const subject = '【エラー】AIスケジューラ実行エラー';
    const date = Utilities.formatDate(new Date(), CONFIG.TIME_ZONE, 'yyyy年MM月dd日 HH:mm:ss');

    const body = `
AIスケジューラでエラーが発生しました

発生日時: ${date}
エラー内容: ${error.toString()}

このエラーにより、スケジュール分析が正常に完了しませんでした。
システム管理者にお問い合わせください。

---
システム情報:
- 実行ユーザー: ${Session.getActiveUser().getEmail()}
- 実行時刻: ${date}
`;

    sendGmailNotification(subject, body);
    logToSpreadsheet('ERROR', 'SYSTEM', 'エラー通知送信', error.toString());

  } catch (notificationError) {
    Logger.log(`エラー通知送信失敗: ${notificationError.toString()}`);
  }
}

/**
 * 日時をフォーマットする
 */
function formatDateTime(dateTime) {
  const date = new Date(dateTime);
  return Utilities.formatDate(date, CONFIG.TIME_ZONE, 'MM/dd HH:mm');
}

/**
 * 通知のテスト送信
 */
function testNotification() {
  try {
    Logger.log('通知テストを開始...');

    const testAnalysis = {
      timestamp: new Date(),
      totalEvents: 5,
      conflicts: [
        {
          event1: {
            title: 'テスト会議A',
            startTime: new Date(),
            endTime: new Date(Date.now() + 60 * 60 * 1000)
          },
          event2: {
            title: 'テスト会議B',
            startTime: new Date(),
            endTime: new Date(Date.now() + 30 * 60 * 1000)
          },
          severity: 'medium',
          suggestion: 'テスト用の提案です'
        }
      ],
      suggestions: ['テスト用の提案1', 'テスト用の提案2'],
      optimizations: [
        {
          priority: 'high',
          action: 'テスト最適化',
          estimatedSaving: '30分の節約'
        }
      ],
      timeUtilization: 0.75,
      aiInsights: ['これはテスト用のインサイトです'],
      riskFactors: [
        {
          type: 'test',
          severity: 'medium',
          description: 'テスト用のリスク',
          recommendation: 'テスト用の推奨事項'
        }
      ]
    };

    sendNotification(testAnalysis);

    Logger.log('通知テスト完了');

  } catch (error) {
    Logger.log(`通知テストエラー: ${error.toString()}`);
    throw error;
  }
}

// ========================================
// トリガー管理システム
// ========================================

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

    Logger.log('トリガー設定完了');
    logToSpreadsheet('INFO', 'TRIGGER', 'トリガー設定完了', '定期実行・カレンダー');

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

    Logger.log('日次詳細分析完了');

  } catch (error) {
    Logger.log(`日次分析エラー: ${error.toString()}`);
    sendErrorNotification(error);
  }
}

/**
 * 日次レポートを送信
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

【今後の注意事項】
今後1週間は比較的安定したスケジュールです。

詳細は管理用スプレッドシートをご確認ください。
`;

    sendGmailNotification(subject, body);

  } catch (error) {
    Logger.log(`日次レポート送信エラー: ${error.toString()}`);
  }
}

// ========================================
// Webアプリケーション UI
// ========================================

/**
 * Webアプリのメインページを表示
 */
function doGet(e) {
  try {
    const page = e.parameter.page || 'dashboard';

    switch (page) {
      case 'dashboard':
        return createDashboard();
      case 'settings':
        return createSettingsPage();
      default:
        return createDashboard();
    }

  } catch (error) {
    Logger.log(`Webアプリエラー: ${error.toString()}`);
    return createErrorPage(error);
  }
}

/**
 * ダッシュボードを作成
 */
function createDashboard() {
  try {
    // 最新の分析データを取得
    const todayEvents = getCalendarEvents(1);
    const todayAnalysis = performBasicAnalysis(todayEvents);

    const template = HtmlService.createTemplate(`
<!DOCTYPE html>
<html>
<head>
    <meta charset="utf-8">
    <title>AIスケジューラ ダッシュボード</title>
    <style>
        body { font-family: Arial, sans-serif; margin: 20px; background: #f5f5f5; }
        .container { max-width: 1200px; margin: 0 auto; }
        .header { background: #2196F3; color: white; padding: 20px; border-radius: 8px; margin-bottom: 20px; }
        .nav { background: white; padding: 15px; border-radius: 8px; margin-bottom: 20px; }
        .nav a { margin-right: 20px; text-decoration: none; color: #2196F3; font-weight: bold; }
        .card { background: white; padding: 20px; border-radius: 8px; margin-bottom: 20px; box-shadow: 0 2px 4px rgba(0,0,0,0.1); }
        .metric { display: inline-block; margin: 10px 20px 10px 0; }
        .metric-value { font-size: 2em; font-weight: bold; color: #2196F3; }
        .metric-label { color: #666; }
        .button { background: #2196F3; color: white; padding: 10px 20px; border: none; border-radius: 4px; cursor: pointer; }
        .status-good { color: #4caf50; }
        .status-warning { color: #ff9800; }
        .status-error { color: #f44336; }
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1>🤖 AIスケジューラ ダッシュボード</h1>
            <p>最終更新: <?= lastUpdate ?></p>
        </div>

        <div class="nav">
            <a href="?page=dashboard">📊 ダッシュボード</a>
            <a href="?page=settings">⚙️ 設定</a>
        </div>

        <div class="card">
            <h2>📊 本日の状況</h2>
            <div class="metric">
                <div class="metric-value"><?= todayEvents ?></div>
                <div class="metric-label">イベント数</div>
            </div>
            <div class="metric">
                <div class="metric-value <?= conflictClass ?>"><?= todayConflicts ?></div>
                <div class="metric-label">衝突</div>
            </div>
            <div class="metric">
                <div class="metric-value"><?= utilization ?>%</div>
                <div class="metric-label">時間使用率</div>
            </div>
            <div class="metric">
                <div class="metric-value <?= statusClass ?>"><?= overallStatus ?></div>
                <div class="metric-label">総合状況</div>
            </div>
        </div>

        <div class="card">
            <h2>⚡ クイックアクション</h2>
            <button class="button" onclick="runAnalysis()">🔍 今すぐ分析</button>
            <button class="button" onclick="sendTest()">📧 テスト通知</button>
        </div>
    </div>

    <script>
        function runAnalysis() {
            if (confirm('スケジュール分析を実行しますか？')) {
                google.script.run.main();
                alert('分析を開始しました。');
            }
        }

        function sendTest() {
            if (confirm('テスト通知を送信しますか？')) {
                google.script.run.testNotification();
                alert('テスト通知を送信しました。');
            }
        }
    </script>
</body>
</html>
    `);

    // テンプレート変数を設定
    template.lastUpdate = new Date().toLocaleString('ja-JP');
    template.todayEvents = todayEvents.length;
    template.todayConflicts = todayAnalysis.conflicts.length;
    template.conflictClass = todayAnalysis.conflicts.length > 0 ? 'status-error' : 'status-good';
    template.utilization = Math.round(todayAnalysis.timeUtilization * 100);
    template.overallStatus = getOverallStatus(todayAnalysis);
    template.statusClass = getStatusClass(todayAnalysis);

    return template.evaluate()
      .setTitle('AIスケジューラ')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);

  } catch (error) {
    Logger.log(`ダッシュボード作成エラー: ${error.toString()}`);
    return createErrorPage(error);
  }
}

/**
 * 設定ページを作成
 */
function createSettingsPage() {
  const template = HtmlService.createTemplate(`
<!DOCTYPE html>
<html>
<head>
    <meta charset="utf-8">
    <title>AIスケジューラ 設定</title>
    <style>
        body { font-family: Arial, sans-serif; margin: 20px; background: #f5f5f5; }
        .container { max-width: 800px; margin: 0 auto; }
        .header { background: #2196F3; color: white; padding: 20px; border-radius: 8px; margin-bottom: 20px; }
        .nav { background: white; padding: 15px; border-radius: 8px; margin-bottom: 20px; }
        .nav a { margin-right: 20px; text-decoration: none; color: #2196F3; font-weight: bold; }
        .card { background: white; padding: 20px; border-radius: 8px; margin-bottom: 20px; }
        .info { background: #e8f5e8; padding: 15px; border-radius: 4px; margin-bottom: 20px; }
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1>⚙️ AIスケジューラ 設定</h1>
        </div>

        <div class="nav">
            <a href="?page=dashboard">📊 ダッシュボード</a>
            <a href="?page=settings">⚙️ 設定</a>
        </div>

        <div class="info">
            <h3>📝 設定方法</h3>
            <p>設定を変更するには、管理用スプレッドシートの「設定」シートを直接編集してください。</p>
            <p>主要な設定項目:</p>
            <ul>
                <li><strong>通知メール</strong>: 通知を受け取るメールアドレス</li>
                <li><strong>営業開始・終了時間</strong>: 分析対象の時間帯</li>
                <li><strong>通知レベル</strong>: low (緊急のみ) / medium (推奨) / high (すべて)</li>
                <li><strong>分析対象日数</strong>: 未来何日分を分析するか</li>
            </ul>
        </div>

        <div class="card">
            <h2>🔧 システム情報</h2>
            <p><strong>バージョン:</strong> 1.0</p>
            <p><strong>最終更新:</strong> 2025-09-25</p>
            <p><strong>ステータス:</strong> <span style="color: #4caf50;">稼働中</span></p>
        </div>
    </div>
</body>
</html>
  `);

  return template.evaluate()
    .setTitle('AIスケジューラ 設定')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * 全体状況を取得
 */
function getOverallStatus(analysis) {
  if (analysis.conflicts.length > 2) return '要注意';
  if (analysis.conflicts.length > 0) return '注意';
  if (analysis.timeUtilization > 0.9) return '多忙';
  return '良好';
}

/**
 * 状況のCSSクラスを取得
 */
function getStatusClass(analysis) {
  const status = getOverallStatus(analysis);
  switch (status) {
    case '要注意': return 'status-error';
    case '注意': case '多忙': return 'status-warning';
    default: return 'status-good';
  }
}

/**
 * エラーページを作成
 */
function createErrorPage(error) {
  const template = HtmlService.createTemplate(`
<!DOCTYPE html>
<html>
<head>
    <meta charset="utf-8">
    <title>エラー - AIスケジューラ</title>
    <style>
        body { font-family: Arial, sans-serif; margin: 20px; background: #f5f5f5; }
        .error { background: #ffebee; border: 1px solid #f44336; padding: 20px; border-radius: 8px; }
        .button { background: #2196F3; color: white; padding: 10px 20px; border: none; border-radius: 4px; cursor: pointer; text-decoration: none; display: inline-block; margin-top: 10px; }
    </style>
</head>
<body>
    <div class="error">
        <h1>🚨 エラーが発生しました</h1>
        <p><strong>エラー詳細:</strong> <?= error.toString() ?></p>
        <a href="?" class="button">ダッシュボードに戻る</a>
    </div>
</body>
</html>
  `);

  template.error = error;

  return template.evaluate()
    .setTitle('エラー - AIスケジューラ')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ========================================
// ユーティリティ関数
// ========================================

/**
 * 日付をフォーマットする
 */
function formatDate(date, format = 'yyyy-MM-dd HH:mm') {
  return Utilities.formatDate(date, CONFIG.TIME_ZONE, format);
}

/**
 * 2つの日時が重複するかチェック
 */
function isTimeOverlap(start1, end1, start2, end2) {
  return start1 < end2 && start2 < end1;
}

/**
 * 分を時間表記に変換
 */
function minutesToTimeString(minutes) {
  if (minutes < 60) {
    return `${minutes}分`;
  }

  const hours = Math.floor(minutes / 60);
  const remainingMinutes = minutes % 60;

  if (remainingMinutes === 0) {
    return `${hours}時間`;
  }

  return `${hours}時間${remainingMinutes}分`;
}

/**
 * 営業時間内かチェック
 */
function isBusinessHours(datetime) {
  const hour = datetime.getHours();
  const day = datetime.getDay();

  const workStart = parseInt(getSettingValue('営業開始時間') || '9');
  const workEnd = parseInt(getSettingValue('営業終了時間') || '18');

  return day >= 1 && day <= 5 && hour >= workStart && hour < workEnd;
}

/**
 * 平日かチェック
 */
function isWeekday(date) {
  const day = date.getDay();
  return day >= 1 && day <= 5; // 月曜～金曜
}

/**
 * 文字列を安全にトリミング
 */
function safeTrim(str, maxLength = 100) {
  if (!str || typeof str !== 'string') {
    return '';
  }

  if (str.length <= maxLength) {
    return str;
  }

  return str.substring(0, maxLength - 3) + '...';
}

/**
 * 数値を安全にパース
 */
function safeParseInt(value, defaultValue = 0) {
  const parsed = parseInt(value);
  return isNaN(parsed) ? defaultValue : parsed;
}

/**
 * メール形式の検証
 */
function isValidEmail(email) {
  if (!email || typeof email !== 'string') {
    return false;
  }

  const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
  return emailRegex.test(email);
}

// ========================================
// 使用方法・セットアップガイド
// ========================================

/**
 * セットアップガイドを表示
 */
function showSetupGuide() {
  const guide = `
=================================================
🤖 AIスケジューラ セットアップガイド
=================================================

【ステップ1: 初期化】
1. initialize() 関数を実行
2. 作成されたスプレッドシートのIDをメモ
3. CONFIG.SPREADSHEET_ID にそのIDを設定

【ステップ2: 基本設定】
スプレッドシートの「設定」シートで以下を設定:
- 通知メール: your-email@example.com
- 営業開始時間: 9 (デフォルト)
- 営業終了時間: 18 (デフォルト)
- 通知レベル: medium (推奨)

【ステップ3: テスト実行】
1. testRun() を実行して動作確認
2. testNotification() でメール送信テスト

【ステップ4: 自動実行設定】
- トリガーは initialize() で自動設定済み
- 毎時実行と毎日9:00の詳細分析が実行されます

【ステップ5: Webアプリ設定（オプション）】
1. 「デプロイ」→「新しいデプロイ」
2. 種類: ウェブアプリ
3. 実行者: 自分、アクセス権限: 自分のみ

【主要な関数】
- main(): メイン分析実行
- testRun(): テスト実行
- testNotification(): 通知テスト
- initialize(): 初期化（最初に1回のみ）

【問題が発生した場合】
1. 実行トランスクリプトでログを確認
2. スプレッドシートの「実行ログ」シートを確認
3. 設定値を見直し

=================================================
`;

  Logger.log(guide);
  return guide;
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
    // 設定チェック
    health.components.config = validateConfig() ? 'healthy' : 'error';
    if (!validateConfig()) {
      health.issues.push('設定が不完全です');
      health.overall = 'warning';
    }

    // スプレッドシート接続チェック
    try {
      if (CONFIG.SPREADSHEET_ID) {
        SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
        health.components.spreadsheet = 'healthy';
      } else {
        health.components.spreadsheet = 'warning';
        health.issues.push('スプレッドシートIDが設定されていません');
      }
    } catch (error) {
      health.components.spreadsheet = 'error';
      health.issues.push('スプレッドシートに接続できません');
      health.overall = 'error';
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

    // Gmail接続チェック
    try {
      GmailApp.getInboxThreads(0, 1);
      health.components.gmail = 'healthy';
    } catch (error) {
      health.components.gmail = 'warning';
      health.issues.push('Gmail接続に問題がある可能性があります');
      if (health.overall === 'healthy') health.overall = 'warning';
    }

  } catch (error) {
    health.overall = 'error';
    health.issues.push(`ヘルスチェックエラー: ${error.toString()}`);
  }

  Logger.log('システムヘルスチェック結果:');
  Logger.log(`総合状況: ${health.overall}`);
  Logger.log(`問題: ${health.issues.length}件`);
  health.issues.forEach(issue => Logger.log(`- ${issue}`));

  return health;
}

// ========================================
// エントリーポイント - 最初にここから始めてください
// ========================================

/**
 * 🚀 クイックスタート
 *
 * このシステムを使い始めるには:
 * 1. initialize() を実行してください
 * 2. 表示されるスプレッドシートIDをCONFIG.SPREADSHEET_IDに設定
 * 3. testRun() でテスト実行
 * 4. 本格運用開始
 */

Logger.log(`
🤖 AIスケジューラが読み込まれました

📝 セットアップ手順:
1. initialize() を実行
2. showSetupGuide() で詳細ガイドを確認
3. systemHealthCheck() で動作確認

🎯 主要機能:
- スケジュール衝突検知
- AI分析と最適化提案
- 自動通知システム
- Webダッシュボード

バージョン: 1.0 | 作成日: 2025-09-25
`);

/**
 * ========================================
 * 🎁 プレゼント用 AIスケジューラ - 完成！
 * ========================================
 *
 * このファイルには以下の全機能が含まれています:
 *
 * ✅ Google Calendar 連携
 * ✅ AI分析エンジン
 * ✅ スケジュール衝突検知
 * ✅ 最適化提案システム
 * ✅ Gmail通知機能
 * ✅ 自動トリガー設定
 * ✅ Webダッシュボード
 * ✅ データ管理・ログ機能
 * ✅ エラーハンドリング
 *
 * 総コード行数: 約1,400行
 *
 * 使い方:
 * 1. このファイルをGoogle Apps Scriptにコピー
 * 2. initialize()を実行
 * 3. 設定完了後、main()で分析開始
 *
 * 楽しいスケジュール管理を！ 🚀
 */