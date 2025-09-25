/**
 * AI分析エンジン
 * スケジュールの分析と最適化提案を行う
 */

/**
 * AIを使用してスケジュールを分析する
 * @param {Array} events カレンダーイベントの配列
 * @return {Object} 分析結果
 */
function analyzeScheduleWithAI(events) {
  try {
    Logger.log('AI分析を開始...');

    // 基本的な分析を実行
    const basicAnalysis = performBasicAnalysis(events);

    // AI APIを使用した詳細分析
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
    // エラー時は基本分析のみ返す
    return performBasicAnalysis(events);
  }
}

/**
 * 基本的なスケジュール分析を行う
 * @param {Array} events イベント配列
 * @return {Object} 基本分析結果
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
 * @param {Array} events イベント配列
 * @return {Array} 衝突のリスト
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
 * @param {Object} event1 イベント1
 * @param {Object} event2 イベント2
 * @return {boolean} 衝突フラグ
 */
function isTimeConflict(event1, event2) {
  const start1 = new Date(event1.startTime);
  const end1 = new Date(event1.endTime);
  const start2 = new Date(event2.startTime);
  const end2 = new Date(event2.endTime);

  // 完全に重複している場合
  return (start1 < end2 && start2 < end1);
}

/**
 * 衝突の深刻度を計算
 * @param {Object} event1 イベント1
 * @param {Object} event2 イベント2
 * @return {string} 深刻度（high/medium/low）
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
 * @param {Object} event1 イベント1
 * @param {Object} event2 イベント2
 * @return {string} 解決提案
 */
function generateConflictSuggestion(event1, event2) {
  const duration1 = new Date(event1.endTime) - new Date(event1.startTime);
  const duration2 = new Date(event2.endTime) - new Date(event2.startTime);

  // より短いイベントを移動させる提案
  if (duration1 < duration2) {
    return `「${event1.title}」を別の時間に移動することを検討してください`;
  } else {
    return `「${event2.title}」を別の時間に移動することを検討してください`;
  }
}

/**
 * AI APIを呼び出して詳細分析を実行
 * @param {Array} events イベント配列
 * @param {Object} basicAnalysis 基本分析結果
 * @return {Object} AI分析結果
 */
function callAIAnalysisAPI(events, basicAnalysis) {
  try {
    // 現在はダミーデータを返す（実際の実装では外部AI APIを呼び出し）
    // OpenAI、Gemini、ClaudeなどのAPIを使用可能

    const prompt = createAnalysisPrompt(events, basicAnalysis);

    // 実際のAPI呼び出しは flexible_requirements.md の設定に基づいて実装
    // const aiResponse = callExternalAI(prompt);

    // ダミーレスポンス
    const aiResponse = generateDummyAIResponse(events, basicAnalysis);

    return aiResponse;

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
 * AI分析用のプロンプトを作成
 * @param {Array} events イベント配列
 * @param {Object} basicAnalysis 基本分析
 * @return {string} プロンプト
 */
function createAnalysisPrompt(events, basicAnalysis) {
  const eventSummary = events.map(event =>
    `${event.title}: ${event.startTime} - ${event.endTime}`
  ).join('\n');

  return `
スケジュール分析をお願いします：

【イベント一覧】
${eventSummary}

【検出された問題】
- 衝突: ${basicAnalysis.conflicts.length}件
- 時間使用率: ${Math.round(basicAnalysis.timeUtilization * 100)}%

【分析項目】
1. スケジュールの改善提案
2. 時間の最適化方法
3. 生産性向上のアドバイス
4. 潜在的なリスク要因

JSON形式で回答してください。
`;
}

/**
 * ダミーAI応答を生成（開発・テスト用）
 * @param {Array} events イベント配列
 * @param {Object} basicAnalysis 基本分析
 * @return {Object} ダミー応答
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
 * @param {Array} events イベント配列
 * @return {Array} 連続会議のグループ
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
 * @param {Array} events イベント配列
 * @return {number} 使用率（0-1の範囲）
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
 * @param {Array} events イベント配列
 * @return {Object} パターン分析結果
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
 * @param {Array} events イベント配列
 * @return {Array} 多忙な日のリスト
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
 * @param {Array} events イベント配列
 * @return {Object} 空き時間情報
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
 * @param {Array} events イベント配列
 * @param {Object} basicAnalysis 基本分析結果
 * @return {Array} リスク要因のリスト
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