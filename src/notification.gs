/**
 * 通知システム
 * Gmail、Slack等での通知送信を行う
 */

/**
 * 分析結果に基づいて通知を送信する
 * @param {Object} analysis 分析結果
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

    // Slackも設定されていれば送信（将来実装）
    // sendSlackNotification(analysis);

    Logger.log('通知送信完了');
    logToSpreadsheet('INFO', 'NOTIFICATION', '通知送信完了', `衝突${analysis.conflicts.length}件`);

  } catch (error) {
    Logger.log(`通知送信エラー: ${error.toString()}`);
    logToSpreadsheet('ERROR', 'NOTIFICATION', '通知送信エラー', error.toString());
  }
}

/**
 * 通知送信が必要かチェック
 * @param {Object} analysis 分析結果
 * @param {string} level 通知レベル
 * @return {boolean} 送信必要フラグ
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
 * @param {Object} analysis 分析結果
 * @return {string} メール件名
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
 * @param {Object} analysis 分析結果
 * @return {string} メール本文
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
- 分析エンジン: ${analysis.analysisVersion || 'Standard'}
`;

  return emailBody;
}

/**
 * Gmail通知を送信する
 * @param {string} subject 件名
 * @param {string} body 本文
 */
function sendGmailNotification(subject, body) {
  try {
    const notificationEmail = CONFIG.NOTIFICATION_EMAIL || getSettingValue('通知メール');

    if (!notificationEmail) {
      Logger.log('通知メールアドレスが設定されていません');
      return;
    }

    // HTMLバージョンも作成
    const htmlBody = convertToHtml(body);

    GmailApp.sendEmail(
      notificationEmail,
      subject,
      body,
      {
        htmlBody: htmlBody,
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
 * @param {Error} error エラーオブジェクト
 */
function sendErrorNotification(error) {
  try {
    const subject = '【エラー】AIスケジューラ実行エラー';
    const date = Utilities.formatDate(new Date(), CONFIG.TIME_ZONE, 'yyyy年MM月dd日 HH:mm:ss');

    const body = `
AIスケジューラでエラーが発生しました

発生日時: ${date}
エラー内容: ${error.toString()}
スタックトレース: ${error.stack || '利用不可'}

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
 * @param {Date|string} dateTime 日時
 * @return {string} フォーマット済み日時
 */
function formatDateTime(dateTime) {
  const date = new Date(dateTime);
  return Utilities.formatDate(date, CONFIG.TIME_ZONE, 'MM/dd HH:mm');
}

/**
 * プレーンテキストをHTMLに変換
 * @param {string} text プレーンテキスト
 * @return {string} HTML
 */
function convertToHtml(text) {
  let html = text
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/\n/g, '<br>');

  // 絵文字の行を強調
  html = html.replace(/^(🚨.*|💡.*|⚡.*|⚠️.*|🤖.*)/gm, '<strong>$1</strong>');

  // タイトル行を大きく
  html = html.replace(/^([A-Za-z].* (?:\([0-9]+件\))?)$/gm, '<h3>$1</h3>');
  html = html.replace(/^(=+|--+)$/gm, '<hr>');

  // インデントを保持
  html = html.replace(/^   /gm, '&nbsp;&nbsp;&nbsp;');

  return `
<html>
<body style="font-family: Arial, sans-serif; line-height: 1.6; color: #333;">
<div style="max-width: 600px; margin: 0 auto; padding: 20px;">
${html}
</div>
</body>
</html>
`;
}

/**
 * Slack通知を送信する（将来実装）
 * @param {Object} analysis 分析結果
 */
function sendSlackNotification(analysis) {
  // Slack Webhook URLが設定されている場合の実装
  // 現在は flexible_requirements.md で将来実装予定
  Logger.log('Slack通知は将来実装予定です');
}

/**
 * 通知設定をテストする
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

/**
 * 通知履歴を取得する
 * @param {number} days 取得日数
 * @return {Array} 通知履歴
 */
function getNotificationHistory(days = 7) {
  try {
    if (!CONFIG.SPREADSHEET_ID) {
      return [];
    }

    const spreadsheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = spreadsheet.getSheetByName('実行ログ');

    if (!sheet) {
      return [];
    }

    const data = sheet.getDataRange().getValues();
    const cutoffDate = new Date(Date.now() - (days * 24 * 60 * 60 * 1000));

    const notifications = [];

    for (let i = 1; i < data.length; i++) {
      const rowDate = new Date(data[i][0]);
      if (rowDate >= cutoffDate && data[i][2] === 'NOTIFICATION') {
        notifications.push({
          timestamp: rowDate,
          level: data[i][1],
          message: data[i][3],
          details: data[i][4]
        });
      }
    }

    return notifications.sort((a, b) => b.timestamp - a.timestamp);

  } catch (error) {
    Logger.log(`通知履歴取得エラー: ${error.toString()}`);
    return [];
  }
}