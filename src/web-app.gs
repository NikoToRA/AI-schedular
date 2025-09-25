/**
 * Webアプリケーション UI
 * ブラウザからアクセス可能な設定画面とダッシュボード
 */

/**
 * Webアプリのメインページを表示
 * @param {Object} e イベントパラメータ
 * @return {HtmlOutput} HTML出力
 */
function doGet(e) {
  try {
    const page = e.parameter.page || 'dashboard';

    switch (page) {
      case 'dashboard':
        return createDashboard();
      case 'settings':
        return createSettingsPage();
      case 'conflicts':
        return createConflictsPage();
      case 'reports':
        return createReportsPage();
      default:
        return createDashboard();
    }

  } catch (error) {
    Logger.log(`Webアプリエラー: ${error.toString()}`);
    return createErrorPage(error);
  }
}

/**
 * POSTリクエストの処理
 * @param {Object} e イベントパラメータ
 * @return {HtmlOutput} HTML出力
 */
function doPost(e) {
  try {
    const action = e.parameter.action;

    switch (action) {
      case 'updateSettings':
        return handleUpdateSettings(e);
      case 'runAnalysis':
        return handleRunAnalysis(e);
      case 'resolveConflict':
        return handleResolveConflict(e);
      default:
        return createDashboard();
    }

  } catch (error) {
    Logger.log(`POST処理エラー: ${error.toString()}`);
    return createErrorPage(error);
  }
}

/**
 * ダッシュボードを作成
 * @return {HtmlOutput} ダッシュボードHTML
 */
function createDashboard() {
  try {
    // 最新の分析データを取得
    const recentData = getPastAnalysisData(7);
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
        .nav a:hover { text-decoration: underline; }
        .card { background: white; padding: 20px; border-radius: 8px; margin-bottom: 20px; box-shadow: 0 2px 4px rgba(0,0,0,0.1); }
        .metric { display: inline-block; margin: 10px 20px 10px 0; }
        .metric-value { font-size: 2em; font-weight: bold; color: #2196F3; }
        .metric-label { color: #666; }
        .conflict { background: #ffebee; border-left: 4px solid #f44336; padding: 10px; margin: 10px 0; }
        .suggestion { background: #e8f5e8; border-left: 4px solid #4caf50; padding: 10px; margin: 10px 0; }
        .button { background: #2196F3; color: white; padding: 10px 20px; border: none; border-radius: 4px; cursor: pointer; }
        .button:hover { background: #1976D2; }
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
            <a href="?page=conflicts">⚠️ 衝突管理</a>
            <a href="?page=reports">📈 レポート</a>
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
            <h2>🚨 最新の衝突</h2>
            <?= conflictsList ?>
        </div>

        <div class="card">
            <h2>💡 AI提案</h2>
            <?= suggestionsList ?>
        </div>

        <div class="card">
            <h2>📈 週間トレンド</h2>
            <?= weeklyTrend ?>
        </div>

        <div class="card">
            <h2>⚡ クイックアクション</h2>
            <button class="button" onclick="runAnalysis()">🔍 今すぐ分析</button>
            <button class="button" onclick="openSettings()">⚙️ 設定変更</button>
            <button class="button" onclick="sendTestNotification()">📧 テスト通知</button>
        </div>
    </div>

    <script>
        function runAnalysis() {
            if (confirm('スケジュール分析を実行しますか？')) {
                const form = document.createElement('form');
                form.method = 'POST';
                form.action = '';

                const action = document.createElement('input');
                action.type = 'hidden';
                action.name = 'action';
                action.value = 'runAnalysis';
                form.appendChild(action);

                document.body.appendChild(form);
                form.submit();
            }
        }

        function openSettings() {
            window.location.href = '?page=settings';
        }

        function sendTestNotification() {
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
    template.conflictsList = formatConflictsList(todayAnalysis.conflicts);
    template.suggestionsList = formatSuggestionsList(['今日は衝突がありません', 'スケジュールは適切に管理されています']);
    template.weeklyTrend = formatWeeklyTrend(recentData);

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
 * @return {HtmlOutput} 設定ページHTML
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
        .form-group { margin-bottom: 15px; }
        .form-group label { display: block; margin-bottom: 5px; font-weight: bold; }
        .form-group input, .form-group select { width: 100%; padding: 8px; border: 1px solid #ddd; border-radius: 4px; }
        .button { background: #2196F3; color: white; padding: 10px 20px; border: none; border-radius: 4px; cursor: pointer; }
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1>⚙️ AIスケジューラ 設定</h1>
        </div>

        <div class="nav">
            <a href="?page=dashboard">📊 ダッシュボード</a>
            <a href="?page=conflicts">⚠️ 衝突管理</a>
            <a href="?page=reports">📈 レポート</a>
            <a href="?page=settings">⚙️ 設定</a>
        </div>

        <form method="POST">
            <input type="hidden" name="action" value="updateSettings">

            <div class="card">
                <h2>基本設定</h2>
                <div class="form-group">
                    <label>通知メールアドレス:</label>
                    <input type="email" name="notificationEmail" placeholder="your-email@example.com">
                </div>
                <div class="form-group">
                    <label>分析対象日数:</label>
                    <input type="number" name="analysisDays" value="7" min="1" max="30">
                </div>
                <div class="form-group">
                    <label>通知レベル:</label>
                    <select name="notificationLevel">
                        <option value="low">低 (緊急のみ)</option>
                        <option value="medium" selected>中 (衝突・リスク)</option>
                        <option value="high">高 (すべて)</option>
                    </select>
                </div>
            </div>

            <div class="card">
                <h2>スケジュール設定</h2>
                <div class="form-group">
                    <label>営業開始時間:</label>
                    <input type="number" name="workStart" value="9" min="0" max="23">
                </div>
                <div class="form-group">
                    <label>営業終了時間:</label>
                    <input type="number" name="workEnd" value="18" min="1" max="24">
                </div>
                <div class="form-group">
                    <label>会議間バッファー時間 (分):</label>
                    <input type="number" name="bufferTime" value="15" min="0" max="60">
                </div>
            </div>

            <div class="card">
                <button type="submit" class="button">💾 設定を保存</button>
            </div>
        </form>
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
 * @param {Object} analysis 分析結果
 * @return {string} 状況
 */
function getOverallStatus(analysis) {
  if (analysis.conflicts.length > 2) return '要注意';
  if (analysis.conflicts.length > 0) return '注意';
  if (analysis.timeUtilization > 0.9) return '多忙';
  return '良好';
}

/**
 * 状況のCSSクラスを取得
 * @param {Object} analysis 分析結果
 * @return {string} CSSクラス
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
 * 衝突リストをフォーマット
 * @param {Array} conflicts 衝突配列
 * @return {string} HTML文字列
 */
function formatConflictsList(conflicts) {
  if (conflicts.length === 0) {
    return '<div class="suggestion">✅ 現在、衝突はありません</div>';
  }

  return conflicts.map(conflict => `
    <div class="conflict">
      <strong>⚠️ ${conflict.severity === 'high' ? '緊急' : '注意'}</strong><br>
      ${conflict.event1.title} ⇔ ${conflict.event2.title}<br>
      <small>提案: ${conflict.suggestion}</small>
    </div>
  `).join('');
}

/**
 * 提案リストをフォーマット
 * @param {Array} suggestions 提案配列
 * @return {string} HTML文字列
 */
function formatSuggestionsList(suggestions) {
  return suggestions.map(suggestion => `
    <div class="suggestion">💡 ${suggestion}</div>
  `).join('');
}

/**
 * 週間トレンドをフォーマット
 * @param {Array} weekData 週間データ
 * @return {string} HTML文字列
 */
function formatWeeklyTrend(weekData) {
  if (weekData.length === 0) {
    return '<p>データが不十分です</p>';
  }

  const avgEvents = Math.round(weekData.reduce((sum, d) => sum + d.eventCount, 0) / weekData.length);
  const totalConflicts = weekData.reduce((sum, d) => sum + d.conflictCount, 0);

  return `
    <p>📊 過去7日間の平均: <strong>${avgEvents}イベント/日</strong></p>
    <p>⚠️ 総衝突件数: <strong>${totalConflicts}件</strong></p>
    <p>📈 トレンド: ${totalConflicts > 5 ? '改善が必要' : '安定'}</p>
  `;
}

/**
 * エラーページを作成
 * @param {Error} error エラーオブジェクト
 * @return {HtmlOutput} エラーページHTML
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

/**
 * 分析実行のハンドラー
 * @param {Object} e イベントパラメータ
 * @return {HtmlOutput} 結果ページ
 */
function handleRunAnalysis(e) {
  try {
    // 分析を実行
    main();

    // 結果ページを作成
    const template = HtmlService.createTemplate(`
<!DOCTYPE html>
<html>
<head>
    <meta charset="utf-8">
    <title>分析完了 - AIスケジューラ</title>
    <style>
        body { font-family: Arial, sans-serif; margin: 20px; background: #f5f5f5; }
        .success { background: #e8f5e8; border: 1px solid #4caf50; padding: 20px; border-radius: 8px; }
        .button { background: #2196F3; color: white; padding: 10px 20px; border: none; border-radius: 4px; cursor: pointer; text-decoration: none; display: inline-block; margin-top: 10px; }
    </style>
</head>
<body>
    <div class="success">
        <h1>✅ 分析が完了しました</h1>
        <p>スケジュール分析が正常に完了しました。結果はスプレッドシートに保存され、必要に応じて通知が送信されました。</p>
        <a href="?" class="button">ダッシュボードに戻る</a>
    </div>
</body>
</html>
    `);

    return template.evaluate();

  } catch (error) {
    return createErrorPage(error);
  }
}