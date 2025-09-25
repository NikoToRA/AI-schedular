/**
 * ユーティリティ関数
 * 共通的に使用される便利な関数群
 */

/**
 * 日付をフォーマットする
 * @param {Date} date 日付オブジェクト
 * @param {string} format フォーマット文字列
 * @return {string} フォーマット済み日付
 */
function formatDate(date, format = 'yyyy-MM-dd HH:mm') {
  return Utilities.formatDate(date, CONFIG.TIME_ZONE, format);
}

/**
 * 2つの日時が重複するかチェック
 * @param {Date} start1 開始時刻1
 * @param {Date} end1 終了時刻1
 * @param {Date} start2 開始時刻2
 * @param {Date} end2 終了時刻2
 * @return {boolean} 重複フラグ
 */
function isTimeOverlap(start1, end1, start2, end2) {
  return start1 < end2 && start2 < end1;
}

/**
 * 分を時間表記に変換
 * @param {number} minutes 分数
 * @return {string} 時間表記 (例: "1時間30分")
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
 * @param {Date} datetime 日時
 * @return {boolean} 営業時間内フラグ
 */
function isBusinessHours(datetime) {
  const hour = datetime.getHours();
  const day = datetime.getDay();

  // 平日 (月曜=1, 金曜=5) かつ営業時間内
  const workStart = parseInt(getSettingValue('営業開始時間') || '9');
  const workEnd = parseInt(getSettingValue('営業終了時間') || '18');

  return day >= 1 && day <= 5 && hour >= workStart && hour < workEnd;
}

/**
 * 平日かチェック
 * @param {Date} date 日付
 * @return {boolean} 平日フラグ
 */
function isWeekday(date) {
  const day = date.getDay();
  return day >= 1 && day <= 5; // 月曜～金曜
}

/**
 * 日本の祝日かチェック
 * @param {Date} date 日付
 * @return {boolean} 祝日フラグ
 */
function isJapaneseHoliday(date) {
  // 簡易実装：主要な祝日のみ
  const month = date.getMonth() + 1;
  const dayOfMonth = date.getDate();

  const holidays = [
    [1, 1],   // 元日
    [2, 11],  // 建国記念の日
    [4, 29],  // 昭和の日
    [5, 3],   // 憲法記念日
    [5, 4],   // みどりの日
    [5, 5],   // こどもの日
    [8, 11],  // 山の日
    [11, 3],  // 文化の日
    [11, 23], // 勤労感謝の日
    [12, 23], // 天皇誕生日
  ];

  return holidays.some(([m, d]) => month === m && dayOfMonth === d);
}

/**
 * 文字列を安全にトリミング
 * @param {string} str 文字列
 * @param {number} maxLength 最大長
 * @return {string} トリミング済み文字列
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
 * @param {*} value 値
 * @param {number} defaultValue デフォルト値
 * @return {number} パース済み数値
 */
function safeParseInt(value, defaultValue = 0) {
  const parsed = parseInt(value);
  return isNaN(parsed) ? defaultValue : parsed;
}

/**
 * 配列を安全に取得
 * @param {*} value 値
 * @return {Array} 配列
 */
function safeArray(value) {
  return Array.isArray(value) ? value : [];
}

/**
 * オブジェクトを安全に取得
 * @param {*} value 値
 * @return {Object} オブジェクト
 */
function safeObject(value) {
  return (value && typeof value === 'object' && !Array.isArray(value)) ? value : {};
}

/**
 * 実行時間を測定するデコレータ
 * @param {Function} func 対象関数
 * @param {string} name 関数名
 * @return {Function} 測定付き関数
 */
function measureExecutionTime(func, name) {
  return function(...args) {
    const startTime = Date.now();
    const result = func.apply(this, args);
    const executionTime = Date.now() - startTime;

    Logger.log(`${name} 実行時間: ${executionTime}ms`);
    logToSpreadsheet('INFO', 'PERFORMANCE', `${name} 実行完了`, `${executionTime}ms`);

    return result;
  };
}

/**
 * リトライ機能付き関数実行
 * @param {Function} func 実行する関数
 * @param {number} maxRetries 最大リトライ回数
 * @param {number} delayMs 遅延ミリ秒
 * @return {*} 関数の戻り値
 */
function retryFunction(func, maxRetries = 3, delayMs = 1000) {
  for (let attempt = 1; attempt <= maxRetries; attempt++) {
    try {
      return func();
    } catch (error) {
      Logger.log(`実行失敗 (${attempt}/${maxRetries}): ${error.toString()}`);

      if (attempt === maxRetries) {
        throw error;
      }

      if (delayMs > 0) {
        Utilities.sleep(delayMs);
      }
    }
  }
}

/**
 * APIレート制限を考慮した実行
 * @param {Function} func API呼び出し関数
 * @param {number} delayMs 実行間隔
 */
function rateLimitedExecution(func, delayMs = 1000) {
  const lastExecution = PropertiesService.getScriptProperties().getProperty('lastApiCall');
  const now = Date.now();

  if (lastExecution) {
    const timeSince = now - parseInt(lastExecution);
    if (timeSince < delayMs) {
      const waitTime = delayMs - timeSince;
      Logger.log(`レート制限により${waitTime}ms待機`);
      Utilities.sleep(waitTime);
    }
  }

  const result = func();
  PropertiesService.getScriptProperties().setProperty('lastApiCall', now.toString());

  return result;
}

/**
 * メール形式の検証
 * @param {string} email メールアドレス
 * @return {boolean} 有効フラグ
 */
function isValidEmail(email) {
  if (!email || typeof email !== 'string') {
    return false;
  }

  const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
  return emailRegex.test(email);
}

/**
 * URL形式の検証
 * @param {string} url URL
 * @return {boolean} 有効フラグ
 */
function isValidUrl(url) {
  if (!url || typeof url !== 'string') {
    return false;
  }

  try {
    new URL(url);
    return true;
  } catch {
    return false;
  }
}

/**
 * 設定値を型安全に取得
 * @param {string} key 設定キー
 * @param {*} defaultValue デフォルト値
 * @param {string} type 期待する型
 * @return {*} 型変換済み設定値
 */
function getTypedSetting(key, defaultValue, type = 'string') {
  const value = getSettingValue(key) || defaultValue;

  switch (type) {
    case 'number':
      return safeParseInt(value, defaultValue);
    case 'boolean':
      return value === 'true' || value === true;
    case 'array':
      return typeof value === 'string' ? value.split(',').map(s => s.trim()) : safeArray(value);
    default:
      return value;
  }
}

/**
 * デバッグ情報を記録
 * @param {string} message メッセージ
 * @param {Object} data 追加データ
 */
function debugLog(message, data = null) {
  const debugMode = getTypedSetting('デバッグモード', false, 'boolean');

  if (debugMode) {
    Logger.log(`[DEBUG] ${message}`);
    if (data) {
      Logger.log(`[DEBUG] データ: ${JSON.stringify(data, null, 2)}`);
    }

    logToSpreadsheet('DEBUG', 'SYSTEM', message, data ? JSON.stringify(data) : '');
  }
}

/**
 * パフォーマンス統計を収集
 * @return {Object} パフォーマンス統計
 */
function getPerformanceStats() {
  try {
    const stats = {
      timestamp: new Date(),
      memoryUsage: DriveApp.getStorageUsed(),
      scriptRuntime: 0, // GASでは直接取得不可
      apiQuotaUsage: getApiQuotaUsage(),
      errorCount: getRecentErrorCount(),
      successRate: getSuccessRate()
    };

    return stats;

  } catch (error) {
    Logger.log(`パフォーマンス統計取得エラー: ${error.toString()}`);
    return null;
  }
}

/**
 * APIクォータ使用量を取得（推定）
 * @return {Object} クォータ情報
 */
function getApiQuotaUsage() {
  // GASのクォータ制限に基づく推定値
  return {
    calendarAPI: 'unknown',
    gmailAPI: 'unknown',
    spreadsheetsAPI: 'unknown',
    estimated: true
  };
}

/**
 * 最近のエラー件数を取得
 * @return {number} エラー件数
 */
function getRecentErrorCount() {
  try {
    if (!CONFIG.SPREADSHEET_ID) return 0;

    const spreadsheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const logSheet = spreadsheet.getSheetByName('実行ログ');

    if (!logSheet) return 0;

    const data = logSheet.getDataRange().getValues();
    const yesterday = new Date(Date.now() - 24 * 60 * 60 * 1000);

    let errorCount = 0;
    for (let i = 1; i < data.length; i++) {
      const logDate = new Date(data[i][0]);
      if (logDate >= yesterday && data[i][1] === 'ERROR') {
        errorCount++;
      }
    }

    return errorCount;

  } catch (error) {
    return -1; // 取得エラー
  }
}

/**
 * 成功率を計算
 * @return {number} 成功率 (0-1)
 */
function getSuccessRate() {
  try {
    if (!CONFIG.SPREADSHEET_ID) return -1;

    const spreadsheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const logSheet = spreadsheet.getSheetByName('実行ログ');

    if (!logSheet) return -1;

    const data = logSheet.getDataRange().getValues();
    const yesterday = new Date(Date.now() - 24 * 60 * 60 * 1000);

    let totalCount = 0;
    let errorCount = 0;

    for (let i = 1; i < data.length; i++) {
      const logDate = new Date(data[i][0]);
      if (logDate >= yesterday) {
        totalCount++;
        if (data[i][1] === 'ERROR') {
          errorCount++;
        }
      }
    }

    return totalCount > 0 ? (totalCount - errorCount) / totalCount : 1;

  } catch (error) {
    return -1;
  }
}

/**
 * システムヘルスチェック
 * @return {Object} ヘルス状況
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

    // エラー率チェック
    const successRate = getSuccessRate();
    if (successRate >= 0 && successRate < 0.8) {
      health.issues.push('エラー率が高すぎます');
      health.overall = 'warning';
    }

  } catch (error) {
    health.overall = 'error';
    health.issues.push(`ヘルスチェックエラー: ${error.toString()}`);
  }

  return health;
}

/**
 * キャッシュを利用した設定取得
 * @param {string} key 設定キー
 * @param {number} cacheMinutes キャッシュ有効時間（分）
 * @return {string} 設定値
 */
function getCachedSetting(key, cacheMinutes = 10) {
  const cacheKey = `setting_${key}`;
  const cache = CacheService.getScriptCache();

  const cachedValue = cache.get(cacheKey);
  if (cachedValue) {
    return cachedValue;
  }

  const value = getSettingValue(key);
  if (value) {
    cache.put(cacheKey, value, cacheMinutes * 60);
  }

  return value;
}