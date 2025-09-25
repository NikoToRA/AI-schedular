/**
 * ========================================
 * Notion Calendar Sync モジュール
 * ========================================
 *
 * GoogleカレンダーとNotionデータベースを双方向で同期
 *
 * 必要な設定:
 * - NOTION_INTEGRATION_TOKEN: Notion インテグレーションのトークン
 * - NOTION_DATABASE_ID: 同期対象のNotionデータベースID
 */

// ========================================
// 🔧 NOTION設定 - ここに入力してください！
// ========================================
const NOTION_CONFIG = {
  // 👇 ここにNotion Integration Token を入力（必須）
  // 取得方法: https://www.notion.so/my-integrations で作成
  // 例: 'secret_ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijk123'
  INTEGRATION_TOKEN: '★★★ここにIntegration Tokenを入力★★★',

  // 👇 ここにNotion Database ID を入力（必須）
  // 取得方法: NotionデータベースのURLから取得
  // URL例: https://notion.so/workspace/1234567890abcdef?v=...
  // 例: '1234567890abcdef1234567890abcdef'
  DATABASE_ID: '★★★ここにDatabase IDを入力★★★',

  API_VERSION: '2022-06-28'
};

// ========================================
// Notion API連携
// ========================================

/**
 * Notionデータベースのプロパティ構造を取得
 */
function getNotionDatabaseProperties() {
  try {
    // 設定チェック
    if (!NOTION_CONFIG.INTEGRATION_TOKEN || NOTION_CONFIG.INTEGRATION_TOKEN.includes('★')) {
      throw new Error('❌ NOTION設定エラー: INTEGRATION_TOKENが設定されていません。\n' +
                      '👉 https://www.notion.so/my-integrations でIntegrationを作成し、\n' +
                      '   TokenをNOTION_CONFIG.INTEGRATION_TOKENに設定してください。');
    }

    if (!NOTION_CONFIG.DATABASE_ID || NOTION_CONFIG.DATABASE_ID.includes('★')) {
      throw new Error('❌ NOTION設定エラー: DATABASE_IDが設定されていません。\n' +
                      '👉 NotionデータベースのURLからIDを取得し、\n' +
                      '   NOTION_CONFIG.DATABASE_IDに設定してください。\n' +
                      '   URL例: https://notion.so/workspace/【ここがDatabase ID】?v=...');
    }

    const url = `https://api.notion.com/v1/databases/${NOTION_CONFIG.DATABASE_ID}`;
    const response = UrlFetchApp.fetch(url, {
      method: 'GET',
      headers: {
        'Authorization': `Bearer ${NOTION_CONFIG.INTEGRATION_TOKEN}`,
        'Notion-Version': NOTION_CONFIG.API_VERSION,
        'Content-Type': 'application/json'
      }
    });

    if (response.getResponseCode() !== 200) {
      throw new Error(`Notion API エラー: ${response.getResponseCode()} - ${response.getContentText()}`);
    }

    const data = JSON.parse(response.getContentText());
    Logger.log('Notionデータベースプロパティ取得完了');

    return data.properties;

  } catch (error) {
    Logger.log(`Notionデータベースプロパティ取得エラー: ${error.toString()}`);
    throw error;
  }
}

/**
 * Notionデータベースからページ（イベント）を取得
 */
function getNotionPages(filter = null) {
  try {
    const url = `https://api.notion.com/v1/databases/${NOTION_CONFIG.DATABASE_ID}/query`;

    const payload = {
      page_size: 100
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
      throw new Error(`Notion API エラー: ${response.getResponseCode()} - ${response.getContentText()}`);
    }

    const data = JSON.parse(response.getContentText());
    Logger.log(`Notionページ取得完了: ${data.results.length}件`);

    return data.results;

  } catch (error) {
    Logger.log(`Notionページ取得エラー: ${error.toString()}`);
    throw error;
  }
}

/**
 * Notionに新しいページ（イベント）を作成
 */
function createNotionPage(eventData) {
  try {
    const url = 'https://api.notion.com/v1/pages';

    const payload = {
      parent: {
        database_id: NOTION_CONFIG.DATABASE_ID
      },
      properties: createNotionProperties(eventData)
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
      throw new Error(`Notion API エラー: ${response.getResponseCode()} - ${response.getContentText()}`);
    }

    const data = JSON.parse(response.getContentText());
    Logger.log(`Notionページ作成完了: ${eventData.title}`);

    return data;

  } catch (error) {
    Logger.log(`Notionページ作成エラー: ${error.toString()}`);
    throw error;
  }
}

/**
 * Notionページを更新
 */
function updateNotionPage(pageId, eventData) {
  try {
    const url = `https://api.notion.com/v1/pages/${pageId}`;

    const payload = {
      properties: createNotionProperties(eventData)
    };

    const response = UrlFetchApp.fetch(url, {
      method: 'PATCH',
      headers: {
        'Authorization': `Bearer ${NOTION_CONFIG.INTEGRATION_TOKEN}`,
        'Notion-Version': NOTION_CONFIG.API_VERSION,
        'Content-Type': 'application/json'
      },
      payload: JSON.stringify(payload)
    });

    if (response.getResponseCode() !== 200) {
      throw new Error(`Notion API エラー: ${response.getResponseCode()} - ${response.getContentText()}`);
    }

    const data = JSON.parse(response.getContentText());
    Logger.log(`Notionページ更新完了: ${eventData.title}`);

    return data;

  } catch (error) {
    Logger.log(`Notionページ更新エラー: ${error.toString()}`);
    throw error;
  }
}

// ========================================
// データ変換・マッピング
// ========================================

/**
 * CalendarイベントからNotionプロパティを作成
 */
function createNotionProperties(eventData) {
  const properties = {
    // タイトル（必須）
    'Name': {
      title: [
        {
          text: {
            content: eventData.title || 'Untitled Event'
          }
        }
      ]
    },

    // 日付範囲
    'Date': {
      date: {
        start: eventData.startTime ? eventData.startTime.toISOString() : null,
        end: eventData.endTime ? eventData.endTime.toISOString() : null,
        time_zone: 'Asia/Tokyo'
      }
    },

    // GoogleカレンダーイベントID（重複防止用）
    'Calendar Event ID': {
      rich_text: [
        {
          text: {
            content: eventData.id || ''
          }
        }
      ]
    },

    // 説明
    'Description': {
      rich_text: [
        {
          text: {
            content: eventData.description || ''
          }
        }
      ]
    },

    // 場所
    'Location': {
      rich_text: [
        {
          text: {
            content: eventData.location || ''
          }
        }
      ]
    },

    // ステータス
    'Status': {
      select: {
        name: eventData.status || 'Scheduled'
      }
    },

    // 同期日時（最終更新）
    'Last Synced': {
      date: {
        start: new Date().toISOString()
      }
    },

    // 全日フラグ
    'All Day': {
      checkbox: eventData.isAllDay || false
    }
  };

  // 参加者情報（存在する場合）
  if (eventData.attendees && eventData.attendees.length > 0) {
    properties['Attendees'] = {
      rich_text: [
        {
          text: {
            content: eventData.attendees.join(', ')
          }
        }
      ]
    };
  }

  return properties;
}

/**
 * NotionページからCalendarイベントデータを作成
 */
function parseNotionPage(notionPage) {
  const props = notionPage.properties;

  const eventData = {
    notionPageId: notionPage.id,
    title: getNotionText(props['Name']),
    description: getNotionText(props['Description']),
    location: getNotionText(props['Location']),
    calendarEventId: getNotionText(props['Calendar Event ID']),
    startTime: props['Date']?.date?.start ? new Date(props['Date'].date.start) : null,
    endTime: props['Date']?.date?.end ? new Date(props['Date'].date.end) : null,
    isAllDay: props['All Day']?.checkbox || false,
    status: props['Status']?.select?.name || 'Scheduled',
    attendees: getNotionText(props['Attendees']).split(', ').filter(email => email.trim()),
    lastSynced: props['Last Synced']?.date?.start ? new Date(props['Last Synced'].date.start) : null
  };

  return eventData;
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

// ========================================
// 双方向同期メイン機能
// ========================================

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
    const pagesByEventId = {};

    existingPages.forEach(page => {
      const eventId = getNotionText(page.properties['Calendar Event ID']);
      if (eventId) {
        existingEventIds.add(eventId);
        pagesByEventId[eventId] = page;
      }
    });

    let created = 0;
    let updated = 0;
    let skipped = 0;

    for (const event of calendarEvents) {
      try {
        if (existingEventIds.has(event.id)) {
          // 既存のページを更新
          const existingPage = pagesByEventId[event.id];
          const lastSynced = existingPage.properties['Last Synced']?.date?.start ?
            new Date(existingPage.properties['Last Synced'].date.start) : new Date(0);

          // 最終更新から1時間以上経過している場合のみ更新
          if (new Date() - lastSynced > 60 * 60 * 1000) {
            updateNotionPage(existingPage.id, event);
            updated++;
          } else {
            skipped++;
          }
        } else {
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

    Logger.log(`カレンダー→Notion同期完了: 作成${created}件、更新${updated}件、スキップ${skipped}件`);

    return {
      created: created,
      updated: updated,
      skipped: skipped,
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
      or: [
        {
          property: 'Calendar Event ID',
          rich_text: {
            is_empty: true
          }
        },
        {
          property: 'Calendar Event ID',
          rich_text: {
            equals: ''
          }
        }
      ]
    };

    const notionPages = getNotionPages(filter);
    let created = 0;
    let errors = 0;

    for (const page of notionPages) {
      try {
        const eventData = parseNotionPage(page);

        // 必須フィールドのチェック
        if (!eventData.title || !eventData.startTime) {
          Logger.log(`不完全なイベントをスキップ: ${eventData.title}`);
          continue;
        }

        // Googleカレンダーにイベントを作成
        const createdEvent = createCalendarEvent({
          title: eventData.title,
          description: eventData.description,
          location: eventData.location,
          startTime: eventData.startTime,
          endTime: eventData.endTime || new Date(eventData.startTime.getTime() + 60 * 60 * 1000), // デフォルト1時間
          isAllDay: eventData.isAllDay,
          attendees: eventData.attendees
        });

        // NotionページにカレンダーイベントIDを設定
        updateNotionPage(page.id, {
          ...eventData,
          id: createdEvent.id
        });

        created++;
        Logger.log(`Notion→カレンダー作成完了: ${eventData.title}`);

        // API制限対策で少し待機
        Utilities.sleep(200);

      } catch (error) {
        Logger.log(`Notion→カレンダー同期エラー: ${error.toString()}`);
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
 * 双方向同期の実行
 */
function runBidirectionalSync() {
  try {
    Logger.log('双方向同期を開始...');

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
    Logger.log(`サマリー: ${JSON.stringify(summary, null, 2)}`);

    // 同期結果をスプレッドシートに保存（既存の機能を拡張）
    saveSyncResults(summary);

    return summary;

  } catch (error) {
    Logger.log(`双方向同期エラー: ${error.toString()}`);
    throw error;
  }
}

// ========================================
// 結果保存・レポート
// ========================================

/**
 * 同期結果をスプレッドシートに保存
 */
function saveSyncResults(syncResult) {
  try {
    if (!CONFIG.SPREADSHEET_ID) {
      Logger.log('SPREADSHEET_IDが設定されていません');
      return;
    }

    const spreadsheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    let sheet = spreadsheet.getSheetByName('Notion同期結果');

    // 同期結果シートが存在しない場合は作成
    if (!sheet) {
      sheet = createNotionSyncSheet(spreadsheet);
    }

    // データを準備
    const rowData = [
      syncResult.timestamp,
      syncResult.calendarToNotion.created,
      syncResult.calendarToNotion.updated,
      syncResult.calendarToNotion.skipped,
      syncResult.notionToCalendar.created,
      syncResult.notionToCalendar.errors,
      syncResult.calendarToNotion.total + syncResult.notionToCalendar.total,
      'Success'
    ];

    // 新しい行として追加
    sheet.appendRow(rowData);

    Logger.log('同期結果保存完了');

  } catch (error) {
    Logger.log(`同期結果保存エラー: ${error.toString()}`);
  }
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
    'カレンダー→Notionスキップ',
    'Notion→カレンダー作成',
    'エラー数',
    '総処理数',
    'ステータス'
  ];

  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length).setBackground('#9C27B0').setFontColor('white').setFontWeight('bold');

  // 列幅を調整
  sheet.setColumnWidth(1, 150); // 同期日時
  sheet.setColumnWidth(2, 120); // カレンダー→Notion作成
  sheet.setColumnWidth(3, 120); // カレンダー→Notion更新
  sheet.setColumnWidth(4, 130); // カレンダー→Notionスキップ
  sheet.setColumnWidth(5, 120); // Notion→カレンダー作成
  sheet.setColumnWidth(6, 80);  // エラー数
  sheet.setColumnWidth(7, 80);  // 総処理数
  sheet.setColumnWidth(8, 100); // ステータス

  Logger.log('Notion同期結果シート作成完了');
  return sheet;
}

// ========================================
// 🔧 簡単設定・セットアップ関数
// ========================================

/**
 * 🎯 簡単設定関数 - この関数で設定を入力してください！
 * @param {string} integrationToken - Notion Integration Token
 * @param {string} databaseId - Notion Database ID
 */
function setNotionConfig(integrationToken, databaseId) {
  try {
    Logger.log('🔧 Notion設定を開始...');

    // 入力チェック
    if (!integrationToken || !databaseId) {
      throw new Error('❌ 設定エラー: TokenとDatabase IDの両方を入力してください');
    }

    if (integrationToken.length < 10 || databaseId.length < 10) {
      throw new Error('❌ 設定エラー: TokenまたはIDが短すぎます。正しい値を入力してください');
    }

    // 実際のコード内で設定を変更するためのガイド表示
    Logger.log('✅ 入力された設定:');
    Logger.log(`   Integration Token: ${integrationToken.substring(0, 20)}...`);
    Logger.log(`   Database ID: ${databaseId}`);
    Logger.log('');
    Logger.log('👉 次の手順:');
    Logger.log('1. コード内のNOTION_CONFIGセクションを開く');
    Logger.log('2. INTEGRATION_TOKEN に以下を設定:');
    Logger.log(`   '${integrationToken}'`);
    Logger.log('3. DATABASE_ID に以下を設定:');
    Logger.log(`   '${databaseId}'`);
    Logger.log('4. 保存後、testNotionConnection() でテスト実行');

    return true;

  } catch (error) {
    Logger.log(`❌ 設定エラー: ${error.toString()}`);
    return false;
  }
}

/**
 * 📋 設定ガイドを表示
 */
function showNotionConfigGuide() {
  const guide = `
🔧 Notion設定ガイド - 3ステップで完了！
=======================================

【ステップ1: Integration Token取得】
1. https://www.notion.so/my-integrations を開く
2. 「New integration」をクリック
3. 名前を「AI Scheduler」等に設定して作成
4. 「Internal Integration Token」をコピー
   （secret_で始まる長い文字列）

【ステップ2: Database ID取得】
1. Notionでスケジュール用データベースを作成
2. データベースを開いた状態でURLをコピー
3. URLの形式: https://notion.so/workspace/★ここがID★?v=...
4. ★の部分（32文字）がDatabase ID

【ステップ3: 設定入力】
コード内のNOTION_CONFIGセクション（14-29行目）で:
- INTEGRATION_TOKEN: '取得したToken'
- DATABASE_ID: '取得したID'
に置き換えてください

【確認】
設定完了後、testNotionConnection() を実行してテスト！

💡 困ったら setNotionConfig('Token', 'ID') で値を確認できます
`;

  Logger.log(guide);
  return guide;
}

// ========================================
// セットアップ・テスト関数
// ========================================

/**
 * Notion連携のセットアップガイドを表示
 */
function showNotionSetupGuide() {
  const guide = `
=================================================
🔗 Notion Calendar Sync セットアップガイド
=================================================

【Step 1: Notion Integration作成】
1. https://www.notion.so/my-integrations にアクセス
2. 「New integration」をクリック
3. 名前を「AI Scheduler」等に設定
4. 作成後、「Internal Integration Token」をコピー
5. NOTION_CONFIG.INTEGRATION_TOKEN に設定

【Step 2: Notionデータベース準備】
1. Notionで新しいデータベースを作成
2. 以下のプロパティを追加:
   - Name (Title) - イベント名
   - Date (Date) - 日付範囲
   - Calendar Event ID (Text) - GoogleカレンダーID（自動設定）
   - Description (Text) - 説明
   - Location (Text) - 場所
   - Status (Select) - ステータス
   - Last Synced (Date) - 最終同期日時（自動設定）
   - All Day (Checkbox) - 全日フラグ
   - Attendees (Text) - 参加者

3. データベースURLからIDを取得:
   https://notion.so/workspace/DATABASE_ID?v=...
   ↑この部分をコピー

4. NOTION_CONFIG.DATABASE_ID に設定

【Step 3: 権限設定】
1. データベースの「Share」をクリック
2. 作成したIntegrationを追加
3. 「Edit」権限を付与

【Step 4: テスト実行】
1. testNotionConnection() - 接続テスト
2. runBidirectionalSync() - 同期実行

【主要な関数】
- runBidirectionalSync(): 双方向同期実行
- syncCalendarToNotion(): カレンダー→Notion一方向同期
- syncNotionToCalendar(): Notion→カレンダー一方向同期

=================================================
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
    if (!NOTION_CONFIG.INTEGRATION_TOKEN) {
      throw new Error('NOTION_INTEGRATION_TOKENが設定されていません');
    }

    if (!NOTION_CONFIG.DATABASE_ID) {
      throw new Error('NOTION_DATABASE_IDが設定されていません');
    }

    // データベース情報取得テスト
    const properties = getNotionDatabaseProperties();
    Logger.log(`データベースプロパティ: ${Object.keys(properties).join(', ')}`);

    // ページ取得テスト
    const pages = getNotionPages();
    Logger.log(`既存ページ数: ${pages.length}`);

    Logger.log('✅ Notion接続テスト成功！');
    return true;

  } catch (error) {
    Logger.log(`❌ Notion接続テスト失敗: ${error.toString()}`);
    return false;
  }
}

/**
 * サンプルイベントでテスト同期
 */
function testNotionSync() {
  try {
    Logger.log('Notion同期テストを開始...');

    // テスト用のイベントデータ
    const testEvent = {
      id: 'test_' + Date.now(),
      title: 'テスト同期イベント',
      description: 'Notion連携のテストイベントです',
      location: 'テスト会議室',
      startTime: new Date(Date.now() + 60 * 60 * 1000), // 1時間後
      endTime: new Date(Date.now() + 2 * 60 * 60 * 1000), // 2時間後
      isAllDay: false,
      status: 'confirmed'
    };

    // Notionにテストページを作成
    const notionPage = createNotionPage(testEvent);
    Logger.log(`テストページ作成完了: ${notionPage.id}`);

    Logger.log('✅ Notion同期テスト成功！');
    return true;

  } catch (error) {
    Logger.log(`❌ Notion同期テスト失敗: ${error.toString()}`);
    return false;
  }
}

// ========================================
// 自動実行トリガー（既存システムに追加）
// ========================================

/**
 * 同期用トリガーをセットアップ
 */
function setupNotionSyncTriggers() {
  try {
    Logger.log('Notion同期トリガー設定を開始...');

    // 30分ごとに双方向同期を実行
    ScriptApp.newTrigger('runBidirectionalSync')
      .timeBased()
      .everyMinutes(30)
      .create();

    Logger.log('Notion同期トリガー設定完了: 30分間隔');

  } catch (error) {
    Logger.log(`Notion同期トリガー設定エラー: ${error.toString()}`);
    throw error;
  }
}

// ========================================
// エントリーポイント
// ========================================

Logger.log(`
🔗 Notion Calendar Sync モジュールが読み込まれました

📝 セットアップ手順:
1. showNotionSetupGuide() でガイドを確認
2. NOTION_CONFIG に設定を入力
3. testNotionConnection() で接続テスト
4. runBidirectionalSync() で同期開始

🎯 主要機能:
- GoogleカレンダーとNotionの双方向同期
- 重複防止（Event IDベース）
- 自動トリガー対応
- 詳細ログ・レポート

バージョン: 1.0 | Notion API: ${NOTION_CONFIG.API_VERSION}
`);

/**
 * ========================================
 * 🔗 Notion Calendar Sync - 完成！
 * ========================================
 *
 * このモジュールは以下の機能を提供します:
 *
 * ✅ GoogleカレンダーからNotionへの同期
 * ✅ NotionからGoogleカレンダーへの同期
 * ✅ 重複防止（Calendar Event IDベース）
 * ✅ 双方向同期の自動実行
 * ✅ 詳細なログとレポート機能
 * ✅ エラーハンドリングと復旧
 *
 * 必要な設定:
 * - NOTION_INTEGRATION_TOKEN
 * - NOTION_DATABASE_ID
 *
 * Notionデータベースに必要なプロパティ:
 * - Name (Title) - イベント名
 * - Date (Date) - 日付範囲
 * - Calendar Event ID (Text) - 重複防止用
 * - Description, Location, Status, etc.
 *
 * 使い方:
 * 1. 設定完了後、runBidirectionalSync()を実行
 * 2. 自動トリガーで30分ごとに同期
 */