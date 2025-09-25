# 🔧 Notion連携 設定ガイド

**3ステップでNotion連携を設定できます！**

---

## 📋 必要なもの

1. **Notionアカウント** （無料でOK）
2. **Google Apps Script** にデプロイ済みのAIスケジューラ
3. **5分の時間** ⏰

---

## 🚀 3ステップ設定

### Step 1: Notion Integration を作成

#### 1-1. Notion Integrations ページを開く
```
🔗 https://www.notion.so/my-integrations
```

#### 1-2. New integration をクリック
- **Name**: `AI Scheduler` （任意の名前でOK）
- **Associated workspace**: 使用するワークスペースを選択
- **Submit** をクリック

#### 1-3. Integration Token をコピー
```
✅ 「Internal Integration Token」をコピー
形式: secret_ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijk123
```

---

### Step 2: Notion Database を準備

#### 2-1. スケジュール用データベースを作成
Notionで新しいページを作成し、「Database」を選択

#### 2-2. 必要なプロパティを追加
以下のプロパティを作成してください：

| プロパティ名 | 型 | 必須 |
|------------|-----|------|
| **Name** | Title | ✅ |
| **Date** | Date | ✅ |
| **Calendar Event ID** | Text | ✅ |
| **Description** | Text | - |
| **Location** | Text | - |
| **Status** | Select | - |
| **Last Synced** | Date | - |
| **All Day** | Checkbox | - |
| **Attendees** | Text | - |

#### 2-3. Integration に権限を付与
1. データベースの **「Share」** ボタンをクリック
2. **「Add people, emails, groups, or integrations」** をクリック
3. 作成した **「AI Scheduler」** Integration を検索・選択
4. 権限を **「Can edit」** に設定

#### 2-4. Database ID を取得
```
1. データベースを開く
2. URLをコピー
3. URLの形式: https://notion.so/workspace/★ここがDatabase ID★?v=...
4. ★の部分（32文字の英数字）をコピー
```

**例:**
```
URL: https://notion.so/myworkspace/1a2b3c4d5e6f7g8h9i0j1k2l3m4n5o6p?v=abc123
Database ID: 1a2b3c4d5e6f7g8h9i0j1k2l3m4n5o6p
```

---

### Step 3: Google Apps Script で設定

#### 3-1. `notion-calendar-sync.gs` ファイルを開く

#### 3-2. 設定セクションを編集
**14〜29行目** の `NOTION_CONFIG` を編集：

```javascript
const NOTION_CONFIG = {
  // 👇 Step 1で取得したTokenを入力
  INTEGRATION_TOKEN: 'secret_あなたのToken',

  // 👇 Step 2で取得したDatabase IDを入力
  DATABASE_ID: 'あなたのDatabaseID',

  API_VERSION: '2022-06-28'
};
```

#### 3-3. 保存してテスト
```javascript
// 設定ガイドを表示
showNotionConfigGuide()

// 接続テスト実行
testNotionConnection()
```

---

## ✅ 設定確認

### 成功した場合
```
✅ Notion接続テスト成功！
データベースプロパティ: Name, Date, Calendar Event ID, ...
既存ページ数: 0
```

### エラーの場合
よくあるエラーと解決方法：

#### ❌ Integration Token エラー
```
解決方法:
1. Tokenが正しくコピーされているか確認
2. secret_ で始まっているか確認
3. Integrationが正しいワークスペースで作成されているか確認
```

#### ❌ Database ID エラー
```
解決方法:
1. Database IDが32文字の英数字か確認
2. URLから正しい部分をコピーしているか確認
3. データベースが削除されていないか確認
```

#### ❌ 権限エラー
```
解決方法:
1. データベースでIntegrationが追加されているか確認
2. 権限が「Can edit」に設定されているか確認
3. 正しいIntegrationが選択されているか確認
```

---

## 🚀 同期実行

設定完了後、以下の関数で同期を実行：

```javascript
// 双方向同期実行
runBidirectionalSync()

// カレンダー → Notion のみ
syncCalendarToNotion()

// Notion → カレンダー のみ
syncNotionToCalendar()
```

---

## 🎯 便利な関数

### 設定支援
```javascript
// 設定ガイド表示
showNotionConfigGuide()

// 設定値確認（実行前に値をチェック）
setNotionConfig('your_token', 'your_database_id')
```

### テスト・確認
```javascript
// 接続テスト
testNotionConnection()

// サンプル同期テスト
testNotionSync()
```

### 自動実行設定
```javascript
// 30分間隔の自動同期を設定
setupNotionSyncTriggers()
```

---

## 🆘 サポート

設定でお困りの場合：

1. **Google Apps Script の実行トランスクリプト** でエラー詳細を確認
2. **`showNotionConfigGuide()`** で設定手順を再確認
3. **`testNotionConnection()`** で段階的に問題を特定

---

## 🎊 設定完了！

**おめでとうございます！** 🎉

GoogleカレンダーとNotionが自動同期されるようになりました：

- ✅ カレンダーの予定がNotionに自動保存
- ✅ Notionで作成した予定がカレンダーに反映
- ✅ 重複防止機能で安全に同期
- ✅ 30分ごとの自動実行

**素晴らしいスケジュール管理ライフをお楽しみください！** 🚀

---

*💡 このガイドで解決しない問題があれば、いつでもサポートいたします！*