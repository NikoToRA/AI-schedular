# AIスケジューラ デプロイメントガイド

## 1. 事前準備

### 必要なアカウント・権限
- Googleアカウント
- Google Apps Script の利用権限
- Google Calendar の読み書き権限
- Google Sheets の作成・編集権限
- Gmail の送信権限

### 必要なAPI設定
1. Google Apps Script API の有効化
2. Google Calendar API の有効化
3. Google Sheets API の有効化
4. Gmail API の有効化

## 2. Google Apps Script プロジェクトの作成

### Step 1: 新しいプロジェクトを作成
1. [Google Apps Script](https://script.google.com) にアクセス
2. 「新しいプロジェクト」をクリック
3. プロジェクト名を「AIスケジューラ」に変更

### Step 2: ソースコードをアップロード
以下のファイルをGASプロジェクトに追加:

```
src/
├── main.gs          → Code.gs (デフォルトファイルを置き換え)
├── calendar.gs      → 新規ファイル
├── ai-analysis.gs   → 新規ファイル
├── data-manager.gs  → 新規ファイル
├── notification.gs  → 新規ファイル
├── triggers.gs      → 新規ファイル
├── web-app.gs       → 新規ファイル
└── utils.gs         → 新規ファイル
```

### Step 3: 必要な権限を設定
```json
{
  "timeZone": "Asia/Tokyo",
  "dependencies": {
    "enabledAdvancedServices": [
      {
        "userSymbol": "Calendar",
        "serviceId": "calendar",
        "version": "v3"
      },
      {
        "userSymbol": "Gmail",
        "serviceId": "gmail",
        "version": "v1"
      }
    ]
  },
  "webapp": {
    "access": "MYSELF",
    "executeAs": "USER_DEPLOYING"
  }
}
```

## 3. 初期設定の実行

### Step 1: 初期化関数の実行
1. GASエディタで `initialize` 関数を選択
2. 「実行」ボタンをクリック
3. 権限の承認を行う
4. 管理用スプレッドシートが作成される

### Step 2: 設定の更新
作成されたスプレッドシートの「設定」シートで以下を設定:

```
設定項目                | 設定値
--------------------|------------------
SPREADSHEET_ID      | (自動設定済み)
通知メール             | your-email@example.com
AI_API_KEY         | (AI APIキーを設定)
営業開始時間            | 9
営業終了時間            | 18
通知レベル             | medium
```

### Step 3: トリガーの設定
```javascript
// GASエディタで実行
setupTriggers();
```

## 4. Webアプリの設定

### Step 1: Webアプリとしてデプロイ
1. GASエディタの「デプロイ」→「新しいデプロイ」
2. 種類: ウェブアプリ
3. 実行者: 自分
4. アクセス権限: 自分のみ
5. 「デプロイ」をクリック

### Step 2: WebアプリURLの取得
デプロイ完了後、WebアプリのURLをコピーして保存

## 5. テスト実行

### Step 1: 基本機能テスト
```javascript
// 各関数を個別に実行してテスト
testRun();              // メイン機能
testNotification();     // 通知機能
```

### Step 2: Webアプリテスト
1. WebアプリURLにアクセス
2. ダッシュボードが正常に表示されることを確認
3. 設定ページで設定変更をテスト

## 6. 本格運用開始

### Step 1: カレンダーにテストイベントを追加
1. Googleカレンダーに複数のイベントを作成
2. 意図的に時間が重複するイベントを作成

### Step 2: 分析実行
```javascript
main(); // メイン分析を実行
```

### Step 3: 結果確認
1. スプレッドシートに分析結果が保存されることを確認
2. 衝突が検出された場合、通知メールが送信されることを確認

## 7. トラブルシューティング

### よくあるエラーと対処法

#### エラー: "SPREADSHEET_IDが設定されていません"
- **原因**: main.gs の CONFIG でスプレッドシートIDが設定されていない
- **対処**: initialize() を実行してスプレッドシートを作成し、IDを設定

#### エラー: "カレンダーにアクセスできません"
- **原因**: Calendar API の権限が不足
- **対処**: GASプロジェクトで権限を再承認

#### エラー: "通知メールが送信されない"
- **原因**: Gmail API の権限不足または設定エラー
- **対処**:
  1. Gmail API の権限を確認
  2. 通知メールアドレスが正しく設定されているか確認

#### エラー: "AI分析でエラーが発生"
- **原因**: AI API キーが未設定または無効
- **対処**:
  1. flexible_requirements.md でAI API選択を確認
  2. 適切なAPIキーを設定

### デバッグ方法
1. GASエディタの「実行トランスクリプト」でログを確認
2. スプレッドシートの「実行ログ」シートでエラー履歴を確認
3. `debugLog()` 関数を使用して詳細ログを有効化

## 8. セキュリティ設定

### 推奨設定
1. **Webアプリアクセス**: 「自分のみ」に設定
2. **API キー**: スプレッドシートに保存（コードに直接記載しない）
3. **通知メール**: 信頼できるアドレスのみ設定
4. **定期バックアップ**: 週次でスプレッドシートをバックアップ

### 注意事項
- API キーや認証情報をGitHubなどに公開しない
- 不要な権限は付与しない
- 定期的にアクセスログを確認する

## 9. 運用・保守

### 定期メンテナンス
- **週次**: バックアップの確認
- **月次**: エラーログの確認
- **四半期**: パフォーマンス分析と最適化

### 機能拡張
`flexible_requirements.md` を更新して新機能を追加:
1. AI API の変更
2. 新しい通知方法の追加
3. レポート機能の強化

## 10. サポート

### ドキュメント
- `design.md`: システム設計
- `require.md`: 要件定義
- `tasks.md`: 開発タスク
- `flexible_requirements.md`: 調整可能要件

### 問題報告
システムに問題がある場合は、以下の情報を収集:
1. エラーメッセージ
2. 実行時刻
3. 実行トランスクリプトのログ
4. スプレッドシートの実行ログ