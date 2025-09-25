# 調整可能要件定義 (Flexible Requirements)

## 現在の開発方針
**Version**: 1.0
**最終更新**: 2025-09-25

## 実装優先度（調整可能）

### 最優先実装 (Priority 1)
1. **基本カレンダー連携**
   - Google Calendar API の基本操作
   - イベントの取得・作成・更新
   - 認証システム

2. **データ管理**
   - Googleスプレッドシートでの設定管理
   - ログ記録機能

### 次優先実装 (Priority 2)
3. **AI分析機能**
   - AI API選択：OpenAI GPT / Gemini / Claude
   - スケジュール衝突検知
   - 最適化提案生成

4. **通知システム**
   - Gmail通知
   - 基本アラート機能

### 将来実装 (Priority 3)
5. **高度機能**
   - 移動時間計算
   - Webアプリ UI
   - レポート機能

## 技術選択（調整可能）

### AI API 選択
- **候補1**: OpenAI GPT-4 API
- **候補2**: Google Gemini API
- **候補3**: Anthropic Claude API
- **現在選択**: 開発中に決定

### データ保存方式
- **メイン**: Googleスプレッドシート
- **サブ**: GAS Properties Service
- **ログ**: Drive テキストファイル

### 通知方式
- **Primary**: Gmail API
- **Secondary**: Google Chat API（将来）

## 開発制約（現在の制限）
- GAS実行時間制限: 6分
- API呼び出し制限: 各サービスの制限内
- 同時実行制限: 30秒間隔

## 調整履歴
- 2025-09-25: 初版作成