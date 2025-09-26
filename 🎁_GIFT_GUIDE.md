# 🎁 AIスケジューラ - プレゼント用クイックガイド（最新版）

この1ファイルで完成する「AIスケジューラ」を、すぐ使える形で贈れます。Googleカレンダーを真実源とし、Notionへ自動で整理・同期します（片方向）。衝突検知・分析、ログ記録、安定稼働の工夫も搭載済みです。

---

## ✅ これは何？（1分で把握）
- 1ファイル完結: `AIScheduler_Ultimate.gs` を Google Apps Script に貼るだけ
- カレンダー主導: カレンダー→Notion の片方向同期（Notionの変更は無視）
- 賢い同期キー: どのカレンダーでも同一予定を識別（iCalUID / originalStart）
- 終日対応: 日付のみでNotionに登録、複数日は最終日を“含む”形に補正
- ログ充実: 作成/変更/重複/失敗件数を実行ログとシートに記録
- 安定運用: 排他ロック、指数バックオフ、スキーマ検証、動的トリガー（5→30分）

---

## 🚀 5分セットアップ手順（贈られた方の操作）
1) Google Apps Script を開く: https://script.google.com
2) 新規プロジェクト作成 → `コード.gs` を削除 → `AIScheduler_Ultimate.gs` の内容を丸ごと貼り付け
3) メニューから関数 `initialize()` を選び実行（初回は権限承認）
   - 管理スプレッドシートが自動作成され、IDが自動設定されます
   - トリガーが5分間隔で作成されます（後述の自動調整あり）
4) 動作確認: `testRun()` → ログに取得件数や分析結果が出ます
5) （任意）Notion連携を使う場合のみ、次の「Notion設定」を実施

---

## 🔗 Notion設定（任意・使う場合のみ）
1) Notion Integration を作成してトークン取得（https://www.notion.so/my-integrations）
2) Notionデータベースを用意し、次のプロパティを作成
   - Name（Title）
   - 日付（Date）
   - Calendar Event ID（Rich Text）
3) Notionで対象データベースを「Share」→ 作成したIntegrationを追加（Can edit）
4) スクリプトの NOTION_CONFIG に `INTEGRATION_TOKEN` と `DATABASE_ID` を設定
5) `testNotionConnection()` を実行 → 接続OKか確認

補足:
- 同期開始時にスキーマ検証を実施します（上記プロパティと型が必須）。不一致ならログにガイドを出し、同期を中止します。

---

## ⚙️ どう同期される？（重要な挙動）
- 片方向のみ: カレンダー → Notion（Notionの編集はカレンダーへ反映しません）
- 既存更新/新規作成の判定キー（カレンダー横断で重複なし）
  - 単発: 正規ID＝`iCalUID`
  - 繰り返しの例外インスタンス: 正規ID＝`iCalUID::originalStartTime`
  - 旧キー互換: `iCalUID::currentStart` / `event.id` / `calendarId::event.id` でも突合し、正規IDに移行
- 終日イベント: Notionには時間なしの日付で保存。複数日はGoogleの“排他的end”を1日戻し、最終日を含む範囲に補正
- ログ/記録:
  - 実行ログ: 「新規追加タスク:X件, 変更タスク:Y件, 重複タスク:Z件, 失敗:E件」
  - スプレッドシート: 管理用スプレッドシートに「分析結果」「スケジュール衝突」「Notion同期結果」を保存
- 安定運用:
  - 排他ロック: 並行実行を防止
  - 指数バックオフ: 429/5xx で自動リトライ（最大3回）
  - 動的トリガー: 変化なし（作成0/変更0/失敗0）が連続すると30分間隔へ、変化があれば5分に戻る

---

## 📅 日々の使い方
- ふだんはカレンダーを編集するだけ（移動・時間変更OK）
  - 同じNotionページが自動更新されます（新規は作られません）
  - 別カレンダーに“コピー”した場合は別タスクとして扱われる設計です
- 手動で同期したいとき
  - `main()`（通常はこちら。分析＋同期）
  - `runCalendarToNotionOnly()`（同期だけを実行）

---

## 🧪 うまく動かないとき（チェックリスト）
1) Notion設定は正しい？
   - `INTEGRATION_TOKEN`/`DATABASE_ID` が設定済みか
   - DBのプロパティ名・型が合っているか（Name=Title, 日付=Date, Calendar Event ID=Rich Text）
2) 同期ログは出ている？
   - 実行ログに「新規/変更/重複/失敗」のサマリが出るか
3) スプレッドシートは一致？
   - `CONFIG.SPREADSHEET_ID` のIDのシートを開いているか
   - `Notion同期結果` シートに行が追記されているか
4) 実行時間・レート制限
   - 大量データ時は時間切れや429が起きることがあります（自動リトライ・次回再試行に期待）

便利関数:
- `systemHealthCheck()` … 設定/接続のヘルスをまとめて確認
- `showNotionConfigGuide()` … Notion設定の簡易ガイドをログ表示

---

## 📌 要件定義（Requirements）

### 機能要件（Functional）
- カレンダー→Notion片方向同期（Notionの変更は無視）
- 既存ページ更新ロジック（新規作成の抑制）
  - 単発: 一意キー＝`iCalUID`
  - 繰り返し例外: 一意キー＝`iCalUID::originalStartTime`
  - 旧キー互換: `iCalUID::currentStart` / `event.id` / `calendarId::event.id` でヒット時は正規IDへ移行
- 終日/複数日対応
  - 終日はNotionに時間なしで登録
  - 複数日終日はGoogleの排他的endを1日戻し、最終日を含む範囲で表現
- 差分更新
  - タイトル/日付/Calendar Event ID に差分がある場合のみNotionを更新
- 記録・ログ
  - 実行ログに「新規/変更/重複/失敗」を出力
  - スプレッドシートには変更があった回のみ追記

### 非機能要件（Non-Functional）
- 可用性/安定運用
  - 排他ロック: 並行実行を防止
  - 指数バックオフ: 429/5xx は最大3回再試行
  - アダプティブ間隔: 5→15→30→60分、変更時は5分×3回のバースト
- 性能
  - デフォルト同期範囲は30日（設定可能）
  - Notionのページ取得はページネーション対応
- 監査性
  - スプレッドシート「分析結果」「スケジュール衝突」「Notion同期結果」に履歴
- 運用
  - 初回のみ権限承認
  - Notionのスキーマ検証（Name=Title, 日付=Date, Calendar Event ID=Rich Text）

### 制約・既知事項（Constraints）
- 真実源はカレンダー（Notionからの更新や削除は同期しない）
- カレンダーを“コピー”して別イベントを作ると（iCalUIDが変われば）別タスクとして扱われる
- タイムゾーンはJST基準（Asia/Tokyo）。他TZの時刻イベントはISOで保存（表示差に注意）
- 削除同期は未実装（Notion側に残る）

### セキュリティ/設定
- Notionトークン/DB ID はコード内設定（Script Properties化も可能）
- 権限: Apps Script（Calendar/UrlFetch/Spreadsheet）への承認が必要

### 検収項目（Acceptance Criteria）
- 同一予定の重複Notionページが作成されない（新規0で安定）
- カレンダーで日付/時刻を変更すると、同じNotionページが更新される
- 終日・複数日の表現が意図どおり（最終日を含む）
- 変更がないときはシート追記がされない
- アダプティブ間隔が段階的に延長（最終60分）し、変更検出で5分×3回の追従になる

---

## 🛡️ 設計上の割り切り（仕様）
- 真実源はカレンダー（Notionの編集は反映しない）
- 削除は未反映（Notion側は残ります）
- タイムゾーンはJST基準（Asia/Tokyo）。終日は日付のみで安全に同期

---

## 🔧 カスタマイズ（任意）
- 同期期間: `SYNC.CAL_TO_NOTION_DAYS`（既定30日）
- 双方向に戻す: `SYNC.ENABLE_BIDIRECTIONAL = true`（非推奨。現在は片方向推奨）
- トリガー調整: `SYNC.SHORT_INTERVAL_MIN` / `SYNC.LONG_INTERVAL_MIN` / `SYNC.STABLE_THRESHOLD`

---

## ✉️ プレゼント時のメッセージ例
```
🎁 あなたのカレンダーをNotionに自動整理する「AIスケジューラ」を贈ります！
1ファイルをGoogle Apps Scriptに貼って initialize() を実行するだけで、
日々の予定変更がそのままNotionに反映され、重複や取りこぼしを防ぎます。
終日予定や複数日にも対応。ログやスプレッドシートで履歴も見えます。
```

---

## 📞 サポート
- セットアップの質問、Notion設定、カスタマイズの相談に対応可能です。
- まずは `testRun()` と `systemHealthCheck()` のログをご共有ください。

素敵なプレゼント体験になりますように！ ✨
