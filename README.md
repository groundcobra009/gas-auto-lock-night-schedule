# 夜の予定自動ブロックスクリプト

## 概要
このGoogle Apps Scriptプロジェクトは、指定したキーワードを含む予定がある日の夜の時間帯を自動的にブロックするためのツールです。飲み会や懇親会などの予定がある日を自動的に検出し、その日の夜の時間帯を「予定あり」としてブロックすることで、他の予定が入らないようにします。

## 主な機能
- 指定したキーワードを含む予定の自動検出
- 20:30以降の予定がある日の夜の時間帯を自動ブロック
- 該当日の夜の時間帯を自動ブロック
- 前日分のブロック予定の自動削除
- 3時間ごとの自動ブロック更新
- 手動でのブロック設定と削除
- トリガーの初期設定機能

## セットアップ手順
1. Google Spreadsheetを新規作成
2. スクリプトエディタを開く（ツール > スクリプトエディタ）
3. コードをコピー＆ペースト
4. 保存して実行権限を承認
5. スプレッドシートに戻り、メニューから「設定シートを作成」を実行
6. 設定シートに必要な情報を入力
   - カレンダーID: 対象のGoogleカレンダーID
   - キーワード: ブロック対象となる予定のキーワード（カンマ区切り）
   - ブロック開始: ブロック開始時間（HH:MM形式）
   - ブロック終了: ブロック終了時間（HH:MM形式）
   - 検索日数（未来）: 未来何日分の予定を検索するか

## 使用方法
### メニュー項目
- **設定シートを作成**: 初期設定用のシートを作成
- **予定自動ブロック実行**: 手動で予定のブロックを実行
- **3時間ごと自動ブロックON/OFF**: 自動更新の有効/無効を切り替え
- **朝5時ロック削除ON/OFF**: 前日分のブロック自動削除の有効/無効を切り替え
- **前日の予定を手動削除**: 前日分のブロック予定を手動で削除
- **トリガー初期設定**: 既存トリガーをリセットして再設定

### 設定項目の説明
1. **カレンダーID**
   - 対象となるGoogleカレンダーのID
   - カレンダーの設定から取得可能

2. **キーワード**
   - ブロック対象となる予定のキーワード
   - カンマ区切りで複数指定可能
   - デフォルト: 飲,懇親,宴,パーティ,会食,交流,親睦,打ち上げ

3. **ブロック時間**
   - ブロック開始時間（デフォルト: 18:30）
   - ブロック終了時間（デフォルト: 21:00）
   - HH:MM形式で指定

4. **夜予定ブロック時間**
   - 夜予定ブロック開始時間（デフォルト: 18:30）
   - 夜予定ブロック終了時間（デフォルト: 20:00）
   - 20:30以降の予定がある場合に使用される時間帯

5. **検索日数（未来）**
   - 未来何日分の予定を検索するか
   - デフォルト: 30日

## 自動化機能
1. **3時間ごとの自動ブロック**
   - 3時間ごとに予定を自動チェック
   - 新しい予定があれば自動的にブロック

2. **朝5時の自動削除**
   - 毎朝5時に前日分のブロック予定を自動削除
   - 不要なブロック予定を自動クリーンアップ

## 注意事項
- カレンダーIDは必ず正しく設定してください
- キーワードは適切に設定し、誤検出を防いでください
- ブロック時間は実際の予定に合わせて調整してください
- 自動化機能は必要に応じてON/OFFを切り替えてください

## トラブルシューティング
1. **カレンダーが見つからない場合**
   - カレンダーIDが正しいか確認
   - カレンダーへのアクセス権限を確認

2. **予定がブロックされない場合**
   - キーワードが正しく設定されているか確認
   - 検索日数が適切か確認
   - ブロック時間が正しく設定されているか確認

3. **自動化が動作しない場合**
   - トリガーが正しく設定されているか確認
   - スクリプトの実行権限を確認

## 更新履歴
- 2025-01-17: v1.1.0リリース
  - 20:30以降の予定に対する専用ブロック時間設定を追加
  - トリガー初期設定機能を追加
  - トリガーからの実行時のUI呼び出しエラーを修正
  - スプレッドシートURL取得機能を追加
- 2024-03-21: v1.0.0初版リリース
  - 基本的な自動ブロック機能
  - 自動削除機能
  - 手動操作機能 