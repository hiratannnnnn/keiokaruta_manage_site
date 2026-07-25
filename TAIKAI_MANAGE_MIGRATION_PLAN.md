# `taikai_manage` 新DB/API対応計画

このドキュメントは、`keiokaruta_manage_site` を、近年更新された
`taikai_manage` のDB・API構成へ対応させるための変更予定をまとめたものです。

初版のAPI接続実装を反映済みです。残りの段階は、API側に保持されていない
フォーム回答の詳細項目やメール運用を含むため、引き続き段階的に切り替えます。

## 実装済み（初版）

- `server/TaikaiApi.js` にBearer認証付きの共通APIクライアントを追加
- `server/Results.js` を出場大会履歴APIへ切り替え
- `server/CountMatches.js` を同APIへ切り替え
- `server/RegisterDatabase.js` と旧メニュー入口を新API登録へ集約
- フォーム回答のインストール型トリガーから`/registrations`へ登録
- フォーム作成時に大会を作成し、開催日保存時に級別日程を同期
- 振込状態・公認状態の操作をAPI更新後にシートへ反映

利用前にScript Propertiesへ次を設定します。

- `TAIKAI_API_BASE_URL`（例: `https://example.com/api/v1`）
- `TAIKAI_API_TOKEN`

未設定のままフォーム作成・登録・履歴検索を実行すると、処理はエラーになります。

## 1. 現状

このプロジェクトは、Google Apps Script（GAS）とスプレッドシートを中心に動作し、
出場大会履歴・大会登録について旧PHP APIへ直接アクセスしています。

なお、画面上の表示名は「大会結果」ですが、現在の`Results.js`が呼び出す
`search_results2.php`は、対戦ごとの勝敗ではなく、選手が出場した大会の履歴を返します。
対戦単位の結果検索（`search_results.php`）は、このプロジェクトの現在のGAS画面では
使用していません。

主な関連箇所:

| 現在のファイル | 役割 | 移行方針 |
|---|---|---|
| `server/config.js` | 旧API URLなどの固定設定 | 新APIのURL・設定プロパティ参照へ変更 |
| `server/RegisterDatabase.js` | 大会・参加者の一括登録 | 新APIの大会・日程・出場登録へ置換 |
| `server/CountMatches.js` | 選手の出場履歴・公認大会回数 | `/players/{id}/participations`へ置換 |
| `server/Results.js` | 出場大会履歴検索（画面表示名は「大会結果」） | 新APIの出場履歴へ移行 |
| `server/FormCreate.js` | フォーム・回答シート作成 | 大会・級別日程のAPI登録情報を保持できるよう拡張 |
| `server/TournamentDetail.js` | 大会シートの参加者・振込管理 | APIのentry状態との同期対象 |
| `server/Calendar.js` | 大会一覧・完了処理 | API上の大会状態との関係を整理 |
| `server/MailManagement.js` | メール予定・送信管理 | 将来的に`email_jobs`等との同期対象 |

## 2. 新APIとの対応方針

`taikai_manage`のAPIを、GASから呼び出す共通クライアント層を追加します。

### 認証・設定

GASのScript Propertiesから次を読み込みます。

- `TAIKAI_API_BASE_URL`
- `TAIKAI_API_TOKEN`

トークンはソースコード、スプレッドシート、Gitへ保存しません。
HTTPSを必須とし、APIエラーは利用者向けの安全なメッセージへ変換します。

### 主なAPI対応

| 業務 | 利用予定API | 対応する現行機能 |
|---|---|---|
| 大会作成・照合 | `GET/POST /tournaments` | フォーム作成、大会登録 |
| 級別日程作成・照合 | `GET/POST /schedules` | A〜E級の日付・公認状態 |
| 選手照合 | `GET /players` | 氏名・メールアドレスによる特定 |
| 出場登録 | `POST /registrations` | `RegisterDatabase.js`の一括登録 |
| 出場履歴 | `GET /players/{id}/participations` | `Results.js`、`CountMatches.js` |
| 振込確認 | `PUT /entries/{entry_id}/payment` | 大会詳細・振込管理 |
| キャンセル | `PUT /entries/{entry_id}/cancellation` | 将来のキャンセル処理 |

## 3. 実装予定の段階

### Phase 1: 共通APIクライアント

新規の共通モジュールを追加し、以下を担当させます。

- URL・トークンの取得
- クエリ文字列の生成
- JSONリクエスト・レスポンス処理
- HTTPエラーの標準化
- APIのIDを文字列として扱う処理

この段階では、旧API呼び出しは残します。

### Phase 2: 出場大会履歴の読み取り移行

`server/Results.js`と`server/CountMatches.js`の読み取り処理を新APIへ切り替えます。

- 氏名またはメールアドレスで選手を照合
- 同姓同名の場合は自動決定せずエラーにする
- `participations`を現在の`date：location：raffleDate`形式へ変換
- 公認・非公認、団体・職域などの既存フィルタを維持
- 旧APIとの結果比較期間を設けてから切り替える

`Results.js`では、現在の画面が期待するレスポンス形式（開催日・大会名など）へ
互換変換して返します。画面の「大会結果」というラベルは、実態に合わせて
「出場大会」または「出場履歴」へ変更する候補とします。

### Phase 3: 大会・級別日程の登録

`server/FormCreate.js`または専用処理で、フォーム作成時に次をAPIへ登録します。

- 大会名を`tournaments`へ登録
- A〜E級ごとの開催日、抽選日、申込期限、振込期限を`schedules`へ登録
- 公認・非公認を`is_sanctioned`へ反映
- 会場・受付情報・支払時期は、入力欄追加後に保存

既存のフォームシートは当面残し、シートを表示用・運用用のローカルキャッシュとして扱います。

### Phase 4: 参加者登録の移行

`server/RegisterDatabase.js`を、新APIの`POST /registrations`中心へ変更します。

- 既存の登録対象判定（振込状態・繰越状態）は維持
- フォーム回答シートのメールアドレスを`players.email`へ使用
- 氏名を名字・名前へ分割して送信
- 大会名・級・開催日から日程を一意に照合
- 同じ登録を再実行しても重複しないようにする
- 登録済み表示は、API成功結果を確認してからシートへ書き込む

### Phase 5: フォーム回答の即時登録

現在は大会登録ボタンによる後処理が中心なので、フォーム送信トリガーを追加し、
回答時に`/registrations`へ登録する方式を検討します。

ただし、抽選日前の申込者もDBには保存し、抽選後に`canceled_at`や対象状態で判定する
新DBの仕様に合わせます。既存運用との違いが大きいため、Phase 4の検証後に実施します。

### Phase 6: 振込・キャンセル状態の同期

大会詳細画面の操作をAPIへ反映します。

- 「済」への変更 → `PUT /entries/{entry_id}/payment`
- キャンセル・落選 → `PUT /entries/{entry_id}/cancellation`
- API更新成功後にシートを更新
- API障害時はシートを変更せず、再実行可能なエラーにする

## 4. メール機能の扱い

`taikai_manage`の新DBでは、自動メールを`email_jobs`・`email_deliveries`で管理します。
このプロジェクトのメール管理はスプレッドシートとGmail下書きを中心に構成されているため、
参加登録の移行と同時に全面移行はしません。

当面は以下の方針とします。

- 手動の案内メール・Gmail下書きは現行機能を維持
- リマインダー・振込確認のAPI連携は別フェーズで設計
- 新API側のメールジョブと二重送信しないよう、移行期間は片方だけを有効化

## 5. 初期移行の対象外

以下は初期移行の対象外です。

- 対戦単位の結果検索（旧`search_results.php`）
- 対戦結果用の旧DBテーブル・登録画面
- スプレッドシートを即時に廃止すること
- メール機能全体の一括置換
- 過去データの無検証な一括投入

対戦単位の結果機能は、`taikai_manage`側でもレガシー機能として扱われているため、
新参加管理APIとは分離して維持します。一方、出場大会履歴は新APIの
`/players/{id}/participations`へ移行します。

## 6. データ移行・検証方針

切り替え前に、次の検証を行います。

1. 既存大会シートとAPI上の大会・日程の対応表を作る
2. 氏名・メールアドレスの照合失敗を一覧化する
3. 旧APIと新APIで出場履歴の件数・内容を比較する
4. 同じ大会・級・選手を再登録しても重複しないことを確認する
5. API障害・認証失敗時にシートが壊れないことを確認する
6. 振込・キャンセル状態の片方向同期を確認する
7. 十分な検証後に旧API呼び出しを削除する

## 7. 完了条件

- 新APIの認証情報がScript Propertiesだけで設定できる
- 大会・級別日程・選手・出場登録を新APIで一意に扱える
- 出場履歴表示と公認大会回数が従来と一致する
- 再実行による重複登録が発生しない
- API障害時に既存シートの状態が不整合にならない
- 旧APIを停止しても、対戦結果検索を除く主要運用が継続できる
